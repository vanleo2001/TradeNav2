VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActiveExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActiveExit.cls
'' Description: Class to handle active exit storage, retrieval, and operation
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/21/2009   DAJ         When checking whether to warn the user about too close
''                          of a trail, check the right thing (Aardvark Issue #4704)
'' 03/27/2009   DAJ         Dump equation that made break even move the stop
'' 04/03/2009   DAJ         Log bar loading information and entry price/time
'' 04/20/2009   DAJ         Put in fixes for highest high/lowest low when invalid
'' 08/13/2009   DAJ         Added the IsAutoExit flag for orders
'' 10/07/2009   DAJ         Allow for broker held linked orders for simple auto exits
'' 12/01/2009   DAJ         Enhancements for Auto Exits held at broker
'' 01/04/2010   DAJ         Several enhancments
'' 01/13/2010   DAJ         Only consider new session date if greater than old
'' 01/14/2010   DAJ         Cancel orders even if broker held if manually deactivated
'' 03/11/2010   DAJ         Use global collections
'' 04/28/2010   DAJ         Fix Stop Orders being generated when Stop Loss set to None (#5717)
'' 05/17/2010   DAJ         Added Trade Sense orders
'' 05/20/2010   DAJ         Fixed issues when no TS order specified for one side
'' 05/24/2010   DAJ         Make sure to explicitly set profits to limits for non-TS orders
'' 05/25/2010   DAJ         Fixes for the TradeSense auto exits
'' 06/03/2010   DAJ         Fix for cancelling a TradeSense auto exit order (#5739)
'' 06/17/2010   DAJ         Fix for having a stop only on a TradeSense auto exit (#5805)
'' 08/12/2010   DAJ         New method for checking streaming availability for a symbol
'' 09/07/2010   DAJ         Fixed wrong side of market check for TradeSense Stop only
'' 09/20/2010   DAJ         When setting up "Saved" auto exit, explicitly set order type for stop (#5928)
'' 09/20/2010   DAJ         Fix for auto exits/TradeSense order groups when stream stops
'' 09/27/2010   DAJ         Fixed UpdateBars when TS orders only specified for one side
'' 11/09/2010   DAJ         Changed the way that Session Date is calculated on activate
'' 11/15/2010   DAJ         Don't count TSOG position for auto exits
'' 12/08/2010   DAJ         Only reactivate PFG FX auto exit during valid times
'' 01/07/2011   DAJ         Changed activate call for TradeSense orders
'' 02/03/2011   DAJ         Moved PFG Forex trading time code to mDataNav
'' 05/16/2011   DAJ         Set OrderDate correctly in case delayed streaming
'' 05/27/2011   DAJ         Set the IsAutomated flag for orders that get created
'' 06/08/2011   DAJ         TradeSense order enhancements
'' 08/17/2011   DAJ         Upon activation, use initial entry price instead of average
'' 10/04/2011   DAJ         Don't handle fill or position if cancelling all
'' 01/30/2012   DAJ         User configure timeout on auto exit reject
'' 02/28/2012   DAJ         Only submit/amend TradeSense orders if UpdateBars is true
'' 07/13/2012   DAJ         Added check for valid trading time for IB Forex
'' 08/17/2012   DAJ         Utilize generic valid forex trading time function
'' 09/07/2012   DAJ         Reset HH and LL to entry price before CreateOrders
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 12/17/2012   DAJ         Fix for setting order quantities when one auto exit order fills
'' 01/10/2013   DAJ         Fix for stop loss not cancelling when "partner" profit target fills
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 04/17/2013   DAJ         Flatten if Stop on wrong side, Cancel manual orders on opposite side
'' 07/22/2013   DAJ         Round the entry price to the nearest tick
'' 11/11/2013   DAJ         Fix for using saved information for a stop order
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/23/2014   DAJ         Fix for using second level instead of first; Using wrong data for HH and LL
'' 11/28/2014   DAJ         Allow second profit target to be zero quantity in SetThreeLevels
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kTrailingStopWait = 5000#

Private Type mPrivate
    lAccountID As Long                  ' ID for the account being traded
    nAcctType As eTT_AccountType        ' Account type for the account ID
    lSymbolID As Long                   ' ID for the symbol being traded
    strSymbol As String                 ' Symbol being traded
    strExitStrategyFile As String       ' Filename of the exit strategy being run
    bActive As Boolean                  ' Is this exit strategy currently active?
    lCurrentPosition As Long            ' Current position for the account and symbol
    dEntryTime As Double                ' Time that the position was entered
    dEntryPrice As Double               ' Entry price for the position
    bDelete As Boolean                  ' Do we want this item deleted?
    lSessionDate As Long                ' Current session date
    
    dHighestHigh As Double              ' Highest high since entry
    dTimeOfHighestHigh As Double        ' Time that the highest high since entry checked
    dLowestLow As Double                ' Lowest low since entry
    dTimeOfLowestLow As Double          ' Time that the lowest low since entry checked
    
    Bars As cGdBars                     ' Bars object
    ExitStrategy As cExitStrategy       ' Exit strategy object
    
    Target1 As cActiveExitOrder         ' First profit target order
    Target2 As cActiveExitOrder         ' Second profit target order
    Target3 As cActiveExitOrder         ' Third profit target order
    Stop1 As cActiveExitOrder           ' First stop loss order
    Stop2 As cActiveExitOrder           ' Second stop loss order
    Stop3 As cActiveExitOrder           ' Third stop loss order
    StopRest As cActiveExitOrder        ' Rest of the quantity not in other three
    
    'bCancellingAll As Boolean           ' This object is cancelling all orders
    'bFlattening As Boolean              ' This object is flattening the position
    bDeactivating As Boolean            ' Is this object being deactivated?
    bStrategyDirty As Boolean           ' Does the strategy need to be reloaded?
    bWaitingOnData As Boolean           ' Are we waiting on data before activating?
    bManualDeactivate As Boolean        ' Is this a manual deactivation?
    
    bBreakEvenMoved As Boolean          ' Have we already moved the break-even stop order?
    dTrailPrice As Double               ' Trailing stop price
    dTrailWithLimitPrice As Double      ' Trailing stop with limit price
    dLastTrailTime As Double            ' Time we did the last trail
End Type
Private m As mPrivate

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nAcctType = g.Broker.AccountTypeForID(lAccountID)
    SetAccountIds
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    SetSymbols
End Property

Public Property Get ExitStrategyFile() As String
    ExitStrategyFile = m.strExitStrategyFile
End Property
Public Property Let ExitStrategyFile(ByVal strExitStrategyFile As String)
    m.strExitStrategyFile = strExitStrategyFile
End Property

Public Property Get Active() As Boolean
    Active = m.bActive
End Property
Private Property Let Active(ByVal bActive As Boolean)
    m.bActive = bActive
End Property

Public Property Get Deactivating() As Boolean
    Deactivating = m.bDeactivating
End Property

Public Property Get Key() As String
    Key = "Exit" & vbTab & Str(AccountID) & vbTab & Str(SymbolOrSymbolID)
End Property

Public Property Get Name() As String
    Name = m.ExitStrategy.StrategyName
End Property

Public Property Get ExitStrategy() As cExitStrategy
    Set ExitStrategy = m.ExitStrategy
End Property

Public Property Get ToDelete() As Boolean
    ToDelete = m.bDelete
End Property
Public Property Let ToDelete(ByVal bDelete As Boolean)
    m.bDelete = bDelete
End Property

Public Property Get SessionDate() As Long
    SessionDate = m.lSessionDate
End Property

Public Property Get Runner() As Boolean
    Runner = ((m.ExitStrategy.SpecifyLots = True) And (m.ExitStrategy.ExitEntirePosition = False))
End Property

Public Property Get WaitingOnData() As Boolean
    WaitingOnData = m.bWaitingOnData
End Property

Private Property Get TickMove() As Double
    ' Don't use MinMove here because we already accomodate for the half ticks and quarter ticks
    ' in the user interface and using MinMove here would break that (DAJ: 07/23/2008)...
    TickMove = m.Bars.TickMove
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up the active exit object from the database
'' Inputs:      Account, Symbol or SymbolID, Exit Strategy ID
'' Returns:     True if loaded successfully, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lAccountID As Long, ByVal vSymbolOrSymbolID As Variant) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    If VarType(vSymbolOrSymbolID) = vbString Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits] " & _
                    "WHERE [AccountID]=" & Str(AccountID) & " AND [Symbol]='" & Str(vSymbolOrSymbolID) & ";", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits] " & _
                    "WHERE [AccountID]=" & Str(AccountID) & " AND [SymbolID]=" & Str(vSymbolOrSymbolID) & ";", dbOpenDynaset)
    End If
    
    Load = LoadFromRecordset(rs)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFromRecordset
'' Description: Load up the active exit object from the database
'' Inputs:      Recordset
'' Returns:     True if loaded successfully, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadFromRecordset(ByVal rs As Recordset) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        AccountID = rs!AccountID
        m.nAcctType = g.Broker.AccountTypeForID(AccountID)
        m.lSymbolID = rs!SymbolID
        m.strSymbol = rs!Symbol
        m.strExitStrategyFile = rs!ExitStrategyFile
        Active = False 'rs!Active
        m.lCurrentPosition = rs!CurrentPosition
        m.dEntryPrice = NearestTick(rs!EntryPrice)
        m.Target1.Load rs!Target1OrderID
        m.Target2.Load rs!Target2OrderID
        m.Target3.Load rs!Target3OrderID
        m.Stop1.Load rs!StopOrderId
        m.Stop2.Load rs!Stop2OrderID
        m.Stop3.Load rs!Stop3OrderID
        m.StopRest.Load rs!StopROrderID
        
        DumpDebug "Loading active exit for " & Str(AccountID) & "/" & Symbol & ": '" & ToString & "'"
        
        SetSymbols
        LoadStrategy
        LoadBars
        
        bReturn = True
    End If
    
    LoadFromRecordset = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.LoadFromRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategy
'' Description: Load up the exit order strategy
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadStrategy() As Boolean
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the account positions collection
    Dim lNumTicksStopBuffer As Long     ' Number of ticks for the stop buffer
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    
    Set m.ExitStrategy = New cExitStrategy
    m.ExitStrategy.Load m.strExitStrategyFile
    m.bStrategyDirty = False
    
    If g.Broker.DontAllowStopMove Then
        lNumTicksStopBuffer = g.Broker.NumTicksStopBuffer
        
        If (g.bStarting = False) And (g.bUnloading = False) Then
            If (m.ExitStrategy.StopLossType = eGDStopLossType_Trail) Then
                ' Use StopLossTicks in this case instead of TrailTicks because that is where the
                ' trailing stop amount is stored in this case (Aardvark Issue #4704)...
                If m.ExitStrategy.StopLossTicks <= lNumTicksStopBuffer Then
                    If InfBox("Your current settings state that a Stop order cannot be moved closer than " & Str(lNumTicksStopBuffer) & " tick(s) away from the market, but your trailing stop is set to trail behind the market by " & Str(m.ExitStrategy.TrailTicks) & " tick(s).  This is going to cause your trailing stop not to trail.||Do you want to continue?", "?", "+Yes|-No", "Auto Exit Load Warning") = "N" Then
                        bReturn = False
                    End If
                    DumpDebug "Num Ticks Stop Buffer = " & Str(lNumTicksStopBuffer) & ", Trail Ticks = " & Str(m.ExitStrategy.TrailTicks) & ", Return = " & Str(bReturn)
                End If
            ElseIf (m.ExitStrategy.StopLossType = eGDStopLossType_BreakEven) And (m.ExitStrategy.UseTrail = True) Then
                If m.ExitStrategy.TrailTicks <= lNumTicksStopBuffer Then
                    If InfBox("Your current settings state that a Stop order cannot be moved closer than " & Str(lNumTicksStopBuffer) & " tick(s) away from the market, but your trailing stop is set to trail behind the market by " & Str(m.ExitStrategy.TrailTicks) & " tick(s).  This is going to cause your trailing stop not to trail.||Do you want to continue?", "?", "+Yes|-No", "Auto Exit Load Warning") = "N" Then
                        bReturn = False
                    End If
                    DumpDebug "Num Ticks Stop Buffer = " & Str(lNumTicksStopBuffer) & ", Trail Ticks = " & Str(m.ExitStrategy.TrailTicks) & ", Return = " & Str(bReturn)
                End If
            End If
        End If
    End If

    If bReturn = True Then
        g.Broker.CreateFillSummaryForAutoExit Me
    End If
    
    LoadStrategy = bReturn
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.LoadStrategy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load up the local version of the bars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadBars()
On Error GoTo ErrSection:

    Dim lBar As Long                    ' Bar for the entry time
    Dim Bars As cGdBars                 ' Temporary bars object

    Set m.Bars = New cGdBars
    
    m.Bars.ArrayMask = eBARS_TickByTick

    If (m.ExitStrategy.StopLossType = eGDStopLossType_BreakEven) Or (m.ExitStrategy.StopLossType = eGDStopLossType_Trail) Then
        If m.dEntryTime <> 0# Then
            If DM_GetBars(m.Bars, SymbolOrSymbolID, ePRD_EachTick, Int(m.dEntryTime)) = True Then
                DumpDebug "Tick Data Loaded from " & DateFormat(Int(m.dEntryTime), MM_DD_YYYY) & "; Size = " & Str(m.Bars.Size)
            End If
        Else
            SetBarProperties m.Bars, SymbolOrSymbolID
        End If
        
        g.RealTime.AddTickBuffer m.Bars
        g.RealTime.SpliceBars m.Bars
        
        DumpDebug "Number of Bars = " & Str(m.Bars.Size)
        
        If m.dEntryTime <> 0# Then
            lBar = m.Bars.FindDateTime(m.dEntryTime)
            If lBar > 0 Then
                m.Bars.DeleteFirstBars lBar
                DumpDebug "Bars Chopped at " & DateFormat(m.Bars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & "; Size = " & Str(m.Bars.Size)
            End If
        End If
    Else
        ' Make sure to let the GetBars add the symbol to realtime, because otherwise this could lead
        ' to the Bars for this symbol on the Trade Console to get "stuck" and not update...
        Set Bars = GetTradeBars(SymbolOrSymbolID)
        If Not Bars Is Nothing Then
            Set m.Bars = Bars.MakeCopy
        End If
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.LoadBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the active exit object to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    If m.lSymbolID = 0& Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits] " & _
                    "WHERE [AccountID]=" & Str(AccountID) & " AND [Symbol]='" & Symbol & "';", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits] " & _
                    "WHERE [AccountID]=" & Str(AccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & ";", dbOpenDynaset)
    End If
    
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!AccountID = AccountID
    rs!SymbolID = m.lSymbolID
    rs!Symbol = Symbol
    rs!ExitStrategyFile = m.strExitStrategyFile
    rs!Active = m.bActive
    rs!CurrentPosition = m.lCurrentPosition
    rs!EntryPrice = m.dEntryPrice
    rs!Target1OrderID = m.Target1.OrderID
    rs!Target2OrderID = m.Target2.OrderID
    rs!Target3OrderID = m.Target3.OrderID
    rs!StopOrderId = m.Stop1.OrderID
    rs!Stop2OrderID = m.Stop2.OrderID
    rs!Stop3OrderID = m.Stop3.OrderID
    rs!StopROrderID = m.StopRest.OrderID
    
    rs.Update
    
    DumpDebug "Saved: '" & ToString & "'"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the active exit from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    If (m.lSymbolID = 0&) Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits] " & _
                    "WHERE [AccountID]=" & Str(AccountID) & " AND [Symbol]='" & Str(Symbol) & "';", dbOpenDynaset)
    Else
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActiveExits] " & _
                    "WHERE [AccountID]=" & Str(AccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & ";", dbOpenDynaset)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOrder
'' Description: Does the given order exist in this item?
'' Inputs:      Order ID
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOrder(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    Select Case lOrderID
        Case m.Target1.Order.OrderID, m.Target2.Order.OrderID, m.Target3.Order.OrderID
            bReturn = True
            
        Case m.Stop1.Order.OrderID, m.Stop2.Order.OrderID, m.Stop3.Order.OrderID, m.StopRest.Order.OrderID
            bReturn = True
            
        Case Else
            bReturn = False
    End Select

    HasOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.HasOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrders
'' Description: Does the active exit have any currently working orders?
'' Inputs:      Skip Order ID
'' Returns:     True if has working orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrders(Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal bIgnoreBrokerHeld As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    If (bIgnoreBrokerHeld = True) And (HoldAtBroker = True) Then
        bReturn = False
    ElseIf (m.Stop1.OrderID > 0&) And (m.Stop1.OrderID <> lSkipOrderID) Then
        bReturn = True
    ElseIf (m.Stop2.OrderID > 0&) And (m.Stop2.OrderID <> lSkipOrderID) Then
        bReturn = True
    ElseIf (m.Stop3.OrderID > 0&) And (m.Stop3.OrderID <> lSkipOrderID) Then
        bReturn = True
    ElseIf (m.StopRest.OrderID > 0&) And (m.StopRest.OrderID <> lSkipOrderID) Then
        bReturn = True
    ElseIf (m.Target1.OrderID > 0&) And (m.Target1.OrderID <> lSkipOrderID) Then
        bReturn = True
    ElseIf (m.Target2.OrderID > 0&) And (m.Target2.OrderID <> lSkipOrderID) Then
        bReturn = True
    ElseIf (m.Target3.OrderID > 0&) And (m.Target3.OrderID <> lSkipOrderID) Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    HasWorkingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.HasWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasPendingOrders
'' Description: Does the active exit have any currently pending orders?
'' Inputs:      None
'' Returns:     True if has pending orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasPendingOrders() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    If (bReturn = False) And (Not m.Stop1.Order Is Nothing) Then
        bReturn = OrderIsPending(m.Stop1.Order)
    End If

    If (bReturn = False) And (Not m.Stop2.Order Is Nothing) Then
        bReturn = OrderIsPending(m.Stop2.Order)
    End If
    
    If (bReturn = False) And (Not m.Stop3.Order Is Nothing) Then
        bReturn = OrderIsPending(m.Stop3.Order)
    End If
    
    If (bReturn = False) And (Not m.StopRest.Order Is Nothing) Then
        bReturn = OrderIsPending(m.StopRest.Order)
    End If
    
    If (bReturn = False) And (Not m.Target1.Order Is Nothing) Then
        bReturn = OrderIsPending(m.Target1.Order)
    End If

    If (bReturn = False) And (Not m.Target2.Order Is Nothing) Then
        bReturn = OrderIsPending(m.Target2.Order)
    End If

    If (bReturn = False) And (Not m.Target3.Order Is Nothing) Then
        bReturn = OrderIsPending(m.Target3.Order)
    End If
    
    HasPendingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.HasPendingOrders"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Activate
'' Description: Activate the active exit
'' Inputs:      Current Information
'' Returns:     True if Activated, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Activate(Optional ByVal strCurrentInfo As String = "") As Boolean
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the account positions collection
    Dim AcctPos As cAccountPosition     ' Account position
    Dim strSavedInfo As String          ' Saved information from previous session
    Dim astrSavedInfo As cGdArray       ' Saved information in an array
    Dim bReturn As Boolean              ' Return value for the function
    Dim frm As frmAlertPopup            ' Alert popup message
    Static bInProgress As Boolean       ' Are we already in progress?
    
    Set astrSavedInfo = New cGdArray
    astrSavedInfo.Create eGDARRAY_Strings
    
    bReturn = True
        
    If bInProgress = False Then
        bInProgress = True
        
        If (m.bActive = False) And (CanActivate = True) Then
            ' Get the current position for the symbol in this account...
            g.Broker.CreateFillSummaryForAutoExit Me
                
            Set AcctPos = g.Broker.FillSummary(AccountID, SymbolOrSymbolID, 0&)
            If Not AcctPos Is Nothing Then
                With AcctPos
                    m.lCurrentPosition = .CurrentPositionSnapshot - TradeSenseGroupPosition
                    'm.dEntryPrice = .AverageEntrySnapshot
                    m.dEntryPrice = NearestTick(.FirstOpenEntryPriceSnapshot)
                    m.dTrailPrice = 0#
                    m.dTrailWithLimitPrice = 0#
                    m.dLastTrailTime = 0#
                    m.dLowestLow = m.dEntryPrice
                    m.dHighestHigh = m.dEntryPrice
                    
                    If m.lCurrentPosition <> 0& Then
                        m.dEntryTime = ConvertBrokerDate(.LastTradedSnapshot, m.nAcctType, Symbol, False)
                    Else
                        m.dEntryTime = 0#
                    End If
                    
                    If m.Stop1.OrderID < 0& Then m.Stop1.OrderID = 0&
                    If m.Stop2.OrderID < 0& Then m.Stop2.OrderID = 0&
                    If m.Stop3.OrderID < 0& Then m.Stop3.OrderID = 0&
                    If m.StopRest.OrderID < 0& Then m.StopRest.OrderID = 0&
                    If m.Target1.OrderID < 0& Then m.Target1.OrderID = 0&
                    If m.Target2.OrderID < 0& Then m.Target2.OrderID = 0&
                    If m.Target3.OrderID < 0& Then m.Target3.OrderID = 0&
                    
                    LoadBars
                    
                    Save
                End With
                
                ActivateTsOrders AcctPos
                m.lSessionDate = m.Bars.SessionDateForTime(ConvertTimeZone(g.RealTime.FeedTime(m.strSymbol), "NY", m.Bars.Prop(eBARS_ExchangeTimeZoneInf)), False)
                
                Active = True
                DumpDebug "Activate(" & Str(SymbolOrSymbolID) & ", " & Str(AccountID) & ") - Current Position: " & Str(m.lCurrentPosition) & ", Entry Price: " & Str(m.dEntryPrice) & ", Entry Date: " & DateFormat(m.dEntryTime, MM_DD_YY, HH_MM_SS) & ", Session Date: " & DateFormat(m.lSessionDate, MM_DD_YYYY)
                DumpDebug vbTab & "CancelOpposite = " & Str(m.ExitStrategy.CancelOpposite) & "; MarketIfWrongSide = " & Str(m.ExitStrategy.MarketIfWrongSide)
                
                SetOrderQuantities kNullData, strCurrentInfo, "Activate"
                
                If Len(strCurrentInfo) = 0 Then
                    strSavedInfo = GetIniFileProperty(Str(AccountID) & "|" & Symbol, "", "AutoExits", g.strIniFile)
                    If Len(strSavedInfo) > 0 Then
                        astrSavedInfo.SplitFields strSavedInfo, "|"
                        If (astrSavedInfo(0) <> m.strExitStrategyFile) Or (astrSavedInfo(1) <> Str(m.lCurrentPosition)) Or (astrSavedInfo(2) <> Str(m.dEntryPrice)) Then
                            strSavedInfo = ""
                            SetIniFileProperty Str(AccountID) & "|" & Symbol, "", "AutoExits", g.strIniFile
                            DumpDebug "SaveLastKnown - ''"
                        End If
                    End If
                Else
                    DumpDebug "Activating from " & strCurrentInfo
                    strSavedInfo = CurrentInfoToLastKnown(strCurrentInfo)
                End If
                
                If VerifyOrders(strSavedInfo) Then
                    CreateOrders True, strSavedInfo, "auto exit being activated"
                    bReturn = True
                Else
                    Set frm = New frmAlertPopup
                    frm.ShowMessageBox "|The auto exit for " & Symbol & " could not be activated because one or more of the orders were going to be placed on the wrong side of the market.", "Auto Exit Activation Failure", vbCenter
                    Deactivate True, "Activation Failed due to one or more orders being on wrong side of market", False
                    bReturn = False
                End If
            End If
        End If
        
        bInProgress = False
    End If
    
    Activate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    bInProgress = False
    RaiseError "cActiveExit.Activate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Deactivate
'' Description: Deactivate the active exit
'' Inputs:      Manual Deactivate?, Reason, Save Last Known?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Deactivate(Optional ByVal bManualDeactivate As Boolean = False, Optional ByVal strReason As String = "", Optional ByVal bSaveLastKnown As Boolean = True)
On Error GoTo ErrSection:
    
    If Len(strReason) > 0 Then
        DumpDebug "Deactivate: Active = " & Str(m.bActive) & " ; Deactivating = " & Str(m.bDeactivating) & " ; Manual Deactivate = " & Str(bManualDeactivate) & " ; Reason = " & strReason
    Else
        DumpDebug "Deactivate: Active = " & Str(m.bActive) & " ; Deactivating = " & Str(m.bDeactivating) & " ; Manual Deactivate = " & Str(bManualDeactivate)
    End If
    
    If (m.bActive = True) And (m.bDeactivating = False) Then
        m.bManualDeactivate = bManualDeactivate
        
        If bSaveLastKnown Then
            SaveLastKnown bManualDeactivate
        End If
        
        If HasWorkingOrders(-1&, Not bManualDeactivate) Then
            m.bDeactivating = True
            CancelAllOrders -1&, False, Not bManualDeactivate
        Else
            Active = False
            m.bDeactivating = False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.Deactivate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Update the local bars object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim bNewBar As Boolean              ' Has a new bar started?
    Dim bMoveOrder As Boolean           ' Do we want to move the stop order?
    Dim dNewPrice As Double             ' New price for the order
    Dim dNewWithLimitPrice As Double    ' New with limit price for the order
    Dim strEquation As String           ' Equation for moving order
    Dim lBar As Long                    ' Bar number for the given date/time
    Dim bConditionTrue As Boolean       ' Is the condition true?
    Static bProfitCondTrue As Boolean   ' Is the profit condition true?
    Static bStopCondTrue As Boolean     ' Is the stop condition true?

    If (m.bActive = True) And (m.lCurrentPosition <> 0) Then
        If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
            If m.lCurrentPosition > 0 Then
                If Not m.ExitStrategy.TsProfitLong Is Nothing Then
                    With m.ExitStrategy.TsProfitLong
                        If Len(.ConditionCoded) > 0 Then
                            If .UpdateBars(bNewBar) Then
                                If (bNewBar = True) And (.Price <> 0) Then
                                    DumpDebug vbTab & "TS Profit Long New Bar: Price = " & Str(.Price) & " (" & .PriceEnglish & "), With Limit = " & Str(.WithLimitPrice) & " (" & .WithLimitEnglish & ")"
                                End If
                                If (Flattening = False) And (m.Target1.CancelRequested = False) Then
                                    If .ConditionTrue Then
                                        If m.Target1.OrderID <= 0 Then
                                            If CreateTargetOrder(1) Then
                                                m.Target1.SubmitOrder "Condition True: '" & .ConditionEnglish & "'"
                                            End If
                                        Else
                                            m.Target1.MoveTradeSense .Price, .WithLimitPrice, .EnglishOrderPrice
                                        End If
                                    Else
                                        m.Target1.CancelWorkingOrder False, , , , .ConditionEnglish
                                    End If
                                End If
                            End If
                        End If
                    End With
                End If
            ElseIf m.lCurrentPosition < 0 Then
                If Not m.ExitStrategy.TsProfitShort Is Nothing Then
                    With m.ExitStrategy.TsProfitShort
                        If Len(.ConditionCoded) > 0 Then
                            If .UpdateBars(bNewBar) Then
                                If (bNewBar = True) And (.Price <> 0) Then
                                    DumpDebug vbTab & "TS Profit Short New Bar: Price = " & Str(.Price) & " (" & .PriceEnglish & "), With Limit = " & Str(.WithLimitPrice) & " (" & .WithLimitEnglish & ")"
                                End If
                                If (Flattening = False) And (m.Target1.CancelRequested = False) Then
                                    If .ConditionTrue Then
                                        If m.Target1.OrderID <= 0 Then
                                            If CreateTargetOrder(1) Then
                                                m.Target1.SubmitOrder "Condition True: '" & .ConditionEnglish & "'"
                                            End If
                                        Else
                                            m.Target1.MoveTradeSense .Price, .WithLimitPrice, .EnglishOrderPrice
                                        End If
                                    Else
                                        m.Target1.CancelWorkingOrder False, , , , .ConditionEnglish
                                    End If
                                End If
                            End If
                        End If
                    End With
                End If
            End If
        End If
        
        If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
            If m.lCurrentPosition > 0 Then
                If Not m.ExitStrategy.TsStopLong Is Nothing Then
                    With m.ExitStrategy.TsStopLong
                        If Len(.ConditionCoded) > 0 Then
                            If .UpdateBars(bNewBar) Then
                                If (bNewBar = True) And (.Price <> 0) Then
                                    DumpDebug vbTab & "TS Stop Long New Bar: Price = " & Str(.Price) & " (" & .PriceEnglish & "), With Limit = " & Str(.WithLimitPrice) & " (" & .WithLimitEnglish & ")"
                                End If
                                If Flattening = False Then
                                    If .ConditionTrue Then
                                        If m.StopRest.CancelRequested = False Then
                                            If m.StopRest.Quantity > 0 Then
                                                If m.StopRest.OrderID <= 0 Then
                                                    If CreateStopLossOrder(0) Then
                                                        m.StopRest.SubmitOrder "Condition True: '" & .ConditionEnglish & "'"
                                                    End If
                                                Else
                                                    m.StopRest.MoveTradeSense .Price, .WithLimitPrice, .EnglishOrderPrice
                                                End If
                                            End If
                                        End If
                                        If m.Stop1.CancelRequested = False Then
                                            If m.Stop1.Quantity > 0 Then
                                                If m.Stop1.OrderID <= 0 Then
                                                    If CreateStopLossOrder(1) Then
                                                        m.Stop1.SubmitOrder "Condition True: '" & .ConditionEnglish & "'"
                                                    End If
                                                Else
                                                    m.Stop1.MoveTradeSense .Price, .WithLimitPrice, .EnglishOrderPrice
                                                End If
                                            End If
                                        End If
                                    Else
                                        m.Stop1.CancelWorkingOrder False, , , , .ConditionEnglish
                                        m.StopRest.CancelWorkingOrder False, , , , .ConditionEnglish
                                    End If
                                End If
                            End If
                        End If
                    End With
                End If
            ElseIf m.lCurrentPosition < 0 Then
                If Not m.ExitStrategy.TsStopShort Is Nothing Then
                    With m.ExitStrategy.TsStopShort
                        If Len(.ConditionCoded) > 0 Then
                            If .UpdateBars(bNewBar) Then
                                If (bNewBar = True) And (.Price <> 0) Then
                                    DumpDebug vbTab & "TS Stop Short New Bar: Price = " & Str(.Price) & " (" & .PriceEnglish & "), With Limit = " & Str(.WithLimitPrice) & " (" & .WithLimitEnglish & ")"
                                End If
                                If Flattening = False Then
                                    If .ConditionTrue Then
                                        If m.StopRest.CancelRequested = False Then
                                            If m.StopRest.Quantity > 0 Then
                                                If m.StopRest.OrderID <= 0 Then
                                                    If CreateStopLossOrder(0) Then
                                                        m.StopRest.SubmitOrder "Condition True: '" & .ConditionEnglish & "'"
                                                    End If
                                                Else
                                                    m.StopRest.MoveTradeSense .Price, .WithLimitPrice, .EnglishOrderPrice
                                                End If
                                            End If
                                        End If
                                        If m.Stop1.CancelRequested = False Then
                                            If m.Stop1.Quantity > 0 Then
                                                If m.Stop1.OrderID <= 0 Then
                                                    If CreateStopLossOrder(1) Then
                                                        m.Stop1.SubmitOrder "Condition True: '" & .ConditionEnglish & "'"
                                                    End If
                                                Else
                                                    m.Stop1.MoveTradeSense .Price, .WithLimitPrice, .EnglishOrderPrice
                                                End If
                                            End If
                                        End If
                                    Else
                                        m.Stop1.CancelWorkingOrder False, , , , .ConditionEnglish
                                        m.StopRest.CancelWorkingOrder False, , , , .ConditionEnglish
                                    End If
                                End If
                            End If
                        End If
                    End With
                End If
            End If
        ElseIf (m.ExitStrategy.StopLossType = eGDStopLossType_BreakEven) Or (m.ExitStrategy.StopLossType = eGDStopLossType_Trail) Then
            If g.RealTime.UpdateBars(m.Bars, bNewBar) Then
                If bNewBar = True Then
                    DumpDebug vbTab & "NewBar Flag True -- Reloading Data..."
                    g.RealTime.SpliceBars m.Bars
                    
                    DumpDebug vbTab & "Number of Bars = " & Str(m.Bars.Size)
                    If m.dEntryTime <> 0# Then
                        lBar = m.Bars.FindDateTime(m.dEntryTime)
                        If lBar > 0 Then
                            m.Bars.DeleteFirstBars lBar
                            DumpDebug vbTab & "Bars Chopped at " & DateFormat(m.Bars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & "; Size = " & Str(m.Bars.Size)
                        End If
                    End If
                End If
                
                HighestHighSinceEntry
                LowestLowSinceEntry
                
                If (m.Bars.Size > 0) And OneStopHasValidID Then
                    Select Case m.ExitStrategy.StopLossType
                        Case eGDStopLossType_BreakEven
                            If (m.bBreakEvenMoved = False) Then
                                If m.lCurrentPosition < 0 Then
                                    bMoveOrder = (m.dLowestLow <= (m.dEntryPrice - (m.ExitStrategy.AfterTicks * TickMove)))
                                    If bMoveOrder Then
                                        strEquation = Str(m.dLowestLow) & " <= " & Str(m.dEntryPrice) & " - (" & Str(m.ExitStrategy.AfterTicks) & " * " & Str(TickMove) & ")"
                                        dNewPrice = m.dEntryPrice - (m.ExitStrategy.MoveToTicks * TickMove)
                                        If m.ExitStrategy.UseWithLimit Then
                                            dNewWithLimitPrice = dNewPrice + (m.ExitStrategy.WithLimitTicks * TickMove)
                                        Else
                                            dNewWithLimitPrice = 0#
                                        End If
                                        
                                        DumpDebug "MoveOrder: '" & strEquation & "', NewPrice = " & Str(dNewPrice) & ", New With Limit = " & Str(dNewWithLimitPrice)
                                    End If
                                Else
                                    bMoveOrder = (m.dHighestHigh >= (m.dEntryPrice + (m.ExitStrategy.AfterTicks * TickMove)))
                                    If bMoveOrder Then
                                        strEquation = Str(m.dHighestHigh) & " >= " & Str(m.dEntryPrice) & " + (" & Str(m.ExitStrategy.AfterTicks) & " * " & Str(TickMove) & ")"
                                        dNewPrice = m.dEntryPrice + (m.ExitStrategy.MoveToTicks * TickMove)
                                        If m.ExitStrategy.UseWithLimit Then
                                            dNewWithLimitPrice = dNewPrice - (m.ExitStrategy.WithLimitTicks * TickMove)
                                        Else
                                            dNewWithLimitPrice = 0#
                                        End If
                                        
                                        DumpDebug "MoveOrder: '" & strEquation & "', NewPrice = " & Str(dNewPrice) & ", New With Limit = " & Str(dNewWithLimitPrice)
                                    End If
                                End If
                                
                                If (bMoveOrder = True) Then
                                    If (m.Stop1.Waiting = False) And (m.Stop1.BreakEvenMoved = False) Then
                                        m.Stop1.MoveBreakEven dNewPrice, dNewWithLimitPrice, strEquation
                                    Else
                                        DumpDebug m.Stop1.Name & " not moved: Waiting = " & Str(m.Stop1.Waiting) & ", Break Even Moved = " & Str(m.Stop1.BreakEvenMoved)
                                    End If
                                    If (m.Stop2.Waiting = False) And (m.Stop2.BreakEvenMoved = False) Then
                                        m.Stop2.MoveBreakEven dNewPrice, dNewWithLimitPrice, strEquation
                                    Else
                                        DumpDebug m.Stop2.Name & " not moved: Waiting = " & Str(m.Stop2.Waiting) & ", Break Even Moved = " & Str(m.Stop2.BreakEvenMoved)
                                    End If
                                    If (m.Stop3.Waiting = False) And (m.Stop3.BreakEvenMoved = False) Then
                                        m.Stop3.MoveBreakEven dNewPrice, dNewWithLimitPrice, strEquation
                                    Else
                                        DumpDebug m.Stop3.Name & " not moved: Waiting = " & Str(m.Stop3.Waiting) & ", Break Even Moved = " & Str(m.Stop3.BreakEvenMoved)
                                    End If
                                    If (m.StopRest.Waiting = False) And (m.StopRest.BreakEvenMoved = False) Then
                                        m.StopRest.MoveBreakEven dNewPrice, dNewWithLimitPrice, strEquation
                                    Else
                                        DumpDebug m.StopRest.Name & " not moved: Waiting = " & Str(m.StopRest.Waiting) & ", Break Even Moved = " & Str(m.StopRest.BreakEvenMoved)
                                    End If
                                    
                                    If m.Stop1.BreakEvenMoved And m.Stop2.BreakEvenMoved And m.Stop3.BreakEvenMoved And m.StopRest.BreakEvenMoved Then
                                        m.bBreakEvenMoved = True
                                        m.dLastTrailTime = gdTickCount
                                    End If
                                End If
                            ElseIf m.ExitStrategy.UseTrail Then
                                TrailOrder m.ExitStrategy.TrailTicks
                            End If
                            
                        Case eGDStopLossType_Trail
                            TrailOrder m.ExitStrategy.StopLossTicks
                    
                    End Select
                End If
            End If
        End If
        
        ' DAJ: Lets try not to do this here and see what happens...
        ''ModifyOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.UpdateBars"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderCallback
'' Description: Update the appropriate order if it exists in the object
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderCallback(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim bFound As Boolean               ' Was the order found?
    Dim strOrderName As String          ' Name of the order
    Dim lReturn As Long                 ' Return value from the OrderCallback call
    
    DumpDebug "OrderCallback ('" & Order.OrderText & "') -- Active = " & Str(m.bActive)
    
    If m.bActive Then
        
        bFound = False
        Select Case Order.OrderID
            Case m.Target1.OrderID
                bFound = True
                strOrderName = m.Target1.Name
                lReturn = m.Target1.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                
            Case m.Target2.OrderID
                bFound = True
                strOrderName = m.Target2.Name
                lReturn = m.Target2.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                
            Case m.Target3.OrderID
                bFound = True
                strOrderName = m.Target3.Name
                lReturn = m.Target3.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                
            Case m.Stop1.OrderID
                bFound = True
                strOrderName = m.Stop1.Name
                lReturn = m.Stop1.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                m.Stop1.Waiting = False
                
            Case m.Stop2.OrderID
                bFound = True
                strOrderName = m.Stop2.Name
                lReturn = m.Stop2.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                m.Stop2.Waiting = False
                
            Case m.Stop3.OrderID
                bFound = True
                strOrderName = m.Stop3.Name
                lReturn = m.Stop3.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                m.Stop3.Waiting = False
                
            Case m.StopRest.OrderID
                bFound = True
                strOrderName = m.StopRest.Name
                lReturn = m.StopRest.OrderCallback(Order, strOrderName, m.lCurrentPosition)
                m.StopRest.Waiting = False
                
        End Select
        
        If m.bDeactivating Then
            If HasWorkingOrders(-1&, Not m.bManualDeactivate) = False Then
                DumpDebug vbTab & "Deactivating (No Working Orders)"
                Active = False
                m.bDeactivating = False
            Else
                DumpDebug vbTab & "Deactivating (Working Orders)"
                CancelAllOrders -1&, False, Not m.bManualDeactivate
            End If
        Else
            If bFound Then
                ' If this order just got rejected, then perform the appropriate activity based on
                ' the user option (either flatten the position and notify the user, ask the user
                ' if they want to flatten, or do nothing)...
                If Order.Status = eTT_OrderStatus_Rejected Then
                    If Flattening = False Then
                        DumpDebug vbTab & "Rejected (Not Flattening)"
                        Select Case g.OrderStrategies.RejectOption
                            Case eGDOrderRejectOption_Flatten
                                Flatten "Rejected Order"
                                InfBox "The position for symbol '" & Symbol & "' in account '" & g.Broker.AccountNameForID(AccountID) & "' is being flattened because an order for the '" & m.ExitStrategy.StrategyName & "' automated exit has been rejected.|", "!", , m.ExitStrategy.StrategyName, True
                                
                            Case eGDOrderRejectOption_Ask
                                If InfBox("The " & strOrderName & " order for the '" & m.ExitStrategy.StrategyName & "' automated exit for symbol '" & Symbol & "' in account '" & g.Broker.AccountNameForID(AccountID) & "' has been rejected.||This may leave your position unprotected.||Would you like to flatten your position?|", "?", "+Flatten|-Do Nothing", m.ExitStrategy.StrategyName, , g.OrderStrategies.RejectTimeout) = "F" Then
                                    Flatten "Rejected Order"
                                    InfBox "The position for symbol '" & Symbol & "' in account '" & g.Broker.AccountNameForID(AccountID) & "' is being flattened because an order for the '" & m.ExitStrategy.StrategyName & "' automated exit has been rejected.|", "!", , m.ExitStrategy.StrategyName, True
                                End If
                                
                            Case eGDOrderRejectOption_Nothing
                        End Select
                    End If
                
                ' If this order just got an unsolicited cancel (or it expired because it was at the
                ' end of the day), set the appropriate quantity to this orders remaining quantity so
                ' that when a new order gets generated during the next session, it gets the correct
                ' quantity (we do not want this to happen, however, if the order has gone filled or
                ' has just gone amended because the remaning quantity will not necessarily reflect
                ' the actual position)...
                ElseIf (UnsolicitedCancel(Order) = True) And (Flattening = False) Then
                    DumpDebug vbTab & "Unsolicited Cancel (Not Flattening)"
                    SetOrderQuantities Order.OrderID, "", "Order Callback: Unsolicited Cancel"
                    ModifyOrders
                
                ' If the order just went working, but needs the quantity changed, do it (e.g. a Break Even
                ' or Trailing Stop order that was in Amend Pending status when the position changed)...
                ElseIf lReturn = 2& Then
                    If CanModify(Order) = True Then
                        DumpDebug vbTab & "Can Modify"
                        SetOrderQuantities kNullData, "", "Order Callback: Can Modify"
                        ModifyOrders
                    End If
                End If
            End If
        End If
        
        Save
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.OrderCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Update the orders in the object based on the given fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FillCallback(ByVal Fill As cPtFill, ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lOldPosition As Long            ' Previous position with the fill
    Dim lNewPosition As Long            ' New position with the fill
    Dim AcctPos As cAccountPosition     ' Account position object
    
    If m.bActive And (Not CancellingAll) Then
        lOldPosition = m.lCurrentPosition
        lNewPosition = g.Broker.CurrentPosition(AccountID, SymbolOrSymbolID, 0&) - TradeSenseGroupPosition
        
        Set AcctPos = g.Broker.FillSummary(AccountID, SymbolOrSymbolID, 0&)
        
        Select Case Order.OrderID
            Case m.Target1.OrderID, m.Target1.PrevOrderID
                m.Target1.Order = Order
                
            Case m.Target2.OrderID, m.Target2.PrevOrderID
                m.Target2.Order = Order
                
            Case m.Target3.OrderID, m.Target3.PrevOrderID
                m.Target3.Order = Order
                
            Case m.Stop1.OrderID, m.Stop1.PrevOrderID
                m.Stop1.Order = Order
                
            Case m.Stop2.OrderID, m.Stop2.PrevOrderID
                m.Stop2.Order = Order
                
            Case m.Stop3.OrderID, m.Stop3.PrevOrderID
                m.Stop3.Order = Order
                
            Case m.StopRest.OrderID, m.StopRest.PrevOrderID
                m.StopRest.Order = Order
                
        End Select
        
        If lNewPosition <> lOldPosition Then
            ' Determine the new position...
            m.lCurrentPosition = lNewPosition
            
            ' The current position just went non-zero...
            If (lOldPosition = 0&) And (m.lCurrentPosition <> 0&) Then
                DumpDebug "FillCallback - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition) & "; Remaining Quantity = " & Str(Order.RemainingQuantity)
                
                ' DAJ 04/17/2013: If the user chose to cancel all manual working orders in the opposite
                ' direction when the auto exit gets initiated, do it here...
                CancelOrdersInOppositeDirection
                
                ' DAJ 09/07/2012: Make sure to reset all of the variables (including setting the HH and LL
                ' to the entry price) before doing CreateOrders.  In some circumstances, CreateOrders allows
                ' for a time slice which allows UpdateBars to happen.  That then wants to reset the HH and LL...
                ResetVariables Fill.Price, ConvertBrokerDate(Fill.FillDate, m.nAcctType, Symbol, False)
                ClearBreakEvenMovedFlag
                ResetOrderIDs
                ActivateTsOrders AcctPos
                
                SetOrderQuantities Fill.OrderID, "", "FillCallback: Non-Zero"
                CreateOrders True, "", "position has been entered"
            
            ' The current position just went to zero...
            ElseIf (lOldPosition <> 0&) And (m.lCurrentPosition = 0&) Then
                DumpDebug "FillCallback - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition) & "; Remaining Quantity = " & Str(Order.RemainingQuantity)
                
                CancelAllOrders Order.OrderID
                            
                If m.bStrategyDirty Then
                    LoadStrategy
                    DumpDebug vbTab & "CancelOpposite = " & Str(m.ExitStrategy.CancelOpposite) & "; MarketIfWrongSide = " & Str(m.ExitStrategy.MarketIfWrongSide)
                End If
                
                m.dEntryTime = 0#
                ClearBreakEvenMovedFlag
                m.dLastTrailTime = 0#
                
                SetOrderQuantities Fill.OrderID, "", "FillCallback: Zero"
                
                m.dHighestHigh = kNullData
                m.dLowestLow = kNullData
                                
                m.Bars.Size = 0&
            
            ' The current position just changed direction...
            ElseIf ((lOldPosition < 0&) And (m.lCurrentPosition > 0&)) Or ((lOldPosition > 0&) And (m.lCurrentPosition < 0&)) Then
                DumpDebug "FillCallback - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition) & "; Remaining Quantity = " & Str(Order.RemainingQuantity)
                
                CancelAllOrders Order.OrderID
                
                If m.bStrategyDirty Then
                    LoadStrategy
                    DumpDebug vbTab & "CancelOpposite = " & Str(m.ExitStrategy.CancelOpposite) & "; MarketIfWrongSide = " & Str(m.ExitStrategy.MarketIfWrongSide)
                End If
                
                ' DAJ 04/17/2013: If the user chose to cancel all manual working orders in the opposite
                ' direction when the auto exit gets initiated, do it here...
                CancelOrdersInOppositeDirection
                
                ' DAJ 09/07/2012: Make sure to reset all of the variables (including setting the HH and LL
                ' to the entry price) before doing CreateOrders.  In some circumstances, CreateOrders allows
                ' for a time slice which allows UpdateBars to happen.  That then wants to reset the HH and LL...
                ResetVariables Fill.Price, ConvertBrokerDate(Fill.FillDate, m.nAcctType, Symbol, False)
                ClearBreakEvenMovedFlag
                ResetOrderIDs
                ActivateTsOrders AcctPos
                
                SetOrderQuantities Fill.OrderID, "", "FillCallback: Changed Directions"
                CreateOrders True, "", "position has changed direction"
                        
            ' The current position changed, but in the same direction...
            Else
                DumpDebug "FillCallback - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition) & "; Remaining Quantity = " & Str(Order.RemainingQuantity)
                
                m.ExitStrategy.UpdateTsOrderPos AcctPos
                
                SetOrderQuantities Fill.OrderID, "", "FillCallback: Same Direction"
                ModifyOrders
            End If
        
            ' Make sure to close out the appropriate order ID if applicable...
            If Order.Status = eTT_OrderStatus_Filled Then
                Select Case Order.OrderID
                    Case m.Stop1.OrderID
                        m.Stop1.OrderID = -2&
                    Case m.Stop2.OrderID
                        m.Stop2.OrderID = -2&
                    Case m.Stop3.OrderID
                        m.Stop3.OrderID = -2&
                    Case m.StopRest.OrderID
                        m.StopRest.OrderID = -2&
                    Case m.Target1.OrderID
                        m.Target1.OrderID = -2&
                    Case m.Target2.OrderID
                        m.Target2.OrderID = -2&
                    Case m.Target3.OrderID
                        m.Target3.OrderID = -2&
                End Select
                
                Save
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.FillCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPosition
'' Description: Update the active exit based on the current position
'' Inputs:      None
'' Returns:     None
''
'' FillSumm:    Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshPosition()
On Error GoTo ErrSection:

    Dim lOldPosition As Long            ' Previous position with the fill
    Dim lNewPosition As Long            ' New position with the fill
    Dim AcctPos As cAccountPosition     ' Fill summary object
    
    If m.bActive And (Not CancellingAll) Then
        lOldPosition = m.lCurrentPosition
        
        Set AcctPos = g.Broker.FillSummary(AccountID, SymbolOrSymbolID, 0&)
        If Not AcctPos Is Nothing Then
            lNewPosition = AcctPos.CurrentPositionSnapshot - TradeSenseGroupPosition
            
            If lNewPosition <> lOldPosition Then
                ' Determine the new position...
                m.lCurrentPosition = lNewPosition
                
                ' The current position just went non-zero...
                If (lOldPosition = 0&) And (m.lCurrentPosition <> 0&) Then
                    DumpDebug "RefreshPosition (Going Non-Zero) - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition)
                    
                    ' DAJ 04/17/2013: If the user chose to cancel all manual working orders in the opposite
                    ' direction when the auto exit gets initiated, do it here...
                    CancelOrdersInOppositeDirection
                    
                    ' DAJ 09/07/2012: Make sure to reset all of the variables (including setting the HH and LL
                    ' to the entry price) before doing CreateOrders.  In some circumstances, CreateOrders allows
                    ' for a time slice which allows UpdateBars to happen.  That then wants to reset the HH and LL...
                    ResetVariables AcctPos.FirstOpenEntryPriceSnapshot, ConvertBrokerDate(AcctPos.FirstOpenEntryDateSnapshot, m.nAcctType, Symbol, False)
                    ClearBreakEvenMovedFlag
                    ResetOrderIDs
                    LoadBars
                    ActivateTsOrders AcctPos
                    
                    SetOrderQuantities kNullData, "", "RefreshPosition: Non-Zero"
                    CreateOrders True, "", "position has been entered"
                    
                ' The current position just went to zero...
                ElseIf (lOldPosition <> 0&) And (m.lCurrentPosition = 0&) Then
                    DumpDebug "RefreshPosition (Going Flat) - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition)
                    
                    CancelAllOrders
                                
                    If m.bStrategyDirty Then
                        LoadStrategy
                        DumpDebug vbTab & "CancelOpposite = " & Str(m.ExitStrategy.CancelOpposite) & "; MarketIfWrongSide = " & Str(m.ExitStrategy.MarketIfWrongSide)
                    End If
                    
                    m.dEntryTime = 0#
                    ClearBreakEvenMovedFlag
                    m.dLastTrailTime = 0#
                    
                    SetOrderQuantities kNullData, "", "RefreshPosition: Zero"
                    
                    m.dHighestHigh = kNullData
                    m.dLowestLow = kNullData
                    
                    g.RealTime.RemoveTickBuffer m.Bars
                    m.Bars.Size = 0&
                    
                    InfBox ""
                
                ' The current position just changed direction...
                ElseIf ((lOldPosition < 0&) And (m.lCurrentPosition > 0&)) Or ((lOldPosition > 0&) And (m.lCurrentPosition < 0&)) Then
                    DumpDebug "RefreshPosition (Reverse) - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition)
                    
                    CancelAllOrders
                    
                    If m.bStrategyDirty Then
                        LoadStrategy
                        DumpDebug vbTab & "CancelOpposite = " & Str(m.ExitStrategy.CancelOpposite) & "; MarketIfWrongSide = " & Str(m.ExitStrategy.MarketIfWrongSide)
                    End If
                    
                    ' DAJ 04/17/2013: If the user chose to cancel all manual working orders in the opposite
                    ' direction when the auto exit gets initiated, do it here...
                    CancelOrdersInOppositeDirection
                    
                    ' DAJ 09/07/2012: Make sure to reset all of the variables (including setting the HH and LL
                    ' to the entry price) before doing CreateOrders.  In some circumstances, CreateOrders allows
                    ' for a time slice which allows UpdateBars to happen.  That then wants to reset the HH and LL...
                    ResetVariables AcctPos.FirstOpenEntryPriceSnapshot, ConvertBrokerDate(AcctPos.FirstOpenEntryDateSnapshot, m.nAcctType, Symbol, False)
                    ClearBreakEvenMovedFlag
                    ResetOrderIDs
                    LoadBars
                    ActivateTsOrders AcctPos
                    
                    SetOrderQuantities kNullData, "", "RefreshPosition: Changed Direction"
                    CreateOrders True, "", "position has changed direction"
                
                ' The current position changed, but in the same direction...
                Else
                    DumpDebug "RefreshPosition (Same Direction) - OldPosition: " & Str(lOldPosition) & "; New Position: " & Str(m.lCurrentPosition)
                    
                    m.ExitStrategy.UpdateTsOrderPos AcctPos
            
                    SetOrderQuantities kNullData, "", "RefreshPosition: Same Direction"
                    ModifyOrders
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.RefreshPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Select Case Order.OrderID
        Case m.Stop1.OrderID
            m.Stop1.Order = Order
        
        Case m.Stop2.OrderID
            m.Stop2.Order = Order
        
        Case m.Stop3.OrderID
            m.Stop3.Order = Order
        
        Case m.StopRest.OrderID
            m.StopRest.Order = Order
        
        Case m.Target1.OrderID
            m.Target1.Order = Order
            
        Case m.Target2.OrderID
            m.Target2.Order = Order
            
        Case m.Target3.OrderID
            m.Target3.Order = Order
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIDChanged
'' Description: Notify the appropriate item that an order ID may have changed
'' Inputs:      Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderIDChanged(ByVal lOldOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object
    
    If m.bActive Then
        Select Case lOldOrderID
            Case m.Stop1.OrderID
                m.Stop1.OrderIDChanged lNewOrderID
                
            Case m.Stop2.OrderID
                m.Stop2.OrderIDChanged lNewOrderID
                
            Case m.Stop3.OrderID
                m.Stop3.OrderIDChanged lNewOrderID
                
            Case m.StopRest.OrderID
                m.StopRest.OrderIDChanged lNewOrderID
                
            Case m.Target1.OrderID
                m.Target1.OrderIDChanged lNewOrderID
            
            Case m.Target2.OrderID
                m.Target2.OrderIDChanged lNewOrderID
            
            Case m.Target3.OrderID
                m.Target3.OrderIDChanged lNewOrderID
                
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cActiveExit.OrderIDChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshStrategy
'' Description: Refresh the exit strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshStrategy()
On Error GoTo ErrSection:

    If (m.lCurrentPosition = 0&) Then
        LoadStrategy
        DumpDebug vbTab & "CancelOpposite = " & Str(m.ExitStrategy.CancelOpposite) & "; MarketIfWrongSide = " & Str(m.ExitStrategy.MarketIfWrongSide)
    Else
        m.bStrategyDirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.RefreshStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StrategyRenamed
'' Description: The strategy has been renamed
'' Inputs:      New Strategy Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StrategyRenamed(ByVal strNewFileName As String)
On Error GoTo ErrSection:

    m.strExitStrategyFile = strNewFileName
    Save
    
    RefreshStrategy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.StrategyRenamed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flatten
'' Description: Flatten the strategy
'' Inputs:      Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Flatten(Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    If Len(strReason) > 0 Then
        DumpDebug "Flattening (" & strReason & ")"
    Else
        DumpDebug "Flattening"
    End If
    g.FlattenQueue.AddToFlattenQueue g.Broker.AccountNumberForID(AccountID), Symbol, 0&, eGDFlattenQueueOperation_Flatten

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.Flatten"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAllOrders
'' Description: Cancel all of the orders that are open and not pending
'' Inputs:      Order ID to skip, Reset ID's?, Ignore Broker Held?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelAllOrders(Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal bResetIDs As Boolean = True, Optional ByVal bIgnoreBrokerHeld As Boolean = True)
On Error GoTo ErrSection:

    Dim strAccountNumber As String      ' Account number

    DumpDebug "CancelAllOrders(" & Str(lSkipOrderID) & ", " & Str(bResetIDs) & ", " & Str(bIgnoreBrokerHeld) & ")"

    If (Flattening = False) And (HasWorkingOrders(-1&, bIgnoreBrokerHeld) = True) Then
        If (m.Target3.OrderID <> 0) And (Not m.Target3.Order Is Nothing) And (m.Target3.OrderID <> lSkipOrderID) Then
            CancelRequested m.Target3.Order
        End If
        If (m.Target2.OrderID <> 0) And (Not m.Target2.Order Is Nothing) And (m.Target2.OrderID <> lSkipOrderID) Then
            CancelRequested m.Target2.Order
        End If
        If (m.Target1.OrderID <> 0) And (Not m.Target1.Order Is Nothing) And (m.Target1.OrderID <> lSkipOrderID) Then
            CancelRequested m.Target1.Order
        End If
        If (m.Stop1.OrderID <> 0) And (Not m.Stop1.Order Is Nothing) And (m.Stop1.OrderID <> lSkipOrderID) Then
            CancelRequested m.Stop1.Order
        End If
        If (m.Stop2.OrderID <> 0) And (Not m.Stop2.Order Is Nothing) And (m.Stop2.OrderID <> lSkipOrderID) Then
            CancelRequested m.Stop2.Order
        End If
        If (m.Stop3.OrderID <> 0) And (Not m.Stop3.Order Is Nothing) And (m.Stop3.OrderID <> lSkipOrderID) Then
            CancelRequested m.Stop3.Order
        End If
        If (m.StopRest.OrderID <> 0) And (Not m.StopRest.Order Is Nothing) And (m.StopRest.OrderID <> lSkipOrderID) Then
            CancelRequested m.StopRest.Order
        End If
    
        strAccountNumber = g.Broker.AccountNumberForID(AccountID)
        g.FlattenQueue.AddToFlattenQueue strAccountNumber, Symbol, 0&, eGDFlattenQueueOperation_CancelAll, lSkipOrderID, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.CancelAllOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelWorkingOrders
'' Description: Cancel the working orders that are not pending
'' Inputs:      Order ID to skip
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelWorkingOrders(Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal bResetIDs As Boolean = True)
On Error GoTo ErrSection:

    If m.bActive Then
        m.Target3.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
        m.Target2.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
        m.Target1.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
        
        m.Stop1.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
        m.Stop2.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
        m.Stop3.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
        m.StopRest.CancelWorkingOrder m.bDeactivating, lSkipOrderID, bResetIDs ', True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.CancelWorkingOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSessionDate
'' Description: Check to see if we are in a new session
'' Inputs:      Session Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateSessionDate(ByVal lSessionDate As Long)
On Error GoTo ErrSection:
    
    Dim strSavedInfo As String          ' Saved information from previous session
    
    If lSessionDate > m.lSessionDate Then
        If CanUpdateSession Then
            If (m.bActive = True) And (m.lCurrentPosition <> 0) Then
                DumpDebug "New Session Date = " & DateFormat(lSessionDate, MM_DD_YYYY)
                SaveLastKnown False
                
                m.Target3.CancelFromUpdateSession
                m.Target2.CancelFromUpdateSession
                m.Target1.CancelFromUpdateSession
                
                m.Stop1.CancelFromUpdateSession
                m.Stop2.CancelFromUpdateSession
                m.Stop3.CancelFromUpdateSession
                m.StopRest.CancelFromUpdateSession
            
                SetOrderQuantities kNullData, "", "UpdateSessionDate"
                strSavedInfo = GetIniFileProperty(Str(AccountID) & "|" & Symbol, "", "AutoExits", g.strIniFile)
                CreateOrders True, strSavedInfo, "new session has started"
            End If
            
            m.lSessionDate = lSessionDate
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.UpdateSessionDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelRequested
'' Description: Cancel was requested by Trade Navigator for the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelRequested(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Select Case Order.OrderID
        Case m.Stop1.OrderID
            m.Stop1.CancelRequested = True
            DumpDebug "Cancel Requested on Stop Loss = True"
            
        Case m.Stop2.OrderID
            m.Stop2.CancelRequested = True
            DumpDebug "Cancel Requested on Second Stop Loss = True"
            
        Case m.Stop3.OrderID
            m.Stop3.CancelRequested = True
            DumpDebug "Cancel Requested on Third Stop Loss = True"
            
        Case m.StopRest.OrderID
            m.StopRest.CancelRequested = True
            DumpDebug "Cancel Requested on Rest Stop Loss = True"
            
        Case m.Target1.OrderID
            m.Target1.CancelRequested = True
            DumpDebug "Cancel Requested on First Profit Target = True"
        
        Case m.Target2.OrderID
            m.Target2.CancelRequested = True
            DumpDebug "Cancel Requested on Second Profit Target = True"
        
        Case m.Target3.OrderID
            m.Target3.CancelRequested = True
            DumpDebug "Cancel Requested on Third Profit Target = True"

    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cActiveExit.CancelRequested"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckModifyOrders
'' Description: Check to see if any orders need to (and can be) modified
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckModifyOrders()
On Error GoTo ErrSection:

    Dim bModifyOrders As Boolean        ' Do we need to call modify orders?

    bModifyOrders = False
    If m.bActive Then
        bModifyOrders = m.Stop1.CheckModify
        If bModifyOrders = False Then
            bModifyOrders = m.Stop2.CheckModify
        End If
        If bModifyOrders = False Then
            bModifyOrders = m.Stop3.CheckModify
        End If
        If bModifyOrders = False Then
            bModifyOrders = m.StopRest.CheckModify
        End If
        If bModifyOrders = False Then
            bModifyOrders = m.Target1.CheckModify
        End If
        If bModifyOrders = False Then
            bModifyOrders = m.Target2.CheckModify
        End If
        If bModifyOrders = False Then
            bModifyOrders = m.Target3.CheckModify
        End If
        
        If bModifyOrders = True Then
            DumpDebug "One or more orders needs to be modified, but not because pending"
            ModifyOrders
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.CheckModifyOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollPosition
'' Description: Roll the position and orders into another contract
'' Inputs:      New Contract, Confirm?
'' Returns:     True if rolled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RollPosition(ByVal vNewSymbolOrSymbolID As Variant, Optional ByVal bConfirm As Boolean = True) As Boolean
On Error GoTo ErrSection:

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.RollPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCurrentInfo
'' Description: Retrieve the current information on the state of the auto exit
'' Inputs:      None
'' Returns:     Current Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCurrentInfo() As String
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Array of information about the auto exit
    Dim dCurrentPrice As Double         ' Current price for the symbol
    
    Set astrInfo = New cGdArray
    astrInfo.Create eGDARRAY_Strings
    
    dCurrentPrice = g.RealTime.LastKnownPrice(SymbolOrSymbolID)
    
    astrInfo(0) = Str(m.Stop1.Quantity)
    If m.Stop1.OrderID > 0 Then
        astrInfo(1) = Str(Abs(m.Stop1.Order.StopPrice - dCurrentPrice))
        astrInfo(2) = Str(Abs(m.Stop1.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(1) = "0"
        astrInfo(2) = "0"
    End If
    
    astrInfo(3) = Str(m.Target1.Quantity)
    If m.Target1.Quantity > 0 Then
        astrInfo(4) = Str(Abs(m.Target1.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(4) = "0"
    End If

    astrInfo(5) = Str(m.Target2.Quantity)
    If m.Target2.Quantity > 0 Then
        astrInfo(6) = Str(Abs(m.Target2.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(6) = "0"
    End If

    astrInfo(7) = Str(m.Target3.Quantity)
    If m.Target3.Quantity > 0 Then
        astrInfo(8) = Str(Abs(m.Target3.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(8) = "0"
    End If

    astrInfo(9) = Str(m.Stop2.Quantity)
    If m.Stop2.OrderID > 0 Then
        astrInfo(10) = Str(Abs(m.Stop2.Order.StopPrice - dCurrentPrice))
        astrInfo(11) = Str(Abs(m.Stop2.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(10) = "0"
        astrInfo(11) = "0"
    End If
    
    astrInfo(12) = Str(m.Stop3.Quantity)
    If m.Stop3.OrderID > 0 Then
        astrInfo(13) = Str(Abs(m.Stop3.Order.StopPrice - dCurrentPrice))
        astrInfo(14) = Str(Abs(m.Stop3.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(13) = "0"
        astrInfo(14) = "0"
    End If
    
    astrInfo(15) = Str(m.StopRest.Quantity)
    If m.StopRest.OrderID > 0 Then
        astrInfo(16) = Str(Abs(m.StopRest.Order.StopPrice - dCurrentPrice))
        astrInfo(17) = Str(Abs(m.StopRest.Order.LimitPrice - dCurrentPrice))
    Else
        astrInfo(16) = "0"
        astrInfo(17) = "0"
    End If
    
    GetCurrentInfo = astrInfo.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.GetCurrentInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsSimpleBracket
'' Description: Determine if the auto exit is a simple bracket or not
'' Inputs:      None
'' Returns:     True if Simple Bracket, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSimpleBracket() As Boolean
On Error GoTo ErrSection:

    IsSimpleBracket = m.ExitStrategy.IsSimpleBracket

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.IsSimpleBracket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HoldAtBroker
'' Description: Hold the OCO link at the broker?
'' Inputs:      None
'' Returns:     True if Hold at Broker, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HoldAtBroker() As Boolean
On Error GoTo ErrSection:
    
    'Dim bHasTrail As Boolean            ' Does this have a trailing stop?
    
    'bHasTrail = (m.ExitStrategy.StopLossType = eGDStopLossType_Trail) Or (m.ExitStrategy.UseTrail = True)
    HoldAtBroker = g.Broker.HoldOcoAtBroker(AccountID) ' And (bHasTrail = False) ' And IsSimpleBracket

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cActiveExit.HoldAtBroker"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOrderQuantities
'' Description: Set the quantities for all of the applicable orders
'' Inputs:      Order ID, Current Info
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOrderQuantities(Optional ByVal lOrderID As Long = kNullData, Optional ByVal strCurrentInfo As String = "", Optional ByVal strCalledFrom As String = "")
On Error GoTo ErrSection:

    Dim astrCurrentInfo As cGdArray     ' Current information split out into an array

    DumpDebug "SetOrderQuantities(" & Str(lOrderID) & ", '" & strCurrentInfo & "', '" & strCalledFrom & "')"
    If Len(strCalledFrom) = 0 Then
        DumpDebug "Order IDs - T1:" & Str(m.Target1.OrderID) & ", T2:" & Str(m.Target2.OrderID) & ", T3:" & Str(m.Target3.OrderID) & ", S1:" & Str(m.Stop1.OrderID) & ", S2:" & Str(m.Stop2.OrderID) & ", S3:" & Str(m.Stop3.OrderID) & ", SR:" & Str(m.StopRest.OrderID)
    Else
        DumpDebug "Order IDs - T1:" & Str(m.Target1.OrderID) & ", T2:" & Str(m.Target2.OrderID) & ", T3:" & Str(m.Target3.OrderID) & ", S1:" & Str(m.Stop1.OrderID) & ", S2:" & Str(m.Stop2.OrderID) & ", S3:" & Str(m.Stop3.OrderID) & ", SR:" & Str(m.StopRest.OrderID) & " (" & strCalledFrom & ")"
    End If
    
    ' If the current position is flat, make sure that all quantities are set to zero...
    If m.lCurrentPosition = 0& Then
        DumpDebug vbTab & "Current Position is Flat, so set quantities to zero"
        SetNoTargetQuantities
        SetNoStopQuantities
    
    ElseIf Len(strCurrentInfo) > 0 Then
        DumpDebug vbTab & "Setting quantities from saved information: '" & strCurrentInfo & "'"
        Set astrCurrentInfo = New cGdArray
        astrCurrentInfo.SplitFields strCurrentInfo, vbTab
        m.Target1.Quantity = CLng(Val(astrCurrentInfo(3)))
        m.Target2.Quantity = CLng(Val(astrCurrentInfo(5)))
        m.Target3.Quantity = CLng(Val(astrCurrentInfo(7)))
        m.Stop1.Quantity = CLng(Val(astrCurrentInfo(0)))
        m.Stop2.Quantity = CLng(Val(astrCurrentInfo(9)))
        m.Stop3.Quantity = CLng(Val(astrCurrentInfo(12)))
        m.StopRest.Quantity = CLng(Val(astrCurrentInfo(15)))
        
    ElseIf IsTargetOrder(lOrderID) Then
        DumpDebug vbTab & "Only update given target order: " & Str(lOrderID)
        Select Case lOrderID
            Case m.Target1.Order.OrderID
                DumpDebug vbTab & vbTab & "T1; Remaining = " & Str(m.Target1.Order.RemainingQuantity)
                m.Target1.Quantity = m.Target1.Order.RemainingQuantity
                If UseStopLoss Then
                    m.Stop1.Quantity = m.Target1.Quantity
                End If
            Case m.Target2.Order.OrderID
                DumpDebug vbTab & vbTab & "T2; Remaining = " & Str(m.Target2.Order.RemainingQuantity)
                m.Target2.Quantity = m.Target2.Order.RemainingQuantity
                If UseStopLoss Then
                    m.Stop2.Quantity = m.Target1.Quantity
                End If
            Case m.Target3.Order.OrderID
                DumpDebug vbTab & vbTab & "T3; Remaining = " & Str(m.Target3.Order.RemainingQuantity)
                m.Target3.Quantity = m.Target3.Order.RemainingQuantity
                If UseStopLoss Then
                    m.Stop3.Quantity = m.Target1.Quantity
                End If
        End Select
    
    ElseIf IsStopOrder(lOrderID) Then
        DumpDebug vbTab & "Only update given stop order: " & Str(lOrderID)
        Select Case lOrderID
            Case m.Stop1.Order.OrderID
                DumpDebug vbTab & vbTab & "S1; Remaining = " & Str(m.Stop1.Order.RemainingQuantity)
                m.Stop1.Quantity = m.Stop1.Order.RemainingQuantity
                If m.Target1.Quantity > 0 Then
                    m.Target1.Quantity = m.Stop1.Quantity
                End If
            Case m.Stop2.Order.OrderID
                DumpDebug vbTab & vbTab & "S2; Remaining = " & Str(m.Stop2.Order.RemainingQuantity)
                m.Stop2.Quantity = m.Stop2.Order.RemainingQuantity
                If m.Target2.Quantity > 0 Then
                    m.Target2.Quantity = m.Stop2.Quantity
                End If
            Case m.Stop3.Order.OrderID
                DumpDebug vbTab & vbTab & "S3; Remaining = " & Str(m.Stop3.Order.RemainingQuantity)
                m.Stop3.Quantity = m.Stop3.Order.RemainingQuantity
                If m.Target3.Quantity > 0 Then
                    m.Target3.Quantity = m.Stop3.Quantity
                End If
            Case m.StopRest.Order.OrderID
                DumpDebug vbTab & vbTab & "SR; Remaining = " & Str(m.StopRest.Order.RemainingQuantity)
                m.StopRest.Quantity = m.StopRest.Order.RemainingQuantity
        End Select
    
    ElseIf (RemainingTargetQuantity = Abs(m.lCurrentPosition)) Then
        DumpDebug vbTab & "Current Position = RemainingTargetQuantity  (" & Str(m.lCurrentPosition) & ")"
        If m.Target1.OrderID > 0 Then
            If Not m.Target1.Order Is Nothing Then
                m.Target1.Quantity = m.Target1.Order.RemainingQuantity
            End If
        Else
            m.Target1.Quantity = 0&
        End If
        If m.Target2.OrderID > 0 Then
            If Not m.Target2.Order Is Nothing Then
                m.Target2.Quantity = m.Target2.Order.RemainingQuantity
            End If
        Else
            m.Target2.Quantity = 0&
        End If
        If m.Target3.OrderID > 0 Then
            If Not m.Target3.Order Is Nothing Then
                m.Target3.Quantity = m.Target3.Order.RemainingQuantity
            End If
        Else
            m.Target3.Quantity = 0&
        End If
        
        If UseStopLoss Then
            m.Stop1.Quantity = m.Target1.Quantity
            m.Stop2.Quantity = m.Target2.Quantity
            m.Stop3.Quantity = m.Target3.Quantity
            m.StopRest.Quantity = Abs(m.lCurrentPosition) - (m.Stop1.Quantity + m.Stop2.Quantity + m.Stop3.Quantity)
        End If
        
    ElseIf (RemainingStopQuantity = Abs(m.lCurrentPosition)) Then
        DumpDebug vbTab & "Current Position = RemainingStopQuantity  (" & Str(m.lCurrentPosition) & ")"
        If m.Stop1.OrderID > 0 Then
            If Not m.Stop1.Order Is Nothing Then
                m.Stop1.Quantity = m.Stop1.Order.RemainingQuantity
            End If
        Else
            m.Stop1.Quantity = 0&
        End If
        If m.Stop2.OrderID > 0 Then
            If Not m.Stop2.Order Is Nothing Then
                m.Stop2.Quantity = m.Stop2.Order.RemainingQuantity
            End If
        Else
            m.Stop2.Quantity = 0&
        End If
        If m.Stop3.OrderID > 0 Then
            If Not m.Stop3.Order Is Nothing Then
                m.Stop3.Quantity = m.Stop3.Order.RemainingQuantity
            End If
        Else
            m.Stop3.Quantity = 0&
        End If
        If m.StopRest.OrderID > 0 Then
            If Not m.StopRest.Order Is Nothing Then
                m.StopRest.Quantity = m.StopRest.Order.RemainingQuantity
            End If
        Else
            m.StopRest.Quantity = 0&
        End If
        
        If UseProfitTarget Then
            m.Target1.Quantity = m.Stop1.Quantity
            m.Target2.Quantity = m.Stop2.Quantity
            m.Target3.Quantity = m.Stop3.Quantity
        End If
    Else
        
        ' No profit targets are selected...
        If (m.ExitStrategy.UseTarget1 = False) And (m.ExitStrategy.UseTarget2 = False) And (m.ExitStrategy.UseTarget3 = False) Then
            SetNoLevels
            
        ' Only the first profit target is selected...
        ElseIf (m.ExitStrategy.UseTarget2 = False) And (m.ExitStrategy.UseTarget3 = False) Then
            If (m.Target1.OrderID < -1&) Then
                SetNoLevels
            Else
                SetOneLevel 1
            End If
            
        ' Only the first and second profit targets are selected...
        ElseIf (m.ExitStrategy.UseTarget3 = False) Then
            If (m.Target1.OrderID < -1&) And (m.Target2.OrderID < -1&) Then
                SetNoLevels
            ElseIf (m.Target1.OrderID > -2&) And (m.Target2.OrderID > -2&) Then
                SetTwoLevels 1, 2
            ElseIf (m.Target1.OrderID < -1&) Then
                SetOneLevel 2
            Else
                SetOneLevel 1
            End If
        
        ' All three profit targets are selected...
        Else
            If (CancellingAll) Or ((m.Target1.OrderID < -1&) And (m.Target2.OrderID < -1&) And (m.Target3.OrderID < -1&)) Then
                SetNoLevels
            ElseIf (m.Target1.OrderID > -2&) And (m.Target2.OrderID < -1&) And (m.Target3.OrderID < -1&) Then
                SetOneLevel 1
            ElseIf (m.Target1.OrderID < -1&) And (m.Target2.OrderID > -2&) And (m.Target3.OrderID < -1&) Then
                SetOneLevel 2
            ElseIf (m.Target1.OrderID < -1&) And (m.Target2.OrderID < -1&) And (m.Target3.OrderID > -2&) Then
                SetOneLevel 3
            ElseIf (m.Target1.OrderID > -2&) And (m.Target2.OrderID > -2&) And (m.Target3.OrderID < -1&) Then
                SetTwoLevels 1, 2
            ElseIf (m.Target1.OrderID < -1&) And (m.Target2.OrderID > -2&) And (m.Target3.OrderID > -2&) Then
                SetTwoLevels 2, 3
            ElseIf (m.Target1.OrderID > -2&) And (m.Target2.OrderID < -1&) And (m.Target3.OrderID > -2&) Then
                SetTwoLevels 1, 3
            Else
                SetThreeLevels
            End If
        End If
    End If
    
    If Len(strCalledFrom) = 0 Then
        DumpDebug "SetOrderQuantities - T1:" & Str(m.Target1.Quantity) & ", T2:" & Str(m.Target2.Quantity) & ", T3:" & Str(m.Target3.Quantity) & ", S1:" & Str(m.Stop1.Quantity) & ", S2:" & Str(m.Stop2.Quantity) & ", S3:" & Str(m.Stop3.Quantity) & ", SR:" & Str(m.StopRest.Quantity)
    Else
        DumpDebug "SetOrderQuantities - T1:" & Str(m.Target1.Quantity) & ", T2:" & Str(m.Target2.Quantity) & ", T3:" & Str(m.Target3.Quantity) & ", S1:" & Str(m.Stop1.Quantity) & ", S2:" & Str(m.Stop2.Quantity) & ", S3:" & Str(m.Stop3.Quantity) & ", SR:" & Str(m.StopRest.Quantity) & " (" & strCalledFrom & ")"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetOrderQuantities"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetNoTargetQuantities
'' Description: Reset quantities for the profit levels
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetNoTargetQuantities()
On Error GoTo ErrSection:

    m.Target1.Quantity = 0&
    m.Target2.Quantity = 0&
    m.Target3.Quantity = 0&
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetNoTargetQuantities"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetNoStopQuantities
'' Description: Reset quantities for the stop loss orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetNoStopQuantities()
On Error GoTo ErrSection:

    m.Stop1.Quantity = 0&
    m.Stop2.Quantity = 0&
    m.Stop3.Quantity = 0&
    m.StopRest.Quantity = 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetNoStopQuantities"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetNoLevels
'' Description: Set no profit target levels, but a stop if appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetNoLevels()
On Error GoTo ErrSection:

    SetNoTargetQuantities
    SetNoStopQuantities
    
    If UseStopLoss Then
        If (CancellingAll = False) And (m.StopRest.OrderID > -1&) Then
            m.StopRest.Quantity = Abs(m.lCurrentPosition)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetNoLevels"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOneLevel
'' Description: Set quantities for one profit target level
'' Inputs:      Profit Level
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOneLevel(ByVal lLevel As Long)
On Error GoTo ErrSection:

    Select Case lLevel
        Case 1
            If UseProfitTarget Then
                If Runner Then
                    m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
                Else
                    m.Target1.Quantity = Abs(m.lCurrentPosition)
                End If
                m.Target2.Quantity = 0&
                m.Target3.Quantity = 0&
            Else
                SetNoTargetQuantities
            End If
            
            If UseStopLoss Then
                m.Stop1.Quantity = m.Target1.Quantity
                m.Stop2.Quantity = 0&
                m.Stop3.Quantity = 0&
                m.StopRest.Quantity = Abs(m.lCurrentPosition) - m.Target1.Quantity
            Else
                SetNoStopQuantities
            End If
            
        Case 2
            If UseProfitTarget Then
                m.Target1.Quantity = 0&
                If Runner Then
                    m.Target2.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target2Quantity)
                Else
                    m.Target2.Quantity = Abs(m.lCurrentPosition)
                End If
                m.Target3.Quantity = 0&
            Else
                SetNoTargetQuantities
            End If
            
            If UseStopLoss Then
                m.Stop1.Quantity = 0&
                m.Stop2.Quantity = m.Target2.Quantity
                m.Stop3.Quantity = 0&
                m.StopRest.Quantity = Abs(m.lCurrentPosition) - m.Target2.Quantity
            Else
                SetNoStopQuantities
            End If
            
        Case 3
            If UseProfitTarget Then
                m.Target1.Quantity = 0&
                m.Target2.Quantity = 0&
                If Runner Then
                    m.Target3.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target3Quantity)
                Else
                    m.Target3.Quantity = Abs(m.lCurrentPosition)
                End If
            Else
                SetNoTargetQuantities
            End If
            
            If UseStopLoss Then
                m.Stop1.Quantity = 0&
                m.Stop2.Quantity = 0&
                m.Stop3.Quantity = m.Target3.Quantity
                m.StopRest.Quantity = Abs(m.lCurrentPosition) - m.Target3.Quantity
            Else
                SetNoStopQuantities
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetOneLevel"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTwoLevels
'' Description: Set quantities for two profit target levels
'' Inputs:      First Profit Level, Second Profit Level
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTwoLevels(ByVal lLevel1 As Long, ByVal llevel2 As Long)
On Error GoTo ErrSection:

    If UseProfitTarget Then
        If (lLevel1 = 1) And (llevel2 = 2) Then
            If m.ExitStrategy.SpecifyLots = False Then
                m.Target1.Quantity = RoundQuantity(Int(Abs(m.lCurrentPosition) / 2), False)
                m.Target2.Quantity = Abs(m.lCurrentPosition) - m.Target1.Quantity
            ElseIf m.ExitStrategy.ExitEntirePosition = True Then
                m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
                m.Target2.Quantity = Abs(m.lCurrentPosition) - m.Target1.Quantity
            Else
                m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
                m.Target2.Quantity = MinLong(Abs(m.lCurrentPosition) - m.Target1.Quantity, m.ExitStrategy.Target2Quantity)
            End If
            m.Target3.Quantity = 0&
            
        ElseIf (lLevel1 = 1) And (llevel2 = 3) Then
            If m.ExitStrategy.SpecifyLots = False Then
                m.Target1.Quantity = RoundQuantity(Int(Abs(m.lCurrentPosition) / 2), False)
                m.Target3.Quantity = Abs(m.lCurrentPosition) - m.Target1.Quantity
            ElseIf m.ExitStrategy.ExitEntirePosition = True Then
                m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
                m.Target3.Quantity = Abs(m.lCurrentPosition) - m.Target1.Quantity
            Else
                m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
                m.Target3.Quantity = MinLong(Abs(m.lCurrentPosition) - m.Target1.Quantity, m.ExitStrategy.Target3Quantity)
            End If
            m.Target2.Quantity = 0&
            
        ElseIf (lLevel1 = 2) And (llevel2 = 3) Then
            If m.ExitStrategy.SpecifyLots = False Then
                m.Target2.Quantity = RoundQuantity(Int(Abs(m.lCurrentPosition) / 2), False)
                m.Target3.Quantity = Abs(m.lCurrentPosition) - m.Target2.Quantity
            ElseIf m.ExitStrategy.ExitEntirePosition = True Then
                m.Target2.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target2Quantity)
                m.Target3.Quantity = Abs(m.lCurrentPosition) - m.Target2.Quantity
            Else
                m.Target2.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target2Quantity)
                m.Target3.Quantity = MinLong(Abs(m.lCurrentPosition) - m.Target2.Quantity, m.ExitStrategy.Target3Quantity)
            End If
            m.Target1.Quantity = 0&
            
        End If
    Else
        SetNoTargetQuantities
    End If
    

    If UseStopLoss Then
        m.Stop1.Quantity = m.Target1.Quantity
        m.Stop2.Quantity = m.Target2.Quantity
        m.Stop3.Quantity = m.Target3.Quantity
        m.StopRest.Quantity = Abs(m.lCurrentPosition) - (m.Stop1.Quantity + m.Stop2.Quantity + m.Stop3.Quantity)
    Else
        SetNoStopQuantities
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetTwoLevels"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetThreeLevels
'' Description: Set quantities for the three profit target levels
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetThreeLevels()
On Error GoTo ErrSection:

    If UseProfitTarget Then
        If m.ExitStrategy.SpecifyLots = False Then
            m.Target1.Quantity = RoundQuantity(Int(Abs(m.lCurrentPosition) / 3), False)
            m.Target2.Quantity = RoundQuantity(Int(((Abs(m.lCurrentPosition) - m.Target1.Quantity) / 2)), True)
            m.Target3.Quantity = Abs(m.lCurrentPosition) - (m.Target1.Quantity + m.Target2.Quantity)
        ElseIf m.ExitStrategy.ExitEntirePosition Then
            m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
            m.Target2.Quantity = MinLong(Abs(m.lCurrentPosition) - m.Target1.Quantity, m.ExitStrategy.Target2Quantity)
            m.Target3.Quantity = Abs(m.lCurrentPosition) - (m.Target1.Quantity + m.Target2.Quantity)
        Else
            m.Target1.Quantity = MinLong(Abs(m.lCurrentPosition), m.ExitStrategy.Target1Quantity)
            m.Target2.Quantity = MinLong(Abs(m.lCurrentPosition) - m.Target1.Quantity, m.ExitStrategy.Target2Quantity)
            m.Target3.Quantity = MinLong(Abs(m.lCurrentPosition) - (m.Target1.Quantity + m.Target2.Quantity), m.ExitStrategy.Target3Quantity)
        End If
    Else
        SetNoTargetQuantities
    End If

    If UseStopLoss Then
        m.Stop1.Quantity = m.Target1.Quantity
        m.Stop2.Quantity = m.Target2.Quantity
        m.Stop3.Quantity = m.Target3.Quantity
        m.StopRest.Quantity = Abs(m.lCurrentPosition) - (m.Stop1.Quantity + m.Stop2.Quantity + m.Stop3.Quantity)
    Else
        SetNoStopQuantities
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetThreeLevels"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateTargetOrder
'' Description: Create a profit target order for the given profit level
'' Inputs:      Profit Level, Saved Information
'' Returns:     True if Order Created, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateTargetOrder(ByVal lProfitLevel As Long, Optional ByVal strSavedInfo As String = "") As Boolean
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object to create
    Dim strEquation As String           ' Equation for calculating the limit price
    Dim bReturn As Boolean              ' Return value for the function
    
    DumpDebug "CreateTargetOrder(" & Str(lProfitLevel) & ", '" & strSavedInfo & "')"
    
    bReturn = True
    With Order
        .AccountID = AccountID
        .IsAutoExit = True
        .IsAutomated = True
        
        If m.lCurrentPosition < 0 Then
            .Buy = True
            
            If m.ExitStrategy.ProfitTargetType <> eGDProfitTargetType_TradeSense Then
                .OrderType = eTT_OrderType_Limit
                .LimitPrice = Val(Parse(strSavedInfo, ";", 2))
                strEquation = Str(.LimitPrice) & " (Saved Info)"
                If .LimitPrice = 0# Then
                    Select Case lProfitLevel
                        Case 1
                            .LimitPrice = m.dEntryPrice - (m.ExitStrategy.Target1Ticks * TickMove)
                            strEquation = Str(m.dEntryPrice) & " - (" & Str(m.ExitStrategy.Target1Ticks) & " * " & Str(TickMove) & ")"
                        Case 2
                            .LimitPrice = m.dEntryPrice - (m.ExitStrategy.Target2Ticks * TickMove)
                            strEquation = Str(m.dEntryPrice) & " - (" & Str(m.ExitStrategy.Target2Ticks) & " * " & Str(TickMove) & ")"
                        Case 3
                            .LimitPrice = m.dEntryPrice - (m.ExitStrategy.Target3Ticks * TickMove)
                            strEquation = Str(m.dEntryPrice) & " - (" & Str(m.ExitStrategy.Target3Ticks) & " * " & Str(TickMove) & ")"
                    End Select
                End If
            Else
                With m.ExitStrategy.TsProfitShort
                    If .ConditionTrue = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because condition false"
                    ElseIf .AllRtDataAvailable = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because waiting for data"
                    Else
                        strEquation = .EnglishOrderPrice
                        Order.OrderType = .OrderType
                        Order.StopPrice = .StopPrice
                        Order.LimitPrice = .LimitPrice
                    End If
                End With
            End If
        Else
            .Buy = False
            
            If m.ExitStrategy.ProfitTargetType <> eGDProfitTargetType_TradeSense Then
                .OrderType = eTT_OrderType_Limit
                .LimitPrice = Val(Parse(strSavedInfo, ";", 2))
                strEquation = Str(.LimitPrice) & " (Saved Info)"
                If .LimitPrice = 0# Then
                    Select Case lProfitLevel
                        Case 1
                            .LimitPrice = m.dEntryPrice + (m.ExitStrategy.Target1Ticks * TickMove)
                            strEquation = Str(m.dEntryPrice) & " + (" & Str(m.ExitStrategy.Target1Ticks) & " * " & Str(TickMove) & ")"
                        Case 2
                            .LimitPrice = m.dEntryPrice + (m.ExitStrategy.Target2Ticks * TickMove)
                            strEquation = Str(m.dEntryPrice) & " + (" & Str(m.ExitStrategy.Target2Ticks) & " * " & Str(TickMove) & ")"
                        Case 3
                            .LimitPrice = m.dEntryPrice + (m.ExitStrategy.Target3Ticks * TickMove)
                            strEquation = Str(m.dEntryPrice) & " + (" & Str(m.ExitStrategy.Target3Ticks) & " * " & Str(TickMove) & ")"
                    End Select
                End If
            Else
                With m.ExitStrategy.TsProfitLong
                    If .ConditionTrue = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because condition false"
                    ElseIf .AllRtDataAvailable = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because waiting for data"
                    Else
                        strEquation = .EnglishOrderPrice
                        Order.OrderType = .OrderType
                        Order.StopPrice = .StopPrice
                        Order.LimitPrice = .LimitPrice
                    End If
                End With
            End If
        End If
        
        If bReturn = True Then
            .Enter = False
            
            ' Make sure that the order is a Day order...
            .Expiration = -1&
            
            .OrderDate = .BrokerDate(CurrentTime("", Symbol))
            Select Case lProfitLevel
                Case 1
                    .Quantity = m.Target1.Quantity
                Case 2
                    .Quantity = m.Target2.Quantity
                Case 3
                    .Quantity = m.Target3.Quantity
            End Select
            .Status = eTT_OrderStatus_Open
            .StatusDate = ConvertToBrokerDate(CurrentTime, g.Broker.AccountTypeForID(AccountID), Symbol)
            .SymbolOrSymbolID = SymbolOrSymbolID
            
            ' If the OCO link needs to be held at the broker end and we have the other order ID,
            ' then set up the link...
            If HoldAtBroker Then
                Select Case lProfitLevel
                    Case 1
                        If m.Stop1.OrderID > 0 Then
                            .BrokerCancelOrderID = m.Stop1.OrderID * -1&
                        End If
                        
                    Case 2
                        If m.Stop2.OrderID > 0 Then
                            .BrokerCancelOrderID = m.Stop2.OrderID * -1&
                        End If
                        
                    Case 3
                        If m.Stop3.OrderID > 0 Then
                            .BrokerCancelOrderID = m.Stop3.OrderID * -1&
                        End If
                End Select
            End If
            
            .Save
            
            Select Case lProfitLevel
                Case 1
                    m.Target1.OrderID = .OrderID
                    m.Target1.Order = Order
                    DumpDebug "T1 Order Created: " & m.Target1.Order.OrderText & " (" & m.Target1.Order.GenesisOrderID & ") = " & strEquation
                Case 2
                    m.Target2.OrderID = .OrderID
                    m.Target2.Order = Order
                    DumpDebug "T2 Order Created: " & m.Target2.Order.OrderText & " (" & m.Target2.Order.GenesisOrderID & ") = " & strEquation
                Case 3
                    m.Target3.OrderID = .OrderID
                    m.Target3.Order = Order
                    DumpDebug "T3 Order Created: " & m.Target3.Order.OrderText & " (" & m.Target3.Order.GenesisOrderID & ") = " & strEquation
            End Select
            
            Save
        End If
    End With
    
    CreateTargetOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.CreateTargetOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateStopLossOrder
'' Description: Create the stop loss order
'' Inputs:      Level, Saved Information
'' Returns:     True if order created, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateStopLossOrder(ByVal lStopLevel As Long, Optional ByVal strSavedInfo As String = "") As Boolean
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object to create
    Dim strEquation As String           ' Equation for calculating the stop price
    Dim bReturn As Boolean              ' Return value for the function
    
    DumpDebug "CreateStopLossOrder(" & Str(lStopLevel) & ", '" & strSavedInfo & "')"
    
    bReturn = True
    With Order
        .AccountID = AccountID
        .IsAutoExit = True
        .IsAutomated = True
        
        If m.lCurrentPosition < 0 Then
            .Buy = True
            
            If m.ExitStrategy.StopLossType <> eGDStopLossType_TradeSense Then
                .StopPrice = Val(Parse(strSavedInfo, ";", 3))
                strEquation = Str(.StopPrice) & " (Saved Info)"
                If .StopPrice = 0# Then
                    .StopPrice = m.dEntryPrice + (m.ExitStrategy.StopLossTicks * TickMove)
                    strEquation = Str(m.dEntryPrice) & " + (" & Str(m.ExitStrategy.StopLossTicks) & " * " & Str(TickMove) & ")"
                    If m.ExitStrategy.UseWithLimit Then
                        .OrderType = eTT_OrderType_StopWithLimit
                        .LimitPrice = .StopPrice + (m.ExitStrategy.WithLimitTicks * TickMove)
                    Else
                        .OrderType = eTT_OrderType_Stop
                        .LimitPrice = 0#
                    End If
                Else
                    .LimitPrice = Val(Parse(strSavedInfo, ";", 2))
                    If .LimitPrice = kNullData Then
                        .OrderType = eTT_OrderType_Stop
                    Else
                        .OrderType = eTT_OrderType_StopWithLimit
                    End If
                End If
            Else
                With m.ExitStrategy.TsStopShort
                    If .ConditionTrue = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because condition false"
                    ElseIf .AllRtDataAvailable = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because waiting for data"
                    Else
                        strEquation = .EnglishOrderPrice
                        Order.OrderType = .OrderType
                        Order.StopPrice = .StopPrice
                        Order.LimitPrice = .LimitPrice
                    End If
                End With
            End If
        Else
            .Buy = False
            
            If m.ExitStrategy.StopLossType <> eGDStopLossType_TradeSense Then
                .StopPrice = Val(Parse(strSavedInfo, ";", 3))
                strEquation = Str(.StopPrice) & " (Saved Info)"
                If .StopPrice = 0# Then
                    .StopPrice = m.dEntryPrice - (m.ExitStrategy.StopLossTicks * TickMove)
                    strEquation = Str(m.dEntryPrice) & " - (" & Str(m.ExitStrategy.StopLossTicks) & " * " & Str(TickMove) & ")"
                    If m.ExitStrategy.UseWithLimit Then
                        .OrderType = eTT_OrderType_StopWithLimit
                        .LimitPrice = .StopPrice - (m.ExitStrategy.WithLimitTicks * TickMove)
                    Else
                        .OrderType = eTT_OrderType_Stop
                        .LimitPrice = 0#
                    End If
                Else
                    .LimitPrice = Val(Parse(strSavedInfo, ";", 2))
                    If .LimitPrice = kNullData Then
                        .OrderType = eTT_OrderType_Stop
                    Else
                        .OrderType = eTT_OrderType_StopWithLimit
                    End If
                End If
            Else
                With m.ExitStrategy.TsStopLong
                    If .ConditionTrue = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because condition false"
                    ElseIf .AllRtDataAvailable = False Then
                        bReturn = False
                        DumpDebug vbTab & "Not created because waiting for data"
                    Else
                        strEquation = .EnglishOrderPrice
                        Order.OrderType = .OrderType
                        Order.StopPrice = .StopPrice
                        Order.LimitPrice = .LimitPrice
                    End If
                End With
            End If
        End If
        
        If bReturn = True Then
            .Enter = False
            
            ' Make sure that the order is a Day order...
            .Expiration = -1&
            
            .OrderDate = .BrokerDate(CurrentTime("", Symbol))
            Select Case lStopLevel
                Case 0
                    .Quantity = m.StopRest.Quantity
                Case 1
                    .Quantity = m.Stop1.Quantity
                Case 2
                    .Quantity = m.Stop2.Quantity
                Case 3
                    .Quantity = m.Stop3.Quantity
            End Select
            .Status = eTT_OrderStatus_Open
            .StatusDate = ConvertToBrokerDate(CurrentTime, g.Broker.AccountTypeForID(AccountID), Symbol)
            .SymbolOrSymbolID = SymbolOrSymbolID
            
            ' If the OCO link needs to be held at the broker end and we have the other order ID,
            ' then set up the link...
            If HoldAtBroker Then
                Select Case lStopLevel
                    Case 0
                        .BrokerCancelOrderID = 0&
                        
                    Case 1
                        If m.Target1.OrderID > 0 Then
                            .BrokerCancelOrderID = m.Target1.OrderID * -1&
                        End If
                    
                    Case 2
                        If m.Target2.OrderID > 0 Then
                            .BrokerCancelOrderID = m.Target2.OrderID * -1&
                        End If
                    
                    Case 3
                        If m.Target3.OrderID > 0 Then
                            .BrokerCancelOrderID = m.Target3.OrderID * -1&
                        End If
                End Select
            End If
            
            ' DAJ 04/17/2013: If the stop order is going to go in on the wrong side of the market
            ' and the user chose to flatten if that is the case, go ahead and submit a market order
            ' here instead of a stop...
            If (Order.WrongSideOfMarket = True) And (m.ExitStrategy.MarketIfWrongSide = True) Then
                DumpDebug vbTab & "Stop Order is on wrong side of market, so changing it to a market order"
                .StopPrice = 0#
                .LimitPrice = 0#
                .OrderType = eTT_OrderType_Market
            End If
            
            .Save
            
            Select Case lStopLevel
                Case 0
                    m.StopRest.OrderID = .OrderID
                    m.StopRest.Order = Order
                    DumpDebug "SR Order Created: " & m.StopRest.Order.OrderText & " (" & m.StopRest.Order.GenesisOrderID & ") = " & strEquation
                Case 1
                    m.Stop1.OrderID = .OrderID
                    m.Stop1.Order = Order
                    DumpDebug "S1 Order Created: " & m.Stop1.Order.OrderText & " (" & m.Stop1.Order.GenesisOrderID & ") = " & strEquation
                Case 2
                    m.Stop2.OrderID = .OrderID
                    m.Stop2.Order = Order
                    DumpDebug "S2 Order Created: " & m.Stop2.Order.OrderText & " (" & m.Stop2.Order.GenesisOrderID & ") = " & strEquation
                Case 3
                    m.Stop3.OrderID = .OrderID
                    m.Stop3.Order = Order
                    DumpDebug "S3 Order Created: " & m.Stop3.Order.OrderText & " (" & m.Stop3.Order.GenesisOrderID & ") = " & strEquation
            End Select
            
            Save
        End If
    End With
    
    CreateStopLossOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.CreateStopLossOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrders
'' Description: Create (and possibly submit) all of the orders applicable
'' Inputs:      Submit Orders?, Saved Information, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateOrders(Optional ByVal bSubmitOrders As Boolean = True, Optional ByVal strSavedInfo As String = "", Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    Dim astrSavedInfo As cGdArray       ' Saved information from previous session

    Set astrSavedInfo = New cGdArray
    astrSavedInfo.Create eGDARRAY_Strings
    astrSavedInfo.SplitFields strSavedInfo, "|"
    
    DumpDebug "CreateOrders(" & Str(bSubmitOrders) & ", '" & strSavedInfo & "', '" & strReason & "')"

    If m.Stop1.Quantity > 0& Then
        If m.Stop1.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating First Stop Order"
            If CreateStopLossOrder(1, astrSavedInfo(3)) Then
                If bSubmitOrders Then
                    m.Stop1.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create First Stop Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create First Stop Order because one already exists"
        End If
    End If
    
    If m.Stop2.Quantity > 0& Then
        If m.Stop2.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating Second Stop Order"
            If CreateStopLossOrder(2, astrSavedInfo(7)) Then
                If bSubmitOrders Then
                    m.Stop2.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create Second Stop Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create Second Stop Order because one already exists"
        End If
    End If
    
    If m.Stop3.Quantity > 0& Then
        If m.Stop3.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating Third Stop Order"
            If CreateStopLossOrder(3, astrSavedInfo(8)) Then
                If bSubmitOrders Then
                    m.Stop3.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create Third Stop Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create Third Stop Order because one already exists"
        End If
    End If
    
    If m.StopRest.Quantity > 0& Then
        If m.StopRest.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating Rest Stop Order"
            If CreateStopLossOrder(0, astrSavedInfo(9)) Then
                If bSubmitOrders Then
                    m.StopRest.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create Rest Stop Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create Rest Stop Order because one already exists"
        End If
    End If
    
    If m.Target1.Quantity > 0& Then
        If m.Target1.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating Target1 Order"
            If CreateTargetOrder(1, astrSavedInfo(4)) Then
                If bSubmitOrders Then
                    m.Target1.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create Target1 Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create Target1 Order because one already exists"
        End If
    End If
    
    If m.Target2.Quantity > 0& Then
        If m.Target2.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating Target2 Order"
            If CreateTargetOrder(2, astrSavedInfo(5)) Then
                If bSubmitOrders Then
                    m.Target2.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create Target2 Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create Target2 Order because one already exists"
        End If
    End If
    
    If m.Target3.Quantity > 0& Then
        If m.Target3.OrderID = 0& Then
            DumpDebug "Create Orders --> Creating Target3 Order"
            If CreateTargetOrder(3, astrSavedInfo(6)) Then
                If bSubmitOrders Then
                    m.Target3.SubmitOrder strReason
                End If
            Else
                DumpDebug "Create Orders --> Did not create Target3 Order because Trade Sense condition is False"
            End If
        Else
            DumpDebug "Create Orders --> Could not create Target3 Order because one already exists"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.CreateOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyOrders
'' Description: Verify that all orders are on the right side of the market
'' Inputs:      Saved Info
'' Returns:     True if all OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyOrders(Optional ByVal strSavedInfo As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrSavedInfo As cGdArray       ' Saved information from previous session
    Dim dPrice As Double                ' Order price

    Set astrSavedInfo = New cGdArray
    astrSavedInfo.Create eGDARRAY_Strings
    astrSavedInfo.SplitFields strSavedInfo, "|"

    bReturn = True

    If m.lCurrentPosition < 0 Then
        If m.Stop1.Quantity > 0 Then
            If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
                If (m.ExitStrategy.TsStopShort.OrderType = eTT_OrderType_Stop) Or (m.ExitStrategy.TsStopShort.OrderType = eTT_OrderType_StopWithLimit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopShort.StopPrice, SymbolOrSymbolID) <= 2 Then
                        DumpDebug vbTab & "S1: " & Str(m.ExitStrategy.TsStopShort.StopPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                ElseIf (m.ExitStrategy.TsStopShort.OrderType = eTT_OrderType_Limit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopShort.LimitPrice, SymbolOrSymbolID) >= -2 Then
                        DumpDebug vbTab & "S1: " & Str(m.ExitStrategy.TsStopShort.LimitPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                End If
            Else
                dPrice = Val(Parse(astrSavedInfo(3), ";", 3))
                If dPrice = 0# Then
                    dPrice = m.dEntryPrice + (m.ExitStrategy.StopLossTicks * TickMove)
                End If
            
                If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                    DumpDebug vbTab & "S1: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                    bReturn = False
                End If
            End If
        End If
        
        If (m.Stop2.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(7), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice + (m.ExitStrategy.StopLossTicks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                DumpDebug vbTab & "S2: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
        
        If (m.Stop3.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(8), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice + (m.ExitStrategy.StopLossTicks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                DumpDebug vbTab & "S3: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
        
        If (m.StopRest.Quantity > 0) And (bReturn = True) Then
            If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
                If (m.ExitStrategy.TsStopShort.OrderType = eTT_OrderType_Stop) Or (m.ExitStrategy.TsStopShort.OrderType = eTT_OrderType_StopWithLimit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopShort.StopPrice, SymbolOrSymbolID) <= 2 Then
                        DumpDebug vbTab & "SR: " & Str(m.ExitStrategy.TsStopShort.StopPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                ElseIf (m.ExitStrategy.TsStopShort.OrderType = eTT_OrderType_Limit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopShort.LimitPrice, SymbolOrSymbolID) >= -2 Then
                        DumpDebug vbTab & "SR: " & Str(m.ExitStrategy.TsStopShort.LimitPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                End If
            Else
                dPrice = Val(Parse(astrSavedInfo(9), ";", 3))
                If dPrice = 0# Then
                    dPrice = m.dEntryPrice + (m.ExitStrategy.StopLossTicks * TickMove)
                End If
                
                If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                    DumpDebug vbTab & "SR: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                    bReturn = False
                End If
            End If
        End If
        
        If (m.Target1.Quantity > 0) And (bReturn = True) Then
            If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
                If (m.ExitStrategy.TsProfitShort.OrderType = eTT_OrderType_Stop) Or (m.ExitStrategy.TsProfitShort.OrderType = eTT_OrderType_StopWithLimit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsProfitShort.StopPrice, SymbolOrSymbolID) <= 2 Then
                        DumpDebug vbTab & "T1: " & Str(m.ExitStrategy.TsProfitShort.StopPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                ElseIf (m.ExitStrategy.TsProfitShort.OrderType = eTT_OrderType_Limit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsProfitShort.LimitPrice, SymbolOrSymbolID) >= -2 Then
                        DumpDebug vbTab & "T1: " & Str(m.ExitStrategy.TsProfitShort.LimitPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                End If
            Else
                dPrice = Val(Parse(astrSavedInfo(4), ";", 3))
                If dPrice = 0# Then
                    dPrice = m.dEntryPrice - (m.ExitStrategy.Target1Ticks * TickMove)
                End If
                
                If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                    DumpDebug vbTab & "T1: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                    bReturn = False
                End If
            End If
        End If
        
        If (m.Target2.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(5), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice - (m.ExitStrategy.Target2Ticks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                DumpDebug vbTab & "T2: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
        
        If (m.Target3.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(6), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice - (m.ExitStrategy.Target3Ticks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                DumpDebug vbTab & "T3: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
    ElseIf m.lCurrentPosition > 0 Then
        If m.Stop1.Quantity > 0 Then
            If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
                If (m.ExitStrategy.TsStopLong.OrderType = eTT_OrderType_Stop) Or (m.ExitStrategy.TsStopLong.OrderType = eTT_OrderType_StopWithLimit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopLong.StopPrice, SymbolOrSymbolID) >= -2 Then
                        DumpDebug vbTab & "S1: " & Str(m.ExitStrategy.TsStopLong.StopPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                ElseIf (m.ExitStrategy.TsStopLong.OrderType = eTT_OrderType_Limit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopLong.LimitPrice, SymbolOrSymbolID) <= 2 Then
                        DumpDebug vbTab & "S1: " & Str(m.ExitStrategy.TsStopLong.LimitPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                End If
            Else
                dPrice = Val(Parse(astrSavedInfo(3), ";", 3))
                If dPrice = 0# Then
                    dPrice = m.dEntryPrice - (m.ExitStrategy.StopLossTicks * TickMove)
                End If
                
                If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                    DumpDebug vbTab & "S1: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                    bReturn = False
                End If
            End If
        End If
        
        If (m.Stop2.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(7), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice - (m.ExitStrategy.StopLossTicks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                DumpDebug vbTab & "S2: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
        
        If (m.Stop3.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(8), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice - (m.ExitStrategy.StopLossTicks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                DumpDebug vbTab & "S3: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
        
        If (m.StopRest.Quantity > 0) And (bReturn = True) Then
            If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
                If (m.ExitStrategy.TsStopLong.OrderType = eTT_OrderType_Stop) Or (m.ExitStrategy.TsStopLong.OrderType = eTT_OrderType_StopWithLimit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopLong.StopPrice, SymbolOrSymbolID) >= -2 Then
                        DumpDebug vbTab & "SR: " & Str(m.ExitStrategy.TsStopLong.StopPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                ElseIf (m.ExitStrategy.TsStopLong.OrderType = eTT_OrderType_Limit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsStopLong.LimitPrice, SymbolOrSymbolID) <= 2 Then
                        DumpDebug vbTab & "SR: " & Str(m.ExitStrategy.TsStopLong.LimitPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                End If
            Else
                dPrice = Val(Parse(astrSavedInfo(9), ";", 3))
                If dPrice = 0# Then
                    dPrice = m.dEntryPrice - (m.ExitStrategy.StopLossTicks * TickMove)
                End If
                
                If NumTicksFromMarket(dPrice, SymbolOrSymbolID) >= -2 Then
                    DumpDebug vbTab & "SR: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                    bReturn = False
                End If
            End If
        End If
        
        If (m.Target1.Quantity > 0) And (bReturn = True) Then
            If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
                If (m.ExitStrategy.TsProfitLong.OrderType = eTT_OrderType_Stop) Or (m.ExitStrategy.TsProfitLong.OrderType = eTT_OrderType_StopWithLimit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsProfitLong.StopPrice, SymbolOrSymbolID) >= -2 Then
                        DumpDebug vbTab & "T1: " & Str(m.ExitStrategy.TsProfitLong.StopPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                ElseIf (m.ExitStrategy.TsProfitLong.OrderType = eTT_OrderType_Limit) Then
                    If NumTicksFromMarket(m.ExitStrategy.TsProfitLong.LimitPrice, SymbolOrSymbolID) <= 2 Then
                        DumpDebug vbTab & "T1: " & Str(m.ExitStrategy.TsProfitLong.LimitPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                        bReturn = False
                    End If
                End If
            Else
                dPrice = Val(Parse(astrSavedInfo(4), ";", 3))
                If dPrice = 0# Then
                    dPrice = m.dEntryPrice + (m.ExitStrategy.Target1Ticks * TickMove)
                End If
                
                If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                    DumpDebug vbTab & "T1: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                    bReturn = False
                End If
            End If
        End If
        
        If (m.Target2.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(5), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice + (m.ExitStrategy.Target2Ticks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                DumpDebug vbTab & "T2: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
        
        If (m.Target3.Quantity > 0) And (bReturn = True) Then
            dPrice = Val(Parse(astrSavedInfo(6), ";", 3))
            If dPrice = 0# Then
                dPrice = m.dEntryPrice + (m.ExitStrategy.Target3Ticks * TickMove)
            End If
            
            If NumTicksFromMarket(dPrice, SymbolOrSymbolID) <= 2 Then
                DumpDebug vbTab & "T3: " & Str(dPrice) & " is less than two ticks away from the market(" & Str(g.RealTime.LastKnownPrice(SymbolOrSymbolID)) & ")"
                bReturn = False
            End If
        End If
    End If
    
    VerifyOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.VerifyOrders"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrders
'' Description: Modify the orders with the new position quantity
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ModifyOrders()
On Error GoTo ErrSection:

    If (Flattening = False) Then
        If UseStopLoss Then
            If StopLossConditionTrue Then
                If m.Stop1.ModifyTheOrder(m.Target1.OrderID) = True Then
                    If CreateStopLossOrder(1) Then
                        m.Stop1.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                        SubmitOrder m.Stop1.Order
                    
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal m.Stop1.Order.OrderID, CurrentTime, m.Stop1.Name & " generated by auto exit at " & m.Stop1.Order.StopPriceString & " because it doesn't exist yet"
                        End If
                    Else
                        DumpDebug "ModifyOrders - Did not create first stop loss order because condition was False"
                    End If
                End If
            
                If m.Stop2.ModifyTheOrder(m.Target2.OrderID) = True Then
                    If CreateStopLossOrder(2) Then
                        m.Stop2.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                        SubmitOrder m.Stop2.Order
                    
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal m.Stop2.Order.OrderID, CurrentTime, m.Stop2.Name & " generated by auto exit at " & m.Stop2.Order.StopPriceString & " because it doesn't exist yet"
                        End If
                    Else
                        DumpDebug "ModifyOrders - Did not create second stop loss order because condition was False"
                    End If
                End If
            
                If m.Stop3.ModifyTheOrder(m.Target3.OrderID) = True Then
                    If CreateStopLossOrder(3) Then
                        m.Stop3.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                        SubmitOrder m.Stop3.Order
                    
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal m.Stop3.Order.OrderID, CurrentTime, m.Stop3.Name & " generated by auto exit at " & m.Stop3.Order.StopPriceString & " because it doesn't exist yet"
                        End If
                    Else
                        DumpDebug "ModifyOrders - Did not create third stop loss order because condition was False"
                    End If
                End If
            
                If m.StopRest.ModifyTheOrder = True Then
                    If CreateStopLossOrder(0) Then
                        m.StopRest.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                        SubmitOrder m.StopRest.Order
                    
                        If g.Broker.AutoJournalAutomated Then
                            g.JournalBridge.CreateOrderJournal m.StopRest.Order.OrderID, CurrentTime, m.StopRest.Name & " generated by auto exit at " & m.StopRest.Order.StopPriceString & " because it doesn't exist yet"
                        End If
                    Else
                        DumpDebug "ModifyOrders - Did not create rest stop loss order because condition was False"
                    End If
                End If
            End If
        End If
    
        If UseProfitTarget Then
            If ProfitTargetConditionTrue Then
                If m.ExitStrategy.UseTarget1 Then
                    If m.Target1.ModifyTheOrder(m.Stop1.OrderID) = True Then
                        If CreateTargetOrder(1) Then
                            m.Target1.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                            SubmitOrder m.Target1.Order
                        
                            If g.Broker.AutoJournalAutomated Then
                                g.JournalBridge.CreateOrderJournal m.Target1.Order.OrderID, CurrentTime, m.Target1.Name & " generated by auto exit at " & m.Target1.Order.LimitPriceString & " because it doesn't exist yet"
                            End If
                        Else
                            DumpDebug "ModifyOrders - Did not create first profit target order because condition was False"
                        End If
                    End If
                End If
            
                If m.ExitStrategy.UseTarget2 Then
                    If m.Target2.ModifyTheOrder(m.Stop2.OrderID) = True Then
                        If CreateTargetOrder(2) Then
                            m.Target2.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                            SubmitOrder m.Target2.Order
                        
                            If g.Broker.AutoJournalAutomated Then
                                g.JournalBridge.CreateOrderJournal m.Target2.Order.OrderID, CurrentTime, m.Target2.Name & " generated by auto exit at " & m.Target2.Order.LimitPriceString & " because it doesn't exist yet"
                            End If
                        Else
                            DumpDebug "ModifyOrders - Did not create second profit target order because condition was False"
                        End If
                    End If
                End If
            
                If m.ExitStrategy.UseTarget3 Then
                    If m.Target3.ModifyTheOrder(m.Stop3.OrderID) = True Then
                        If CreateTargetOrder(3) Then
                            m.Target3.Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(AccountID))
                            SubmitOrder m.Target3.Order
                        
                            If g.Broker.AutoJournalAutomated Then
                                g.JournalBridge.CreateOrderJournal m.Target3.Order.OrderID, CurrentTime, m.Target3.Name & " generated by auto exit at " & m.Target3.Order.LimitPriceString & " because it doesn't exist yet"
                            End If
                        Else
                            DumpDebug "ModifyOrders - Did not create third profit target order because condition was False"
                        End If
                    End If
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.ModifyOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsExternalOrder
'' Description: Is the given order not one associated with this exit?
'' Inputs:      Order ID
'' Returns:     True if external, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsExternalOrder(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    
    If lOrderID = m.Stop1.OrderID Then
        bReturn = False
    ElseIf lOrderID = m.Stop2.OrderID Then
        bReturn = False
    ElseIf lOrderID = m.Stop3.OrderID Then
        bReturn = False
    ElseIf lOrderID = m.StopRest.OrderID Then
        bReturn = False
    ElseIf lOrderID = m.Target1.OrderID Then
        bReturn = False
    ElseIf lOrderID = m.Target2.OrderID Then
        bReturn = False
    ElseIf lOrderID = m.Target3.OrderID Then
        bReturn = False
    Else
        bReturn = True
    End If
    
    IsExternalOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.IsExternalOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTargetOrder
'' Description: Is the given order one of the targets for this exit?
'' Inputs:      Order ID
'' Returns:     True if target, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsTargetOrder(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    
    If (lOrderID = m.Target1.OrderID) Or ((m.Target1.OrderID = -2) And (lOrderID = m.Target1.PrevOrderID)) Then
        bReturn = True
    ElseIf (lOrderID = m.Target2.OrderID) Or ((m.Target2.OrderID = -2) And (lOrderID = m.Target2.PrevOrderID)) Then
        bReturn = True
    ElseIf (lOrderID = m.Target3.OrderID) Or ((m.Target3.OrderID = -2) And (lOrderID = m.Target3.PrevOrderID)) Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    IsTargetOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.IsTargetOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsStopOrder
'' Description: Is the given order the stop order for this exit?
'' Inputs:      Order ID
'' Returns:     True if stop, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsStopOrder(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    
    If (lOrderID = m.Stop1.OrderID) Or ((m.Stop1.OrderID = -2) And (lOrderID = m.Stop1.PrevOrderID)) Then
        bReturn = True
    ElseIf (lOrderID = m.Stop2.OrderID) Or ((m.Stop2.OrderID = -2) And (lOrderID = m.Stop2.PrevOrderID)) Then
        bReturn = True
    ElseIf (lOrderID = m.Stop3.OrderID) Or ((m.Stop3.OrderID = -2) And (lOrderID = m.Stop3.PrevOrderID)) Then
        bReturn = True
    ElseIf (lOrderID = m.StopRest.OrderID) Or ((m.StopRest.OrderID = -2) And (lOrderID = m.StopRest.PrevOrderID)) Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    IsStopOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.IsStopOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetOrderIDs
'' Description: Reset the order IDs back to zero
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetOrderIDs()
On Error GoTo ErrSection:

    m.Stop1.ResetInfo
    m.Stop2.ResetInfo
    m.Stop3.ResetInfo
    m.StopRest.ResetInfo
    
    m.Target1.ResetInfo
    m.Target2.ResetInfo
    m.Target3.ResetInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.ResetOrderIDs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrder
'' Description: Attempt to amend the given order
'' Inputs:      Order, Show Can Modify Message?
'' Returns:     True if submitted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ModifyOrder(Order As cPtOrder, Optional ByVal bShowCanModifyMsg As Boolean = True) As Boolean
On Error GoTo ErrSection:

    DumpDebug "Modifying Order: " & Order.OrderText & " (" & Order.GenesisOrderID & ", " & Order.BrokerID & ")"
    ModifyOrder = mTradeTracker.EditOrder(Order, , False, eGDEditOrderReturn_Submit, bShowCanModifyMsg)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.ModifyOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HighestHighSinceEntry
'' Description: Determine the highest high since the initial entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HighestHighSinceEntry()
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim lPos As Long                    ' Position in bars to start
    
    lPos = m.Bars.FindDateTime(m.dTimeOfHighestHigh)
    dReturn = gdMaxValue(m.Bars.ArrayHandle(eBARS_Close), lPos, m.Bars.Size - 1)
    m.dTimeOfHighestHigh = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
    
    ' 04/20/2009 DAJ: If the return is invalid (<=0) then set the highest high to
    ' the entry price as long as it is valid...
    If (dReturn <= 0) And (m.dEntryPrice > 0) Then
        DumpDebug "Highest high set to entry price (" & Str(m.dEntryPrice) & ") because return was <= 0"
        dReturn = m.dEntryPrice
    End If
    
    If (dReturn > m.dHighestHigh) Or (m.dHighestHigh = kNullData) Then
        DumpDebug "Highest high moved from " & Str(m.dHighestHigh) & " to " & Str(dReturn)
        m.dHighestHigh = dReturn
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.HighestHighSinceEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LowestLowSinceEntry
'' Description: Determine the lowest low since the initial entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LowestLowSinceEntry()
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value from the function
    Dim lPos As Long                    ' Position in bars to start
    Dim lIndex As Long                  ' Index into a for loop
    Dim fh As Integer                   ' File handle to a log file
    Static bDumped As Boolean           ' Have ticks been dumped?
    
    lPos = m.Bars.FindDateTime(m.dTimeOfLowestLow)
    dReturn = gdMinValue(m.Bars.ArrayHandle(eBARS_Close), lPos, m.Bars.Size - 1)
    m.dTimeOfLowestLow = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
    
    ' 04/20/2009 DAJ: We have one customer who is running into dReturn being
    ' zero which should be invalid.  We think that the only way this could happen
    ' is if there is a tick with a price of zero, so do a one time dump here...
    If dReturn = 0 Then
        If bDumped = False Then
            bDumped = True
            DumpDebug "Ticks being dumped because return is zero"
            
            fh = FreeFile
            Open AddSlash(App.Path) & "XOSTICKS.LOG" For Output As #fh
            
            For lIndex = 0 To m.Bars.Size - 1
                If lIndex = lPos Then
                    Print #fh, "lPos --> ";
                End If
                Print #fh, DateFormat(m.Bars(eBARS_Close, lIndex), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " " & Str(m.Bars(eBARS_Close, lIndex))
            Next lIndex
            
            Close #fh
        End If
    End If
    
    ' 04/20/2009 DAJ: If the return is invalid (<=0) then set the lowest low to
    ' the entry price as long as it is valid...
    If (dReturn <= 0) And (m.dEntryPrice > 0) Then
        DumpDebug "Lowest low set to entry price (" & Str(m.dEntryPrice) & ") because return was <= 0"
        dReturn = m.dEntryPrice
    End If
    
    If (dReturn < m.dLowestLow) Or (m.dLowestLow = kNullData) Then
        DumpDebug "Lowest low moved from " & Str(m.dLowestLow) & " to " & Str(dReturn)
        m.dLowestLow = dReturn
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.LowestLowSinceEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TrailOrder
'' Description: Trail the order using the given trail amount
'' Inputs:      Trail Amount
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TrailOrder(ByVal dTrailTicks As Double)
On Error GoTo ErrSection:

    Dim dTrailPrice As Double           ' Trailing stop price
    
    If (gdTickCount >= m.dLastTrailTime + kTrailingStopWait) Then
        If m.lCurrentPosition < 0 Then
            If m.dLowestLow <> kNullData Then
                dTrailPrice = m.dLowestLow + (dTrailTicks * TickMove)
                If (dTrailPrice < m.dTrailPrice) Or (m.dTrailPrice = 0) Then
                    m.dTrailPrice = dTrailPrice
                    If m.ExitStrategy.UseWithLimit Then
                        m.dTrailWithLimitPrice = m.dTrailPrice + (m.ExitStrategy.WithLimitTicks * TickMove)
                    Else
                        m.dTrailWithLimitPrice = 0#
                    End If
                End If
            End If
        Else
            If m.dHighestHigh <> kNullData Then
                dTrailPrice = m.dHighestHigh - (dTrailTicks * TickMove)
                If (dTrailPrice > m.dTrailPrice) Or (m.dTrailPrice = 0) Then
                    m.dTrailPrice = dTrailPrice
                    If m.ExitStrategy.UseWithLimit Then
                        m.dTrailWithLimitPrice = m.dTrailPrice - (m.ExitStrategy.WithLimitTicks * TickMove)
                    Else
                        m.dTrailWithLimitPrice = 0#
                    End If
                End If
            End If
        End If
        
        m.dLastTrailTime = gdTickCount
    End If
    
    If (m.lCurrentPosition <> 0) And (m.dTrailPrice <> 0) Then
        m.Stop1.TrailOrder m.dTrailPrice, m.dTrailWithLimitPrice
        m.Stop2.TrailOrder m.dTrailPrice, m.dTrailWithLimitPrice
        m.Stop3.TrailOrder m.dTrailPrice, m.dTrailWithLimitPrice
        m.StopRest.TrailOrder m.dTrailPrice, m.dTrailWithLimitPrice
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.TrailOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanActivate
'' Description: Determine whether we can activate the active exit (must have
''              streaming on and connected to broker)
'' Inputs:      None
'' Returns:     True if can activate, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanActivate() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim nAcctType As eTT_AccountType    ' Account type for the account

    bReturn = False
    If (g.Broker.ConnectionStatusForAccount(AccountID) = eGDConnectionStatus_Connected) Then
        If g.RealTime.RtDataAvailable(SymbolOrSymbolID, ePRD_Days) Then
            m.bWaitingOnData = False
            bReturn = Not g.Broker.PositionVerify(m.nAcctType)
        Else
            m.bWaitingOnData = True
        End If
    End If
    
    CanActivate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.CanActivate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveLastKnown
'' Description: Save off the last known information for restoration purposes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveLastKnown(ByVal bReset As Boolean)
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Information to save off
    
    If (m.lCurrentPosition <> 0&) And (m.bActive = True) And (bReset = False) Then
        Set astrInfo = New cGdArray
        astrInfo.Create eGDARRAY_Strings
        
        astrInfo(0) = m.strExitStrategyFile
        astrInfo(1) = Str(m.lCurrentPosition)
        astrInfo(2) = Str(m.dEntryPrice)
        
        If (Not m.Stop1.Order Is Nothing) And (m.Stop1.OrderID <> -2&) Then
            astrInfo(3) = Str(m.Stop1.Order.Quantity) & ";" & Str(m.Stop1.Order.LimitPrice) & ";" & Str(m.Stop1.Order.StopPrice)
        End If
        
        If (Not m.Target1.Order Is Nothing) And (m.Target1.OrderID <> -2&) Then
            astrInfo(4) = Str(m.Target1.Order.Quantity) & ";" & Str(m.Target1.Order.LimitPrice)
        End If
    
        If (Not m.Target2.Order Is Nothing) And (m.Target2.OrderID <> -2&) Then
            astrInfo(5) = Str(m.Target2.Order.Quantity) & ";" & Str(m.Target2.Order.LimitPrice)
        End If
    
        If (Not m.Target3.Order Is Nothing) And (m.Target3.OrderID <> -2&) Then
            astrInfo(6) = Str(m.Target3.Order.Quantity) & ";" & Str(m.Target3.Order.LimitPrice)
        End If
        
        If (Not m.Stop2.Order Is Nothing) And (m.Stop2.OrderID <> -2&) Then
            astrInfo(7) = Str(m.Stop2.Order.Quantity) & ";" & Str(m.Stop2.Order.LimitPrice) & ";" & Str(m.Stop2.Order.StopPrice)
        End If
        
        If (Not m.Stop3.Order Is Nothing) And (m.Stop3.OrderID <> -2&) Then
            astrInfo(8) = Str(m.Stop3.Order.Quantity) & ";" & Str(m.Stop3.Order.LimitPrice) & ";" & Str(m.Stop3.Order.StopPrice)
        End If
        
        If (Not m.StopRest.Order Is Nothing) And (m.StopRest.OrderID <> -2&) Then
            astrInfo(9) = Str(m.StopRest.Order.Quantity) & ";" & Str(m.StopRest.Order.LimitPrice) & ";" & Str(m.StopRest.Order.StopPrice)
        End If
        
        SetIniFileProperty Str(AccountID) & "|" & Symbol, astrInfo.JoinFields("|"), "AutoExits", g.strIniFile
        DumpDebug "SaveLastKnown - '" & astrInfo.JoinFields("|") & "'"
    Else
        SetIniFileProperty Str(AccountID) & "|" & Symbol, "", "AutoExits", g.strIniFile
        DumpDebug "SaveLastKnown - ''"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SaveLastKnown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentInfoToLastKnown
'' Description: Convert the current info string to a last known string
'' Inputs:      Current Info
'' Returns:     Last Known
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CurrentInfoToLastKnown(ByVal strCurrentInfo As String) As String
On Error GoTo ErrSection:

    Dim astrCurrentInfo As cGdArray     ' Current Information split out into an array
    Dim astrLastKnown As cGdArray       ' Last known information split out into an array
    Dim dLastPrice As Double            ' Last known price
    Dim strReturn As String             ' Return value from function
    
    If Len(strCurrentInfo) > 0 Then
        Set astrCurrentInfo = New cGdArray
        astrCurrentInfo.SplitFields strCurrentInfo, vbTab
    
        Set astrLastKnown = New cGdArray
        astrLastKnown.Create eGDARRAY_Strings
        
        astrLastKnown(0) = m.strExitStrategyFile
        astrLastKnown(1) = Str(m.lCurrentPosition)
        astrLastKnown(2) = Str(m.dEntryPrice)
        
        dLastPrice = g.RealTime.LastKnownPrice(SymbolOrSymbolID)
        If m.lCurrentPosition > 0 Then
            astrLastKnown(3) = astrCurrentInfo(0) & ";" & Str(dLastPrice - Val(astrCurrentInfo(2))) & ";" & Str(dLastPrice - Val(astrCurrentInfo(1)))
            astrLastKnown(4) = astrCurrentInfo(3) & ";" & Str(dLastPrice + Val(astrCurrentInfo(4)))
            astrLastKnown(5) = astrCurrentInfo(5) & ";" & Str(dLastPrice + Val(astrCurrentInfo(6)))
            astrLastKnown(6) = astrCurrentInfo(7) & ";" & Str(dLastPrice + Val(astrCurrentInfo(8)))
            astrLastKnown(7) = astrCurrentInfo(9) & ";" & Str(dLastPrice - Val(astrCurrentInfo(11))) & ";" & Str(dLastPrice - Val(astrCurrentInfo(10)))
            astrLastKnown(8) = astrCurrentInfo(12) & ";" & Str(dLastPrice - Val(astrCurrentInfo(14))) & ";" & Str(dLastPrice - Val(astrCurrentInfo(13)))
            astrLastKnown(9) = astrCurrentInfo(15) & ";" & Str(dLastPrice - Val(astrCurrentInfo(17))) & ";" & Str(dLastPrice - Val(astrCurrentInfo(16)))
        ElseIf m.lCurrentPosition < 0 Then
            astrLastKnown(3) = astrCurrentInfo(0) & ";" & Str(dLastPrice + Val(astrCurrentInfo(2))) & ";" & Str(dLastPrice + Val(astrCurrentInfo(1)))
            astrLastKnown(4) = astrCurrentInfo(3) & ";" & Str(dLastPrice - Val(astrCurrentInfo(4)))
            astrLastKnown(5) = astrCurrentInfo(5) & ";" & Str(dLastPrice - Val(astrCurrentInfo(6)))
            astrLastKnown(6) = astrCurrentInfo(7) & ";" & Str(dLastPrice - Val(astrCurrentInfo(8)))
            astrLastKnown(7) = astrCurrentInfo(9) & ";" & Str(dLastPrice + Val(astrCurrentInfo(11))) & ";" & Str(dLastPrice + Val(astrCurrentInfo(10)))
            astrLastKnown(8) = astrCurrentInfo(12) & ";" & Str(dLastPrice + Val(astrCurrentInfo(14))) & ";" & Str(dLastPrice + Val(astrCurrentInfo(13)))
            astrLastKnown(9) = astrCurrentInfo(15) & ";" & Str(dLastPrice + Val(astrCurrentInfo(17))) & ";" & Str(dLastPrice + Val(astrCurrentInfo(16)))
        End If
        
        strReturn = astrLastKnown.JoinFields("|")
    Else
        strReturn = ""
    End If
    
    CurrentInfoToLastKnown = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.CurrentInfoToLastKnown"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemainingTargetQuantity
'' Description: Calculate the sum of the remaining quantities on the targets
'' Inputs:      None
'' Returns:     Remaining Quantity of the Targets
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemainingTargetQuantity() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 0&
    If m.Target1.OrderID > 0 Then
        If Not m.Target1.Order Is Nothing Then
            lReturn = lReturn + m.Target1.Order.RemainingQuantity
        End If
    End If
    
    If m.Target2.OrderID > 0 Then
        If Not m.Target2.Order Is Nothing Then
            lReturn = lReturn + m.Target2.Order.RemainingQuantity
        End If
    End If
    
    If m.Target3.OrderID > 0 Then
        If Not m.Target3.Order Is Nothing Then
            lReturn = lReturn + m.Target3.Order.RemainingQuantity
        End If
    End If
    
    RemainingTargetQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.RemainingTargetQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemainingStopQuantity
'' Description: Calculate the sum of the remaining quantities on the stops
'' Inputs:      None
'' Returns:     Remaining Quantity of the Stops
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemainingStopQuantity() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 0&
    If m.Stop1.OrderID > 0 Then
        If Not m.Stop1.Order Is Nothing Then
            lReturn = lReturn + m.Stop1.Order.RemainingQuantity
        End If
    End If
    
    If m.Stop2.OrderID > 0 Then
        If Not m.Stop2.Order Is Nothing Then
            lReturn = lReturn + m.Stop2.Order.RemainingQuantity
        End If
    End If
    
    If m.Stop3.OrderID > 0 Then
        If Not m.Stop3.Order Is Nothing Then
            lReturn = lReturn + m.Stop3.Order.RemainingQuantity
        End If
    End If
    
    If m.StopRest.OrderID > 0 Then
        If Not m.StopRest.Order Is Nothing Then
            lReturn = lReturn + m.StopRest.Order.RemainingQuantity
        End If
    End If
    
    RemainingStopQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.RemainingStopQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flattening
'' Description: Are we currently being flattened?
'' Inputs:      None
'' Returns:     True if flattening, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Flattening() As Boolean
On Error GoTo ErrSection:

    Flattening = g.FlattenQueue.IsGettingFlattened(g.Broker.AccountNumberForID(AccountID), Symbol, 0&, eGDFlattenQueueOperation_Flatten)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.Flattening"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancellingAll
'' Description: Are we currently cancelling all orders?
'' Inputs:      None
'' Returns:     True if cancelling all orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancellingAll() As Boolean
On Error GoTo ErrSection:

    CancellingAll = g.FlattenQueue.IsGettingFlattened(g.Broker.AccountNumberForID(AccountID), Symbol, 0&, eGDFlattenQueueOperation_CancelAll)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.CancellingAll"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OtherSideFilled
'' Description: Did the opposite order fill?
'' Inputs:      Order
'' Returns:     True if other order filled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OtherSideFilled(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    Select Case Order.OrderID
        Case m.Stop1.Order.OrderID
            If Not m.Target1 Is Nothing Then
                DumpDebug vbTab & vbTab & "OtherSideFilled(S1) = " & Str(m.Target1.OrderID)
                bReturn = (m.Target1.OrderID = -2)
            End If
        Case m.Stop2.Order.OrderID
            If Not m.Target2 Is Nothing Then
                DumpDebug vbTab & vbTab & "OtherSideFilled(S2) = " & Str(m.Target2.OrderID)
                bReturn = (m.Target2.OrderID = -2)
            End If
        Case m.Stop3.Order.OrderID
            If Not m.Target3 Is Nothing Then
                DumpDebug vbTab & vbTab & "OtherSideFilled(S3) = " & Str(m.Target3.OrderID)
                bReturn = (m.Target3.OrderID = -2)
            End If
        Case m.Target1.Order.OrderID
            If Not m.Stop1 Is Nothing Then
                DumpDebug vbTab & vbTab & "OtherSideFilled(T1) = " & Str(m.Stop1.OrderID)
                bReturn = (m.Stop1.OrderID = -2)
            End If
        Case m.Target2.Order.OrderID
            If Not m.Stop2 Is Nothing Then
                DumpDebug vbTab & vbTab & "OtherSideFilled(T2) = " & Str(m.Stop2.OrderID)
                bReturn = (m.Stop2.OrderID = -2)
            End If
        Case m.Target3.Order.OrderID
            If Not m.Stop3 Is Nothing Then
                DumpDebug vbTab & vbTab & "OtherSideFilled(T3) = " & Str(m.Stop3.OrderID)
                bReturn = (m.Stop3.OrderID = -2)
            End If
    End Select
    
    OtherSideFilled = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.OtherSideFilled"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnsolicitedCancel
'' Description: Determine if this is an unsolicited cancel that is not related
''              to the other side of an OCO being filled
'' Inputs:      Order
'' Returns:     True if Unsolicited Cancel, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UnsolicitedCancel(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (IsOpenOrder(Order.Status) = False) And ((Order.Status <> eTT_OrderStatus_Filled) And (Order.Status <> eTT_OrderStatus_Amended)) Then
        DumpDebug vbTab & vbTab & "UnsolicitedCancel('" & Order.OrderText & "'): " & OrderStatus(Order.Status)
        'If (Order.BrokerCancelOrderID = 0) Or (OtherSideFilled(Order) = False) Then
        '    DumpDebug vbTab & vbTab & "UnsolicitedCancel('" & Order.OrderText & "'): BOCO = " & Str(Order.BrokerCancelOrderID)
        '    bReturn = True
        'End If
        bReturn = Not OtherSideFilled(Order)
    End If
    
    UnsolicitedCancel = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.UnsolicitedCancel"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OneStopHasValidID
'' Description: Does at least one of the stop orders have a valid ID?
'' Inputs:      None
'' Returns:     True if valid ID, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OneStopHasValidID() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If m.Stop1.OrderID > 0 Then
        bReturn = True
    ElseIf m.Stop2.OrderID > 0 Then
        bReturn = True
    ElseIf m.Stop3.OrderID > 0 Then
        bReturn = True
    ElseIf m.StopRest.OrderID > 0 Then
        bReturn = True
    End If
    
    OneStopHasValidID = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.OneStopHasValidID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearBreakEvenMovedFlag
'' Description: Clear the break-even moved flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearBreakEvenMovedFlag()
On Error GoTo ErrSection:

    If Not m.Stop1 Is Nothing Then
        m.Stop1.BreakEvenMoved = False
    End If
    If Not m.Stop2 Is Nothing Then
        m.Stop2.BreakEvenMoved = False
    End If
    If Not m.Stop3 Is Nothing Then
        m.Stop3.BreakEvenMoved = False
    End If
    If Not m.StopRest Is Nothing Then
        m.StopRest.BreakEvenMoved = False
    End If

    m.bBreakEvenMoved = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.ClearBreakEvenMovedFlag"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateTsOrders
'' Description: Activate the appropriate Trade Sense orders
'' Inputs:      Account Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ActivateTsOrders(ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:

    Dim strEvaluate As String           ' Evaluation string

    If m.lCurrentPosition > 0 Then
        If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
            If Len(m.ExitStrategy.TsProfitLong.ConditionCoded) > 0 Then
                m.ExitStrategy.TsProfitLong.ActivateWithAcctPos m.strSymbol, AcctPos
                
                If m.ExitStrategy.TsProfitLong.IntraBar Then
                    strEvaluate = "Evaluate Intra-Bar"
                Else
                    strEvaluate = "Evaluate Every New '" & GetPeriodStr(m.ExitStrategy.TsProfitLong.BarPeriod) & "' Bar"
                End If
                DumpDebug vbTab & "Long Profit Target (" & m.ExitStrategy.TsProfitLong.EnglishPreview(False) & ") Activated -- " & strEvaluate
            Else
                DumpDebug vbTab & "Not Activating Long Profit Target because none specified"
            End If
        End If
        If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
            If Len(m.ExitStrategy.TsStopLong.ConditionCoded) > 0 Then
                m.ExitStrategy.TsStopLong.ActivateWithAcctPos m.strSymbol, AcctPos
                
                If m.ExitStrategy.TsStopLong.IntraBar Then
                    strEvaluate = "Evaluate Intra-Bar"
                Else
                    strEvaluate = "Evaluate Every New '" & GetPeriodStr(m.ExitStrategy.TsStopLong.BarPeriod) & "' Bar"
                End If
                DumpDebug vbTab & "Long Stop Loss (" & m.ExitStrategy.TsStopLong.EnglishPreview(False) & ") Activated -- " & strEvaluate
            Else
                DumpDebug vbTab & "Not Activating Long Stop Loss because none specified"
            End If
        End If
    ElseIf m.lCurrentPosition < 0 Then
        If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
            If Len(m.ExitStrategy.TsProfitShort.ConditionCoded) > 0 Then
                m.ExitStrategy.TsProfitShort.ActivateWithAcctPos m.strSymbol, AcctPos
                                
                If m.ExitStrategy.TsProfitShort.IntraBar Then
                    strEvaluate = "Evaluate Intra-Bar"
                Else
                    strEvaluate = "Evaluate Every New '" & GetPeriodStr(m.ExitStrategy.TsProfitShort.BarPeriod) & "' Bar"
                End If
                DumpDebug vbTab & "Short Profit Target (" & m.ExitStrategy.TsProfitShort.EnglishPreview(False) & ") Activated -- " & strEvaluate
            Else
                DumpDebug vbTab & "Not Activating Short Profit Target because none specified"
            End If
        End If
        If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
            If Len(m.ExitStrategy.TsStopShort.ConditionCoded) > 0 Then
                m.ExitStrategy.TsStopShort.ActivateWithAcctPos m.strSymbol, AcctPos
                                
                If m.ExitStrategy.TsStopShort.IntraBar Then
                    strEvaluate = "Evaluate Intra-Bar"
                Else
                    strEvaluate = "Evaluate Every New '" & GetPeriodStr(m.ExitStrategy.TsStopShort.BarPeriod) & "' Bar"
                End If
                DumpDebug vbTab & "Short Stop Loss (" & m.ExitStrategy.TsStopShort.EnglishPreview(False) & ") Activated -- " & strEvaluate
            Else
                DumpDebug vbTab & "Not Activating Short Stop Loss because none specified"
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.ActivateTsOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UseStopLoss
'' Description: Do we want to use a stop loss?
'' Inputs:      None
'' Returns:     True if use stop loss, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UseStopLoss() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If m.ExitStrategy.StopLossType = eGDStopLossType_None Then
        bReturn = False
    ElseIf m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
        If m.lCurrentPosition < 0 Then
            bReturn = (Len(m.ExitStrategy.TsStopShort.ConditionCoded) > 0)
        ElseIf m.lCurrentPosition > 0 Then
            bReturn = (Len(m.ExitStrategy.TsStopLong.ConditionCoded) > 0)
        Else
            bReturn = False
        End If
    Else
        bReturn = True
    End If
    
    UseStopLoss = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.UseStopLoss"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UseProfitTarget
'' Description: Do we want to use a profit target?
'' Inputs:      None
'' Returns:     True if use profit target, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UseProfitTarget() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
        If m.lCurrentPosition < 0 Then
            bReturn = (Len(m.ExitStrategy.TsProfitShort.ConditionCoded) > 0)
        ElseIf m.lCurrentPosition > 0 Then
            bReturn = (Len(m.ExitStrategy.TsProfitLong.ConditionCoded) > 0)
        Else
            bReturn = False
        End If
    Else
        bReturn = True
    End If
    
    UseProfitTarget = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.UseProfitTarget"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StopLossConditionTrue
'' Description: Is the stop loss condition true?
'' Inputs:      None
'' Returns:     True if condition true, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StopLossConditionTrue() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If m.ExitStrategy.StopLossType = eGDStopLossType_TradeSense Then
        If m.lCurrentPosition > 0 Then
            bReturn = m.ExitStrategy.TsStopLong.ConditionTrue
        ElseIf m.lCurrentPosition < 0 Then
            bReturn = m.ExitStrategy.TsStopShort.ConditionTrue
        End If
    End If
    
    StopLossConditionTrue = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.StopLossConditionTrue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ProfitTargetConditionTrue
'' Description: Is the profit target condition true?
'' Inputs:      None
'' Returns:     True if condition true, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ProfitTargetConditionTrue() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If m.ExitStrategy.ProfitTargetType = eGDProfitTargetType_TradeSense Then
        If m.lCurrentPosition > 0 Then
            bReturn = m.ExitStrategy.TsProfitLong.ConditionTrue
        ElseIf m.lCurrentPosition < 0 Then
            bReturn = m.ExitStrategy.TsProfitShort.ConditionTrue
        End If
    End If
    
    ProfitTargetConditionTrue = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.ProfitTargetConditionTrue"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountIds
'' Description: Set the account IDs on the active exit orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAccountIds()
On Error GoTo ErrSection:

    m.Target1.AccountID = AccountID
    m.Target2.AccountID = AccountID
    m.Target3.AccountID = AccountID
    
    m.Stop1.AccountID = AccountID
    m.Stop2.AccountID = AccountID
    m.Stop3.AccountID = AccountID
    m.StopRest.AccountID = AccountID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetAccountIds"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetSymbols
'' Description: Set the symbols on the active exit orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetSymbols()
On Error GoTo ErrSection:

    m.Target1.Symbol = Symbol
    m.Target2.Symbol = Symbol
    m.Target3.Symbol = Symbol
    
    m.Stop1.Symbol = Symbol
    m.Stop2.Symbol = Symbol
    m.Stop3.Symbol = Symbol
    m.StopRest.Symbol = Symbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.SetSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Come up with a string with all of the active exit information
'' Inputs:      None
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = AccountID
    strReturn = strReturn & "|" & Str(m.lSymbolID)
    strReturn = strReturn & "|" & Symbol
    strReturn = strReturn & "|" & m.strExitStrategyFile
    strReturn = strReturn & "|" & Str(m.bActive)
    strReturn = strReturn & "|" & Str(m.lCurrentPosition)
    strReturn = strReturn & "|" & Str(m.dEntryPrice)
    strReturn = strReturn & "|" & Str(m.Target1.OrderID)
    strReturn = strReturn & "|" & Str(m.Target2.OrderID)
    strReturn = strReturn & "|" & Str(m.Target3.OrderID)
    strReturn = strReturn & "|" & Str(m.Stop1.OrderID)
    strReturn = strReturn & "|" & Str(m.Stop2.OrderID)
    strReturn = strReturn & "|" & Str(m.Stop3.OrderID)
    strReturn = strReturn & "|" & Str(m.StopRest.OrderID)
    
    ToString = strReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cActiveExit.ToString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeSenseGroupPosition
'' Description: Determine the total position for TradeSense order groups for
''              the symbol and account for this auto exit
'' Inputs:      None
'' Returns:     Total Position of TradeSense order groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TradeSenseGroupPosition() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 0&
    If Not g.TsoGroups Is Nothing Then
        lReturn = g.TsoGroups.TotalPosition(m.lAccountID, SymbolOrSymbolID)
    End If
    
    TradeSenseGroupPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.TradeSenseGroupPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanUpdateSession
'' Description: Determine if the session can be updated
'' Inputs:      None
'' Returns:     True if session can be updated, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanUpdateSession() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If IsForex(m.strSymbol) Then
        If InStr(m.strSymbol, "@") <> 0 Then
            bReturn = mDataNav.ValidForexTradingTime(m.strSymbol, CurrentTime("NY", m.strSymbol, True))
        End If
    End If
    
    CanUpdateSession = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.CanUpdateSession"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetVariables
'' Description: Reset the variables
'' Inputs:      Entry Price, Entry Time
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetVariables(ByVal dEntryPrice As Double, ByVal dEntryTime As Double)
On Error GoTo ErrSection:

    m.dEntryPrice = NearestTick(dEntryPrice)
    m.dEntryTime = dEntryTime
    DumpDebug "Entry Price: " & Str(m.dEntryPrice) & "; Entry Time: " & DateFormat(m.dEntryTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
    m.dTrailPrice = 0#
    m.dTrailWithLimitPrice = 0#
    m.dLastTrailTime = 0#
    m.dHighestHigh = m.dEntryPrice
    m.dLowestLow = m.dEntryPrice
                    
    ' 09/23/2014 DAJ: Call UpdateBars here to clear the tick buffer before we reset
    ' the bars so that when UpdateBars gets called after this, we only get new ticks...
    g.RealTime.UpdateBars m.Bars
                    
    m.Bars.Size = 1&
    m.Bars(eBARS_Close, 0) = m.dEntryPrice
    m.Bars(eBARS_DateTime, 0) = RoundToMinute(m.dEntryTime)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.ResetVariables"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RoundQuantity
'' Description: Round the given quantity to the next lot size up
'' Inputs:      Quantity, Allow Zero?
'' Returns:     Rounded quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RoundQuantity(ByVal lQuantity As Long, ByVal bAllowZero As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lMinLotSize As Long             ' Minimum lot size for the given symbol/account
    Dim lMinQuantity As Long            ' Minimum quantity for the given symbol/account
    
    lMinLotSize = g.Broker.MinimumLotSize(AccountID, SymbolOrSymbolID)
    lMinQuantity = g.Broker.MinimumOrderQuantity(AccountID, SymbolOrSymbolID)
    
    lReturn = RoundToMinMove(lQuantity, lMinLotSize)
    If (lReturn <> 0&) Or (bAllowZero = False) Then
        If (lReturn < lMinQuantity) Then
            lReturn = lMinQuantity
        ElseIf (lReturn < lQuantity) Then
            lReturn = lReturn + lMinLotSize
        End If
    End If
    
    RoundQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.RoundQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrdersInOppositeDirection
'' Description: Cancel manual working orders in the opposite direction
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelOrdersInOppositeDirection()
On Error GoTo ErrSection:

    Dim nDirection As eGDFilterDirection ' Direction to cancel orders

    If m.ExitStrategy.CancelOpposite Then
        If m.lCurrentPosition <> 0 Then
            DumpDebug "Cancelling all working manual orders in opposite direction"
            
            If m.lCurrentPosition > 0 Then
                nDirection = eGDFilterDirection_Shorts
            ElseIf m.lCurrentPosition < 0 Then
                nDirection = eGDFilterDirection_Longs
            End If
            
            g.Broker.CancelWorkingOrders AccountID, SymbolOrSymbolID, 0&, , nDirection
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.CancelOrdersInOppositeDirection"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NearestTick
'' Description: Round the given price to the nearest tick
'' Inputs:      Price
'' Returns:     Rounded price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NearestTick(ByVal dPrice As Double) As Double
On Error GoTo ErrSection:

    NearestTick = m.Bars.RoundToPrice(dPrice)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveExit.NearestTick"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send the given debug text to the appropriate place(s)
'' Inputs:      Debug Text
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strText As String)
On Error GoTo ErrSection:

    Dim strToDump As String             ' String to dump to the log
    
    strToDump = "XOS (" & g.Broker.AccountNameForID(AccountID) & ", " & Symbol & "): " & strText

    g.Broker.BrokerDebug m.nAcctType, strToDump
    frmTest2.AddList strToDump
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize members of the class when the object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Bars = New cGdBars
    Set m.ExitStrategy = New cExitStrategy
    
    Set m.Target1 = New cActiveExitOrder
    m.Target1.Init "First Profit Target", "T1"
    
    Set m.Target2 = New cActiveExitOrder
    m.Target2.Init "Second Profit Target", "T2"
    
    Set m.Target3 = New cActiveExitOrder
    m.Target3.Init "Third Profit Target", "T3"
    
    Set m.Stop1 = New cActiveExitOrder
    m.Stop1.Init "First Stop Loss", "S1"
    
    Set m.Stop2 = New cActiveExitOrder
    m.Stop2.Init "Second Stop Loss", "S2"
    
    Set m.Stop3 = New cActiveExitOrder
    m.Stop3.Init "Third Stop Loss", "S3"
    
    Set m.StopRest = New cActiveExitOrder
    m.StopRest.Init "Rest Stop Loss", "SR"
    
    m.bDelete = False
    m.bWaitingOnData = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate members of the class when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Bars = Nothing
    Set m.ExitStrategy = Nothing
    
    Set m.Target1 = Nothing
    Set m.Target2 = Nothing
    Set m.Target3 = Nothing
    Set m.Stop1 = Nothing
    Set m.Stop2 = Nothing
    Set m.Stop3 = Nothing
    Set m.StopRest = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveExit.Class_Terminate"
    
End Sub
