VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActiveTsOrderGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cActiveTsOrderGroup.cls
'' Description: Class that handles an active trade sense order group
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/17/2010   DAJ         Changed filenames to ID instead of name
'' 06/24/2010   DAJ         Expose some display stuff for orders that get displayed
'' 06/28/2010   DAJ         Added Use flag for turning off orders
'' 06/30/2010   DAJ         Fix for using TradeSense order groups through PFG
'' 07/05/2010   DAJ         Added CanActivate routine
'' 07/09/2010   DAJ         Fixes regarding resubmitting parked group
'' 07/20/2010   DAJ         Added capabilities for inputs
'' 07/22/2010   DAJ         Prepend $ on Forex for Daniel Code stuff
'' 08/05/2010   DAJ         Moved Daniel Code stuff to the confirmation form
'' 08/05/2010   DAJ         Only change Genesis Forex to PFG Forex if in a PFG account
'' 08/11/2010   DAJ         Only dump to log on new bar for "Check Intraday"
'' 08/12/2010   DAJ         Check for streaming availability
'' 08/12/2010   DAJ         Display order text in the action column
'' 08/12/2010   DAJ         Only log display stuff when it changes
'' 08/17/2010   DAJ         Fixed issue with loading group with inputs from a string
'' 08/19/2010   DAJ         Added capability for TradeSense order expire at end of session
'' 09/14/2010   DAJ         Fix for checking orders before bars updated
'' 09/23/2010   DAJ         Cancel old associated orders when activate (#5930)
'' 09/29/2010   DAJ         Added confirmation flag for TradeSense orders
'' 09/30/2010   DAJ         Don't submit exit order if not in a position
'' 09/30/2010   DAJ         Added log for data not being available yet
'' 10/04/2010   DAJ         Make sure to do a CheckOrder on exits that may not have been submitted yet (#5924)
'' 10/08/2010   DAJ         Added HasOrder functions, fixes for exits when no position
'' 10/14/2010   DAJ         Enhanced logging for expiration
'' 10/15/2010   DAJ         Don't update UI for Condition False unless it just went False
'' 10/20/2010   DAJ         Continuous Loop, Clean up if flat and all exits
'' 10/25/2010   DAJ         Additional logging for position calculation/fill loading
'' 10/27/2010   DAJ         Log reactivate value at activation
'' 11/02/2010   DAJ         Make sure orders get checked after data arrives
'' 11/05/2010   DAJ         If 100% Exit at Market submitted, don't submit other exits
'' 11/11/2010   DAJ         Upon reactivation, pass the fills to the orders a little later
'' 11/11/2010   DAJ         If a working order expires, just clear it and let the code do the rest
'' 11/15/2010   DAJ         Only OCO groups in same session, Don't count TSOG position for auto exits
'' 11/17/2010   DAJ         Don't cancel Market orders, log evaluation on activation
'' 11/30/2010   DAJ         Don't check orders upon reactivation
'' 12/08/2010   DAJ         Modify quantity of stop and reverse orders as position changes
'' 12/16/2010   DAJ         Display order text in "wrong side of market" error
'' 12/17/2010   DAJ         Tell order to update BarsSince functions on last fillcallback on load
'' 01/07/2011   DAJ         Changed activate call for TradeSense orders
'' 02/01/2011   DAJ         If order submission fails, close the tree
'' 02/03/2011   DAJ         Park group on Error/Reject, Display fill prices in UI
'' 02/03/2011   DAJ         Added the "Market Closed" and "Not connected to broker" display status'
'' 03/02/2011   DAJ         Notify global collection when Submitted changes
'' 03/17/2011   DAJ         Added flatten capability
'' 03/29/2011   DAJ         When parking all, queue up the dialog so we do not block any processes
'' 03/29/2011   DAJ         When an order is rejected or error, make sure to park it too
'' 04/04/2011   DAJ         Don't check order that just got a fill
'' 04/04/2011   DAJ         Don't allow an activate if live account and delayed stream
'' 04/28/2011   DAJ         Submit PFG Forex orders as GTC orders
'' 04/29/2011   DAJ         Log current market prices when submit or cancel caused by condition
'' 05/04/2011   DAJ         When receiving a fill, check OTO orders even if working in case position changed
'' 05/18/2011   DAJ         Added custom start/stop time for Market1
'' 05/23/2011   DAJ         When flattening a position, make sure to use TradeSymbol instead of symbol
'' 05/27/2011   DAJ         Set the IsAutomated flag for orders that get created
'' 06/01/2011   DAJ         Call g.Broker.SendOrder direct because we don't want orders merged anyway
'' 06/09/2011   DAJ         Implemented auto-breakout period and persisting trade symbol
'' 07/13/2011   DAJ         Only allow PFG FX group to run from 4pm central Sunday to 4pm central Friday
'' 07/13/2011   DAJ         Allow for true continuous looping if ProjectX
'' 08/23/2011   DAJ         Adjust quantity on submitted exits as position changes
'' 09/08/2011   DAJ         Moved auto breakout calculation down to cRunExpression
'' 09/28/2011   DAJ         Added automated journaling if user wants it
'' 10/07/2011   DAJ         Run expressions on a newly activated order because of OTO
'' 10/25/2011   DAJ         Add order to collection before submit
'' 12/21/2011   DAJ         Check order again after market reopens
'' 01/31/2012   DAJ         Handle unsolicted cancels in TradeSense order groups
'' 07/13/2012   DAJ         Added check for valid trading time for IB Forex
'' 08/17/2012   DAJ         Utilize generic valid forex trading time function
'' 08/17/2012   DAJ         Fixes for IB and Currenex forex
'' 09/26/2012   DAJ         On a new session, stop using old expired orders and resubmit new ones if applicable
'' 09/27/2012   DAJ         Fix for last change -- don't do anything if no order there
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 10/10/2012   DAJ         Added activation error if errors out, added check for valid quantity/lot size
'' 10/24/2012   DAJ         Always dump expression to log when condition false
'' 11/28/2012   DAJ         Submit Market if Stop on wrong side of market flag
'' 02/06/2013   DAJ         Don't submit orders between the close on Friday and Open on Sunday
'' 02/11/2013   DAJ         Added CancelLeftOverOrder functions
'' 03/18/2013   DAJ         Fix for ValidTradingTime function not allowing trading on Friday
'' 04/26/2013   DAJ         Don't mark an order expired if the session date of the order is the current session
'' 05/15/2013   DAJ         Submit Market order if amend on a stop is on wrong side ( if flag set )
'' 02/12/2014   DAJ         Persist the lot size
'' 06/26/2014   DAJ         When trade symbol changes and running on IB account, get contract info from IB
'' 06/26/2014   DAJ         Roll position on a TradeSense order group
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Fill Display
'' 01/29/2015   DAJ         Modified the call to GetContractInfo
'' 02/05/2015   DAJ         Re-run expressions on TradeSense orders after a fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGD_TsoStatus
    eGD_TsoStatus_Open = 0
    eGD_TsoStatus_Trigger
    eGD_TsoStatus_Closed
    eGD_TsoStatus_Parked
End Enum

Private Enum eGDRollPhases
    eGDRollPhase_None = 0
    eGDRollPhase_Parking = 1
    eGDRollPhase_Flattening = 2
    eGDRollPhase_Entering = 3
End Enum

Private Type mPrivate
    lID As Long                         ' Group identifier

    strSymbol As String                 ' Symbol to run the group on
    lSymbolID As Long                   ' Symbol to run the group on
    lAccountID As Long                  ' Account ID to run the group on
    lQuantity As Long                   ' Quantity for the orders
    bSubmitted As Boolean               ' Is this submitted or parked?
    bRtDataAvailable As Boolean         ' Real-time data available?
    lSessionActivated As Long           ' Session that the group was originally activated in
    dCustomStartTime As Double          ' Custom start time for the Market1 data
    dCustomStopTime As Double           ' Custom stop time for the Market1 data
    strTradeSymbol As String            ' Trade symbol for the initial activation
    strNewTradeSymbol As String         ' New trade symbol after a reload of data
    'strAutoBreakoutPeriod As String     ' Auto breakout period for the initial activation
    lLotSize As Long                    ' Lot size for adjusting positions
    
    tsOrderGrp As cTradeSenseOrderGroup ' Trade Sense Order Group to run
    Orders As cGdTree                   ' Actual orders associated with the group
    anStatus As cGdArray                ' Array of status for each order
    astrToDo As cGdArray                ' Array of commands to do when order goes working
    anLeaveOpen As cGdArray             ' Array of flags whether to put order back open after cancel
    astrStatus As cGdArray              ' Array of display status for each order
    alLevel As cGdArray                 ' Array of display levels for each order
    astrAction As cGdArray              ' Array of action strings for each order
    RollOrder As cPtOrder               ' Roll order
    
    nBroker As eTT_AccountType          ' Broker for the given account
    Fills As cGdTree                    ' Collection of fills received for this group
    astrFillIds As cGdArray             ' Array of Fill ID's
    alOrderIds As cGdArray              ' Array of Order ID's
    lPosition As Long                   ' Current position for this order group
    lPositionBeforeRoll As Long         ' Position before the roll
    Bars As cGdBars                     ' Bar properties for the main symbol
    
    dLastLogWait As Double              ' Tick count of the last time we logged we were waiting for data
    
    bAutoReactivate As Boolean          ' Reactivate group after closed?
    dAutoExpiration As Double           ' Don't automatically reactivate the group after this time
    bCanReactivate As Boolean           ' Can this group be reactivated (has it been manually cancelled?)
    lMarketExit As Long                 ' Index of a market exit order that has been submitted
    bFlattening As Boolean              ' Is the group in the process of being flattened?
    bActivationError As Boolean         ' Did we error trying to activate?
    nRollingPhase As eGDRollPhases      ' Phase of the rolling process
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property
Public Property Let ID(ByVal lID As Long)
    m.lID = lID
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0 Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property

Public Property Get TradeSymbol() As String
    TradeSymbol = m.strTradeSymbol
End Property
Private Property Let TradeSymbol(ByVal strTradeSymbol As String)
    If strTradeSymbol <> m.strTradeSymbol Then
        DumpDebug "TradeSymbol changing from '" & m.strTradeSymbol & "' to '" & strTradeSymbol & "'"
        m.strTradeSymbol = strTradeSymbol
        
        g.Broker.GetContractInfo m.nBroker, m.strTradeSymbol
    End If
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nBroker = g.Broker.AccountTypeForID(lAccountID)
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get Quantity() As Long
    Quantity = m.lQuantity
End Property
Public Property Let Quantity(ByVal lQuantity As Long)
    m.lQuantity = lQuantity
End Property

Public Property Get Submitted() As Boolean
    Submitted = m.bSubmitted
End Property
Public Property Let Submitted(ByVal bSubmitted As Boolean)
    If bSubmitted <> m.bSubmitted Then
        m.bSubmitted = bSubmitted
        g.TsoGroups.UpdateLastModified
    End If
End Property

Public Property Get tsOrderGroup() As cTradeSenseOrderGroup
    Set tsOrderGroup = m.tsOrderGrp
End Property
Public Property Let tsOrderGroup(ByVal tsOrderGrp As cTradeSenseOrderGroup)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumItems As Long               ' Number of items

    Set m.tsOrderGrp = tsOrderGrp
    lNumItems = tsOrderGrp.OrderCount + 1
    
    m.Orders.Clear
    m.anStatus.Create eGDARRAY_Longs, lNumItems
    m.astrToDo.Create eGDARRAY_Strings, lNumItems
    m.anLeaveOpen.Create eGDARRAY_TinyInts, lNumItems
    m.astrStatus.Create eGDARRAY_Strings, lNumItems
    m.alLevel.Create eGDARRAY_Longs, lNumItems
    m.astrAction.Create eGDARRAY_Strings, lNumItems
    m.alOrderIds.Create eGDARRAY_Longs
    
    Status(0) = eGD_TsoStatus_Open
    For lIndex = 1 To lNumItems - 1
        m.Orders.Add Nothing, Str(lIndex)
        If tsOrderGrp.Order(lIndex).OTO = 0 Then
            DisplayLevel(lIndex) = 1
            If tsOrderGrp.Order(lIndex).Use = False Then
                Status(lIndex) = eGD_TsoStatus_Closed
                DisplayStatus(lIndex) = "Not Used"
            Else
                Status(lIndex) = eGD_TsoStatus_Open
                DisplayStatus(lIndex) = "Open"
            End If
        Else
            DisplayLevel(lIndex) = DisplayLevel(tsOrderGrp.Order(lIndex).OTO) + 1
            If tsOrderGrp.Order(lIndex).Use = False Then
                Status(lIndex) = eGD_TsoStatus_Closed
                DisplayStatus(lIndex) = "Not Used"
            Else
                Status(lIndex) = eGD_TsoStatus_Trigger
                DisplayStatus(lIndex) = "Pending Trigger"
            End If
        End If
        
        LeaveOpen(lIndex) = 3
        Action(lIndex) = ""
    Next lIndex
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.TsOrderGroup.Let"
    
End Property

Public Property Get StatusForOrders() As String
    StatusForOrders = m.anStatus.JoinFields(",")
End Property
Public Property Let StatusForOrders(ByVal strStatusForOrders As String)
    m.anStatus.SplitFields strStatusForOrders, ","
End Property

Public Property Get DisplayStatusForOrders() As String
    DisplayStatusForOrders = m.astrStatus.JoinFields(",")
End Property
Public Property Let DisplayStatusForOrders(ByVal strStatusForOrders As String)
    m.astrStatus.SplitFields strStatusForOrders, ","
End Property

Public Property Get Key() As String
    Key = m.strSymbol & vbTab & Str(m.lAccountID) & vbTab & m.tsOrderGrp.ID
End Property

Public Property Get DisplayStatus(ByVal lIndex As Long) As String
    DisplayStatus = m.astrStatus(lIndex)
End Property
Private Property Let DisplayStatus(ByVal lIndex As Long, ByVal strStatus As String)
    If strStatus <> m.astrStatus(lIndex) Then
        m.astrStatus(lIndex) = strStatus
        DumpDebug vbTab & "DisplayStatus(" & Str(lIndex) & ") = " & strStatus
    End If
End Property

Public Property Get DisplayLevel(ByVal lIndex As Long) As Long
    DisplayLevel = m.alLevel(lIndex)
End Property
Private Property Let DisplayLevel(ByVal lIndex As Long, ByVal lLevel As Long)
    m.alLevel(lIndex) = lLevel
End Property

Public Property Get Status(ByVal lIndex As Long) As eGD_TsoStatus
    Status = m.anStatus(lIndex)
End Property
Private Property Let Status(ByVal lIndex As Long, ByVal nStatus As eGD_TsoStatus)
    If nStatus <> m.anStatus(lIndex) Then
        m.anStatus(lIndex) = nStatus
        DumpDebug vbTab & "Status(" & Str(lIndex) & ") = " & StatusText(lIndex)
    End If
End Property

Private Property Get LeaveOpen(ByVal lIndex As Long) As Byte
    LeaveOpen = m.anLeaveOpen(lIndex)
End Property
Private Property Let LeaveOpen(ByVal lIndex As Long, ByVal nLeaveOpen As Byte)
    m.anLeaveOpen(lIndex) = nLeaveOpen
End Property

Public Property Get Action(ByVal lIndex As Long) As String
    Action = m.astrAction(lIndex)
End Property
Private Property Let Action(ByVal lIndex As Long, ByVal strAction As String)
    If strAction <> m.astrAction(lIndex) Then
        m.astrAction(lIndex) = strAction
        DumpDebug vbTab & "Action(" & Str(lIndex) & ") = '" & strAction & "'"
    End If
End Property

Private Property Get LogText() As String
    LogText = "TsOrderGroup (" & m.tsOrderGrp.Name & ", " & g.Broker.AccountNameForID(m.lAccountID) & ", " & m.strSymbol & "): "
End Property

Public Property Get AutoReactivate() As Boolean
    AutoReactivate = m.bAutoReactivate
End Property
Public Property Let AutoReactivate(ByVal bAutoReactivate As Boolean)
    m.bAutoReactivate = bAutoReactivate
End Property

Public Property Get AutoExpiration() As Double
    AutoExpiration = m.dAutoExpiration
End Property
Public Property Let AutoExpiration(ByVal dAutoExpiration As Double)
    m.dAutoExpiration = dAutoExpiration
End Property

Public Property Get SessionActivated() As Long
    SessionActivated = m.lSessionActivated
End Property
Private Property Let SessionActivated(ByVal lSessionActivated As Long)
    m.lSessionActivated = lSessionActivated
End Property

Public Property Get Position() As Long
    Position = m.lPosition
End Property

Public Property Get FlattenOpposite() As Boolean
    FlattenOpposite = m.tsOrderGrp.FlattenOpposite
End Property
Public Property Let FlattenOpposite(ByVal bFlattenOpposite As Boolean)
    m.tsOrderGrp.FlattenOpposite = bFlattenOpposite
End Property

Public Property Get CustomStartTime() As Double
    CustomStartTime = m.dCustomStartTime
End Property
Public Property Let CustomStartTime(ByVal dCustomStartTime As Double)
    m.dCustomStartTime = dCustomStartTime
End Property

Public Property Get CustomStopTime() As Double
    CustomStopTime = m.dCustomStopTime
End Property
Public Property Let CustomStopTime(ByVal dCustomStopTime As Double)
    m.dCustomStopTime = dCustomStopTime
End Property

Public Property Get LotSize() As Long
    LotSize = m.lLotSize
End Property
Public Property Let LotSize(ByVal lLotSize As Long)
    m.lLotSize = lLotSize
End Property

Public Property Get ActivationError() As Boolean
    ActivationError = m.bActivationError
End Property

Private Property Get RollingPhase() As eGDRollPhases
    RollingPhase = m.nRollingPhase
End Property
Private Property Let RollingPhase(ByVal nRollingPhase As eGDRollPhases)
    If nRollingPhase <> m.nRollingPhase Then
        DumpDebug "RollingPhase changing from '" & RollingPhaseToString(m.nRollingPhase) & "' to '" & RollingPhaseToString(nRollingPhase) & "'"
        m.nRollingPhase = nRollingPhase
    End If
End Property

Private Property Get PositionBeforeRoll() As Long
    PositionBeforeRoll = m.lPositionBeforeRoll
End Property
Private Property Let PositionBeforeRoll(ByVal lPositionBeforeRoll As Long)
    If lPositionBeforeRoll <> m.lPositionBeforeRoll Then
        DumpDebug "PositionBeforeRoll changing from " & Str(m.lPositionBeforeRoll) & " to " & Str(lPositionBeforeRoll)
        m.lPositionBeforeRoll = lPositionBeforeRoll
    End If
End Property

Private Property Get NewTradeSymbol() As String
    NewTradeSymbol = m.strNewTradeSymbol
End Property
Private Property Let NewTradeSymbol(ByVal strNewTradeSymbol As String)
    If strNewTradeSymbol <> m.strNewTradeSymbol Then
        DumpDebug "NewTradeSymbol changing from '" & m.strNewTradeSymbol & "' to '" & strNewTradeSymbol & "'"
        m.strNewTradeSymbol = strNewTradeSymbol
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Activate
'' Description: Activate the orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Activate(Optional ByVal bReactivate As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    Dim strReason As String             ' Reason why it cannot be activated
    Dim Order As cPtOrder               ' Order object
    Dim strEvaluate As String           ' Evaluation string
    Dim strLog As String                ' String to dump to log

    m.bActivationError = False
    If CanActivate(strReason) Then
        If bReactivate Then
            If m.dAutoExpiration <> kNullData Then
                DumpDebug "Reactivating because it is before " & DateFormat(m.dAutoExpiration, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
                tsOrderGroup = tsOrderGroup
            
                m.astrFillIds.Clear
                m.Fills.Clear
                m.lPosition = 0&
            Else
                DumpDebug "Reactivating without an expiration time"
                tsOrderGroup = tsOrderGroup
            End If
        ElseIf m.bAutoReactivate Then
            If m.dAutoExpiration <> kNullData Then
                DumpDebug "Set to reactivate at " & DateFormat(m.dAutoExpiration, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
            Else
                DumpDebug "Set to reactivate"
            End If
        Else
            DumpDebug "Not set to reactivate"
        End If
        
        If (m.astrFillIds.Size > 0) And (m.Fills.Count = 0) Then
            LoadFills
        End If
        
        If m.lMarketExit <> kNullData Then
            m.lMarketExit = kNullData
            DumpDebug vbTab & "Market Exit Cleared"
        End If
        
        If (m.dCustomStartTime <> kNullData) Or (m.dCustomStopTime <> kNullData) Then
            If m.dCustomStartTime <> kNullData Then
                strLog = "Custom Start Time: " & DateFormat(m.dCustomStartTime / 1440#, NO_DATE, HH_MM, AMPM_UPPER) & " exchange time"
            End If
            If m.dCustomStopTime <> kNullData Then
                If Len(strLog) > 0 Then
                    strLog = strLog & "; "
                End If
                strLog = strLog & "Custom Stop Time: " & DateFormat(m.dCustomStopTime / 1440#, NO_DATE, HH_MM, AMPM_UPPER) & " exchange time"
            End If
            
            DumpDebug strLog
        End If
        
        SetBarProperties m.Bars, SymbolOrSymbolID
        
        If SessionActivated = kNullData Then
            SessionActivated = m.Bars.SessionDateForTradeTime(CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True))
            
            TradeSymbol = DetermineTradeSymbol(True)
            'm.strAutoBreakoutPeriod = AutoBreakoutPeriod(m.strSymbol)
        Else
            If Len(TradeSymbol) = 0 Then
                TradeSymbol = DetermineTradeSymbol(True)
            End If
            'If Len(m.strAutoBreakoutPeriod) = 0 Then
            '    m.strAutoBreakoutPeriod = AutoBreakoutPeriod(m.strSymbol)
            'End If
        End If
        NewTradeSymbol = DetermineTradeSymbol(False)
        DumpDebug vbTab & "Session Activated = " & DateFormat(SessionActivated, MM_DD_YYYY) & "; Trade Symbol = '" & TradeSymbol & "'; New Trade Symbol = '" & NewTradeSymbol & "'; Quantity = " & Str(m.lQuantity) & "; Lot Size = " & Str(m.lLotSize) ' ; Auto Period = '" & m.strAutoBreakoutPeriod & "'"
        
        For lIndex = 1 To m.tsOrderGrp.OrderCount
            If Status(lIndex) <> eGD_TsoStatus_Closed Then
                m.tsOrderGrp.Order(lIndex).ActivateWithFills m.strSymbol, m.Fills, m.nBroker, LogText, m.dCustomStartTime, m.dCustomStopTime, m.lLotSize ',  m.strAutoBreakoutPeriod
                If m.tsOrderGrp.Order(lIndex).IntraBar Then
                    strEvaluate = "Evaluate Intra Bar"
                Else
                    strEvaluate = "Evaluate Every New '" & GetPeriodStr(m.tsOrderGrp.Order(lIndex).BarPeriod) & "' Bar"
                End If
                DumpDebug "Order " & Str(lIndex) & " (" & m.tsOrderGrp.Order(lIndex).EnglishPreview(False) & ") Activated (" & StatusText(lIndex) & ") -- " & strEvaluate
                
                ' 11/11/2010 DAJ: Upon activation, send in each of the fills so that the order
                ' can calculate late calculating values...
                ' 12/17/2010 DAJ: Update the BarsSince functions after the last fill...
                ' 01/06/2010 DAJ: Send fills into the Activate call so that Late Calculating/BarsSince
                ' functions can be set before the SetValues is called...
                'For lIndex2 = 1 To m.Fills.Count
                '    m.tsOrderGrp.Order(lIndex).FillCallback m.Fills(lIndex2), (lIndex2 = m.Fills.Count)
                'Next lIndex2
                
                CancelLeftOverOrder lIndex
                
                If m.tsOrderGrp.Order(lIndex).OTO = 0 Then
                    Status(lIndex) = eGD_TsoStatus_Open
                    If m.tsOrderGrp.Order(lIndex).AllRtDataAvailable Then
                        DisplayStatus(lIndex) = "Open"
                    Else
                        DisplayStatus(lIndex) = "Waiting for Data"
                    End If
                ElseIf Status(m.tsOrderGrp.Order(lIndex).OTO) = eGD_TsoStatus_Closed Then
                    Status(lIndex) = eGD_TsoStatus_Open
                    If m.tsOrderGrp.Order(lIndex).AllRtDataAvailable Then
                        DisplayStatus(lIndex) = "Open"
                    Else
                        DisplayStatus(lIndex) = "Waiting for Data"
                    End If
                Else
                    Status(lIndex) = eGD_TsoStatus_Trigger
                    DisplayStatus(lIndex) = "Pending Trigger"
                End If
                
                Submitted = True
            End If
        Next lIndex
        
        UpdateGroupUI
        m.bCanReactivate = True
        m.bFlattening = False
        
        ' Only check orders here if we are not reactivating.  In a reactivate situation, let
        ' the orders be checked when they normally would (e.g. next new bar or intra bar)...
        If bReactivate = False Then
            ' CheckOrder will only check the orders that are open and have the Use flag turned on...
            For lIndex = 1 To m.tsOrderGrp.OrderCount
                CheckOrder lIndex
            Next lIndex
        End If
        
        bReturn = True
    Else
        InfBox "Could not activate '" & m.tsOrderGrp.Name & "' because " & strReason, "!", , "Activation Error"
        bReturn = False
    End If
    
    Activate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    m.bActivationError = True
    RaiseError "cActiveTsOrderGroup.Activate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Update the bars for the orders and check expressions if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Trade symbol
    
    If m.bSubmitted = True Then
        If ValidRunTime Then
            For lIndex = 1 To m.tsOrderGrp.OrderCount
                UpdateBarsForOrder lIndex
            Next lIndex
            
            If (TradeSymbol <> NewTradeSymbol) And (RollingPhase = eGDRollPhase_None) Then
                If ValidTradingTime = True Then
                    RollSymbol
                End If
            End If
        Else
            ParkAll "Not a valid time to run this order group", True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the given order in the collection and check OTO/OCO
'' Inputs:      Order, Old Order ID
'' Returns:     True if order found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RefreshOrder(ByVal Order As cPtOrder, Optional ByVal lOldOrderID As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim PrevOrder As cPtOrder           ' Order out of the collection
    Dim astrCommand As New cGdArray     ' Command split out
    
    bReturn = False
    For lIndex = 1 To m.Orders.Count
        If Not m.Orders(lIndex) Is Nothing Then
            Set PrevOrder = m.Orders(lIndex)
            
            If (PrevOrder.OrderID = Order.OrderID) Or (PrevOrder.OrderID = lOldOrderID) Then
                If PrevOrder.Status <> Order.Status Then
                    DumpDebug "Order Status Change for " & Str(lIndex) & " (" & Order.OrderText & "): " & OrderStatus(Order.Status)
                    Set m.Orders(lIndex) = Order
                    DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to '" & Order.OrderText & "' (" & OrderStatus(Order.Status) & ")"
                    DisplayStatus(lIndex) = OrderStatus(Order.Status)
                    
                    If Order.Status = eTT_OrderStatus_Filled Then
                        Action(lIndex) = Order.FillsText
                    ElseIf Order.Fills.Count > 0 Then
                        Action(lIndex) = Order.OrderText & "(ID = " & Order.BrokerID & "; Status = " & OrderStatus(Order.Status) & ") ; " & Order.FillsText
                    Else
                        Action(lIndex) = Order.OrderText & "(ID = " & Order.BrokerID & "; Status = " & OrderStatus(Order.Status) & ")"
                    End If
                    
                    Select Case Order.Status
                        Case eTT_OrderStatus_Partial
                            ' Let the fill callback take care of this...
                            'HandleOrderWithFills lIndex
                            'Status(lIndex) = eGD_TsoStatus_Open
                            
                        Case eTT_OrderStatus_Filled
                            ' Let the fill callback take care of this...
                            'HandleOrderWithFills lIndex
                            'Status(lIndex) = eGD_TsoStatus_Closed
                        
                        Case eTT_OrderStatus_Cancelled
                            Select Case LeaveOpen(lIndex)
                                Case 0
                                    DumpDebug "Closing Tree for " & Str(lIndex) & " (" & Order.OrderText & ") because order has been closed (" & OrderStatus(Order.Status) & ")"
                                    CloseTree lIndex
                                
                                Case 1
                                    Status(lIndex) = eGD_TsoStatus_Open
                                    DisplayStatus(lIndex) = "Condition False"
                                    LeaveOpen(lIndex) = 3
                                    Action(lIndex) = ""
                                    Set m.Orders(lIndex) = Nothing
                                    DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                                
                                Case 2
                                    Status(lIndex) = eGD_TsoStatus_Parked
                                    DisplayStatus(lIndex) = "Parked"
                                    LeaveOpen(lIndex) = 3
                                    Action(lIndex) = ""
                                    Set m.Orders(lIndex) = Nothing
                                    DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                                    
                                    If AllParked Then
                                        Submitted = False
                                    End If
                                
                                Case 3
                                    Status(lIndex) = eGD_TsoStatus_Open
                                    DisplayStatus(lIndex) = "Unsolicited Cancel"
                                    LeaveOpen(lIndex) = 3
                                    Action(lIndex) = ""
                                    Set m.Orders(lIndex) = Nothing
                                    DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                                    
                                Case 4
                                    Status(lIndex) = eGD_TsoStatus_Open
                                    DisplayStatus(lIndex) = "Condition True"
                                    LeaveOpen(lIndex) = 3
                                    Action(lIndex) = ""
                                    Set m.Orders(lIndex) = Nothing
                                    DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                                    SubmitOrder lIndex
                            
                            End Select
                        
                        Case eTT_OrderStatus_Expired
                            ' 11/11/2010 DAJ: If the order has expired, clear it out.  If the TradeSense
                            ' order is supposed to expire, that will be handled elsewhere.  If the TradeSense
                            ' order is not supposed to expire, let the order be resubmitted...
                            Set m.Orders(lIndex) = Nothing
                            DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                        
                        Case eTT_OrderStatus_Rejected
                            Set m.Orders(lIndex) = Nothing
                            DumpDebug "Parking Group because order " & Str(lIndex) & " was rejected (" & Order.Message & ")"
                            m.bCanReactivate = False
                            ParkAll "Order " & Str(lIndex) & " has been rejected (" & Order.Message & ")", True
                            
                        Case eTT_OrderStatus_Error
                            Set m.Orders(lIndex) = Nothing
                            DumpDebug "Parking Group because order " & Str(lIndex) & " could not be submitted (" & Order.Message & ")"
                            m.bCanReactivate = False
                            ParkAll "Order " & Str(lIndex) & " could not be submitted (" & Order.Message & ")", True
                        
                        Case eTT_OrderStatus_BalCancelled
                            Status(lIndex) = eGD_TsoStatus_Closed
                            Set m.Orders(lIndex) = Nothing
                            DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                        
                        Case eTT_OrderStatus_Working
                            If Len(m.astrToDo(lIndex)) > 0 Then
                                astrCommand.SplitFields m.astrToDo(lIndex), vbTab
                                
                                Select Case UCase(astrCommand(0))
                                    Case "CANCEL"
                                        CancelOrder lIndex, "no longer pending", CBool(astrCommand(1))
                                        m.astrToDo(lIndex) = ""
                                    Case "PARK"
                                        ParkOrder lIndex
                                        m.astrToDo(lIndex) = ""
                                    Case "MODIFY"
                                        Order.StopPrice = Val(astrCommand(1))
                                        Order.LimitPrice = Val(astrCommand(2))
                                        Order.Quantity = CLng(Val(astrCommand(3)))
                                        
                                        ModifyOrder lIndex, Order, "no longer pending"
                                        m.astrToDo(lIndex) = ""
                                End Select
                            End If
                        
                        Case eTT_OrderStatus_Parked
                            Status(lIndex) = eGD_TsoStatus_Parked
                            If AllParked Then
                                Submitted = False
                            End If
                            
                        'Case eTT_OrderStatus_Amended
                    End Select
                    
                    UpdateOrderUI lIndex
                End If
                
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    ' If we are flattening, when the last order is closed, close the position...
    If bReturn And m.bFlattening Then
        If AllClosed Then
            ClosePosition "Flattening"
        End If
    ElseIf (bReturn = True) And (RollingPhase = eGDRollPhase_Parking) Then
        RollSymbol
    End If
    
    RefreshOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.RefreshOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillCallback
'' Description: Handle the fill for one of the orders
'' Inputs:      Fill, Order, Entered a Position?
'' Returns:     True if order found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillCallback(ByVal Fill As cPtFill, ByVal FillOrder As cPtOrder, Optional bEnteredPosition As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order object
    Dim bUpdateBars As Boolean          ' Update bars?
    Dim lPreviousPosition As Long       ' Position before this fill

    bReturn = False
    
    If (RollingPhase <> eGDRollPhase_None) And (Not m.RollOrder Is Nothing) Then
        If Fill.OrderID = m.RollOrder.OrderID Then
            UpdateFill Fill, True
            CalculatePosition
            
            RollSymbol
            
            bReturn = True
        End If
    End If
    
    If bReturn = False Then
        For lIndex = 1 To m.Orders.Count
            If Not m.Orders(lIndex) Is Nothing Then
                Set Order = m.Orders(lIndex)
                If Not Order Is Nothing Then
                    If Order.OrderID = Fill.OrderID Then
                        DumpDebug "FillCallback for order " & Str(lIndex) & ": " & mTradeTracker.FillDisplay(Fill)
                        
                        lPreviousPosition = m.lPosition
                        UpdateFill Fill
                        CalculatePosition
                        
                        If FillOrder.Status = eTT_OrderStatus_Filled Then
                            Action(lIndex) = FillOrder.FillsText
                        ElseIf FillOrder.Fills.Count > 0 Then
                            Action(lIndex) = FillOrder.OrderText & "(ID = " & FillOrder.BrokerID & "; Status = " & OrderStatus(FillOrder.Status) & ") ; " & FillOrder.FillsText
                        End If
                        
                        bUpdateBars = HandleOrderWithFills(lIndex)
                        If Status(lIndex) = eGD_TsoStatus_Trigger Then
                            Status(lIndex) = eGD_TsoStatus_Open
                            DisplayStatus(lIndex) = "Open"
                            
                        ElseIf FillOrder.Status = eTT_OrderStatus_Filled Then
                            Status(lIndex) = eGD_TsoStatus_Closed
                            DisplayStatus(lIndex) = "Filled"
                            Set m.Orders(lIndex) = Nothing
                            
                            If m.lMarketExit = lIndex Then
                                m.lMarketExit = kNullData
                                DumpDebug vbTab & "Market Exit Cleared"
                            End If
                            
                            DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                        End If
                        UpdateOrderUI lIndex
                        
                        bEnteredPosition = ((m.lPosition <> 0&) And (lPreviousPosition = 0&))
                        DumpDebug "Entered Position = " & Str(bEnteredPosition) & "; Previous = " & Str(lPreviousPosition) & "; Current = " & Str(m.lPosition)
                        
                        bReturn = True
                        Exit For
                    End If
                End If
            End If
        Next lIndex
    End If
    
    FillCallback = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.FillCallback"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllClosed
'' Description: Determine if all of the orders are closed
'' Inputs:      None
'' Returns:     True if All Closed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllClosed() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    For lIndex = 1 To m.anStatus.Size - 1
        If Status(lIndex) <> eGD_TsoStatus_Closed Then
            bReturn = False
            Exit For
        End If
    Next lIndex
    
    AllClosed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.AllClosed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllParked
'' Description: Determine if all of the orders are closed or parked
'' Inputs:      None
'' Returns:     True if All Parked or Closed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllParked() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    For lIndex = 1 To m.anStatus.Size - 1
        If (Status(lIndex) <> eGD_TsoStatus_Closed) And (Status(lIndex) <> eGD_TsoStatus_Parked) Then
            bReturn = False
            Exit For
        End If
    Next lIndex
    
    AllParked = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.AllParked"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string for persistance
'' Inputs:      None
'' Returns:     Information String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Trade sense order group info
    
    Set astrInfo = New cGdArray
    astrInfo.Create eGDARRAY_Strings, 8
    
    astrInfo(0) = m.tsOrderGrp.FileName
    astrInfo(1) = Str(m.tsOrderGrp.Custom)
    astrInfo(2) = Symbol
    astrInfo(3) = Str(AccountID)
    astrInfo(4) = Str(Quantity)
    astrInfo(5) = StatusForOrders
    astrInfo(6) = DisplayStatusForOrders
    astrInfo(7) = FillIds
    astrInfo(8) = m.tsOrderGrp.Inputs.ToString
    astrInfo(9) = OrderIds
    astrInfo(10) = Str(m.bAutoReactivate)
    astrInfo(11) = Str(m.dAutoExpiration)
    astrInfo(12) = Str(m.lSessionActivated)
    astrInfo(13) = Str(m.dCustomStartTime)
    astrInfo(14) = Str(m.dCustomStopTime)
    astrInfo(15) = TradeSymbol
    astrInfo(16) = Str(m.lLotSize)
    
    ToString = astrInfo.JoinFields(vbTab)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Build up the object from the given string
'' Inputs:      Information String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strInfoString As String)
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' TradeSense order group info
    Dim tsGrp As cTradeSenseOrderGroup  ' TradeSense order group object
    Dim Inputs As cTradeSenseOrderInputs ' TradeSense order inputs
    
    Set astrInfo = New cGdArray
    astrInfo.SplitFields strInfoString, vbTab
    
    Set tsGrp = New cTradeSenseOrderGroup
    tsGrp.FromFile astrInfo(0), CBool(astrInfo(1))
    
    tsOrderGroup = tsGrp
    SymbolOrSymbolID = astrInfo(2)
    AccountID = CLng(Val(astrInfo(3)))
    Quantity = CLng(Val(astrInfo(4)))
    If Len(astrInfo(5)) > 0 Then
        StatusForOrders = astrInfo(5)
    End If
    If Len(astrInfo(6)) > 0 Then
        DisplayStatusForOrders = astrInfo(6)
    End If
    m.astrFillIds.SplitFields astrInfo(7), ","
    
    If Len(astrInfo(8)) > 0 Then
        Set Inputs = New cTradeSenseOrderInputs
        Inputs.ForGroups = True
        Inputs.FromString astrInfo(8)
        tsOrderGroup.SetInputValues Inputs
    End If
    
    If Len(astrInfo(9)) > 0 Then
        m.alOrderIds.SplitFields astrInfo(9), ","
    End If
    
    If Len(astrInfo(10)) > 0 Then
        m.bAutoReactivate = (astrInfo(10) <> "0")
    End If
    
    If Len(astrInfo(11)) > 0 Then
        m.dAutoExpiration = Val(astrInfo(11))
    End If
    
    If Len(astrInfo(12)) > 0 Then
        m.lSessionActivated = CLng(Val(astrInfo(12)))
    End If
    
    If Len(astrInfo(13)) > 0 Then
        m.dCustomStartTime = Val(astrInfo(13))
    End If

    If Len(astrInfo(14)) > 0 Then
        m.dCustomStopTime = Val(astrInfo(14))
    End If
    
    If Len(astrInfo(15)) > 0 Then
        TradeSymbol = astrInfo(15)
    End If
    
    If Len(astrInfo(16)) > 0 Then
        m.lLotSize = CLng(Val(astrInfo(16)))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOrder
'' Description: Does this order group contain the given order?
'' Inputs:      Order
'' Returns:     True if group has order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    HasOrder = HasOrderID(Order.OrderID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.HasOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasOrderID
'' Description: Does this order group contain the given order ID?
'' Inputs:      Order ID
'' Returns:     True if group has order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasOrderID(ByVal lOrderID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.Orders.Count
        If Not m.Orders(lIndex) Is Nothing Then
            If m.Orders(lIndex).OrderID = lOrderID Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasOrderID = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.HasOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelRequested
'' Description: The given order is being cancelled some other place in Trade Navigator
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelRequested(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Orders.Count
        If Not m.Orders(lIndex) Is Nothing Then
            If m.Orders(lIndex).OrderID = Order.OrderID Then
                If LeaveOpen(lIndex) = 3 Then
                    LeaveOpen(lIndex) = 0
                End If
                
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CancelRequested"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelLeftOverOrders
'' Description: Cancel any of the left over orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelLeftOverOrders()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.tsOrderGrp.OrderCount
        CancelLeftOverOrder lIndex
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CancelLeftOverOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReloadData
'' Description: Reload the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReloadData()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    
    For lIndex = 1 To m.tsOrderGrp.OrderCount
        Set tsOrder = m.tsOrderGrp.Order(lIndex)
        If Not tsOrder Is Nothing Then
            tsOrder.ReloadData
        End If
    Next lIndex
    
    NewTradeSymbol = DetermineTradeSymbol(False)
    DumpDebug "Data reloaded"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ReloadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelLeftOverOrder
'' Description: Cancel the given left over order
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelLeftOverOrder(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    
    If m.alOrderIds(lIndex) <> 0 Then
        Set Order = New cPtOrder
        If Order.Load(m.alOrderIds(lIndex)) Then
            If IsOpenOrder(Order.Status) Then
                DumpDebug "Cancelling Order " & Order.BrokerID & ": Left over order from previous instance"
                mTradeTracker.CancelOrder Order, False
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CancelLeftOverOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Submit or Amend the given order as appropriate
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOrder(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    Dim bModify As Boolean              ' Modify the order?
    Dim bContinue As Boolean            ' Continue with the submit?
    Dim strReturn As String             ' Return value from user question
    Dim lNewQty As Long                 ' New quantity for the order
    Dim lOldQty As Long                 ' Old quantity for the order
    Static bInProgress As Boolean       ' Are we already in here?
    
    If bInProgress = False Then
        bInProgress = True
        
        Set tsOrder = m.tsOrderGrp.Order(lIndex)
        If Not tsOrder Is Nothing Then
            If m.Orders(lIndex) Is Nothing Then
                Set Order = New cPtOrder
                With Order
                    .AccountID = m.lAccountID
                    .AutoTradeItemID = 0
                    .Buy = tsOrder.Buy
                    .Enter = tsOrder.Entry
                    .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(m.lAccountID))
                    .OrderType = tsOrder.OrderType
                    .Quantity = OrderQuantity(tsOrder)
                    .SymbolOrSymbolID = TradeSymbol
                    .IsAutomated = True
                    
                    ' DAJ 04/28/2011: Since PFG works off of NY time and Genesis Forex works off of GMT,
                    ' half of the year PFG forex closes an hour earlier.  This will cause day orders to expire
                    ' before the end of the Genesis forex session.  In light of this, we are going to
                    ' try submitting GTC orders for the PFG forex symbols...
                    If UseGtcOrder(.Symbol) Then
                        .Expiration = 0
                    Else
                        .Expiration = -1&
                    End If
                
                    bContinue = False
                    If .Quantity > 0 Then
                        Select Case .OrderType
                            Case eTT_OrderType_Market
                                .StopPrice = 0
                                .LimitPrice = 0
                                bContinue = True
                            Case eTT_OrderType_Limit
                                .StopPrice = 0
                                .LimitPrice = tsOrder.LimitPrice
                                bContinue = (.LimitPrice > 0)
                            Case eTT_OrderType_Stop
                                .StopPrice = tsOrder.StopPrice
                                .LimitPrice = 0
                                bContinue = (.StopPrice > 0)
                            Case eTT_OrderType_StopWithLimit
                                .StopPrice = tsOrder.StopPrice
                                .LimitPrice = tsOrder.LimitPrice
                                bContinue = (.StopPrice > 0) And (.LimitPrice > 0)
                        End Select
                        
                        ' DAJ 11/28/2012: If the order is a Stop or a Stop with Limit, but it is going to
                        ' be submitted on the wrong side of the market AND the user has chosen the option
                        ' to turn it into a Market order, do so here and submit it...
                        If (.OrderType = eTT_OrderType_Stop) Or (.OrderType = eTT_OrderType_StopWithLimit) Then
                            If (.WrongSideOfMarket = True) And (tsOrder.MarketIfWrongSide = True) Then
                                DumpDebug "Order " & Str(lIndex) & " changed to a Market order because Stop was on wrong side of market"
                                .OrderType = eTT_OrderType_Market
                                .StopPrice = 0
                                .LimitPrice = 0
                                bContinue = True
                            End If
                        End If
                    End If
                End With
                
                If bContinue Then
                    If Order.WrongSideOfMarket Then
                        DumpDebug "Not Submitting Order " & Str(lIndex) & " (" & Order.OrderText & ") because condition (" & tsOrder.ConditionEnglish & ") is True but Price is on the wrong side of the market (Price = '" & tsOrder.EnglishOrderPrice & "')"
                        m.bCanReactivate = False
                        ParkAll "Order " & Str(lIndex) & " could not be submitted (Order is on the wrong side of the market)", True
                    ElseIf m.lMarketExit = kNullData Then
                        strReturn = "S"
                        If g.Broker.ConfirmTradeSense Then
                            strReturn = ConfirmOrder(Order, m.tsOrderGrp.Name, "P")
                        End If
                        
                        Select Case strReturn
                            Case "S"
                                DumpDebug "Submitting Order " & Str(lIndex) & " (" & Order.OrderText & ") because condition (" & tsOrder.ConditionEnglish & ") is True (Price = '" & tsOrder.EnglishOrderPrice & "') Current Market = '" & tsOrder.CurrentPrices & "'"
                                
                                Order.Save
                                Set m.Orders(lIndex) = Order
                                DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to '" & Order.OrderText & "' (" & OrderStatus(Order.Status) & ")"
                                
                                If g.Broker.SendOrder(Order) > 0 Then
                                    If Order.Status <> eTT_OrderStatus_Error Then
                                        ' If this is a Market order that is set to exit 100% of the position,
                                        ' don't submit any other exits...
                                        If (Order.Enter = False) And (Order.OrderType = eTT_OrderType_Market) And (tsOrder.ExitPercent = 100) Then
                                            m.lMarketExit = lIndex
                                            DumpDebug vbTab & "Market Exit set to " & Str(lIndex)
                                        End If
                                        
                                        If g.Broker.AutoJournalAutomated Then
                                            g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "TradeSense Order Group (" & m.tsOrderGrp.Name & ") submitting order because condition is True"
                                        End If
                                    End If
                                End If
                                
                            Case "P"
                                ParkOrder lIndex
                            
                            Case "C"
                                CancelOrder lIndex, "User Cancelled from Confirmation", False
                            
                        End Select
                    End If
                ElseIf Order.Quantity = 0 Then
                    If m.lPosition = 0 Then
                        DumpDebug "Not Submitting Order " & Str(lIndex) & " (" & Order.OrderText & ") because no position to exit"
                        DisplayStatus(lIndex) = "No Position to Exit"
                        UpdateGroupUI
                    Else
                        DumpDebug "Not Submitting Order " & Str(lIndex) & " (" & Order.OrderText & ") because exiting wrong side of position"
                        DisplayStatus(lIndex) = "Exits Other Side"
                        UpdateGroupUI
                    End If
                Else
                    DumpDebug "Not Submitting Order " & Str(lIndex) & " (" & Order.OrderText & ") because condition (" & tsOrder.ConditionEnglish & ") is True but Price is invalid (Price = '" & tsOrder.EnglishOrderPrice & "')"
                    DisplayStatus(lIndex) = "Invalid Price"
                    UpdateGroupUI
                End If
            Else
                bModify = False
                Set Order = m.Orders(lIndex).MakeCopy
                With Order
                    Select Case .OrderType
                        Case eTT_OrderType_Limit
                            If tsOrder.LimitPrice <> .LimitPrice Then
                                .StopPrice = 0
                                .LimitPrice = tsOrder.LimitPrice
                                bModify = True
                            End If
                        Case eTT_OrderType_Stop
                            If tsOrder.StopPrice <> .StopPrice Then
                                .StopPrice = tsOrder.StopPrice
                                .LimitPrice = 0
                                bModify = True
                            End If
                        Case eTT_OrderType_StopWithLimit
                            If (tsOrder.LimitPrice <> .LimitPrice) Or (tsOrder.StopPrice <> .StopPrice) Then
                                .StopPrice = tsOrder.StopPrice
                                .LimitPrice = tsOrder.LimitPrice
                                bModify = True
                            End If
                    End Select
                    
                    lNewQty = OrderQuantity(tsOrder)
                    lOldQty = .Quantity
                End With
                
                ' DAJ 08/23/2011: After talking with Pete and Eric about some trade logs,
                ' we are going to go ahead and adjust the quantity on submitted exits here
                ' too...
                If (lOldQty <> lNewQty) Then ' And (tsOrder.Entry = True) Then
                    Order.Quantity = lNewQty
                    If bModify Then
                        ModifyOrder lIndex, Order, "price changed (" & tsOrder.EnglishOrderPrice & ") and quantity changed from " & Str(lOldQty) & " to " & Str(lNewQty)
                    Else
                        ModifyOrder lIndex, Order, "quantity changed from " & Str(lOldQty) & " to " & Str(lNewQty)
                    End If
                ElseIf bModify Then
                    ModifyOrder lIndex, Order, "price changed (" & tsOrder.EnglishOrderPrice & ")"
                End If
            End If
        End If
        
        bInProgress = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "cActiveTsOrderGroup.SubmitOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrder
'' Description: Set the given order up for modification
'' Inputs:      Index, Order, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ModifyOrder(ByVal lIndex As Long, Order As cPtOrder, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim PrevOrder As cPtOrder           ' The existing order
    
    Set PrevOrder = m.Orders(lIndex)
    If Not PrevOrder Is Nothing Then
        If (OrderIsPending(PrevOrder) = True) Or (PrevOrder.Status = eTT_OrderStatus_Amended) Then
            DumpDebug "Setting up Modify Order " & Str(lIndex) & " (" & Order.OrderText & ") because " & strReason & " (Pending)"
            m.astrToDo(lIndex) = "Modify" & vbTab & Str(Order.StopPrice) & vbTab & Str(Order.LimitPrice) & vbTab & Str(Order.Quantity)
        ElseIf IsOpenOrder(PrevOrder.Status, False) Then
            If ((Order.OrderType = eTT_OrderType_Stop) Or (Order.OrderType = eTT_OrderType_StopWithLimit)) And (Order.WrongSideOfMarket = True) Then
                If m.tsOrderGrp.Order(lIndex).MarketIfWrongSide = True Then
                    DumpDebug "Re-submitting order " & Str(lIndex) & " (" & Order.OrderText & ") because stop is on wrong side of market"
                    LeaveOpen(lIndex) = 4
                    mTradeTracker.CancelOrder PrevOrder, False
                Else
                    DumpDebug "Not modifying order " & Str(lIndex) & " (" & Order.OrderText & ") because stop is on wrong side of market"
                End If
            Else
                DumpDebug "Modifying Order " & Str(lIndex) & " (" & Order.OrderText & ") because " & strReason
                mTradeTracker.EditOrder Order, , False, eGDEditOrderReturn_Submit, True
            
                If g.Broker.AutoJournalAutomated Then
                    g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "TradeSense Order Group (" & m.tsOrderGrp.Name & ") modifying order"
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ModifyOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Set the given order up for cancel
'' Inputs:      Index, Reason, Leave Open?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelOrder(ByVal lIndex As Long, ByVal strReason As String, ByVal bLeaveOpen As Boolean)
On Error GoTo ErrSection:

    Dim PrevOrder As cPtOrder           ' The existing order
    
    Set PrevOrder = m.Orders(lIndex)
    If Not PrevOrder Is Nothing Then
        If PrevOrder.OrderType <> eTT_OrderType_Market Then
            If (OrderIsPending(PrevOrder) = True) Or (PrevOrder.Status = eTT_OrderStatus_Amended) Then
                DumpDebug "Setting up Cancel Order " & Str(lIndex) & " (" & PrevOrder.OrderText & ") because " & strReason & " (Pending)"
                If bLeaveOpen Then
                    m.astrToDo(lIndex) = "Cancel" & vbTab & "1"
                Else
                    m.astrToDo(lIndex) = "Cancel" & vbTab & "0"
                End If
            ElseIf IsOpenOrder(PrevOrder.Status, False) Then
                DumpDebug "Cancelling Order " & Str(lIndex) & " (" & PrevOrder.OrderText & ") because " & strReason
                If bLeaveOpen Then
                    LeaveOpen(lIndex) = 1
                Else
                    LeaveOpen(lIndex) = 0
                End If
                mTradeTracker.CancelOrder PrevOrder, False
                
                If g.Broker.AutoJournalAutomated Then
                    g.JournalBridge.CreateOrderJournal PrevOrder.OrderID, CurrentTime, "TradeSense Order Group (" & m.tsOrderGrp.Name & ") cancelling order"
                End If
            End If
        Else
            DumpDebug "Not Cancelling Order " & Str(lIndex) & " because it is a market order"
        End If
    ElseIf (LeaveOpen(lIndex) = 0) Or (bLeaveOpen = False) Then
        CloseTree lIndex, False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Set the given order up for parking
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkOrder(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim PrevOrder As cPtOrder           ' The existing order
    
    Set PrevOrder = m.Orders(lIndex)
    If Not PrevOrder Is Nothing Then
        If (OrderIsPending(PrevOrder) = True) Or (PrevOrder.Status = eTT_OrderStatus_Amended) Then
            DumpDebug "Setting up Park Order " & Str(lIndex) & " (" & PrevOrder.OrderText & ") (Pending)"
            m.astrToDo(lIndex) = "Park"
        ElseIf IsOpenOrder(PrevOrder.Status, False) Then
            DumpDebug "Parking Order " & Str(lIndex) & " (" & PrevOrder.OrderText & ")"
            LeaveOpen(lIndex) = 2
            mTradeTracker.CancelOrder PrevOrder, False
                
            If g.Broker.AutoJournalAutomated Then
                g.JournalBridge.CreateOrderJournal PrevOrder.OrderID, CurrentTime, "TradeSense Order Group (" & m.tsOrderGrp.Name & ") parking order"
            End If
        End If
    ElseIf LeaveOpen(lIndex) = 0 Then
        CloseTree lIndex, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ParkOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateOrder
'' Description: Set the given order up for activating
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ActivateOrder(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim PrevOrder As cPtOrder           ' The existing order
    
    Set PrevOrder = m.Orders(lIndex)
    If Not PrevOrder Is Nothing Then
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ActivateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAll
'' Description: Cancel the entire order group
'' Inputs:      Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CancelAll(ByVal strReason As String, Optional ByVal bCanReactivate As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.tsOrderGrp.OrderCount
        DumpDebug "Cancelling " & Str(lIndex) & "(" & StatusText(lIndex) & "): " & strReason
        If (Status(lIndex) = eGD_TsoStatus_Open) Or (Status(lIndex) = eGD_TsoStatus_Parked) Then
            CancelOrder lIndex, strReason, False
        ElseIf Status(lIndex) = eGD_TsoStatus_Trigger Then
            CloseTree lIndex
        End If
    Next lIndex
    
    UpdateGroupUI
    
    m.bCanReactivate = bCanReactivate

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CancelAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkAll
'' Description: Park the entire order group
'' Inputs:      Reason, Show Message?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ParkAll(Optional ByVal strReason As String = "", Optional ByVal bShowMessage As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.tsOrderGrp.OrderCount
        DumpDebug "Parking " & Str(lIndex) & "(" & StatusText(lIndex) & "): '" & strReason & "'"
        If Status(lIndex) = eGD_TsoStatus_Open Then
            If m.Orders(lIndex) Is Nothing Then
                CloseTree lIndex, True
            Else
                ParkOrder lIndex
            End If
        ElseIf Status(lIndex) = eGD_TsoStatus_Trigger Then
            CloseTree lIndex, True
        End If
    Next lIndex
    
    If AllParked Then
        Submitted = False
    End If
    
    UpdateGroupUI
    
    If (bShowMessage = True) And (Len(strReason) > 0) Then
        frmOnlineBroker.AddDialogMessage "'" & m.tsOrderGrp.Name & "' for symbol '" & m.strSymbol & "' has been parked for the following reason:||" & strReason & "|", "!", , "TradeSense Order Group Parked"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ParkAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingOrders
'' Description: Determine if the group has working orders
'' Inputs:      None
'' Returns:     True if Working Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasWorkingOrders() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    For lIndex = 1 To m.Orders.Count
        If Not m.Orders(lIndex) Is Nothing Then
            If IsOpenOrder(m.Orders(lIndex).Status) Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasWorkingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.HasWorkingOrders"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasPendingOrders
'' Description: Determine if the group has pending orders
'' Inputs:      None
'' Returns:     True if Pending Orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasPendingOrders() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    For lIndex = 1 To m.Orders.Count
        If Not m.Orders(lIndex) Is Nothing Then
            If OrderIsPending(m.Orders(lIndex)) Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasPendingOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.HasPendingOrders"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CloseTree
'' Description: Close the given node and all order triggered by it
'' Inputs:      Index, Park?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CloseTree(ByVal lNode As Long, Optional ByVal bPark As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.tsOrderGrp.OrderCount
        If Status(lIndex) <> eGD_TsoStatus_Closed Then
            If lIndex = lNode Then
                If bPark Then
                    Status(lIndex) = eGD_TsoStatus_Parked
                    DisplayStatus(lIndex) = "Parked"
                    
                    If AllParked Then
                        Submitted = False
                    End If
                Else
                    Status(lIndex) = eGD_TsoStatus_Closed
                    DisplayStatus(lIndex) = "Cancelled"
                    Set m.Orders(lIndex) = Nothing
                    DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                End If
                UpdateOrderUI lIndex
            ElseIf m.tsOrderGrp.Order(lIndex).OTO = lNode Then
                CloseTree lIndex, bPark
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CloseTree"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanActivate
'' Description: Can the group be activated?
'' Inputs:      Reason
'' Returns:     True if can activate, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanActivate(Optional strReason As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nPeriodicity As Long            ' Periodicity
    
    bReturn = False
    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
        If (m.lQuantity > 0&) And (m.lLotSize > 0&) Then
            If g.Broker.ConnectionStatusForAccount(m.lAccountID) = eGDConnectionStatus_Connected Then
                If g.Broker.PositionMatch(m.lAccountID, TradeSymbol) Then
                    If (g.Broker.IsLiveAccount(m.nBroker) = True) And (g.RealTime.SymbolDelay(m.strSymbol) > 0) Then
                        strReason = "'" & m.strSymbol & "' is currently being streamed delayed"
                    Else
                        If ValidRunTime Then
                            bReturn = True
                            strReason = ""
                        Else
                            strReason = "Not a valid time to run this order group"
                        End If
                    End If
                Else
                    strReason = "There is a Position Mismatch for '" & TradeSymbol & "' in account '" & g.Broker.AccountNameForID(m.lAccountID) & "'"
                End If
            Else
                strReason = "Not currently connected to account '" & g.Broker.AccountNameForID(m.lAccountID) & "'"
            End If
        Else
            If m.lQuantity <= 0& Then
                strReason = "Invalid quantity: " & Str(m.lQuantity)
            Else
                strReason = "Invalid lot size: " & Str(m.lLotSize)
            End If
        End If
    Else
        strReason = "Not currently connected to streaming"
    End If
    
    CanActivate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CanActivate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoReactivate
'' Description: Should this group be reactivated?
'' Inputs:      Non
'' Returns:     True if reactivate, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DoReactivate() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If AllClosed = True Then
        If m.bAutoReactivate = True Then
            If (g.RealTime.FeedTime < m.dAutoExpiration) Or (m.dAutoExpiration = kNullData) Then
                bReturn = m.bCanReactivate
            Else
                DumpDebug "Not reactivating because after " & DateFormat(m.dAutoExpiration, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
            End If
        End If
    End If
    
    DoReactivate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.DoReactivate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllRtDataAvailable
'' Description: Determine if streaming data is available for all symbols
'' Inputs:      None
'' Returns:     True if streaming data available, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllRtDataAvailable() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRtDataAvailable As Boolean     ' Streaming data available for all?
    Dim dTickCount As Double            ' Tick count
    Dim bExprIsNothing As Boolean       ' Is the order's expression set to nothing?

    If m.bRtDataAvailable = False Then
        bRtDataAvailable = True
        For lIndex = 1 To m.tsOrderGrp.OrderCount
            If Status(lIndex) <> eGD_TsoStatus_Closed Then
                If m.tsOrderGrp.Order(lIndex).AllRtDataAvailable(bExprIsNothing) = False Then
                    dTickCount = gdTickCount
                    If dTickCount > m.dLastLogWait + 30000 Then
                        DumpDebug "Order " & Str(lIndex) & " waiting for data (ExprIsNothing = " & Str(bExprIsNothing) & ")"
                        m.dLastLogWait = dTickCount
                    End If
                    
                    bRtDataAvailable = False
                    Exit For
                End If
            End If
        Next lIndex
        
        m.bRtDataAvailable = bRtDataAvailable
    End If
    
    AllRtDataAvailable = m.bRtDataAvailable
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.AllRtDataAvailable"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flatten
'' Description: Flatten position (if in one) and close the order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Flatten()
On Error GoTo ErrSection:

    m.bFlattening = True
    CancelAll "Flattening"
    
    If AllClosed Then
        ClosePosition "Flattening"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.Flatten"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClosePosition
'' Description: Close the current position
'' Inputs:      Reason, Rolling?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClosePosition(ByVal strReason As String, Optional ByVal bRolling As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to close the position
    
    If m.lPosition <> 0& Then
        DumpDebug "Closing Position because " & strReason
        Set Order = CreateMarketOrder(m.lAccountID, TradeSymbol, (m.lPosition < 0), Abs(m.lPosition * m.lLotSize))
        
        Order.Enter = False
        Order.Save
        
        If bRolling = True Then
            Set m.RollOrder = Order
        End If
        
        g.Broker.SendOrder Order
                
        If g.Broker.AutoJournalAutomated Then
            g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "TradeSense Order Group (" & m.tsOrderGrp.Name & ") closing position"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ClosePosition"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderWithFills
'' Description: Handle an order that has at least partially filled
'' Inputs:      Index
'' Returns:     True if need to Update Bars, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HandleOrderWithFills(ByVal lNode As Long) As Boolean
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object
    Dim lIndex As Long                  ' Index into a for loop
    Dim bUpdateBars As Boolean          ' Update bars?
    Dim alToCheck As cGdArray           ' Array of indexes to check
    
    Set alToCheck = New cGdArray
    alToCheck.Create eGDARRAY_Longs
    
    bUpdateBars = False
    
    Set Order = m.Orders(lNode)
    If Not Order Is Nothing Then
        For lIndex = 1 To m.tsOrderGrp.OrderCount
            If (Status(lIndex) <> eGD_TsoStatus_Closed) And (Status(lIndex) <> eGD_TsoStatus_Parked) Then
                If m.tsOrderGrp.Order(lIndex).OTO = lNode Then
                    If Status(lIndex) = eGD_TsoStatus_Trigger Then
                        Status(lIndex) = eGD_TsoStatus_Open
                        DisplayStatus(lIndex) = "Open"
                        UpdateOrderUI lIndex
                        alToCheck.Add lIndex
                        
                        ' DAJ 10/06/2011: Need to do a RunExpressions here so that any late calculating
                        ' functions get updated with the fill that just happened...
                        'm.tsOrderGrp.Order(lIndex).RunExpressions
                        
                        bUpdateBars = True
                    ElseIf Not m.Orders(lIndex) Is Nothing Then
                        ' DAJ 05/04/2011: Want to check orders here in case we got more partial fills
                        ' that changed the position on us...
                        alToCheck.Add lIndex
                    End If
                ElseIf m.tsOrderGrp.Order(lIndex).IsOCO(lNode) Then
                    If (Status(lIndex) = eGD_TsoStatus_Trigger) Or (Status(lIndex) = eGD_TsoStatus_Parked) Or (m.Orders(lIndex) Is Nothing) Then
                        DumpDebug "Closing Tree for " & Str(lIndex) & " because " & Str(lNode) & " is filling (OCO)"
                        CloseTree lIndex
                    Else
                        CancelOrder lIndex, "order " & Str(lNode) & " filling (OCO)", False
                    End If
                    
                ' DAJ 04/04/2011: Do not set the order that is currently filling to be checked
                ' here because it could cause it to be modified when we don't want it to be...
                ElseIf (Status(lIndex) = eGD_TsoStatus_Open) And (lIndex <> lNode) Then
                    alToCheck.Add lIndex
                End If
            End If
        Next lIndex
    
        For lIndex = 0 To alToCheck.Size - 1
            ' DAJ 10/06/2011: Need to do a RunExpressions here so that any late calculating
            ' functions get updated with the fill that just happened...
            ' DAJ 02/04/2015: Moved this from the above OTO block so that we update all of
            ' the orders that we are going to check...
            m.tsOrderGrp.Order(alToCheck(lIndex)).RunExpressions
            
            ' DAJ 10/04/2010: In the case of "No Position to Exit" and "Exits other Side" orders, we want
            ' to do a CheckOrder no matter what.  So, if the UpdateBarsForOrder routine didn't do it,
            ' we need to do it here...
            If (UpdateBarsForOrder(alToCheck(lIndex)) = False) And (Status(alToCheck(lIndex)) = eGD_TsoStatus_Open) Then
                CheckOrder alToCheck(lIndex)
            End If
        Next lIndex
    End If
    
    HandleOrderWithFills = bUpdateBars

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.HandleOrderWithFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Update the fill for the orders and check expressions if necessary
'' Inputs:      Fill, Force?, Update TradeSense orders?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateFill(ByVal Fill As cPtFill, Optional ByVal bForce As Boolean = False, Optional ByVal bUpdateTsOrders As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bExists As Boolean              ' Does the Fill ID exist in the array?
    
    If (m.bSubmitted = True) Or (bForce = True) Then
        If m.Fills.Exists(Str(Fill.FillID)) Then
            Set m.Fills(Str(Fill.FillID)) = Fill
        Else
            m.Fills.Add Fill, Str(Fill.FillID)
        End If
        
        bExists = False
        For lIndex = 0 To m.astrFillIds.Size - 1
            If m.astrFillIds(lIndex) = Str(Fill.FillID) Then
                bExists = True
                Exit For
            End If
        Next lIndex
        If bExists = False Then
            m.astrFillIds.Add Str(Fill.FillID)
        End If
        
        If bUpdateTsOrders = True Then
            For lIndex = 1 To m.tsOrderGrp.OrderCount
                'If Status(lIndex) = eGD_TsoStatus_Open Then
                If Status(lIndex) <> eGD_TsoStatus_Closed Then
                    m.tsOrderGrp.Order(lIndex).FillCallback Fill
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.UpdateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOrderUI
'' Description: Update the order in the user interface
'' Inputs:      Order Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateOrderUI(ByVal lOrderNumber As Long)
On Error GoTo ErrSection:

    If FormIsLoaded("frmActiveTsOrderGroups") Then
        frmActiveTsOrderGroups.UpdateGroupOrder Me, lOrderNumber
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.UpdateOrderUI"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateGroupUI
'' Description: Update the group in the user interface
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateGroupUI()
On Error GoTo ErrSection:

    If AllClosed = False Then
        If FormIsLoaded("frmActiveTsOrderGroups") Then
            frmActiveTsOrderGroups.UpdateGroup Me
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.UpdateGroupUI"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBarsForOrder
'' Description: Update the bars for the order at the given index
'' Inputs:      Index
'' Returns:     True if Bars Updated, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UpdateBarsForOrder(ByVal lIndex As Long) As Boolean
On Error GoTo ErrSection:

    Dim bNewBar As Boolean              ' Did the order's bars hit a new bar?
    Dim bExpire As Boolean              ' Did the order expire?
    Dim bNewSession As Boolean          ' Are we in a new session?
    Dim bReturn As Boolean              ' Return value for the function
    Dim OldOrder As cPtOrder            ' Order object
    
    bReturn = False
    If Status(lIndex) = eGD_TsoStatus_Open Then
        With m.tsOrderGrp.Order(lIndex)
            bReturn = .UpdateBars(bNewBar, bExpire, bNewSession)
            If bReturn Then
                If bExpire = False Then
                    If bNewBar Then
                        If .ExpireDay = True Then
                            DumpDebug "Updated bars for order " & Str(.OrderNumber) & " (" & .EnglishPreview & "), Expriration = " & DateFormat(.Expiration, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", Feed Time = " & DateFormat(g.RealTime.FeedTime(m.strSymbol), MM_DD_YY, HH_MM_SS, AMPM_UPPER)
                        Else
                            DumpDebug "Updated bars for order " & Str(.OrderNumber) & " (" & .EnglishPreview & ")"
                        End If
                        
                        If (bNewSession = True) And (Not m.Orders(lIndex) Is Nothing) Then
                            ' DAJ 04/26/2013: Only mark the order as expired if the session date of
                            ' the order is less than the current session date...
                            If m.Orders(lIndex).SessionDate < .SessionDate Then
                                DumpDebug vbTab & "Order(" & Str(lIndex) & ") set to expired: Current Session = " & DateFormat(.SessionDate) & "; Order Session = " & DateFormat(m.Orders(lIndex).SessionDate)
                                
                                Status(lIndex) = eGD_TsoStatus_Open
                                DisplayStatus(lIndex) = "Working Order Expired"
                                Action(lIndex) = ""
                                Set OldOrder = m.Orders(lIndex)
                                Set m.Orders(lIndex) = Nothing
                                DumpDebug vbTab & vbTab & "Order(" & Str(lIndex) & ") set to Nothing"
                                
                                mTradeTracker.CancelOrder OldOrder, False
                            Else
                                DumpDebug vbTab & "Order(" & Str(lIndex) & ") not set to expired: Current Session = " & DateFormat(.SessionDate) & "; Order Session = " & DateFormat(m.Orders(lIndex).SessionDate)
                            End If
                        End If
                    End If
                    
                    CheckOrder lIndex
                End If
                
            ' If there isn't any new data, but we are still waiting for data to come in, if the
            ' data is now available, check the order...
            ElseIf m.bRtDataAvailable = False Then
                If AllRtDataAvailable Then
                    CheckOrder lIndex
                End If
                
            ' If we were not able to submit the order because the market was closed, but we
            ' are now in valid trading hours, call check order again to see if we can submit now...
            ElseIf (UCase(DisplayStatus(lIndex)) = "MARKET CLOSED") And (ValidTradingTime = True) Then
                If AllRtDataAvailable Then
                    CheckOrder lIndex
                End If
            End If
            
            If bExpire = True Then
                CancelOrder lIndex, "TradeSense order expired", False
            End If
        End With
    End If
    
    UpdateBarsForOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.UpdateBarsForOrder"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckOrder
'' Description: Check the order condition and act appropriately
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckOrder(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    Dim bCondition As Boolean           ' Condition true?

    Set tsOrder = m.tsOrderGrp.Order(lIndex)
    If Not tsOrder Is Nothing Then
        If tsOrder.Use = True Then
            If Status(lIndex) = eGD_TsoStatus_Open Then
                If AllRtDataAvailable Then
                    If DisplayStatus(lIndex) = "Waiting for Data" Then
                        DisplayStatus(lIndex) = "Open"
                        UpdateOrderUI lIndex
                    End If
                            
                    bCondition = tsOrder.ConditionTrue
                    If (OnlyExitsWorking = True) And (m.lPosition = 0) Then
                        CancelOrder lIndex, "only exits working", False
                    
                    ElseIf bCondition Then
                        If (tsOrder.Entry = False) And (m.lPosition = 0) Then
                            DisplayStatus(lIndex) = "No Position to Exit"
                            UpdateGroupUI
                        
                            If Not m.Orders(lIndex) Is Nothing Then
                                CancelOrder lIndex, "not in a position", True
                            End If
                        ElseIf (tsOrder.Entry = False) And (m.lPosition > 0) And (tsOrder.Buy = True) Then
                            DisplayStatus(lIndex) = "Exits Other Side"
                            UpdateGroupUI
                        
                            If Not m.Orders(lIndex) Is Nothing Then
                                CancelOrder lIndex, "exits other side", True
                            End If
                        ElseIf (tsOrder.Entry = False) And (m.lPosition < 0) And (tsOrder.Buy = False) Then
                            DisplayStatus(lIndex) = "Exits Other Side"
                            UpdateGroupUI
                        
                            If Not m.Orders(lIndex) Is Nothing Then
                                CancelOrder lIndex, "exits other side", True
                            End If
                        ElseIf (tsOrder.Entry = False) And (m.lMarketExit <> kNullData) And (lIndex <> m.lMarketExit) Then
                            DisplayStatus(lIndex) = "Market Exit Submitted"
                            UpdateGroupUI
                        
                            If Not m.Orders(lIndex) Is Nothing Then
                                CancelOrder lIndex, "market exit submitted", True
                            End If
                        ElseIf ValidTradingTime = False Then
                            DisplayStatus(lIndex) = "Market Closed"
                            UpdateGroupUI
                        ElseIf g.Broker.ConnectionStatusForAccount(m.lAccountID) <> eGDConnectionStatus_Connected Then
                            DisplayStatus(lIndex) = "Not connected to account"
                            UpdateGroupUI
                        Else
                            SubmitOrder lIndex
                        End If
                    ElseIf UCase(DisplayStatus(lIndex)) <> UCase("Condition False") Then
                        DisplayStatus(lIndex) = "Condition False"
                        UpdateOrderUI lIndex
                        
                        DumpDebug vbTab & vbTab & "Condition False ( " & tsOrder.ConditionEnglish & " ) ; Current Market = '" & tsOrder.CurrentPrices & "'"
                        
                        If Not m.Orders(lIndex) Is Nothing Then
                            CancelOrder lIndex, "condition no longer true", True
                        End If
                    End If
                Else
                    DisplayStatus(lIndex) = "Waiting for Data"
                    UpdateOrderUI lIndex
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CheckOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillIds
'' Description: Build a delimited string of Fill ID's
'' Inputs:      None
'' Returns:     String of Fill ID's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillIds() As String
On Error GoTo ErrSection:

    FillIds = m.astrFillIds.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.FillIds"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFills
'' Description: Load the fills collection
'' Inputs:      Fill ID's
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFills()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    
    For lIndex = 0 To m.astrFillIds.Size - 1
        Set Fill = New cPtFill
        If Fill.Load(CLng(Val(m.astrFillIds(lIndex)))) Then
            DumpDebug "Fill " & Str(Fill.FillID) & " loaded ('" & mTradeTracker.FillDisplay(Fill) & "')"
            UpdateFill Fill, True, False
        End If
    Next lIndex
    
    CalculatePosition

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.LoadFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculatePosition
'' Description: Calculate the position for this TradeSense order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculatePosition()
On Error GoTo ErrSection:

    Dim lPosition As Long               ' Position for the group
    Dim lIndex As Long                  ' Index into a for loop
    
    lPosition = 0&
    For lIndex = 1 To m.Fills.Count
        If m.Fills(lIndex).Buy Then
            lPosition = lPosition + m.Fills(lIndex).Quantity
        Else
            lPosition = lPosition - m.Fills(lIndex).Quantity
        End If
    Next lIndex
    
    m.lPosition = lPosition / m.lLotSize
    DumpDebug "Position calculated: " & Str(m.lPosition)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.CalculatePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderQuantity
'' Description: Calculate the quantity for the given order
'' Inputs:      TradeSense Order
'' Returns:     Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderQuantity(ByVal tsOrder As cTradeSenseOrder) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lQuantity As Long               ' Quantity
    Dim lPosition As Long               ' Position

    lQuantity = m.lQuantity * m.lLotSize
    lPosition = m.lPosition * m.lLotSize

    If m.lPosition = 0 Then
        If tsOrder.Entry Then
            lReturn = lQuantity
        Else
            lReturn = 0&
        End If
    ElseIf m.lPosition > 0 Then
        ' Long Entry...
        If (tsOrder.Buy = True) And (tsOrder.Entry = True) Then
            lReturn = lQuantity
        
        ' Long Exit...
        ElseIf (tsOrder.Buy = False) And (tsOrder.Entry = False) Then
            lReturn = (Int(m.lPosition * (tsOrder.ExitPercent / 100) + 0.99999)) * m.lLotSize
        
        ' Short Entry...
        ElseIf (tsOrder.Buy = False) And (tsOrder.Entry = True) Then
            lReturn = lQuantity + lPosition
        
        ' Short Exit...
        Else
            lReturn = 0&
        End If
    Else
        ' Long Entry...
        If (tsOrder.Buy = True) And (tsOrder.Entry = True) Then
            lReturn = lQuantity + Abs(lPosition)
        
        ' Long Exit...
        ElseIf (tsOrder.Buy = False) And (tsOrder.Entry = False) Then
            lReturn = 0&
        
        ' Short Entry...
        ElseIf (tsOrder.Buy = False) And (tsOrder.Entry = True) Then
            lReturn = lQuantity
        
        ' Short Exit...
        Else
            lReturn = (Int(Abs(m.lPosition) * (tsOrder.ExitPercent / 100) + 0.99999)) * m.lLotSize
        End If
    End If
    
    OrderQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.OrderQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIds
'' Description: Return a list of ID's from the order collection
'' Inputs:      None
'' Returns:     Order ID List
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderIds() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Order ID list to return
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, m.Orders.Count
    
    For lIndex = 1 To m.Orders.Count
        If m.Orders(lIndex) Is Nothing Then
            astrReturn(lIndex) = "0"
        ElseIf IsOpenOrder(m.Orders(lIndex).Status) And (m.Orders(lIndex).Status <> eTT_OrderStatus_Parked) Then
            astrReturn(lIndex) = Str(m.Orders(lIndex).OrderID)
        Else
            astrReturn(lIndex) = "0"
        End If
    Next lIndex
    
    OrderIds = astrReturn.JoinFields(",")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.OrderIds"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OnlyExitsWorking
'' Description: Are all of the working orders exits?
'' Inputs:      None
'' Returns:     True if only exits are working, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OnlyExitsWorking() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    For lIndex = 1 To m.Orders.Count
        If m.anStatus(lIndex) = eGD_TsoStatus_Open Then
            If m.tsOrderGrp.Order(lIndex).Entry = True Then
                bReturn = False
                Exit For
            End If
        End If
    Next lIndex
    
    OnlyExitsWorking = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.OnlyExitsWorking"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidTradingTime
'' Description: Is this a valid trading time for the symbol?
'' Inputs:      None
'' Returns:     True if valid trading time, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidTradingTime() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strTradeSymbol As String        ' Trade symbol
    Dim dCurrentTime As Double          ' Current time
    
    bReturn = True
    strTradeSymbol = TradeSymbol
    
    If (IsForex(strTradeSymbol) = True) And (InStr(strTradeSymbol, "@") <> 0) Then
        bReturn = mDataNav.ValidForexTradingTime(strTradeSymbol, CurrentTime("NY", strTradeSymbol, True))
    Else
        dCurrentTime = CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True)
        
        Select Case Weekday(dCurrentTime)
            Case vbFriday
                bReturn = (dCurrentTime <= (Val(Int(dCurrentTime)) + (m.Bars.Prop(eBARS_EndTime) / 1440#)))
            Case vbSunday
                bReturn = (dCurrentTime >= (Val(Int(dCurrentTime)) + (m.Bars.Prop(eBARS_StartTime) / 1440#)))
            Case vbSaturday
                bReturn = False
            Case Else
                bReturn = True
        End Select
    End If
        
    ValidTradingTime = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ValidTradingTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRunTime
'' Description: Is this a valid time to run this order group?
'' Inputs:      None
'' Returns:     True if valid run time, False otherwise
''
'' Note:        PFG Forex orders cannot be cancelled between 5pm Chicago time
''              on Friday and 3:45pm Chicago time on Sunday
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRunTime() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strTradeSymbol As String        ' Trade symbol
    Dim dCurrentTime As Double          ' Current time in Eastern time
    
    bReturn = True
    strTradeSymbol = TradeSymbol
    
    If IsForex(strTradeSymbol) Then
        If InStr(strTradeSymbol, "@") <> 0 Then
            dCurrentTime = CurrentTime("NY", strTradeSymbol, True)
            
            Select Case Weekday(dCurrentTime)
                Case vbFriday, vbSunday
                    bReturn = mDataNav.ValidForexTradingTime(strTradeSymbol, dCurrentTime)
                Case vbSaturday
                    bReturn = False
                Case Else
                    bReturn = True
            End Select
        End If
    End If
    
    ValidRunTime = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ValidRunTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StatusText
'' Description: Return the status as an english word
'' Inputs:      Index
'' Returns:     Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StatusText(ByVal lIndex As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = Str(lIndex)
    Select Case Status(lIndex)
        Case eGD_TsoStatus_Open
            strReturn = "Open"
        Case eGD_TsoStatus_Trigger
            strReturn = "Trigger"
        Case eGD_TsoStatus_Closed
            strReturn = "Closed"
        Case eGD_TsoStatus_Parked
            strReturn = "Parked"
    End Select
    
    StatusText = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.StatusText"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGenesisForex
'' Description: Is the current symbol a Genesis Forex symbol?
'' Inputs:      None
'' Returns:     True if Genesis Forex, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsGenesisForex() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If IsForex(m.strSymbol) Then
        If InStr(m.strSymbol, "@") = 0 Then
            bReturn = True
        End If
    End If
    
    IsGenesisForex = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.IsGenesisForex"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DetermineTradeSymbol
'' Description: Determine the trade symbol
'' Inputs:      None
'' Returns:     Trade Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DetermineTradeSymbol(ByVal bUseSessionActivated As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If (IsGenesisForex = True) And g.Broker.IsPfgBroker(m.nBroker) Then
        strReturn = m.strSymbol & "@PFG"
    ElseIf (IsGenesisForex = True) And g.Broker.IsIbBroker(m.nBroker) Then
        strReturn = m.strSymbol & "@IB"
    ElseIf (IsGenesisForex = True) And g.Broker.IsCurrenexBroker(m.nBroker) Then
        strReturn = m.strSymbol & "@CNX"
    ElseIf bUseSessionActivated Then
        strReturn = ConvertToTradeSymbol(m.strSymbol, SessionActivated)
    Else
        strReturn = ConvertToTradeSymbol(m.strSymbol, m.Bars.SessionDateForTradeTime(CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), m.strSymbol, True)))
    End If
    
    DetermineTradeSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.DetermineTradeSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UseGtcOrder
'' Description: Determine if we want to submit orders as GTC
'' Inputs:      Symbol
'' Returns:     True if GTC, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UseGtcOrder(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If InStr(strSymbol, "@PFG") <> 0 Then
        bReturn = True
    ElseIf InStr(strSymbol, "@IB") <> 0 Then
        bReturn = True
    ElseIf InStr(strSymbol, "@CNX") <> 0 Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    UseGtcOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.UseGtcOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnterPosition
'' Description: Enter a position
'' Inputs:      New Position, Reason, Symbol, Rolling?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnterPosition(ByVal lNewPosition As Long, ByVal strReason As String, Optional ByVal strSymbol As String = "", Optional ByVal bRolling As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order to close the position
    
    If m.lPosition = 0& Then
        DumpDebug "Entering Position because " & strReason
        
        If Len(strSymbol) = 0 Then
            strSymbol = TradeSymbol
        End If
        
        Set Order = CreateMarketOrder(m.lAccountID, strSymbol, (lNewPosition > 0), Abs(lNewPosition * m.lLotSize))
        
        Order.Enter = True
        Order.Save
        
        If bRolling = True Then
            Set m.RollOrder = Order
        End If
        
        g.Broker.SendOrder Order
                
        If g.Broker.AutoJournalAutomated Then
            g.JournalBridge.CreateOrderJournal Order.OrderID, CurrentTime, "TradeSense Order Group (" & m.tsOrderGrp.Name & ") closing position"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ClosePosition"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollingPhaseToString
'' Description: Convert the given rolling phase to a string for logging purposes
'' Inputs:      Rolling Phase
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RollingPhaseToString(ByVal nRollingPhase As eGDRollPhases) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nRollingPhase
        Case eGDRollPhase_None
            strReturn = "eGDRollPhase_None"
        Case eGDRollPhase_Parking
            strReturn = "eGDRollPhase_Parking"
        Case eGDRollPhase_Flattening
            strReturn = "eGDRollPhase_Flattening"
        Case eGDRollPhase_Entering
            strReturn = "eGDRollPhase_Entering"
        Case Else
            strReturn = Str(nRollingPhase)
    End Select
    
    RollingPhaseToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.RollingPhaseToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollSymbol
'' Description: Roll the symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RollSymbol()
On Error GoTo ErrSection:

    Dim bReactivate As Boolean          ' Reactivate the TradeSense order group?
    Dim bCheckParked As Boolean         ' Check to see if all is parked?

    bReactivate = False
    bCheckParked = False

    Select Case RollingPhase
        Case eGDRollPhase_None
            If TradeSymbol <> NewTradeSymbol Then
                PositionBeforeRoll = m.lPosition
                bCheckParked = True
            End If
        
        Case eGDRollPhase_Parking
            bCheckParked = True
        
        Case eGDRollPhase_Flattening
            If m.lPosition = 0& Then
                RollingPhase = eGDRollPhase_Entering
                EnterPosition PositionBeforeRoll, "Rolling", NewTradeSymbol, True
            End If
        
        Case eGDRollPhase_Entering
            If m.lPosition = PositionBeforeRoll Then
                bReactivate = True
            End If
    
    End Select

    If bCheckParked = True Then
        If RollingPhase = eGDRollPhase_None Then
            RollingPhase = eGDRollPhase_Parking
            ParkAll "Rolling"
        End If
    
        If AllParked Then
            If m.lPosition = 0& Then
                bReactivate = True
            Else
                RollingPhase = eGDRollPhase_Flattening
                ClosePosition "Rolling", True
            End If
        End If
    End If

    If bReactivate Then
        RollingPhase = eGDRollPhase_None
        PositionBeforeRoll = kNullData
        Set m.RollOrder = Nothing
        TradeSymbol = NewTradeSymbol
        
        Activate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.RollSymbol"
    
End Sub

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
    
    strToDump = LogText & strText

    g.Broker.BrokerDebug m.nBroker, strToDump
    frmTest2.AddList strToDump
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.bRtDataAvailable = False

    Set m.tsOrderGrp = New cTradeSenseOrderGroup
    Set m.Orders = New cGdTree
    Set m.anStatus = New cGdArray
    Set m.astrToDo = New cGdArray
    Set m.anLeaveOpen = New cGdArray
    Set m.astrStatus = New cGdArray
    Set m.alLevel = New cGdArray
    Set m.Fills = New cGdTree
    Set m.astrFillIds = New cGdArray
    m.astrFillIds.Create eGDARRAY_Strings
    Set m.alOrderIds = New cGdArray
    Set m.astrAction = New cGdArray
    m.astrAction.Create eGDARRAY_Strings
    
    m.lMarketExit = kNullData
    m.lSessionActivated = kNullData
    m.bFlattening = False
    m.dCustomStartTime = kNullData
    m.dCustomStopTime = kNullData
    m.strTradeSymbol = ""
    'm.strAutoBreakoutPeriod = ""
    m.lLotSize = 1&
    m.bActivationError = False
    m.nRollingPhase = eGDRollPhase_None
    
    Set m.Bars = New cGdBars
    Set m.RollOrder = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.tsOrderGrp = Nothing
    Set m.Orders = Nothing
    Set m.anStatus = Nothing
    Set m.astrToDo = Nothing
    Set m.anLeaveOpen = Nothing
    Set m.astrStatus = Nothing
    Set m.alLevel = Nothing
    Set m.Fills = Nothing
    Set m.astrFillIds = Nothing
    Set m.alOrderIds = Nothing
    Set m.astrAction = Nothing
    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.Class_Terminate"

End Sub
