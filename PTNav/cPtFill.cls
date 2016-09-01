VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtFill.cls
'' Description: Wrapper class for fills
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/19/2009   DAJ         Fixed Option Navigator string when there is no order
'' 03/27/2009   DAJ         On Save, include reason for moving fill
'' 05/19/2009   DAJ         Added null checks for broker order ID and broker fill ID
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 12/04/2009   DAJ         Changed automatic commissions for stocks
'' 01/21/2010   DAJ         Changed AutoTradeItem to AutoTradingItem in Save
'' 09/16/2010   DAJ         More information in e-mail order alerts (#5891)
'' 01/13/2011   DAJ         Added the ConvertedFillDate function
'' 03/07/2011   DAJ         Add more fields for Option Nav
'' 03/11/2011   DAJ         Removed redundant fill ID field in message to Option Nav
'' 07/21/2011   DAJ         Enhanced FillText routine
'' 10/07/2011   DAJ         Added a previous broker fill ID property
'' 06/26/2012   DAJ         Send Expiration Date/Lot Size to Option Nav
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 04/16/2013   DAJ         Log when auto trade id changes
'' 07/30/2013   DAJ         Enhanced the FillText routine for automatic journal entries
'' 07/10/2014   DAJ         Enhanced the FillText routine for display purposes
'' 09/04/2014   DAJ         Pulled Option Navigator, SimTrade conversions out of trade objects
'' 10/24/2014   DAJ         Fill Display
'' 12/10/2014   DAJ         Utilize new DateIsSnapshot routines
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lFillID As Long                     ' Fill ID from the database table
    lOrderID As Long                    ' Order ID associated with the fill
    lAccountID As Long                  ' Account ID
    nBroker As eTT_AccountType          ' Broker for the account
    lSymbolID As Long                   ' Order symbol ID
    strSymbol As String                 ' Order symbol
    lAutoTradeItemID As Long            ' Automated Trading Item ID for the order
    strBrokerID As String               ' Broker ID for the fill
    strBrokerOrderID As String          ' Broker Order ID for the fill
    bBuy As Boolean                     ' Buy or Sell?
    lQuantity As Long                   ' Quantity of the fill
    dPrice As Double                    ' Price of the fill
    dFillDate As Date                   ' Date the fill occurred
    lSessionDate As Long                ' Session date for the fill
    dFees As Double                     ' Fees associated with the fill
    strNotes As String                  ' Notes for the fill
    bIsManual As Boolean                ' Is this a manually entered fill?
    bIsSnapshot As Boolean              ' Is this a snapshot fill?
    strPreviousBrokerID As String       ' Previous broker ID for the fill
    bIsSpread As Boolean                ' Is this a spread fill?
    
    ' Calculated numbers for the "total" account position...
    lSequence As Long                   ' Sequence of the fill
    cAction As String                   ' Action of the fill (E=Entry, X=Exit, R=Reverse)
    dClosedProfit As Double             ' Closed out profit for this fill
    dSessionProfit As Double            ' Profit/Loss so far for this session
    dTotalProfit As Double              ' Profit/Loss so far total
    lSessionQuantity As Long            ' Number contracts traded so far this session
    lCurrentPosition As Long            ' Current position after this fill
    dAverageEntry As Double             ' Average entry of position including this fill
    lNumOpen As Long                    ' Number of contracts in this fill still open
    
    ' Calculated numbers for the "category" account position...
    lSequenceCategory As Long           ' Sequence of the fill
    cActionCategory As String           ' Action of the fill (E=Entry, X=Exit, R=Reverse)
    dClosedProfitCategory As Double     ' Closed out profit for this fill
    dSessionProfitCategory As Double    ' Profit/Loss so far for this session
    dTotalProfitCategory As Double      ' Profit/Loss so far total
    lSessionQuantityCategory As Long    ' Number contracts traded so far this session
    lCurrentPositionCategory As Long    ' Current position after this fill
    dAverageEntryCategory As Double     ' Average entry of position including this fill
    lNumOpenCategory As Long            ' Number of contracts in this fill still open
    lEntryRuleIdCategory As Long        ' What is the entry rule ID for this fill?
    lExitRuleIdCategory As Long         ' What is the exit rule ID for this fill?
    nRealSimFlagCategory As Byte        ' Real vs. Sim flag
    
    Bars As cGdBars                     ' Bars object
    
    bRefreshed As Boolean               ' Has this fill been refreshed?
End Type
Private m As mPrivate

Public Property Get FillID() As Long
    FillID = m.lFillID
End Property
Public Property Let FillID(ByVal pData As Long)
    m.lFillID = pData
End Property

Public Property Get OrderID() As Long
    OrderID = m.lOrderID
End Property
Public Property Let OrderID(ByVal pData As Long)
    m.lOrderID = pData
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal pData As Long)
    m.lSymbolID = pData
    SetBarProperties m.Bars, m.lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal pData As String)
    m.strSymbol = pData
    SetBarProperties m.Bars, m.strSymbol
End Property

Public Property Get AutoTradingItemID() As Long
    AutoTradingItemID = m.lAutoTradeItemID
End Property
Public Property Let AutoTradingItemID(ByVal lAutoTradingItemID As Long)
    m.lAutoTradeItemID = lAutoTradingItemID
End Property

Public Property Get BrokerID() As String
    BrokerID = m.strBrokerID
End Property
Public Property Let BrokerID(ByVal pData As String)
    m.strBrokerID = pData
End Property

Public Property Get PreviousBrokerID() As String
    PreviousBrokerID = m.strPreviousBrokerID
End Property
Public Property Let PreviousBrokerID(ByVal strPreviousBrokerID As String)
    m.strPreviousBrokerID = strPreviousBrokerID
End Property

Public Property Get BrokerOrderID() As String
    BrokerOrderID = m.strBrokerOrderID
End Property
Public Property Let BrokerOrderID(ByVal pData As String)
    m.strBrokerOrderID = pData
End Property

Public Property Get Buy() As Boolean
    Buy = m.bBuy
End Property
Public Property Let Buy(ByVal pData As Boolean)
    m.bBuy = pData
End Property

Public Property Get Quantity() As Long
    Quantity = m.lQuantity
End Property
Public Property Let Quantity(ByVal pData As Long)
    m.lQuantity = pData
End Property

Public Property Get Price() As Double
    Price = m.dPrice
End Property
Public Property Get PriceString() As String
    If m.dPrice = 0# Then
        PriceString = ""
    Else
        PriceString = m.Bars.PriceDisplay(m.dPrice)
    End If
End Property
Public Property Let Price(ByVal pData As Double)
    m.dPrice = pData
End Property

Public Property Get FillDate() As Date
    FillDate = m.dFillDate
End Property
Public Property Get FillDateConverted(ByVal strTimeZoneInfo As String) As Date
    FillDateConverted = ConvertBrokerDate(m.dFillDate, m.nBroker, m.strSymbol, False, strTimeZoneInfo)
End Property
Public Property Let FillDate(ByVal pData As Date)
    m.dFillDate = pData
End Property

Public Property Get SessionDate() As Long
    SessionDate = m.lSessionDate
End Property
Public Property Let SessionDate(ByVal lSessionDate As Long)
    m.lSessionDate = lSessionDate
End Property

Public Property Get Fees(Optional ByVal bZeroIfNull As Boolean = True) As Double
    If (bZeroIfNull = True) And (m.dFees = kNullData) Then
        Fees = 0#
    Else
        Fees = m.dFees
    End If
End Property
Public Property Let Fees(Optional ByVal bZeroIfNull As Boolean = False, ByVal dFees As Double)
    If (bZeroIfNull = True) And (dFees = kNullData) Then
        m.dFees = 0#
    Else
        m.dFees = dFees
    End If
End Property

Public Property Get Notes() As String
    Notes = m.strNotes
End Property
Public Property Let Notes(ByVal strNotes As String)
    m.strNotes = strNotes
End Property

Public Property Get IsManual() As Boolean
    IsManual = m.bIsManual
End Property
Public Property Let IsManual(ByVal bIsManual As Boolean)
    m.bIsManual = bIsManual
End Property

Public Property Get IsSnapshot() As Boolean
    IsSnapshot = m.bIsSnapshot
End Property
Public Property Let IsSnapshot(ByVal bIsSnapshot As Boolean)
    m.bIsSnapshot = bIsSnapshot
End Property

Public Property Get IsSpread() As Boolean
    IsSpread = m.bIsSpread
End Property
Public Property Let IsSpread(ByVal bIsSpread As Boolean)
    m.bIsSpread = bIsSpread
End Property

Public Property Get Sequence() As Long
    Sequence = m.lSequence
End Property
Public Property Let Sequence(ByVal lSequence As Long)
    m.lSequence = lSequence
End Property

Public Property Get Action() As String
    Action = m.cAction
End Property
Public Property Let Action(ByVal pData As String)
    m.cAction = pData
End Property

Public Property Get ClosedProfit() As Double
    ClosedProfit = m.dClosedProfit
End Property
Public Property Let ClosedProfit(ByVal dClosedProfit As Double)
    m.dClosedProfit = dClosedProfit
End Property

Public Property Get SessionProfit() As Double
    SessionProfit = m.dSessionProfit
End Property
Public Property Let SessionProfit(ByVal dSessionProfit As Double)
    m.dSessionProfit = dSessionProfit
End Property

Public Property Get TotalProfit() As Double
    TotalProfit = m.dTotalProfit
End Property
Public Property Let TotalProfit(ByVal dTotalProfit As Double)
    m.dTotalProfit = dTotalProfit
End Property

Public Property Get SessionQuantity() As Long
    SessionQuantity = m.lSessionQuantity
End Property
Public Property Let SessionQuantity(ByVal lSessionQuantity As Long)
    m.lSessionQuantity = lSessionQuantity
End Property

Public Property Get CurrentPosition() As Long
    CurrentPosition = m.lCurrentPosition
End Property
Public Property Let CurrentPosition(ByVal lCurrentPosition As Long)
    m.lCurrentPosition = lCurrentPosition
End Property
Public Property Get CurrentPositionString() As String
    CurrentPositionString = g.Broker.TextPosition(m.lCurrentPosition)
End Property

Public Property Get AverageEntry() As Double
    AverageEntry = m.dAverageEntry
End Property
Public Property Get AverageEntryString() As String
    If m.dAverageEntry = 0# Then
        AverageEntryString = ""
    Else
        AverageEntryString = m.Bars.PriceDisplay(m.dAverageEntry)
    End If
End Property
Public Property Let AverageEntry(ByVal dAverageEntry As Double)
    m.dAverageEntry = dAverageEntry
End Property

Public Property Get NumOpen() As Long
    NumOpen = m.lNumOpen
End Property
Public Property Let NumOpen(ByVal lNumOpen As Long)
    m.lNumOpen = lNumOpen
End Property

Public Property Get SequenceCategory() As Long
    SequenceCategory = m.lSequenceCategory
End Property
Public Property Let SequenceCategory(ByVal lSequenceCategory As Long)
    m.lSequenceCategory = lSequenceCategory
End Property

Public Property Get ActionCategory() As String
    ActionCategory = m.cActionCategory
End Property
Public Property Let ActionCategory(ByVal pData As String)
    m.cActionCategory = pData
End Property

Public Property Get ClosedProfitCategory() As Double
    ClosedProfitCategory = m.dClosedProfitCategory
End Property
Public Property Let ClosedProfitCategory(ByVal dClosedProfitCategory As Double)
    m.dClosedProfitCategory = dClosedProfitCategory
End Property

Public Property Get SessionProfitCategory() As Double
    SessionProfitCategory = m.dSessionProfitCategory
End Property
Public Property Let SessionProfitCategory(ByVal dSessionProfitCategory As Double)
    m.dSessionProfitCategory = dSessionProfitCategory
End Property

Public Property Get TotalProfitCategory() As Double
    TotalProfitCategory = m.dTotalProfitCategory
End Property
Public Property Let TotalProfitCategory(ByVal dTotalProfitCategory As Double)
    m.dTotalProfitCategory = dTotalProfitCategory
End Property

Public Property Get SessionQuantityCategory() As Long
    SessionQuantityCategory = m.lSessionQuantityCategory
End Property
Public Property Let SessionQuantityCategory(ByVal lSessionQuantityCategory As Long)
    m.lSessionQuantityCategory = lSessionQuantityCategory
End Property

Public Property Get CurrentPositionCategory() As Long
    CurrentPositionCategory = m.lCurrentPositionCategory
End Property
Public Property Let CurrentPositionCategory(ByVal lCurrentPositionCategory As Long)
    m.lCurrentPositionCategory = lCurrentPositionCategory
End Property
Public Property Get CurrentPositionCategoryString() As String
    CurrentPositionCategoryString = g.Broker.TextPosition(m.lCurrentPositionCategory)
End Property

Public Property Get AverageEntryCategory() As Double
    AverageEntryCategory = m.dAverageEntryCategory
End Property
Public Property Get AverageEntryCategoryString() As String
    If m.dAverageEntryCategory = 0# Then
        AverageEntryCategoryString = ""
    Else
        AverageEntryCategoryString = m.Bars.PriceDisplay(m.dAverageEntryCategory)
    End If
End Property
Public Property Let AverageEntryCategory(ByVal dAverageEntryCategory As Double)
    m.dAverageEntryCategory = dAverageEntryCategory
End Property

Public Property Get NumOpenCategory() As Long
    NumOpenCategory = m.lNumOpenCategory
End Property
Public Property Let NumOpenCategory(ByVal lNumOpenCategory As Long)
    m.lNumOpenCategory = lNumOpenCategory
End Property

Public Property Get EntryRuleIdCategory() As Long
    EntryRuleIdCategory = m.lEntryRuleIdCategory
End Property
Public Property Let EntryRuleIdCategory(ByVal lEntryRuleIdCategory As Long)
    m.lEntryRuleIdCategory = lEntryRuleIdCategory
End Property

Public Property Get ExitRuleIdCategory() As Long
    ExitRuleIdCategory = m.lExitRuleIdCategory
End Property
Public Property Let ExitRuleIdCategory(ByVal lExitRuleIdCategory As Long)
    m.lExitRuleIdCategory = lExitRuleIdCategory
End Property

Public Property Get RealSimFlagCategory() As Byte
    RealSimFlagCategory = m.nRealSimFlagCategory
End Property
Public Property Let RealSimFlagCategory(ByVal nRealSimFlagCategory As Byte)
    m.nRealSimFlagCategory = nRealSimFlagCategory
End Property

Public Property Get NumberOpen(ByVal lAutoTradeItemID As Long) As Long
    If lAutoTradeItemID = -1& Then
        NumberOpen = m.lNumOpen
    Else
        NumberOpen = m.lNumOpenCategory
    End If
End Property
Public Property Let NumberOpen(ByVal lAutoTradeItemID As Long, ByVal lNumOpen As Long)
    If lAutoTradeItemID = -1& Then
        m.lNumOpen = lNumOpen
    Else
        m.lNumOpenCategory = lNumOpen
    End If
End Property

Public Property Get SymbolOrSymbolID() As Variant
On Error GoTo ErrSection:

    If m.lSymbolID = 0 Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPtFill.SymbolOrSymbolID.Get"
    
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    SetBarProperties m.Bars, vSymbolOrSymbolID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cPtFill.SymbolOrSymbolID.Let"
    
End Property

Public Property Get Refreshed() As Boolean
    Refreshed = m.bRefreshed
End Property
Public Property Let Refreshed(ByVal bRefreshed As Boolean)
    m.bRefreshed = bRefreshed
End Property

Public Property Get Key() As String
    Key = Str(m.lAccountID) & vbTab & Str(SymbolOrSymbolID) & vbTab & Str(m.lAutoTradeItemID)
End Property

Public Property Get Bars() As cGdBars
    Set Bars = m.Bars
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the fill
'' Inputs:      Fill ID, Recordset
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lFillID As Long, Optional ByVal rs As Recordset = Nothing, Optional Bars As cGdBars = Nothing, Optional ByVal nBroker As eTT_AccountType = -1) As Boolean
On Error GoTo ErrSection
    
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE [FillID]=" & Str(lFillID) & ";", dbOpenDynaset)
    End If
    
    If Not rs.EOF Then
        m.lFillID = rs!FillID
        m.lOrderID = rs!OrderID
        m.lAccountID = rs!AccountID
        If nBroker = -1 Then
            m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
        Else
            m.nBroker = nBroker
        End If
        m.lSymbolID = rs!SymbolID
        m.strSymbol = rs!Symbol
        m.lAutoTradeItemID = rs!AutoTradeID
        
        m.strBrokerID = NullChk(rs!BrokerFillID, "")
        m.strBrokerOrderID = NullChk(rs!BrokerOrderID, "")
        m.bBuy = rs!BuySell
        m.lQuantity = rs!FillQuantity
        m.dPrice = rs!FillPrice
        m.dFillDate = rs!FillDate
        m.lSessionDate = rs!SessionDate
        m.dFees = rs!Fees
        m.strNotes = rs!Notes
        m.bIsManual = rs!IsManual
        m.bIsSnapshot = rs!IsSnapshot
        m.bIsSpread = rs!IsSpread
        
        m.lSequence = rs!Sequence
        m.cAction = rs!Action
        m.dClosedProfit = rs!ClosedProfit
        m.dSessionProfit = rs!SessionProfit
        m.dTotalProfit = rs!TotalProfit
        m.lSessionQuantity = rs!SessionQuantity
        m.lCurrentPosition = rs!CurrentPosition
        m.dAverageEntry = rs!AverageEntry
        m.lNumOpen = rs!NumOpen
        
        m.lSequenceCategory = rs!SequenceCategory
        m.cActionCategory = rs!ActionCategory
        m.dClosedProfitCategory = rs!ClosedProfitCategory
        m.dSessionProfitCategory = rs!SessionProfitCategory
        m.dTotalProfitCategory = rs!TotalProfitCategory
        m.lSessionQuantityCategory = rs!SessionQuantityCategory
        m.lCurrentPositionCategory = rs!CurrentPositionCategory
        m.dAverageEntryCategory = rs!AverageEntryCategory
        m.lNumOpenCategory = rs!NumOpenCategory
        m.lEntryRuleIdCategory = rs!EntryRuleIdCategory
        m.lExitRuleIdCategory = rs!ExitRuleIdCategory
        m.nRealSimFlagCategory = rs!RealFlagCategory
    End If

    If InStr(m.strSymbol, "-0") <> 0 Then
        m.strSymbol = RollSymbolForDate(m.strSymbol, m.dFillDate)
        m.lSymbolID = g.SymbolPool.SymbolIDforSymbol(m.strSymbol)
    End If

    If Bars Is Nothing Then
        SetBarProperties m.Bars, SymbolOrSymbolID
    Else
        Set m.Bars = Bars
    End If

    Load = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFill.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the fill
'' Inputs:      Update Broker Info?
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save(Optional ByVal bUpdateBrokerInfo As Boolean = True, Optional ByVal strReasonMoved As String = "") As Boolean
On Error GoTo ErrSection

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAcct As Long                   ' Account ID
    Dim lOldAutoTradeID As Long         ' Old automated trading item ID
    Dim Order As cPtOrder               ' Order object
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Create the fill if it does not exist, otherwise update the info...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE [FillID] = " & Str(m.lFillID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        rs.Edit
        If (rs!IsSnapshot = False) And (m.bIsSnapshot = True) Then
            If Len(strReasonMoved) > 0 Then
                g.Broker.BrokerDebug m.nBroker, "Fill " & m.strBrokerID & " changing from history to snapshot (" & strReasonMoved & ")"
            Else
                g.Broker.BrokerDebug m.nBroker, "Fill " & m.strBrokerID & " changing from history to snapshot"
            End If
        ElseIf (rs!IsSnapshot = True) And (m.bIsSnapshot = False) Then
            If Len(strReasonMoved) > 0 Then
                g.Broker.BrokerDebug m.nBroker, "Fill " & m.strBrokerID & " changing from snapshot to history (" & strReasonMoved & ")"
            Else
                g.Broker.BrokerDebug m.nBroker, "Fill " & m.strBrokerID & " changing from snapshot to history"
            End If
        End If
        
        lOldAutoTradeID = rs!AutoTradeID
    Else
        rs.AddNew
        lOldAutoTradeID = m.lAutoTradeItemID
        If SecurityType(SymbolOrSymbolID, True) = "F" Then
            m.dFees = g.Broker.AutomaticCommission(AccountID, SymbolOrSymbolID, Quantity)
        End If
    End If
    
    rs!OrderID = m.lOrderID
    rs!AccountID = m.lAccountID
    rs!SymbolID = m.lSymbolID
    rs!Symbol = m.strSymbol
    rs!AutoTradeID = m.lAutoTradeItemID
        
    rs!BrokerFillID = m.strBrokerID
    rs!BrokerOrderID = m.strBrokerOrderID
    rs!BuySell = m.bBuy
    rs!FillQuantity = m.lQuantity
    rs!FillPrice = m.dPrice
    rs!FillDate = m.dFillDate
    rs!SessionDate = m.lSessionDate
    rs!Fees = m.dFees
    rs!Notes = m.strNotes
    rs!IsManual = m.bIsManual
    rs!IsSnapshot = m.bIsSnapshot
    rs!IsSpread = m.bIsSpread
    
    rs!Sequence = m.lSequence
    rs!Action = m.cAction
    rs!ClosedProfit = m.dClosedProfit
    rs!SessionProfit = m.dSessionProfit
    rs!TotalProfit = m.dTotalProfit
    rs!SessionQuantity = m.lSessionQuantity
    rs!CurrentPosition = m.lCurrentPosition
    rs!AverageEntry = m.dAverageEntry
    rs!NumOpen = m.lNumOpen
    
    rs!SequenceCategory = m.lSequenceCategory
    rs!ActionCategory = m.cActionCategory
    rs!ClosedProfitCategory = m.dClosedProfitCategory
    rs!SessionProfitCategory = m.dSessionProfitCategory
    rs!TotalProfitCategory = m.dTotalProfitCategory
    rs!SessionQuantityCategory = m.lSessionQuantityCategory
    rs!CurrentPositionCategory = m.lCurrentPositionCategory
    rs!AverageEntryCategory = m.dAverageEntryCategory
    rs!NumOpenCategory = m.lNumOpenCategory
    rs!EntryRuleIdCategory = m.lEntryRuleIdCategory
    rs!ExitRuleIdCategory = m.lExitRuleIdCategory
    rs!RealFlagCategory = m.nRealSimFlagCategory
    
    m.lFillID = rs!FillID
    rs.Update
    
    If bUpdateBrokerInfo Then
        g.Broker.RefreshFill Me
    End If
    
    ' If the automated trading item ID changed on this fill, we need to change the order and
    ' the other fills for this order as well...
    If (lOldAutoTradeID <> m.lAutoTradeItemID) And (m.lOrderID > 0&) Then
        g.Broker.BrokerDebug m.nBroker, vbTab & vbTab & mTradeTracker.FillDisplay(Me, True, False, True, True, True, False, False) & ": Auto Trade Item ID changed from " & Str(lOldAutoTradeID) & " to " & Str(m.lAutoTradeItemID)
        
        Set Order = New cPtOrder
        If Order.Load(m.lOrderID) Then
            If Order.AutoTradeItemID <> m.lAutoTradeItemID Then
                Order.AutoTradeItemID = m.lAutoTradeItemID
                For lIndex = 1 To Order.Fills.Count
                    If Order.Fills(lIndex).AutoTradingItemID <> m.lAutoTradeItemID Then
                        Order.Fills(lIndex).AutoTradingItemID = m.lAutoTradeItemID
                    End If
                Next lIndex
                Order.Save
            End If
        End If
    End If
           
    Save = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFill.Save"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the fill from the database
'' Inputs:      Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete(Optional ByVal strReason As String = "")
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As cPtOrder               ' Order object
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE [FillID] = " & Str(m.lFillID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
        
        If Len(strReason) = 0 Then
            g.Broker.BrokerDebug m.nBroker, "Fill " & m.strBrokerID & ": '" & mTradeTracker.FillDisplay(Me) & "' has been deleted"
        Else
            g.Broker.BrokerDebug m.nBroker, "Fill " & m.strBrokerID & ": '" & mTradeTracker.FillDisplay(Me) & "' has been deleted (" & strReason & ")"
        End If
        
        ' Change the order status if necessary and update the order in memory...
        Set Order = New cPtOrder
        If Order.Load(m.lOrderID) Then
            Select Case Order.Status
                Case eTT_OrderStatus_Filled
                    If Order.FillQuantity = 0& Then
                        Order.Status = eTT_OrderStatus_Cancelled
                    ElseIf Order.FillQuantity < Order.Quantity Then
                        Order.Status = eTT_OrderStatus_BalCancelled
                    End If
                    
                Case eTT_OrderStatus_Partial
                    If Order.FillQuantity = 0& Then
                        Order.Status = eTT_OrderStatus_Cancelled
                    End If
                    
                Case eTT_OrderStatus_OverFilled
                    If Order.FillQuantity = Order.Quantity Then
                        Order.Status = eTT_OrderStatus_Filled
                    ElseIf Order.FillQuantity = 0& Then
                        Order.Status = eTT_OrderStatus_Cancelled
                    ElseIf Order.FillQuantity < Order.Quantity Then
                        Order.Status = eTT_OrderStatus_BalCancelled
                    End If
                    
                Case eTT_OrderStatus_BalCancelled
                    If Order.FillQuantity = 0 Then
                        Order.Status = eTT_OrderStatus_Cancelled
                    End If
                    
            End Select
            
            Order.Save
            
            ' Don't add the order to the broker info unless it is already a snapshot order, or it is
            ' now going to be open (2/7/2008 DAJ)...
            If (g.Broker.DateIsSnapshotForOrder(Order)) Or (IsOpenOrder(Order.Status) = True) Then
                g.Broker.AddOrder Order
            End If
            
            OrderCallback Order
        End If
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFill.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEntry
'' Description: Determine if the fill is part of an entry or an exit
'' Inputs:      Category?
'' Returns:     True if entry, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEntry(Optional ByVal bCategory As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    
    If bCategory Then
        bReturn = (m.cActionCategory = "E") Or (m.cActionCategory = "R")
    Else
        bReturn = (m.cAction = "E") Or (m.cAction = "R")
    End If
    
    IsEntry = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFill.IsEntry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenQuantity
'' Description: Determine the portion of the order that is still open
'' Inputs:      Cateogry?
'' Returns:     Quantity still open
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenQuantity(Optional ByVal bCategory As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    If bCategory Then
        lReturn = m.lNumOpenCategory
    Else
        lReturn = m.lNumOpen
    End If
    
    OpenQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFill.OpenQuantity"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the fill
'' Inputs:      None
'' Returns:     Copy of the fill
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cPtFill
On Error GoTo ErrSection:

    Dim aCopy As New cPtFill            ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtFill.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFill.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Intialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Bars = New cGdBars
    m.dFees = kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFill.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtFill.Class_Terminate"
    
End Sub
