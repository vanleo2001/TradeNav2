VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccountPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccountPosition.cls
'' Description: Wrapper class and handler for an account position
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 02/20/2009   DAJ         For Open Profit, change Null price to zero for options
'' 06/02/2009   DAJ         Use Bid/Ask instead of Last for Options P&L
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 02/18/2010   DAJ         Exposed the Entries collection
'' 05/17/2010   DAJ         Added calls for use in late calculating functions
'' 03/18/2011   DAJ         Utilize the CreateMarketOrder function
'' 12/09/2011   DAJ         Added expiration date
'' 11/28/2012   DAJ         Speed enhancements for the Trade Console
'' 03/22/2013   DAJ         Only build the sort key when one of the components changes
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
'' 05/28/2014   DAJ         If current price is null, call open profit null
'' 08/11/2014   DAJ         New flag for how to calculate open equity on options
'' 09/04/2014   DAJ         Pulled Option Navigator conversions out of trade objects
'' 02/13/2015   DAJ         Start with EXACT copy of entries for LIFO snapshot ( same NumberOpen )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lAccountPositionID As Long          ' Unique identifier of the account position
    lAccountID As Long                  ' ID of the associated account
    nBroker As eTT_AccountType          ' Broker for the account
    lSymbolID As Long                   ' ID of the associated symbol
    strSymbol As String                 ' Associated symbol
    lAutoTradeItemID As Long            ' ID of the associated automated trading item
    nFillMatchMode As eTT_FillMatchMode ' Fill match mode
    dStartFillTime As Double            ' Date/Time of the first fill
    lExpirationDate As Long             ' Expiration date for the symbol of the position
    
    lSessionDate As Long                ' Session Date of last trade
    lSessionQuantity As Long            ' Total Quantity traded in session date of last trade
    dSessionProfit As Double            ' Closed out Profit/Loss for session date of last trade
    lNumBuys As Long                    ' Number of non-snapshot buy lots
    lNumSells As Long                   ' Number of non-snapshot sell lots
    lCurrentPosition As Long            ' Non-Snapshot Current Position
    lNumTotal As Long                   ' Number of non-snapshot lots
    dRollingSum As Double               ' Non-Snapshot rolling sum
    dClosedProfit As Double             ' Non-Snapshot closed profit
    dAverageEntry As Double             ' Average entry for the last open non-snapshot position
    dLastTraded As Double               ' Date/Time of last non-snapshot trade
    dTotalFees As Double                ' Total fees for non-snapshot fills
    lLastFillSequence As Long           ' Sequence number of the last non-snapshot fill
    
    lSessionDateSnapshot As Long        ' Session Date of last trade
    lSessionQuantitySnapshot As Long    ' Total Quantity traded in session date of last trade
    dSessionProfitSnapshot As Double    ' Closed out Profit/Loss for session date of last trade
    lNumBuysSnapshot As Long            ' Number of snapshot buy lots
    lNumSellsSnapshot As Long           ' Number of snapshot sell lots
    lCurrentPositionSnapshot As Long    ' Snapshot Current Position
    lNumTotalSnapshot As Long           ' Number of snapshot lots
    dRollingSumSnapshot As Double       ' Snapshot rolling sum
    dClosedProfitSnapshot As Double     ' Snapshot closed profit
    dAverageEntrySnapshot As Double     ' Average entry for the last open snapshot position
    dLastTradedSnapshot As Double       ' Date/Time of last snapshot trade
    dTotalFeesSnapshot As Double        ' Total fees for snapshot fills
    lLastFillSequenceSnapshot As Long   ' Sequence number of the last snapshot fill
    
    dOpenProfit As Double               ' Last known open profit
    bDirty As Boolean                   ' Does this account position need to be recalculated?
    strSortKey As String                ' Sort key
    
    Entries As cGdTree                  ' Collection of non-snapshot entry fills
    EntriesSnapshot As cGdTree          ' Collection of snapshot entry fills
    
    TradeLines As cTradeLines           ' Trade by trade lines for the account position
    
    Bars As cGdBars                     ' Bars object
End Type
Private m As mPrivate

Public Property Get AccountPositionID() As Long
    AccountPositionID = m.lAccountPositionID
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.nBroker = g.Broker.AccountTypeForID(lAccountID)
    BuildSortKey
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
    SetBarProperties m.Bars, m.lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
    SetBarProperties m.Bars, m.strSymbol
    BuildSortKey
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
    SetBarProperties m.Bars, vSymbolOrSymbolID
    BuildSortKey
End Property

Public Property Get AutoTradeItemID() As Long
    AutoTradeItemID = m.lAutoTradeItemID
End Property
Public Property Let AutoTradeItemID(ByVal lAutoTradeItemID As Long)
    m.lAutoTradeItemID = lAutoTradeItemID
    BuildSortKey
End Property

Public Property Get FillMatchMode() As eTT_FillMatchMode
    FillMatchMode = m.nFillMatchMode
End Property
Public Property Let FillMatchMode(ByVal nFillMatchMode As eTT_FillMatchMode)
    m.nFillMatchMode = nFillMatchMode
End Property

Public Property Get StartFillTime() As Double
    StartFillTime = m.dStartFillTime
End Property
Public Property Let StartFillTime(ByVal dStartFillTime As Double)
    m.dStartFillTime = dStartFillTime
End Property

Public Property Get SessionDate() As Long
    SessionDate = m.lSessionDate
End Property
Public Property Let SessionDate(ByVal lSessionDate As Long)
    m.lSessionDate = lSessionDate
End Property

Public Property Get SessionQuantity() As Long
    SessionQuantity = m.lSessionQuantity
End Property
Public Property Let SessionQuantity(ByVal lSessionQuantity As Long)
    m.lSessionQuantity = lSessionQuantity
End Property

Public Property Get SessionProfit() As Double
    SessionProfit = m.dSessionProfit
End Property
Public Property Let SessionProfit(ByVal dSessionProfit As Double)
    m.dSessionProfit = dSessionProfit
End Property

Public Property Get NumBuys() As Long
    NumBuys = m.lNumBuys
End Property
Public Property Let NumBuys(ByVal lNumBuys As Long)
    m.lNumBuys = lNumBuys
End Property

Public Property Get NumSells() As Long
    NumSells = m.lNumSells
End Property
Public Property Let NumSells(ByVal lNumSells As Long)
    m.lNumSells = lNumSells
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

Public Property Get NumTotal() As Long
    NumTotal = m.lNumTotal
End Property
Public Property Let NumTotal(ByVal lNumTotal As Long)
    m.lNumTotal = lNumTotal
End Property

Public Property Get RollingSum() As Double
    RollingSum = m.dRollingSum
End Property
Public Property Let RollingSum(ByVal dRollingSum As Double)
    m.dRollingSum = dRollingSum
End Property

Public Property Get ClosedProfit() As Double
    ClosedProfit = m.dClosedProfit
End Property
Public Property Let ClosedProfit(ByVal dClosedProfit As Double)
    m.dClosedProfit = dClosedProfit
End Property

Public Property Get AverageEntry() As Double
    AverageEntry = m.dAverageEntry
End Property
Public Property Let AverageEntry(ByVal dAverageEntry As Double)
    m.dAverageEntry = dAverageEntry
End Property
Public Property Get AverageEntryString() As String
    If m.dAverageEntry = 0# Then
        AverageEntryString = ""
    Else
        AverageEntryString = m.Bars.PriceDisplay(m.dAverageEntry)
    End If
End Property

Public Property Get LastTraded() As Double
    LastTraded = m.dLastTraded
End Property
Public Property Let LastTraded(ByVal dLastTraded As Double)
    m.dLastTraded = dLastTraded
End Property

Public Property Get TotalFees() As Double
    TotalFees = m.dTotalFees
End Property
Public Property Let TotalFees(ByVal dTotalFees As Double)
    m.dTotalFees = dTotalFees
End Property

Public Property Get LastFillSequence() As Long
    LastFillSequence = m.lLastFillSequence
End Property
Public Property Let LastFillSequence(ByVal lLastFillSequence As Long)
    m.lLastFillSequence = lLastFillSequence
End Property

Public Property Get SessionDateSnapshot() As Long
    SessionDateSnapshot = m.lSessionDateSnapshot
End Property
Public Property Let SessionDateSnapshot(ByVal lSessionDateSnapshot As Long)
    m.lSessionDateSnapshot = lSessionDateSnapshot
End Property

Public Property Get SessionQuantitySnapshot() As Long
    SessionQuantitySnapshot = m.lSessionQuantitySnapshot
End Property
Public Property Let SessionQuantitySnapshot(ByVal lSessionQuantitySnapshot As Long)
    m.lSessionQuantitySnapshot = lSessionQuantitySnapshot
End Property

Public Property Get SessionProfitSnapshot() As Double
    SessionProfitSnapshot = m.dSessionProfitSnapshot
End Property
Public Property Let SessionProfitSnapshot(ByVal dSessionProfitSnapshot As Double)
    m.dSessionProfitSnapshot = dSessionProfitSnapshot
End Property

Public Property Get NumBuysSnapshot() As Long
    NumBuysSnapshot = m.lNumBuysSnapshot
End Property
Public Property Let NumBuysSnapshot(ByVal lNumBuysSnapshot As Long)
    m.lNumBuysSnapshot = lNumBuysSnapshot
End Property

Public Property Get NumSellsSnapshot() As Long
    NumSellsSnapshot = m.lNumSellsSnapshot
End Property
Public Property Let NumSellsSnapshot(ByVal lNumSellsSnapshot As Long)
    m.lNumSellsSnapshot = lNumSellsSnapshot
End Property

Public Property Get CurrentPositionSnapshot() As Long
    CurrentPositionSnapshot = m.lCurrentPositionSnapshot
End Property
Public Property Let CurrentPositionSnapshot(ByVal lCurrentPositionSnapshot As Long)
    m.lCurrentPositionSnapshot = lCurrentPositionSnapshot
End Property
Public Property Get CurrentPositionSnapshotString() As String
    CurrentPositionSnapshotString = g.Broker.TextPosition(m.lCurrentPositionSnapshot)
End Property

Public Property Get NumTotalSnapshot() As Long
    NumTotalSnapshot = m.lNumTotalSnapshot
End Property
Public Property Let NumTotalSnapshot(ByVal lNumTotalSnapshot As Long)
    m.lNumTotalSnapshot = lNumTotalSnapshot
End Property

Public Property Get RollingSumSnapshot() As Double
    RollingSumSnapshot = m.dRollingSumSnapshot
End Property
Public Property Let RollingSumSnapshot(ByVal dRollingSumSnapshot As Double)
    m.dRollingSumSnapshot = dRollingSumSnapshot
End Property

Public Property Get ClosedProfitSnapshot() As Double
    ClosedProfitSnapshot = m.dClosedProfitSnapshot
End Property
Public Property Let ClosedProfitSnapshot(ByVal dClosedProfitSnapshot As Double)
    m.dClosedProfitSnapshot = dClosedProfitSnapshot
End Property

Public Property Get AverageEntrySnapshot() As Double
    AverageEntrySnapshot = m.dAverageEntrySnapshot
End Property
Public Property Let AverageEntrySnapshot(ByVal dAverageEntrySnapshot As Double)
    m.dAverageEntrySnapshot = dAverageEntrySnapshot
End Property
Public Property Get AverageEntrySnapshotString() As String
    If m.dAverageEntrySnapshot = 0# Then
        AverageEntrySnapshotString = ""
    Else
        AverageEntrySnapshotString = m.Bars.PriceDisplay(m.dAverageEntrySnapshot)
    End If
End Property

Public Property Get LastTradedSnapshot() As Double
    LastTradedSnapshot = m.dLastTradedSnapshot
End Property
Public Property Let LastTradedSnapshot(ByVal dLastTradedSnapshot As Double)
    m.dLastTradedSnapshot = dLastTradedSnapshot
End Property

Public Property Get TotalFeesSnapshot() As Double
    TotalFeesSnapshot = m.dTotalFeesSnapshot
End Property
Public Property Let TotalFeesSnapshot(ByVal dTotalFeesSnapshot As Double)
    m.dTotalFeesSnapshot = dTotalFeesSnapshot
End Property

Public Property Get LastFillSequenceSnapshot() As Long
    LastFillSequenceSnapshot = m.lLastFillSequenceSnapshot
End Property
Public Property Let LastFillSequenceSnapshot(ByVal lLastFillSequenceSnapshot As Long)
    m.lLastFillSequenceSnapshot = lLastFillSequenceSnapshot
End Property

Public Property Get TotalFeesTotal() As Double
    TotalFeesTotal = m.dTotalFees + m.dTotalFeesSnapshot
End Property

Public Property Get Dirty() As Boolean
    Dirty = m.bDirty
End Property
Public Property Let Dirty(ByVal bDirty As Boolean)
    m.bDirty = bDirty
End Property

Public Property Get TradeLines() As cTradeLines
    Set TradeLines = m.TradeLines
End Property

Public Property Get Key() As String
    Key = Str(m.lAccountID) & vbTab & Str(SymbolOrSymbolID) & vbTab & Str(m.lAutoTradeItemID)
End Property

Public Property Get SortKey() As String
    SortKey = m.strSortKey
End Property

Public Property Get Entries() As cGdTree
    Set Entries = m.EntriesSnapshot
End Property

Public Property Get ExpirationDate() As Long
    ExpirationDate = m.lExpirationDate
End Property
Public Property Let ExpirationDate(ByVal lExpirationDate As Long)
    m.lExpirationDate = lExpirationDate
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the account position from the database
'' Inputs:      Account Position ID, Recordset, Trade Lines Recordset
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lAccountPositionID As Long, Optional ByVal rs As Recordset = Nothing, Optional ByVal rsTradeLines As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Last fill from the account position fills

    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                    "WHERE [AccountPositionID]=" & Str(lAccountPositionID) & ";", dbOpenDynaset)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        m.lAccountPositionID = rs!AccountPositionID
        m.lAccountID = rs!AccountID
        m.nBroker = g.Broker.AccountTypeForID(m.lAccountID)
        m.lSymbolID = rs!SymbolID
        m.strSymbol = rs!Symbol
        m.lAutoTradeItemID = rs!AutoTradeItemID
        m.nFillMatchMode = rs!FillMatchMode
        m.dStartFillTime = rs!StartFillTime
        m.lExpirationDate = rs!ExpirationDate
        BuildSortKey
        
        m.lSessionDate = rs!SessionDate
        m.lSessionQuantity = rs!SessionQuantity
        m.dSessionProfit = rs!SessionProfit
        m.lNumBuys = rs!NumBuys
        m.lNumSells = rs!NumSells
        m.lCurrentPosition = rs!CurrentPosition
        m.lNumTotal = rs!NumTotal
        m.dRollingSum = rs!RollingSum
        m.dClosedProfit = rs!ClosedProfit
        m.dAverageEntry = rs!AverageEntry
        m.dLastTraded = rs!LastTraded
        m.dTotalFees = rs!TotalFees
        m.lLastFillSequence = rs!LastFillSeq
        
        m.lSessionDateSnapshot = rs!SessionDateSnapshot
        m.lSessionQuantitySnapshot = rs!SessionQuantitySnapshot
        m.dSessionProfitSnapshot = rs!SessionProfitSnapshot
        m.lNumBuysSnapshot = rs!NumBuysSnapshot
        m.lNumSellsSnapshot = rs!NumSellsSnapshot
        m.lCurrentPositionSnapshot = rs!CurrentPositionSnapshot
        m.lNumTotalSnapshot = rs!NumTotalSnapshot
        m.dRollingSumSnapshot = rs!RollingSumSnapshot
        m.dClosedProfitSnapshot = rs!ClosedProfitSnapshot
        m.dAverageEntrySnapshot = rs!AverageEntrySnapshot
        m.dLastTradedSnapshot = rs!LastTradedSnapshot
        m.dTotalFeesSnapshot = rs!TotalFeesSnapshot
        m.lLastFillSequenceSnapshot = rs!LastFillSeqSnapshot
        
        SetBarProperties m.Bars, SymbolOrSymbolID
               
        LoadEntries
        
        m.TradeLines.Load m.lAccountPositionID, SymbolOrSymbolID, rsTradeLines
                
        Load = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the account position to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!AccountID = m.lAccountID
    rs!SymbolID = m.lSymbolID
    rs!Symbol = m.strSymbol
    rs!AutoTradeItemID = m.lAutoTradeItemID
    rs!FillMatchMode = m.nFillMatchMode
    rs!StartFillTime = m.dStartFillTime
    rs!ExpirationDate = m.lExpirationDate
    
    rs!SessionDate = m.lSessionDate
    rs!SessionQuantity = m.lSessionQuantity
    rs!SessionProfit = m.dSessionProfit
    rs!NumBuys = m.lNumBuys
    rs!NumSells = m.lNumSells
    rs!CurrentPosition = m.lCurrentPosition
    rs!NumTotal = m.lNumTotal
    rs!RollingSum = m.dRollingSum
    rs!ClosedProfit = m.dClosedProfit
    rs!AverageEntry = m.dAverageEntry
    rs!LastTraded = m.dLastTraded
    rs!TotalFees = m.dTotalFees
    rs!LastFillSeq = m.lLastFillSequence
    
    rs!SessionDateSnapshot = m.lSessionDateSnapshot
    rs!SessionQuantitySnapshot = m.lSessionQuantitySnapshot
    rs!SessionProfitSnapshot = m.dSessionProfitSnapshot
    rs!NumBuysSnapshot = m.lNumBuysSnapshot
    rs!NumSellsSnapshot = m.lNumSellsSnapshot
    rs!CurrentPositionSnapshot = m.lCurrentPositionSnapshot
    rs!NumTotalSnapshot = m.lNumTotalSnapshot
    rs!RollingSumSnapshot = m.dRollingSumSnapshot
    rs!ClosedProfitSnapshot = m.dClosedProfitSnapshot
    rs!AverageEntrySnapshot = m.dAverageEntrySnapshot
    rs!LastTradedSnapshot = m.dLastTradedSnapshot
    rs!TotalFeesSnapshot = m.dTotalFeesSnapshot
    rs!LastFillSeqSnapshot = m.lLastFillSequenceSnapshot
    
    m.lAccountPositionID = rs!AccountPositionID
    
    rs.Update
    
    SaveEntries
    
    m.TradeLines.Save
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateHistory
'' Description: Recalculate the account position history from the beginning
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateHistory()
On Error GoTo ErrSection:

    If m.nFillMatchMode = eTT_FillMatchMode_Fifo Then
        RecalculateHistoryFifo
    ElseIf m.nFillMatchMode = eTT_FillMatchMode_Lifo Then
        RecalculateHistoryLifo
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateHistory"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateHistoryLifo
'' Description: Recalculate the account position history from the beginning
''              matching fills on a last-in first-out basis
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateHistoryLifo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSequence As Long               ' Fill sequence
    Dim Entry As cPtFill                ' Entry fill object
    Dim dFillProfit As Double           ' Fill profit
    Dim lEntryQuantity As Long          ' Sum of the entry quantities
    Dim dEntrySum As Double             ' Sum of the entry prices
    Dim ForexBaseBars As cGdBars        ' Bars for the forex base pair (if applicable)
    Dim strForexBase As String          ' Forex base symbol
    Dim dExchangeFillTime As Double     ' Exchange fill time
    
    strForexBase = g.Profit.BaseForex(m.strSymbol)
    If Len(strForexBase) = 0 Then
        Set ForexBaseBars = Nothing
    Else
        Set ForexBaseBars = New cGdBars
        DM_GetBars ForexBaseBars, strForexBase
    End If
        
    If m.lAutoTradeItemID = -1& Then
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    Else
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    End If
    
    Clear False
    
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs, m.Bars, m.nBroker) Then
            With Fill
                dExchangeFillTime = ConvertBrokerDate(.FillDate, .Broker, .Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                
                If m.dStartFillTime = 0# Then
                    m.dStartFillTime = Fill.FillDate
                End If
            
                lSequence = lSequence + 1
                If m.lAutoTradeItemID = -1& Then
                    .Sequence = lSequence
                Else
                    .SequenceCategory = lSequence
                End If
                m.lLastFillSequence = lSequence
            
                If .SessionDate <> m.lSessionDate Then
                    m.lSessionDate = .SessionDate
                    m.lSessionQuantity = .Quantity
                    m.dSessionProfit = 0#
                Else
                    m.lSessionQuantity = m.lSessionQuantity + .Quantity
                End If
                dFillProfit = 0#
                
                m.dTotalFees = m.dTotalFees + .Fees
                
                If .FillDate > m.dLastTraded Then m.dLastTraded = .FillDate
                
                If .Buy Then
                    m.lNumBuys = m.lNumBuys + .Quantity
                    m.lNumTotal = m.lNumTotal + .Quantity
                    m.dRollingSum = m.dRollingSum - (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPosition >= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition + .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition + .Quantity
                            .NumOpenCategory = .Quantity
                        End If
                        .Save
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPosition) Then
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.Entries(m.Entries.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
                                        .Save
                                        
                                        m.Entries.Remove m.Entries.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpen = .NumOpen - lIndex
                                        .Save
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
                                        .Save
                                        
                                        m.Entries.Remove m.Entries.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
                                        .Save
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition + .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition + .Quantity
                            .NumOpenCategory = 0&
                        End If
                        .Save
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = m.Entries.Count To 1 Step -1
                            Set Entry = m.Entries(m.Entries.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpen = 0&
                                    .Save
                                    
                                    m.Entries.Remove m.Entries.Count
                                Else
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpenCategory = 0&
                                    .Save
                                    
                                    m.Entries.Remove m.Entries.Count
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition + .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPosition)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition + .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPosition)
                        End If
                        .Save
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPosition = m.lCurrentPosition + .Quantity
                Else
                    m.lNumSells = m.lNumSells + .Quantity
                    m.lNumTotal = m.lNumTotal + .Quantity
                    m.dRollingSum = m.dRollingSum + (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPosition <= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition - .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition - .Quantity
                            .NumOpenCategory = .Quantity
                        End If
                        .Save
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPosition) Then
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.Entries(m.Entries.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
                                        .Save
                                        
                                        m.Entries.Remove m.Entries.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpen = .NumOpen - lIndex
                                        .Save
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
                                        .Save
                                        
                                        m.Entries.Remove m.Entries.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
                                        .Save
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition - .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition - .Quantity
                            .NumOpenCategory = 0&
                        End If
                        .Save
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = m.Entries.Count To 1 Step -1
                            Set Entry = m.Entries(m.Entries.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpen = 0&
                                    .Save
                                    
                                    m.Entries.Remove m.Entries.Count
                                Else
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpenCategory = 0&
                                    .Save
                                    
                                    m.Entries.Remove m.Entries.Count
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition - .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPosition)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition - .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPosition)
                        End If
                        .Save
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPosition = m.lCurrentPosition - .Quantity
                End If
            End With
        End If
        
        rs.MoveNext
    Loop
    
    lEntryQuantity = 0&
    dEntrySum = 0#
    For lIndex = 1 To m.Entries.Count
        If m.lAutoTradeItemID = -1& Then
            lEntryQuantity = lEntryQuantity + m.Entries(lIndex).NumOpen
            dEntrySum = dEntrySum + (m.Entries(lIndex).Price * m.Entries(lIndex).NumOpen)
        Else
            lEntryQuantity = lEntryQuantity + m.Entries(lIndex).NumOpenCategory
            dEntrySum = dEntrySum + (m.Entries(lIndex).Price * m.Entries(lIndex).NumOpenCategory)
        End If
    Next lIndex
    If lEntryQuantity = 0& Then
        m.dAverageEntry = 0#
    Else
        m.dAverageEntry = dEntrySum / Val(lEntryQuantity)
    End If
    
    RecalculateSnapshotLifo False, False
    Save

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateHistoryLifo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateHistoryFifo
'' Description: Recalculate the account position history from the beginning
''              matching fills on a first-in first-out basis
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateHistoryFifo()
On Error GoTo ErrSection:

gdStartProfile 680
gdStartProfile 681

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSequence As Long               ' Fill sequence
    Dim Entry As cPtFill                ' Entry fill object
    Dim dFillProfit As Double           ' Fill profit
    Dim lEntryQuantity As Long          ' Sum of the entry quantities
    Dim dEntrySum As Double             ' Sum of the entry prices
    Dim ForexBaseBars As cGdBars        ' Bars for the forex base pair (if applicable)
    Dim strForexBase As String          ' Forex base symbol
    Dim dExchangeFillTime As Double     ' Exchange fill time
    
    strForexBase = g.Profit.BaseForex(m.strSymbol)
    If Len(strForexBase) = 0 Then
        Set ForexBaseBars = Nothing
    Else
        Set ForexBaseBars = New cGdBars
        DM_GetBars ForexBaseBars, strForexBase
    End If
    
gdStopProfile 681
gdStartProfile 682
    
    If m.lAutoTradeItemID = -1& Then
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    Else
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=0 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    End If
    
gdStopProfile 682
gdStartProfile 683
    
    Clear False
    
gdStopProfile 683
gdStartProfile 684
    
    Do While Not rs.EOF
gdStartProfile 661
        Set Fill = New cPtFill
gdStopProfile 661
gdStartProfile 620
        If Fill.Load(rs!FillID, rs, m.Bars, m.nBroker) Then
gdStopProfile 620
            With Fill
                dExchangeFillTime = ConvertBrokerDate(.FillDate, .Broker, .Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                
                If m.dStartFillTime = 0# Then
                    m.dStartFillTime = Fill.FillDate
                End If
            
                lSequence = lSequence + 1
                If m.lAutoTradeItemID = -1& Then
                    .Sequence = lSequence
                Else
                    .SequenceCategory = lSequence
                End If
                m.lLastFillSequence = lSequence
            
                If .SessionDate <> m.lSessionDate Then
                    m.lSessionDate = .SessionDate
                    m.lSessionQuantity = .Quantity
                    m.dSessionProfit = 0#
                Else
                    m.lSessionQuantity = m.lSessionQuantity + .Quantity
                End If
                dFillProfit = 0#
                
                m.dTotalFees = m.dTotalFees + .Fees
                
                If .FillDate > m.dLastTraded Then m.dLastTraded = .FillDate
                
                If .Buy Then
                    m.lNumBuys = m.lNumBuys + .Quantity
                    m.lNumTotal = m.lNumTotal + .Quantity
                    m.dRollingSum = m.dRollingSum - (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPosition >= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition + .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition + .Quantity
                            .NumOpenCategory = .Quantity
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPosition) Then
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.Entries(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.Entries.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpen = .NumOpen - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.Entries.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition + .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition + .Quantity
                            .NumOpenCategory = 0&
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = m.Entries.Count To 1 Step -1
                            Set Entry = m.Entries(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.Entries.Remove 1
                                Else
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.Entries.Remove 1
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition + .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPosition)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition + .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPosition)
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPosition = m.lCurrentPosition + .Quantity
                Else
                    m.lNumSells = m.lNumSells + .Quantity
                    m.lNumTotal = m.lNumTotal + .Quantity
                    m.dRollingSum = m.dRollingSum + (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPosition <= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition - .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition - .Quantity
                            .NumOpenCategory = .Quantity
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPosition) Then
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.Entries(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.Entries.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpen = .NumOpen - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.Entries.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition - .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition - .Quantity
                            .NumOpenCategory = 0&
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.Entries, Fill, False, ForexBaseBars
                        
                        For lIndex = m.Entries.Count To 1 Step -1
                            Set Entry = m.Entries(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.Entries.Remove 1
                                Else
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.Entries.Remove 1
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfit = m.dSessionProfit + dFillProfit
                        m.dClosedProfit = m.dClosedProfit + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfit
                            .TotalProfit = m.dClosedProfit
                            .SessionQuantity = m.lSessionQuantity
                            .CurrentPosition = m.lCurrentPosition - .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPosition)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfit
                            .TotalProfitCategory = m.dClosedProfit
                            .SessionQuantityCategory = m.lSessionQuantity
                            .CurrentPositionCategory = m.lCurrentPosition - .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPosition)
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.Entries.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPosition = m.lCurrentPosition - .Quantity
                End If
            End With
        End If
        
        rs.MoveNext
    Loop
    
gdStopProfile 684
gdStartProfile 685
    
    lEntryQuantity = 0&
    dEntrySum = 0#
    For lIndex = 1 To m.Entries.Count
        If m.lAutoTradeItemID = -1& Then
            lEntryQuantity = lEntryQuantity + m.Entries(lIndex).NumOpen
            dEntrySum = dEntrySum + (m.Entries(lIndex).Price * m.Entries(lIndex).NumOpen)
        Else
            lEntryQuantity = lEntryQuantity + m.Entries(lIndex).NumOpenCategory
            dEntrySum = dEntrySum + (m.Entries(lIndex).Price * m.Entries(lIndex).NumOpenCategory)
        End If
    Next lIndex
    If lEntryQuantity = 0& Then
        m.dAverageEntry = 0#
    Else
        m.dAverageEntry = dEntrySum / Val(lEntryQuantity)
    End If
    
gdStopProfile 685
gdStartProfile 686
    
    RecalculateSnapshotFifo False, False

gdStopProfile 686
gdStartProfile 687
    
    Save

gdStopProfile 687
gdStopProfile 680

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateHistoryFifo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateSnapshot
'' Description: Recalculate the account position snapshot
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateSnapshot()
On Error GoTo ErrSection:

    If m.nFillMatchMode = eTT_FillMatchMode_Fifo Then
        RecalculateSnapshotFifo
    ElseIf m.nFillMatchMode = eTT_FillMatchMode_Lifo Then
        RecalculateSnapshotLifo
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateSnapshot"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateSnapshotLifo
'' Description: Recalculate the account position snapshot matching the fills
''              on a last-in first-out basis
'' Inputs:      Move to History?, Clear and Save?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateSnapshotLifo(Optional ByVal bMoveSnapshotToHistory As Boolean = False, Optional ByVal bSaveAndClear As Boolean = True)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSequence As Long               ' Fill sequence
    Dim Entry As cPtFill                ' Entry fill object
    Dim dFillProfit As Double           ' Fill profit
    Dim lEntryQuantity As Long          ' Sum of the entry quantities
    Dim dEntrySum As Double             ' Sum of the entry prices
    Dim ForexBaseBars As cGdBars        ' Bars for the forex base pair (if applicable)
    Dim strForexBase As String          ' Forex base symbol
    Dim dExchangeFillTime As Double     ' Exchange fill time
    
    strForexBase = g.Profit.BaseForex(m.strSymbol)
    If Len(strForexBase) = 0 Then
        Set ForexBaseBars = Nothing
    Else
        Set ForexBaseBars = New cGdBars
        DM_GetBars ForexBaseBars, strForexBase
    End If
                
    If m.lAutoTradeItemID = -1& Then
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    Else
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    End If
    
    If bSaveAndClear Then
        Clear True
    End If
    
    ' Carry over information from non-snapshot...
    m.lCurrentPositionSnapshot = m.lCurrentPosition
    For lIndex = 1 To m.Entries.Count
        m.EntriesSnapshot.Add m.Entries(lIndex).MakeCopy, Str(m.Entries(lIndex).FillID)
    Next lIndex
    m.dRollingSumSnapshot = m.dRollingSum
    m.dLastTradedSnapshot = m.dLastTraded
    lSequence = m.lLastFillSequence
    
    m.lSessionDateSnapshot = 0& 'm.lSessionDate
    m.lSessionQuantitySnapshot = 0& 'm.lSessionQuantity
    m.dSessionProfitSnapshot = 0# 'm.dSessionProfit
    m.dClosedProfitSnapshot = m.dClosedProfit
    
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs, m.Bars, m.nBroker) Then
            If m.Entries.Exists(Str(Fill.FillID)) Then
                RecalculateHistoryLifo
                g.Broker.BrokerDebug m.nBroker, "Recalculating Lifo History because a snapshot order exists in history"
                
                Exit Sub
            End If
            
            With Fill
                dExchangeFillTime = ConvertBrokerDate(.FillDate, .Broker, .Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                
                If m.dStartFillTime = 0# Then
                    m.dStartFillTime = Fill.FillDate
                End If
            
                lSequence = lSequence + 1
                If m.lAutoTradeItemID = -1& Then
                    .Sequence = lSequence
                Else
                    .SequenceCategory = lSequence
                End If
                m.lLastFillSequenceSnapshot = lSequence
            
                If .SessionDate <> m.lSessionDateSnapshot Then
                    m.lSessionDateSnapshot = .SessionDate
                    m.lSessionQuantitySnapshot = .Quantity
                    m.dSessionProfitSnapshot = 0#
                Else
                    m.lSessionQuantitySnapshot = m.lSessionQuantitySnapshot + .Quantity
                End If
                dFillProfit = 0#
                
                m.dTotalFeesSnapshot = m.dTotalFeesSnapshot + .Fees
                
                If .FillDate > m.dLastTradedSnapshot Then m.dLastTradedSnapshot = .FillDate
                
                If .Buy Then
                    m.lNumBuysSnapshot = m.lNumBuysSnapshot + .Quantity
                    m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
                    m.dRollingSumSnapshot = m.dRollingSumSnapshot - (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPositionSnapshot >= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpenCategory = .Quantity
                        End If
                        .Save
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
                                        .Save
                                        
                                        m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpen = .NumOpen - lIndex
                                        .Save
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
                                        .Save
                                        
                                        m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
                                        .Save
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpenCategory = 0&
                        End If
                        .Save
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                            Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpen = 0&
                                    .Save
                                    
                                    m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                Else
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpenCategory = 0&
                                    .Save
                                    
                                    m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        End If
                        .Save
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot + .Quantity
                Else
                    m.lNumSellsSnapshot = m.lNumSellsSnapshot + .Quantity
                    m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
                    m.dRollingSumSnapshot = m.dRollingSumSnapshot + (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPositionSnapshot <= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpenCategory = .Quantity
                        End If
                        .Save
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
                                        .Save
                                        
                                        m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpen = .NumOpen - lIndex
                                        .Save
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
                                        .Save
                                        
                                        m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                    Else
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
                                        .Save
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpenCategory = 0&
                        End If
                        .Save
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                            Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpen = 0&
                                    .Save
                                    
                                    m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                Else
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                    .NumOpenCategory = 0&
                                    .Save
                                    
                                    m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        End If
                        .Save
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot - .Quantity
                End If
            End With
        End If
        
        rs.MoveNext
    Loop
    
    lEntryQuantity = 0&
    dEntrySum = 0#
    For lIndex = 1 To m.EntriesSnapshot.Count
        If m.lAutoTradeItemID = -1& Then
            lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpen
            dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpen)
        Else
            lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpenCategory
            dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpenCategory)
        End If
    Next lIndex
    If lEntryQuantity = 0& Then
        m.dAverageEntrySnapshot = 0#
    Else
        m.dAverageEntrySnapshot = dEntrySum / Val(lEntryQuantity)
    End If
    
    CreateTradeLines m.EntriesSnapshot, Nothing, True, ForexBaseBars
    
    If bSaveAndClear = True Then
        Save
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateSnapshotLifo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateSnapshotFifo
'' Description: Recalculate the account position snapshot matching the fills
''              on a first-in first-out basis
'' Inputs:      Move to History?, Clear and Save?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateSnapshotFifo(Optional ByVal bMoveSnapshotToHistory As Boolean = False, Optional ByVal bSaveAndClear As Boolean = True)
On Error GoTo ErrSection:

gdStartProfile 680
gdStartProfile 681

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSequence As Long               ' Fill sequence
    Dim Entry As cPtFill                ' Entry fill object
    Dim dFillProfit As Double           ' Fill profit
    Dim lEntryQuantity As Long          ' Sum of the entry quantities
    Dim dEntrySum As Double             ' Sum of the entry prices
    Dim lQtyLeft As Long                ' Quantity left
    Dim ForexBaseBars As cGdBars        ' Bars for the forex base pair (if applicable)
    Dim strForexBase As String          ' Forex base symbol
    Dim dExchangeFillTime As Double     ' Exchange fill time
    
    strForexBase = g.Profit.BaseForex(m.strSymbol)
    If Len(strForexBase) = 0 Then
        Set ForexBaseBars = Nothing
    Else
        Set ForexBaseBars = New cGdBars
        DM_GetBars ForexBaseBars, strForexBase
    End If
                
gdStopProfile 681
gdStartProfile 682
    
    If m.lAutoTradeItemID = -1& Then
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    Else
        If m.lSymbolID <> 0& Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [SymbolID]=" & Str(m.lSymbolID) & " AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(m.lAccountID) & " AND [Symbol]='" & Str(m.strSymbol) & "' AND [AutoTradeID]=" & Str(m.lAutoTradeItemID) & " AND [IsSnapshot]=-1 " & _
                        "ORDER BY [FillDate],[FillID];", dbOpenDynaset)
        End If
    End If
    
gdStopProfile 682
gdStartProfile 683
    
    If bSaveAndClear Then
        Clear True
    End If
    
    ' Carry over information from non-snapshot...
    m.lCurrentPositionSnapshot = m.lCurrentPosition
    lQtyLeft = Abs(m.lCurrentPosition)
    For lIndex = 1 To m.Entries.Count
        m.EntriesSnapshot.Add m.Entries(lIndex).MakeCopy, Str(m.Entries(lIndex).FillID)
        Set Fill = m.EntriesSnapshot(m.EntriesSnapshot.Count)
        If Fill.Quantity <= lQtyLeft Then
            Fill.NumberOpen(m.lAutoTradeItemID) = Fill.Quantity
            lQtyLeft = lQtyLeft - Fill.Quantity
        Else
            Fill.NumberOpen(m.lAutoTradeItemID) = lQtyLeft
            lQtyLeft = 0&
        End If
    Next lIndex
    m.dRollingSumSnapshot = m.dRollingSum
    m.dLastTradedSnapshot = m.dLastTraded
    lSequence = m.lLastFillSequence
    
    m.lSessionDateSnapshot = 0& 'm.lSessionDate
    m.lSessionQuantitySnapshot = 0& 'm.lSessionQuantity
    m.dSessionProfitSnapshot = 0# 'm.dSessionProfit
    m.dClosedProfitSnapshot = m.dClosedProfit
    
gdStopProfile 683
gdStartProfile 684
    
    Do While Not rs.EOF
        Set Fill = New cPtFill
gdStartProfile 620
        If Fill.Load(rs!FillID, rs, m.Bars, m.nBroker) Then
gdStopProfile 620
            If m.Entries.Exists(Str(Fill.FillID)) Then
                RecalculateHistoryFifo
                g.Broker.BrokerDebug m.nBroker, "Recalculating Fifo History because a snapshot order exists in history"
                
                Exit Sub
            End If
            
            With Fill
                dExchangeFillTime = ConvertBrokerDate(.FillDate, .Broker, .Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                
                If m.dStartFillTime = 0# Then
                    m.dStartFillTime = Fill.FillDate
                End If
            
                lSequence = lSequence + 1
                If m.lAutoTradeItemID = -1& Then
                    .Sequence = lSequence
                Else
                    .SequenceCategory = lSequence
                End If
                m.lLastFillSequenceSnapshot = lSequence
            
                If .SessionDate <> m.lSessionDateSnapshot Then
                    m.lSessionDateSnapshot = .SessionDate
                    m.lSessionQuantitySnapshot = .Quantity
                    m.dSessionProfitSnapshot = 0#
                Else
                    m.lSessionQuantitySnapshot = m.lSessionQuantitySnapshot + .Quantity
                End If
                dFillProfit = 0#
                
                m.dTotalFeesSnapshot = m.dTotalFeesSnapshot + .Fees
                
                If .FillDate > m.dLastTradedSnapshot Then m.dLastTradedSnapshot = .FillDate
                
                If .Buy Then
                    m.lNumBuysSnapshot = m.lNumBuysSnapshot + .Quantity
                    m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
                    m.dRollingSumSnapshot = m.dRollingSumSnapshot - (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPositionSnapshot >= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpenCategory = .Quantity
                        End If
                        
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                    
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.EntriesSnapshot(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.EntriesSnapshot.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpen = .NumOpen - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.EntriesSnapshot.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpenCategory = 0&
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                            Set Entry = m.EntriesSnapshot(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.EntriesSnapshot.Remove 1
                                Else
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.EntriesSnapshot.Remove 1
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot + .Quantity
                Else
                    m.lNumSellsSnapshot = m.lNumSellsSnapshot + .Quantity
                    m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
                    m.dRollingSumSnapshot = m.dRollingSumSnapshot + (.Price * .Quantity)
                    
                    ' Opening a new or adding to an existing position...
                    If m.lCurrentPositionSnapshot <= 0& Then
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "E"
                            .ClosedProfit = 0#
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpen = .Quantity
                        Else
                            .ActionCategory = "E"
                            .ClosedProfitCategory = 0#
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpenCategory = .Quantity
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                        
                    ' Closing part or all of a position...
                    ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = .Quantity To 1 Step -1
                            Set Entry = m.EntriesSnapshot(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
                                    If .NumOpen <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpen + 1
                                        .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.EntriesSnapshot.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpen = .NumOpen - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                        lIndex = 1
                                    End If
                                Else
                                    If .NumOpenCategory <= lIndex Then
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        lIndex = lIndex - .NumOpenCategory + 1
                                        .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        m.EntriesSnapshot.Remove 1
                                    Else
gdStartProfile 670
                                        dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                        
                                        .NumOpenCategory = .NumOpenCategory - lIndex
gdStartProfile 630
                        .Save
gdStopProfile 630
                                        
                                        lIndex = 1
                                    End If
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "X"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpen = 0&
                        Else
                            .ActionCategory = "X"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpenCategory = 0&
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                    ' Reversing the position...
                    Else
                        CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                        
                        For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                            Set Entry = m.EntriesSnapshot(1)
                            With Entry
                                If m.lAutoTradeItemID = -1& Then
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpen = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.EntriesSnapshot.Remove 1
                                Else
gdStartProfile 670
                                    dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
gdStopProfile 670
                                    .NumOpenCategory = 0&
gdStartProfile 630
                        .Save
gdStopProfile 630
                                    
                                    m.EntriesSnapshot.Remove 1
                                End If
                            End With
                        Next lIndex
                        
                        m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                        m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                        
                        If m.lAutoTradeItemID = -1& Then
                            .Action = "R"
                            .ClosedProfit = dFillProfit
                            .SessionProfit = m.dSessionProfitSnapshot
                            .TotalProfit = m.dClosedProfitSnapshot
                            .SessionQuantity = m.lSessionQuantitySnapshot
                            .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        Else
                            .ActionCategory = "R"
                            .ClosedProfitCategory = dFillProfit
                            .SessionProfitCategory = m.dSessionProfitSnapshot
                            .TotalProfitCategory = m.dClosedProfitSnapshot
                            .SessionQuantityCategory = m.lSessionQuantitySnapshot
                            .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                            .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                        End If
gdStartProfile 630
                        .Save
gdStopProfile 630
                        
                        m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                    End If
                
                    m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot - .Quantity
                End If
            End With
        End If
        
        rs.MoveNext
    Loop
    
gdStopProfile 684
gdStartProfile 685
    
    lEntryQuantity = 0&
    dEntrySum = 0#
    For lIndex = 1 To m.EntriesSnapshot.Count
        If m.lAutoTradeItemID = -1& Then
            lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpen
            dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpen)
        Else
            lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpenCategory
            dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpenCategory)
        End If
    Next lIndex
    If lEntryQuantity = 0& Then
        m.dAverageEntrySnapshot = 0#
    Else
        m.dAverageEntrySnapshot = dEntrySum / Val(lEntryQuantity)
    End If
    
    CreateTradeLines m.EntriesSnapshot, Nothing, True, ForexBaseBars
                
gdStopProfile 685
gdStartProfile 687
    
    If bSaveAndClear = True Then
        Save
    End If
    
gdStopProfile 687
gdStopProfile 680

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateSnapshotFifo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSnapshotFill
'' Description: Add the fill to snapshot and recalculate the account position
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSnapshotFill(Fill As cPtFill)
On Error GoTo ErrSection:

    ' If this is a new fill, then we can just add it to the snapshot information
    ' and recalculate from the last known information, but if the fill exists
    ' then we need to call a recalc (otherwise we end up continually adding
    ' the fill as a new fill and throwing the position off).  07/10/2008 DAJ
    If m.nFillMatchMode = eTT_FillMatchMode_Fifo Then
        If m.EntriesSnapshot.Exists(Str(Fill.FillID)) Then
            RecalculateSnapshotFifo
        Else
            AddSnapshotFillFifo Fill
        End If
    ElseIf m.nFillMatchMode = eTT_FillMatchMode_Lifo Then
        If m.EntriesSnapshot.Exists(Str(Fill.FillID)) Then
            RecalculateSnapshotLifo
        Else
            AddSnapshotFillLifo Fill
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.AddSnapshotFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSnapshotFillLifo
'' Description: Add the fill to snapshot and recalculate the account position
''              on a last-in first-out basis
'' Inputs:      Fill, Save?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSnapshotFillLifo(Fill As cPtFill, Optional ByVal bSaveAndClear As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSequence As Long               ' Fill sequence
    Dim Entry As cPtFill                ' Entry fill object
    Dim dFillProfit As Double           ' Fill profit
    Dim lEntryQuantity As Long          ' Sum of the entry quantities
    Dim dEntrySum As Double             ' Sum of the entry prices
    Dim lQtyLeft As Long                ' Quantity left
    Dim ForexBaseBars As cGdBars        ' Bars for the forex base pair (if applicable)
    Dim strForexBase As String          ' Forex base symbol
    Dim dExchangeFillTime As Double     ' Exchange fill time
      
    strForexBase = g.Profit.BaseForex(m.strSymbol)
    If Len(strForexBase) = 0 Then
        Set ForexBaseBars = Nothing
    Else
        Set ForexBaseBars = New cGdBars
        DM_GetBars ForexBaseBars, strForexBase
    End If
    
    If bSaveAndClear Then
        ClearOpenTradeLines
    End If
      
    With Fill
        dExchangeFillTime = ConvertBrokerDate(.FillDate, .Broker, .Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
        
        If m.dStartFillTime = 0# Then
            m.dStartFillTime = Fill.FillDate
        End If
    
        lSequence = m.lLastFillSequenceSnapshot + 1
        If m.lAutoTradeItemID = -1& Then
            .Sequence = lSequence
        Else
            .SequenceCategory = lSequence
        End If
        m.lLastFillSequenceSnapshot = lSequence
    
        If .SessionDate <> m.lSessionDateSnapshot Then
            m.lSessionDateSnapshot = .SessionDate
            m.lSessionQuantitySnapshot = .Quantity
            m.dSessionProfitSnapshot = 0#
        Else
            m.lSessionQuantitySnapshot = m.lSessionQuantitySnapshot + .Quantity
        End If
        dFillProfit = 0#
        
        m.dTotalFeesSnapshot = m.dTotalFeesSnapshot + .Fees
        
        If .FillDate > m.dLastTradedSnapshot Then m.dLastTradedSnapshot = .FillDate
        
        If .Buy Then
            m.lNumBuysSnapshot = m.lNumBuysSnapshot + .Quantity
            m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
            m.dRollingSumSnapshot = m.dRollingSumSnapshot - (.Price * .Quantity)
            
            ' Opening a new or adding to an existing position...
            If m.lCurrentPositionSnapshot >= 0& Then
                If m.lAutoTradeItemID = -1& Then
                    .Action = "E"
                    .ClosedProfit = 0#
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpen = .Quantity
                Else
                    .ActionCategory = "E"
                    .ClosedProfitCategory = 0#
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpenCategory = .Quantity
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                
            ' Closing part or all of a position...
            ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = .Quantity To 1 Step -1
                    Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            If .NumOpen <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpen + 1
                                .NumOpen = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpen = .NumOpen - lIndex
                                .Save
                            
                                lIndex = 1
                            End If
                        Else
                            If .NumOpenCategory <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpenCategory + 1
                                .NumOpenCategory = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpenCategory = .NumOpenCategory - lIndex
                                .Save
                                
                                lIndex = 1
                            End If
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "X"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpen = 0&
                Else
                    .ActionCategory = "X"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpenCategory = 0&
                End If
                .Save
                
            ' Reversing the position...
            Else
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                    Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpen = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                        Else
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpenCategory = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "R"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                Else
                    .ActionCategory = "R"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
            End If
        
            m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot + .Quantity
        Else
            m.lNumSellsSnapshot = m.lNumSellsSnapshot + .Quantity
            m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
            m.dRollingSumSnapshot = m.dRollingSumSnapshot + (.Price * .Quantity)
            
            ' Opening a new or adding to an existing position...
            If m.lCurrentPositionSnapshot <= 0& Then
                If m.lAutoTradeItemID = -1& Then
                    .Action = "E"
                    .ClosedProfit = 0#
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpen = .Quantity
                Else
                    .ActionCategory = "E"
                    .ClosedProfitCategory = 0#
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpenCategory = .Quantity
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                
            ' Closing part or all of a position...
            ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = .Quantity To 1 Step -1
                    Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            If .NumOpen <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpen + 1
                                .NumOpen = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpen = .NumOpen - lIndex
                                .Save
                            
                                lIndex = 1
                            End If
                        Else
                            If .NumOpenCategory <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpenCategory + 1
                                .NumOpenCategory = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpenCategory = .NumOpenCategory - lIndex
                                .Save
                                
                                lIndex = 1
                            End If
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "X"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpen = 0&
                Else
                    .ActionCategory = "X"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpenCategory = 0&
                End If
                .Save
                
            ' Reversing the position...
            Else
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                    Set Entry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpen = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                        Else
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpenCategory = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove m.EntriesSnapshot.Count
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "R"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                Else
                    .ActionCategory = "R"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
            End If
        
            m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot - .Quantity
        End If
    End With
    
    If bSaveAndClear Then
        lEntryQuantity = 0&
        dEntrySum = 0#
        For lIndex = 1 To m.EntriesSnapshot.Count
            If m.lAutoTradeItemID = -1& Then
                lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpen
                dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpen)
            Else
                lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpenCategory
                dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpenCategory)
            End If
        Next lIndex
        If lEntryQuantity = 0& Then
            m.dAverageEntrySnapshot = 0#
        Else
            m.dAverageEntrySnapshot = dEntrySum / Val(lEntryQuantity)
        End If
    
        CreateTradeLines m.EntriesSnapshot, Nothing, True, ForexBaseBars
        
        Save
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.AddSnapshotFillLifo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSnapshotFillFifo
'' Description: Add the fill to snapshot and recalculate the account position
''              on a first-in first-out basis
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSnapshotFillFifo(Fill As cPtFill, Optional ByVal bSaveAndClear As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSequence As Long               ' Fill sequence
    Dim Entry As cPtFill                ' Entry fill object
    Dim dFillProfit As Double           ' Fill profit
    Dim lEntryQuantity As Long          ' Sum of the entry quantities
    Dim dEntrySum As Double             ' Sum of the entry prices
    Dim lQtyLeft As Long                ' Quantity left
    Dim ForexBaseBars As cGdBars        ' Bars for the forex base pair (if applicable)
    Dim strForexBase As String          ' Forex base symbol
    Dim dExchangeFillTime As Double     ' Exchange fill time
    
    strForexBase = g.Profit.BaseForex(m.strSymbol)
    If Len(strForexBase) = 0 Then
        Set ForexBaseBars = Nothing
    Else
        Set ForexBaseBars = New cGdBars
        DM_GetBars ForexBaseBars, strForexBase
    End If
                
    If bSaveAndClear Then
        ClearOpenTradeLines
    End If
    
    With Fill
        dExchangeFillTime = ConvertBrokerDate(.FillDate, .Broker, .Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
        
        If m.dStartFillTime = 0# Then
            m.dStartFillTime = Fill.FillDate
        End If
    
        lSequence = lSequence + 1
        If m.lAutoTradeItemID = -1& Then
            .Sequence = lSequence
        Else
            .SequenceCategory = lSequence
        End If
        m.lLastFillSequenceSnapshot = lSequence
    
        If .SessionDate <> m.lSessionDateSnapshot Then
            m.lSessionDateSnapshot = .SessionDate
            m.lSessionQuantitySnapshot = .Quantity
            m.dSessionProfitSnapshot = 0#
        Else
            m.lSessionQuantitySnapshot = m.lSessionQuantitySnapshot + .Quantity
        End If
        dFillProfit = 0#
        
        m.dTotalFeesSnapshot = m.dTotalFeesSnapshot + .Fees
        
        If .FillDate > m.dLastTradedSnapshot Then m.dLastTradedSnapshot = .FillDate
        
        If .Buy Then
            m.lNumBuysSnapshot = m.lNumBuysSnapshot + .Quantity
            m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
            m.dRollingSumSnapshot = m.dRollingSumSnapshot - (.Price * .Quantity)
            
            ' Opening a new or adding to an existing position...
            If m.lCurrentPositionSnapshot >= 0& Then
                If m.lAutoTradeItemID = -1& Then
                    .Action = "E"
                    .ClosedProfit = 0#
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpen = .Quantity
                Else
                    .ActionCategory = "E"
                    .ClosedProfitCategory = 0#
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpenCategory = .Quantity
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
            
            ' Closing part or all of a position...
            ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = .Quantity To 1 Step -1
                    Set Entry = m.EntriesSnapshot(1)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            If .NumOpen <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpen + 1
                                .NumOpen = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove 1
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpen = .NumOpen - lIndex
                                .Save
                            
                                lIndex = 1
                            End If
                        Else
                            If .NumOpenCategory <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpenCategory + 1
                                .NumOpenCategory = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove 1
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpenCategory = .NumOpenCategory - lIndex
                                .Save
                                
                                lIndex = 1
                            End If
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "X"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpen = 0&
                Else
                    .ActionCategory = "X"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpenCategory = 0&
                End If
                .Save
                
            ' Reversing the position...
            Else
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                    Set Entry = m.EntriesSnapshot(1)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpen = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove 1
                        Else
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, .Price - Fill.Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpenCategory = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove 1
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "R"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                Else
                    .ActionCategory = "R"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot + .Quantity
                    .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
            End If
        
            m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot + .Quantity
        Else
            m.lNumSellsSnapshot = m.lNumSellsSnapshot + .Quantity
            m.lNumTotalSnapshot = m.lNumTotalSnapshot + .Quantity
            m.dRollingSumSnapshot = m.dRollingSumSnapshot + (.Price * .Quantity)
            
            ' Opening a new or adding to an existing position...
            If m.lCurrentPositionSnapshot <= 0& Then
                If m.lAutoTradeItemID = -1& Then
                    .Action = "E"
                    .ClosedProfit = 0#
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpen = .Quantity
                Else
                    .ActionCategory = "E"
                    .ClosedProfitCategory = 0#
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpenCategory = .Quantity
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
                
            ' Closing part or all of a position...
            ElseIf .Quantity <= Abs(m.lCurrentPositionSnapshot) Then
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = .Quantity To 1 Step -1
                    Set Entry = m.EntriesSnapshot(1)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            If .NumOpen <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpen + 1
                                .NumOpen = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove 1
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpen = .NumOpen - lIndex
                                .Save
                            
                                lIndex = 1
                            End If
                        Else
                            If .NumOpenCategory <= lIndex Then
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                lIndex = lIndex - .NumOpenCategory + 1
                                .NumOpenCategory = 0&
                                .Save
                                
                                m.EntriesSnapshot.Remove 1
                            Else
                                dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, lIndex, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                                
                                .NumOpenCategory = .NumOpenCategory - lIndex
                                .Save
                                
                                lIndex = 1
                            End If
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "X"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpen = 0&
                Else
                    .ActionCategory = "X"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpenCategory = 0&
                End If
                .Save
                
            ' Reversing the position...
            Else
                CreateTradeLines m.EntriesSnapshot, Fill, True, ForexBaseBars
                
                For lIndex = m.EntriesSnapshot.Count To 1 Step -1
                    Set Entry = m.EntriesSnapshot(1)
                    With Entry
                        If m.lAutoTradeItemID = -1& Then
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpen, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpen = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove 1
                        Else
                            dFillProfit = dFillProfit + g.Profit.Profit(m.strSymbol, Fill.Price - .Price, .NumOpenCategory, dExchangeFillTime, m.Bars, ForexBaseBars, m.lAccountID)
                            .NumOpenCategory = 0&
                            .Save
                            
                            m.EntriesSnapshot.Remove 1
                        End If
                    End With
                Next lIndex
                
                m.dSessionProfitSnapshot = m.dSessionProfitSnapshot + dFillProfit
                m.dClosedProfitSnapshot = m.dClosedProfitSnapshot + dFillProfit
                
                If m.lAutoTradeItemID = -1& Then
                    .Action = "R"
                    .ClosedProfit = dFillProfit
                    .SessionProfit = m.dSessionProfitSnapshot
                    .TotalProfit = m.dClosedProfitSnapshot
                    .SessionQuantity = m.lSessionQuantitySnapshot
                    .CurrentPosition = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpen = .Quantity - Abs(m.lCurrentPositionSnapshot)
                Else
                    .ActionCategory = "R"
                    .ClosedProfitCategory = dFillProfit
                    .SessionProfitCategory = m.dSessionProfitSnapshot
                    .TotalProfitCategory = m.dClosedProfitSnapshot
                    .SessionQuantityCategory = m.lSessionQuantitySnapshot
                    .CurrentPositionCategory = m.lCurrentPositionSnapshot - .Quantity
                    .NumOpenCategory = .Quantity - Abs(m.lCurrentPositionSnapshot)
                End If
                .Save
                
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
            End If
        
            m.lCurrentPositionSnapshot = m.lCurrentPositionSnapshot - .Quantity
        End If
    End With
    
    If bSaveAndClear = True Then
        lEntryQuantity = 0&
        dEntrySum = 0#
        For lIndex = 1 To m.EntriesSnapshot.Count
            If m.lAutoTradeItemID = -1& Then
                lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpen
                dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpen)
            Else
                lEntryQuantity = lEntryQuantity + m.EntriesSnapshot(lIndex).NumOpenCategory
                dEntrySum = dEntrySum + (m.EntriesSnapshot(lIndex).Price * m.EntriesSnapshot(lIndex).NumOpenCategory)
            End If
        Next lIndex
        If lEntryQuantity = 0& Then
            m.dAverageEntrySnapshot = 0#
        Else
            m.dAverageEntrySnapshot = dEntrySum / Val(lEntryQuantity)
        End If
        
        CreateTradeLines m.EntriesSnapshot, Nothing, True, ForexBaseBars
                
        Save
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.AddSnapshotFillFifo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalculateFromDate
'' Description: Recalculate the account position from a certain date/time
'' Inputs:      Date/Time, Fill Match Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RecalculateFromDate(ByVal dDateTime As Double)
On Error GoTo ErrSection:

    If dDateTime <= m.dLastTraded Then
        RecalculateHistory
    Else
        RecalculateSnapshot
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RecalculateFromDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Update the fill in the account position
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFill(Fill As cPtFill, Optional ByVal bFillChanged As Boolean = True)
On Error GoTo ErrSection:

    If Fill.IsSnapshot = False Then
        RecalculateHistory
    Else
        RecalculateSnapshot
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.UpdateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveFill
'' Description: Remove the fill from the account position
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveFill(Fill As cPtFill)
On Error GoTo ErrSection:

    If Fill.IsSnapshot = False Then
        RecalculateHistory
    Else
        RecalculateSnapshot
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RemoveFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenProfit
'' Description: Calculate the open profit for this position given the current price
'' Inputs:      Current Price
'' Returns:     Open Profit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenProfit(Optional ByVal dCurrentPrice As Double = -99999#) As Double
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dOpenProfit As Double           ' Open profit
    Static dLastPrice As Double         ' Last price we calculated
    
   If dCurrentPrice <> dLastPrice Then
        ' If the current price passed in is Null and the symbol is an option, set the
        ' current price to zero as if the option is expiring worthless.  02/20/2009 DAJ...
        If dCurrentPrice = kNullData Then
            If (InStr(m.strSymbol, " ") <> 0) Then
                dCurrentPrice = 0#
            End If
        End If
            
        If dCurrentPrice = kNullData Then
            m.dOpenProfit = kNullData
        ElseIf dCurrentPrice <> -99999# Then
            dOpenProfit = 0#
            If m.lCurrentPositionSnapshot > 0& Then
                For lIndex = 1 To m.EntriesSnapshot.Count
                    dOpenProfit = dOpenProfit + g.Profit.Profit(m.strSymbol, dCurrentPrice - m.EntriesSnapshot(lIndex).Price, m.EntriesSnapshot(lIndex).NumberOpen(m.lAutoTradeItemID), , m.Bars, , m.lAccountID)
                Next lIndex
                
                m.dOpenProfit = dOpenProfit
            ElseIf m.lCurrentPositionSnapshot < 0& Then
                For lIndex = 1 To m.EntriesSnapshot.Count
                    dOpenProfit = dOpenProfit + g.Profit.Profit(m.strSymbol, m.EntriesSnapshot(lIndex).Price - dCurrentPrice, m.EntriesSnapshot(lIndex).NumberOpen(m.lAutoTradeItemID), , m.Bars, , m.lAccountID)
                Next lIndex
                
                m.dOpenProfit = dOpenProfit
            Else
                m.dOpenProfit = 0#
            End If
            
            For lIndex = 1 To m.TradeLines.Count
                If m.TradeLines(lIndex).IsOpen Then
                    m.TradeLines(lIndex).OpenProfit dCurrentPrice
                End If
            Next lIndex
        
            dLastPrice = dCurrentPrice
        ElseIf m.lCurrentPositionSnapshot = 0& Then
            m.dOpenProfit = 0#
        End If
    End If
    
    OpenProfit = m.dOpenProfit

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.OpenProfit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Convert the account position information to a delimited string
'' Inputs:      None
'' Returns:     Tab Delimited String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of information to return

    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    astrReturn.Add m.lAccountPositionID
    astrReturn.Add m.lAccountID
    astrReturn.Add m.lSymbolID
    astrReturn.Add m.strSymbol
    astrReturn.Add m.lAutoTradeItemID
    astrReturn.Add m.lCurrentPosition
    astrReturn.Add m.lSessionDate
    astrReturn.Add m.lSessionQuantity
    astrReturn.Add m.dSessionProfit
    astrReturn.Add m.dAverageEntry
    astrReturn.Add m.dLastTraded
    
    ToString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFillSummary
'' Description: Convert the account position information to a fill summary string
'' Inputs:      None
'' Returns:     Tab Delimited String
''
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum,
''              Entries, ClosedProfit, AvgEntry, Initial Fill Price,
''              Initial Fill Date, Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToFillSummary() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of information to return
    Dim astrEntries As cGdArray         ' Array of entries
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object

    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    Set astrEntries = New cGdArray
    astrEntries.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.EntriesSnapshot.Count
        Set Fill = m.EntriesSnapshot(lIndex)
        For lIndex2 = 1 To Fill.Quantity
            astrEntries.Add Str(Fill.Price)
        Next lIndex2
    Next lIndex
    
    astrReturn.Add g.Broker.AccountNumberForID(m.lAccountID)
    astrReturn.Add m.strSymbol
    astrReturn.Add Str(m.lAutoTradeItemID)
    astrReturn.Add Str(m.lNumBuysSnapshot)
    astrReturn.Add Str(m.lNumSellsSnapshot)
    astrReturn.Add Str(m.lCurrentPositionSnapshot)
    astrReturn.Add Str(m.lNumTotalSnapshot)
    astrReturn.Add Str(m.dRollingSumSnapshot)
    astrReturn.Add astrEntries.JoinFields(";")
    astrReturn.Add Str(m.dClosedProfitSnapshot)
    astrReturn.Add Str(m.dAverageEntrySnapshot)
    astrReturn.Add Str(FirstOpenEntryPriceSnapshot)
    astrReturn.Add Str(FirstOpenEntryDateSnapshot)
    astrReturn.Add Str(m.lSessionDateSnapshot)
    astrReturn.Add Str(m.dLastTradedSnapshot)
    astrReturn.Add Str(m.lCurrentPosition)
    
    ToFillSummary = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.ToFillSummary"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstOpenEntryPrice
'' Description: Determine the price for the first open non-snapshot entry
'' Inputs:      None
'' Returns:     Price for first open non-snapshot entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstOpenEntryPrice() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double
    
    dReturn = 0#
    If Not m.Entries Is Nothing Then
        If m.Entries.Count > 0 Then
            dReturn = m.Entries(1).Price
        End If
    End If
    
    FirstOpenEntryPrice = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.FirstOpenEntryPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstOpenEntryDate
'' Description: Determine the date/time for the first open non-snapshot entry
'' Inputs:      Convert to Exchange Time?
'' Returns:     Date/Time for first open non-snapshot entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstOpenEntryDate(Optional ByVal bConvertToExchTime As Boolean = False) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double
    
    dReturn = 0#
    If Not m.Entries Is Nothing Then
        If m.Entries.Count > 0 Then
            dReturn = m.Entries(1).FillDate
            If bConvertToExchTime Then
                dReturn = ConvertBrokerDate(dReturn, m.nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            End If
        End If
    End If
    
    FirstOpenEntryDate = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.FirstOpenEntryDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstOpenEntryPriceSnapshot
'' Description: Determine the price for the first open snapshot entry
'' Inputs:      None
'' Returns:     Price for first open snapshot entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstOpenEntryPriceSnapshot() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double
    
    dReturn = 0#
    If Not m.EntriesSnapshot Is Nothing Then
        If m.EntriesSnapshot.Count > 0 Then
            dReturn = m.EntriesSnapshot(1).Price
        End If
    End If
    
    FirstOpenEntryPriceSnapshot = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.FirstOpenEntryPriceSnapshot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstOpenEntryDateSnapshot
'' Description: Determine the date/time for the first open snapshot entry
'' Inputs:      Convert to Exchange Time?
'' Returns:     Date/Time for first open snapshot entry
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstOpenEntryDateSnapshot(Optional ByVal bConvertToExchTime As Boolean = False) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double
    
    dReturn = 0#
    If Not m.EntriesSnapshot Is Nothing Then
        If m.EntriesSnapshot.Count > 0 Then
            dReturn = m.EntriesSnapshot(1).FillDate
            If bConvertToExchTime Then
                dReturn = ConvertBrokerDate(dReturn, m.nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            End If
        End If
    End If
    
    FirstOpenEntryDateSnapshot = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.FirstOpenEntryDateSnapshot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstOpenEntrySnapshot
'' Description: Determine the first open snapshot entry
'' Inputs:      None
'' Returns:     First open snapshot entry (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstOpenEntrySnapshot() As cPtFill
On Error GoTo ErrSection:

    Dim FirstEntry As cPtFill           ' Return value for the function
    
    If m.EntriesSnapshot.Count = 0 Then
        Set FirstEntry = Nothing
    Else
        Set FirstEntry = m.EntriesSnapshot(1)
    End If
    
    Set FirstOpenEntrySnapshot = FirstEntry

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.FirstOpenEntrySnapshot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastOpenEntrySnapshot
'' Description: Determine the last open snapshot entry
'' Inputs:      None
'' Returns:     Last open snapshot entry (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastOpenEntrySnapshot() As cPtFill
On Error GoTo ErrSection:

    Dim LastEntry As cPtFill            ' Return value for the function
    
    If m.EntriesSnapshot.Count = 0 Then
        Set LastEntry = Nothing
    Else
        Set LastEntry = m.EntriesSnapshot(m.EntriesSnapshot.Count)
    End If
    
    Set LastOpenEntrySnapshot = LastEntry

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.LastOpenEntrySnapshot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastExit
'' Description: Retrieve the most recent exit
'' Inputs:      None
'' Returns:     Most recent exit (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastExit() As cPtFill
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Fill = Nothing
    For lIndex = m.TradeLines.Count To 1 Step -1
        If m.TradeLines(lIndex).ExitFillID <> 0& Then
            Set Fill = New cPtFill
            If Fill.Load(m.TradeLines(lIndex).ExitFillID) = False Then
                Set Fill = Nothing
            End If
            
            Exit For
        End If
    Next lIndex
    
    Set LastExit = Fill

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.LastClosedTradeLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReversePosition
'' Description: Reverse the users current position for this symbol/account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReversePosition()
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object
    Dim strAccount As String            ' Account name
    Dim nBroker As eTT_AccountType      ' Broker for the account

    strAccount = g.Broker.AccountNameForID(m.lAccountID)
    nBroker = g.Broker.AccountTypeForID(m.lAccountID)
    
    If g.Broker.PositionMatch(m.lAccountID, SymbolOrSymbolID) = True Then
        If m.lCurrentPositionSnapshot <> 0& Then
            Set Order = CreateMarketOrder(m.lAccountID, SymbolOrSymbolID, (m.lCurrentPositionSnapshot < 0), Abs(m.lCurrentPositionSnapshot) * 2, m.lAutoTradeItemID, True)
            Order.Save
            
            g.Broker.BrokerDebug nBroker, "cAccountPosition.ReversePosition(" & strAccount & ", " & m.strSymbol & ", " & Str(m.lAutoTradeItemID) & ")"
            g.Broker.SendOrder Order
        End If
    Else
        g.Broker.BrokerDebug nBroker, "cAccountPosition.ReversePosition(" & strAccount & ", " & m.strSymbol & ", " & Str(m.lAutoTradeItemID) & "): Aborted due to Position Mismatch"
        InfBox "Trade Navigator has received inconsistent data from the broker for " & m.strSymbol & " in account '" & strAccount & "' and therefore cannot perform a Reverse on the position.||PLEASE CALL YOUR BROKER IMMEDIATELY TO VERIFY YOUR POSITION IN THIS ACCOUNT.||", "!", , "Inconsistent Broker Data"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.Reverse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasCarPosFixEntries
'' Description: Are any of the entry fills Carried Position Fixes?
'' Inputs:      None
'' Returns:     True if has carried position fixes, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasCarPosFixEntries() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.Entries.Count
        If UCase(Parse(m.Entries(lIndex).BrokerID, " ", 1)) = "CARPOSFIX" Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasCarPosFixEntries = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.HasCarPosFixEntries"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTradeLine
'' Description: Get the trade line for the given trade number
'' Inputs:      Trade Number
'' Returns:     Trade Line (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTradeLine(ByVal lTradeNumber As Long) As cTradeLine
On Error GoTo ErrSection:

    Dim TradeLine As cTradeLine         ' Trade Line to return
    
    If lTradeNumber >= 1 And lTradeNumber <= m.TradeLines.Count Then
        Set TradeLine = m.TradeLines(lTradeNumber)
    Else
        Set TradeLine = Nothing
    End If
    
    Set GetTradeLine = TradeLine

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.GetTradeLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradeLine
'' Description: Refresh the given trade line
'' Inputs:      Trade Line
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshTradeLine(ByVal TradeLine As cTradeLine)
On Error GoTo ErrSection:

    m.TradeLines(TradeLine.TradeNumber) = TradeLine

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.RefreshTradeLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPrice
'' Description: Determine the current price to use for open profit
'' Inputs:      Last Price, Bid, Ask
'' Returns:     Current Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentPrice(ByVal dLast As Double, ByVal dBid As Double, ByVal dAsk As Double) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function

    dReturn = dLast
    
    If (InStr(m.strSymbol, " ") <> 0) And (g.Broker.OptionOpenEquity = eGDOptionOpenEquity_UseBidAsk) Then
        If m.lCurrentPositionSnapshot > 0 Then
            If dBid <> kNullData Then
                dReturn = dBid
            End If
        ElseIf m.lCurrentPositionSnapshot < 0 Then
            If dAsk <> kNullData Then
                dReturn = dAsk
            End If
        End If
    End If
    
    CurrentPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountPosition.CurrentPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear out the information
'' Inputs:      Just Snapshot?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Clear(Optional ByVal bJustSnapshot As Boolean = True)
On Error GoTo ErrSection:

    If bJustSnapshot = False Then
        m.dStartFillTime = 0#
        
        m.lNumBuys = 0&
        m.lNumSells = 0&
        m.lCurrentPosition = 0&
        m.lNumTotal = 0&
        m.dRollingSum = 0#
        m.dClosedProfit = 0#
        m.dAverageEntry = 0#
        m.dLastTraded = 0#
        m.dTotalFees = 0#
        m.lLastFillSequence = 0&
        
        m.lSessionDate = 0&
        m.lSessionQuantity = 0&
        m.dSessionProfit = 0#
    
        m.Entries.Clear
        m.TradeLines.Clear
    End If

    m.lNumBuysSnapshot = 0&
    m.lNumSellsSnapshot = 0&
    m.lCurrentPositionSnapshot = 0&
    m.lNumTotalSnapshot = 0&
    m.dRollingSumSnapshot = 0#
    m.dClosedProfitSnapshot = 0#
    m.dAverageEntrySnapshot = 0#
    m.dLastTradedSnapshot = 0#
    m.dTotalFeesSnapshot = 0#
    m.lLastFillSequenceSnapshot = 0&
    
    m.EntriesSnapshot.Clear
    m.TradeLines.ClearSnapshot
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadEntries
'' Description: Load the open entry arrays from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadEntries()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Fill As cPtFill                 ' Fill object to load and store in collection

    ' Clear the entries collections...
    m.Entries.Clear
    m.EntriesSnapshot.Clear
    
    ' Retrieve information from the database for the entry fills...
    Set rs = g.dbPaper.OpenRecordset("SELECT tblFills.*, tblAccountPositionEntries.IsSnapshot AS ApIsSnapshot, tblAccountPositionEntries.AccountPositionID " & _
                "FROM tblAccountPositionEntries INNER JOIN tblFills ON tblAccountPositionEntries.FillID=tblFills.FillID " & _
                "WHERE tblAccountPositionEntries.AccountPositionID=" & Str(m.lAccountPositionID) & " " & _
                "ORDER BY [FillSequence];", dbOpenDynaset)
    Do While Not rs.EOF
        Set Fill = New cPtFill
        If Fill.Load(rs!FillID, rs, m.Bars, m.nBroker) Then
            If rs!ApIsSnapshot = False Then
                m.Entries.Add Fill, Str(Fill.FillID)
            Else
                m.EntriesSnapshot.Add Fill, Str(Fill.FillID)
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.LoadEntries"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveEntries
'' Description: Save the open entry arrays to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveEntries()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    
    ' First, remove all of the current entries from the database...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionEntries] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
    ' Second, add all of the current entries back in...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositionEntries] " & _
                "WHERE [AccountPositionID]=" & Str(m.lAccountPositionID) & ";", dbOpenDynaset)
    For lIndex = 1 To m.Entries.Count
        Set Fill = m.Entries(lIndex)
        With Fill
            rs.AddNew
            rs!AccountPositionID = m.lAccountPositionID
            rs!FillID = .FillID
            If m.lAutoTradeItemID = -1& Then
                rs!FillSequence = .Sequence
            Else
                rs!FillSequence = .SequenceCategory
            End If
            rs!FillQuantity = .Quantity
            rs!FillPrice = .Price
            rs!IsSnapshot = False
            rs.Update
        End With
    Next lIndex
    For lIndex = 1 To m.EntriesSnapshot.Count
        Set Fill = m.EntriesSnapshot(lIndex)
        With Fill
            rs.AddNew
            rs!AccountPositionID = m.lAccountPositionID
            rs!FillID = .FillID
            If m.lAutoTradeItemID = -1& Then
                rs!FillSequence = .Sequence
            Else
                rs!FillSequence = .SequenceCategory
            End If
            rs!FillQuantity = .Quantity
            rs!FillPrice = .Price
            rs!IsSnapshot = True
            rs.Update
        End With
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.SaveEntries"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateTradeLines
'' Description: Create trade line(s) from the given information
'' Inputs:      Entry Fills, Exit Fill, Snapshot?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateTradeLines(Entries As cGdTree, ExitFill As cPtFill, ByVal bIsSnapshot As Boolean, Optional BaseForexBars As cGdBars = Nothing)
On Error GoTo ErrSection:
    
gdStartProfile 610
gdStartProfile 611
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeLine As cTradeLine         ' Trade Line object to add to collection
    Dim Fill As cPtFill                 ' Fill object
    Dim lExitQuantity As Long           ' Quantity left to exit
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim dExchangeExitTime As Double     ' Exit fill time in exchange time zone
    
    nBroker = g.Broker.AccountTypeForID(m.lAccountID)

gdStopProfile 611

    If Not ExitFill Is Nothing Then
gdStartProfile 612
        lExitQuantity = ExitFill.Quantity
        dExchangeExitTime = ConvertBrokerDate(ExitFill.FillDate, ExitFill.Broker, ExitFill.Symbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 612
        
        For lIndex = 1 To Entries.Count
gdStartProfile 613
            If m.nFillMatchMode = eTT_FillMatchMode_Fifo Then
                Set Fill = Entries(lIndex)
            ElseIf m.nFillMatchMode = eTT_FillMatchMode_Lifo Then
                Set Fill = Entries(Entries.Count - lIndex + 1)
            End If
gdStopProfile 613
            If lExitQuantity <= Fill.NumberOpen(m.lAutoTradeItemID) Then
gdStartProfile 614
                Set TradeLine = New cTradeLine
gdStopProfile 614
                With TradeLine
gdStartProfile 615
                    .AccountPositionID = m.lAccountPositionID
                    .SetAccountID m.lAccountID, m.nBroker
                    .TradeNumber = m.TradeLines.Count + 1
                    .SetSymbolOrSymbolID m.lSymbolID, m.strSymbol, m.Bars
                    
                    If Fill.Buy Then .Direction = "L" Else .Direction = "S"
                    .Quantity = lExitQuantity
                    .EntryFillID = Fill.FillID
                    .EntryBrokerFillID = Fill.BrokerID
                    .EntryPrice = Fill.Price
gdStartProfile 618
                    .EntryTime = ConvertBrokerDate(Fill.FillDate, nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 618
                    .ExitFillID = ExitFill.FillID
                    .ExitBrokerFillID = ExitFill.BrokerID
                    .ExitPrice = ExitFill.Price
gdStartProfile 618
                    .ExitTime = ConvertBrokerDate(ExitFill.FillDate, nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 618
gdStartProfile 619
                    If Fill.Buy Then
                        .ClosedProfit = g.Profit.Profit(m.strSymbol, ExitFill.Price - Fill.Price, .Quantity, dExchangeExitTime, m.Bars, BaseForexBars, m.lAccountID)
                    Else
                        .ClosedProfit = g.Profit.Profit(m.strSymbol, Fill.Price - ExitFill.Price, .Quantity, dExchangeExitTime, m.Bars, BaseForexBars, m.lAccountID)
                    End If
gdStopProfile 619
                    .EntryRuleID = Fill.EntryRuleIdCategory
                    .ExitRuleID = ExitFill.ExitRuleIdCategory
                    .RealSimFlag = ExitFill.RealSimFlagCategory
                    .IsOpen = False
                    .IsSnapshot = bIsSnapshot
                    
                    .Fees = ((Fill.Fees / Fill.Quantity) * .Quantity) + ((ExitFill.Fees / ExitFill.Quantity) * .Quantity)
gdStopProfile 615
                End With
                
gdStartProfile 616
                m.TradeLines.Add TradeLine
gdStopProfile 616
gdStartProfile 617
                lExitQuantity = 0&
gdStopProfile 617
            Else
gdStartProfile 614
                Set TradeLine = New cTradeLine
gdStopProfile 614
                With TradeLine
gdStartProfile 615
                    .AccountPositionID = m.lAccountPositionID
                    .SetAccountID m.lAccountID, m.nBroker
                    .TradeNumber = m.TradeLines.Count + 1
                    .SetSymbolOrSymbolID m.lSymbolID, m.strSymbol, m.Bars
                    
                    If Fill.Buy Then .Direction = "L" Else .Direction = "S"
                    .Quantity = Fill.NumberOpen(m.lAutoTradeItemID)
                    .EntryFillID = Fill.FillID
                    .EntryBrokerFillID = Fill.BrokerID
                    .EntryPrice = Fill.Price
gdStartProfile 618
                    .EntryTime = ConvertBrokerDate(Fill.FillDate, nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 618
                    .ExitFillID = ExitFill.FillID
                    .ExitBrokerFillID = ExitFill.BrokerID
                    .ExitPrice = ExitFill.Price
gdStartProfile 618
                    .ExitTime = ConvertBrokerDate(ExitFill.FillDate, nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 618
gdStartProfile 619
                    If Fill.Buy Then
                        .ClosedProfit = g.Profit.Profit(m.strSymbol, ExitFill.Price - Fill.Price, .Quantity, dExchangeExitTime, m.Bars, BaseForexBars, m.lAccountID)
                    Else
                        .ClosedProfit = g.Profit.Profit(m.strSymbol, Fill.Price - ExitFill.Price, .Quantity, dExchangeExitTime, m.Bars, BaseForexBars, m.lAccountID)
                    End If
gdStopProfile 619
                    .EntryRuleID = Fill.EntryRuleIdCategory
                    .ExitRuleID = ExitFill.ExitRuleIdCategory
                    .RealSimFlag = ExitFill.RealSimFlagCategory
                    .IsOpen = False
                    .IsSnapshot = bIsSnapshot
                    
                    .Fees = ((Fill.Fees / Fill.Quantity) * .Quantity) + ((ExitFill.Fees / ExitFill.Quantity) * .Quantity)
gdStopProfile 615
                End With
                
gdStartProfile 616
                m.TradeLines.Add TradeLine
gdStopProfile 616
gdStartProfile 617
                lExitQuantity = lExitQuantity - Fill.NumberOpen(m.lAutoTradeItemID)
gdStopProfile 617
            End If
            
            If lExitQuantity = 0& Then Exit For
        Next lIndex
    ElseIf m.lCurrentPositionSnapshot <> 0& Then
        For lIndex = 1 To Entries.Count
gdStartProfile 613
            Set Fill = Entries(lIndex)
gdStopProfile 613
            
gdStartProfile 614
            Set TradeLine = New cTradeLine
gdStopProfile 614
            With TradeLine
gdStartProfile 615
                .AccountPositionID = m.lAccountPositionID
                .SetAccountID m.lAccountID, m.nBroker
                .TradeNumber = m.TradeLines.Count + 1
                .SetSymbolOrSymbolID m.lSymbolID, m.strSymbol, m.Bars
                
                If Fill.Buy Then .Direction = "L" Else .Direction = "S"
                .Quantity = Fill.NumberOpen(m.lAutoTradeItemID)
                .EntryFillID = Fill.FillID
                .EntryBrokerFillID = Fill.BrokerID
                .EntryPrice = Fill.Price
gdStartProfile 618
                .EntryTime = ConvertBrokerDate(Fill.FillDate, nBroker, m.strSymbol, False, m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
gdStopProfile 618
                .ExitFillID = 0&
                .ExitBrokerFillID = ""
                .ExitPrice = 0#
                .ExitTime = 0#
                .ClosedProfit = 0#
                .EntryRuleID = Fill.EntryRuleIdCategory
                .ExitRuleID = 0&
                .RealSimFlag = 0
                .IsOpen = True
                .IsSnapshot = True
                
                .Fees = (Fill.Fees / Fill.Quantity) * .Quantity
gdStopProfile 615
            End With
            
gdStartProfile 616
            m.TradeLines.Add TradeLine
gdStopProfile 616
        Next lIndex
    End If
    
gdStopProfile 610

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.CreateTradeLines"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearOpenTradeLines
'' Description: Clear any trade lines that are currently marked as open
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearOpenTradeLines()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = m.TradeLines.Count To 1 Step -1
        If m.TradeLines(lIndex).IsOpen Then
            m.TradeLines.Remove lIndex
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.ClearOpenTradeLines"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildSortKey
'' Description: Build the sort key for the object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildSortKey()
On Error GoTo ErrSection:

    m.strSortKey = Pad(m.strSymbol, 20, "L") & Pad(g.Broker.AccountNameForID(m.lAccountID), 50, "L") & Str(m.lAutoTradeItemID)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.BuildSortKey"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Set m.Entries = New cGdTree
    Set m.EntriesSnapshot = New cGdTree
    
    Set m.TradeLines = New cTradeLines
    
    Set m.Bars = New cGdBars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Entries = Nothing
    Set m.EntriesSnapshot = Nothing
    
    Set m.TradeLines = Nothing
    
    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountPosition.Class_Terminate"
    
End Sub
