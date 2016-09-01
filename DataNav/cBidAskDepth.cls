VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBidAskDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Bid As New cGdTable
    Ask As New cGdTable
    BidAsk As New cGdTable
    
    aBestBids As New cGdArray
    aBestAsks As New cGdArray
    
    aUnclaimedData As New cGdArray  'store data passed by real time object when no form has requested data
    
'new tables for salmon
    BidSalmon As New cGdTable
    AskSalmon As New cGdTable
    BidAskSalmon As New cGdTable
    
    aBestBidSalmon As New cGdArray
    aBestAskSalmon As New cGdArray
    
    dMinMove As Double
    dLastNewData As Double          'gdtickcount of when data was last "refreshed" via
    dNewestData As Double           'actual timestamp of most recent bid/ask data
    
    nSessionDate As Long
    nUsageCounter As Long
    
    strSymbol As String
    bIsSpreadSymbol As Boolean
    bDirtyBidAsk As Boolean
    bDumpFileSalmon As Boolean
End Type

Dim m As mPrivate


Private Sub Class_Initialize()

    'create fields for bid data table
    m.Bid.CreateField eGDARRAY_Strings, 0, "MarketMaker"
    m.Bid.CreateField eGDARRAY_Doubles, 1, "Price"
    m.Bid.CreateField eGDARRAY_Longs, 2, "Size"
    m.Bid.CreateField eGDARRAY_Doubles, 3, "Time"
    m.Bid.CreateField eGDARRAY_Strings, 4, "ActiveFlag"
    m.Bid.CreateField eGDARRAY_Strings, 5, "UpDownMarker"

    'create fields for ask data table
    m.Ask.CreateField eGDARRAY_Strings, 0, "MarketMaker"
    m.Ask.CreateField eGDARRAY_Doubles, 1, "Price"
    m.Ask.CreateField eGDARRAY_Longs, 2, "Size"
    m.Ask.CreateField eGDARRAY_Doubles, 3, "Time"
    m.Ask.CreateField eGDARRAY_Strings, 4, "ActiveFlag"
    m.Ask.CreateField eGDARRAY_Strings, 5, "UpDownMarker"
        
    'create fields for combined bid & ask table ordered by price
    m.BidAsk.CreateField eGDARRAY_Doubles, 0, "Price"
    m.BidAsk.CreateField eGDARRAY_Longs, 1, "BidSize"
    m.BidAsk.CreateField eGDARRAY_Longs, 2, "AskSize"
        
    'create array for best 5 bids
    m.aBestBids.Create eGDARRAY_Doubles, 5

    'create array for best 5 asks
    m.aBestAsks.Create eGDARRAY_Doubles, 5

'for salmon
    'create fields for bid data table
    m.BidSalmon.CreateField eGDARRAY_Strings, 0, "MarketMaker"
    m.BidSalmon.CreateField eGDARRAY_Doubles, 1, "Price"
    m.BidSalmon.CreateField eGDARRAY_Longs, 2, "Size"
    m.BidSalmon.CreateField eGDARRAY_Doubles, 3, "Time"
    m.BidSalmon.CreateField eGDARRAY_Strings, 4, "ActiveFlag"
    m.BidSalmon.CreateField eGDARRAY_Strings, 5, "UpDownMarker"

    'create fields for ask data table
    m.AskSalmon.CreateField eGDARRAY_Strings, 0, "MarketMaker"
    m.AskSalmon.CreateField eGDARRAY_Doubles, 1, "Price"
    m.AskSalmon.CreateField eGDARRAY_Longs, 2, "Size"
    m.AskSalmon.CreateField eGDARRAY_Doubles, 3, "Time"
    m.AskSalmon.CreateField eGDARRAY_Strings, 4, "ActiveFlag"
    m.AskSalmon.CreateField eGDARRAY_Strings, 5, "UpDownMarker"
        
    'create fields for combined bid & ask table ordered by price
    m.BidAskSalmon.CreateField eGDARRAY_Doubles, 0, "Price"
    m.BidAskSalmon.CreateField eGDARRAY_Longs, 1, "BidSize"
    m.BidAskSalmon.CreateField eGDARRAY_Longs, 2, "AskSize"
        
    'create array for best 5 bids
    m.aBestBidSalmon.Create eGDARRAY_Doubles, 5

    'create array for best 5 asks
    m.aBestAskSalmon.Create eGDARRAY_Doubles, 5
End Sub

Private Sub Class_Terminate()
    
    'tables
    Set m.Bid = Nothing
    Set m.Ask = Nothing
    Set m.BidAsk = Nothing
    
    'gdArrays
    Set m.aBestBids = Nothing
    Set m.aBestAsks = Nothing
    Set m.aUnclaimedData = Nothing

'for salomn
    'tables
    Set m.BidSalmon = Nothing
    Set m.AskSalmon = Nothing
    Set m.BidAskSalmon = Nothing
    
    'gdArrays
    Set m.aBestBidSalmon = Nothing
    Set m.aBestAskSalmon = Nothing

End Sub

Public Function Update(dLastUpdate As Double) As Boolean
On Error GoTo ErrSection:

    If g.RealTime.SalmonIsRunning Then
        Update = NewQuoteTable()
    ElseIf dLastUpdate < m.dLastNewData Then
        dLastUpdate = m.dLastNewData
        Update = True
    Else
        Update = False
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBidAskDepth.Update", eGDRaiseError_Raise

End Function

Private Function NewQuoteTable() As Boolean
On Error GoTo ErrSection:
    
'JM 08-09-2011: New method simply request bid/ask tables from salmon client dll
'   without sorting & comparing prices before putting it into the local bid/ask tables

    If g.RealTime.SalmonIsRunning Then
        If g.RealTime.SymbolInfo(m.strSymbol).UpdateDOM(m.BidSalmon, m.AskSalmon) Then
            If m.BidSalmon Is Nothing Or m.AskSalmon Is Nothing Then
                DebugLog m.strSymbol & " Salmon returned null pointers for UpdateDOM."
            Else
                NewQuoteTable = True
                UpdateBest True
                m.bDirtyBidAsk = True
            End If
        End If
    End If

ErrExit:
    If m.bDumpFileSalmon Then DumpDepthData
    Exit Function

ErrSection:
    RaiseError "cBidAskDepth.NewQuoteTable", eGDRaiseError_Raise

End Function

Public Sub NewQuote(ByVal strSymbol$, ByVal strMaker$, ByVal dTime#, _
        ByVal dBid#, ByVal nBidSize&, ByVal dAsk#, ByVal nAskSize#, ByVal strActiveFlag$)
On Error GoTo ErrSection:

    Static aIdxBid As cGdArray
    Static aIdxAsk As cGdArray
    Static iCount&
    
    Dim i&, j&, k&
    
    'precautionary
    If strSymbol <> m.strSymbol Then Exit Sub
    
    If g.RealTime.SalmonIsRunning Then
        m.bDirtyBidAsk = NewQuoteTable()
    Else
        ' TLB 1/22/2013: non-salmon GenesisRT doesn't yet support negative bid/ask for spread symbols
        If dBid <= 0 Then dBid = kNullData
        If dAsk <= 0 Then dAsk = kNullData
    
        'JM 08-09-2011: Old method processes strings and put in local table at appropriate locations.
        If aIdxBid Is Nothing Then
            Set aIdxBid = m.Bid.CreateSortedIndex(0)
        ElseIf aIdxBid.Size <> m.Bid.NumRecords Then
            Set aIdxBid = m.Bid.CreateSortedIndex(0)
        End If
        If aIdxAsk Is Nothing Then
            Set aIdxAsk = m.Ask.CreateSortedIndex(0)
        ElseIf aIdxAsk.Size <> m.Ask.NumRecords Then
            Set aIdxAsk = m.Ask.CreateSortedIndex(0)
        End If
        
        dTime = RoundToSecond(dTime)
        If m.dMinMove > 0 Then
            dBid = RoundToMinMove(dBid, m.dMinMove)
            dAsk = RoundToMinMove(dAsk, m.dMinMove)
        End If
   
        If Len(strMaker) > 0 Then
            'bid data
            If m.Bid.SearchAsIndex(aIdxBid, 0, strMaker, j) Then
                SetTableRow m.Bid, strMaker, dBid, nBidSize, -1, dTime, strActiveFlag, aIdxBid(j), True
            Else
                m.Bid.AddRecord ""
                ZeroTableRow m.Bid, m.Bid.NumRecords - 1
                SetTableRow m.Bid, strMaker, dBid, nBidSize, 0, dTime, strActiveFlag, m.Bid.NumRecords - 1, True
                Set aIdxBid = m.Bid.CreateSortedIndex(0)
            End If
            'ask data
            If m.Ask.SearchAsIndex(aIdxAsk, 0, strMaker, j) Then
                SetTableRow m.Ask, strMaker, dAsk, nAskSize, -1, dTime, strActiveFlag, aIdxAsk(j), True
            Else
                m.Ask.AddRecord ""
                ZeroTableRow m.Ask, m.Ask.NumRecords - 1
                SetTableRow m.Ask, strMaker, dAsk, nAskSize, -1, dTime, strActiveFlag, m.Ask.NumRecords - 1, True
                Set aIdxAsk = m.Ask.CreateSortedIndex(0)
            End If
            If dTime > m.dNewestData Then m.dNewestData = dTime
        End If
        
        UpdateBest False
        m.bDirtyBidAsk = True
    End If
    
    m.dLastNewData = gdTickCount
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.NewQuote", eGDRaiseError_Raise

End Sub

Private Sub UpdateByMakers(aNewData As cGdArray)
On Error GoTo ErrSection:

    Static aIdxBid As cGdArray
    Static aIdxAsk As cGdArray
    
    Dim i&, j&, k&
    Dim dBid#, dAsk#, nBidSize&, nAskSize&, dTime#, dPrev#
    Dim strMaker$, strPrevMaker$, strActiveFlag$
    
    If aIdxBid Is Nothing Then
        Set aIdxBid = m.Bid.CreateSortedIndex(0)
    ElseIf aIdxBid.Size <> m.Bid.NumRecords Then
        Set aIdxBid = m.Bid.CreateSortedIndex(0)
    End If
    If aIdxAsk Is Nothing Then
        Set aIdxAsk = m.Ask.CreateSortedIndex(0)
    ElseIf aIdxAsk.Size <> m.Ask.NumRecords Then
        Set aIdxAsk = m.Ask.CreateSortedIndex(0)
    End If
   
    For i = 0 To aNewData.Size - 1
        ParseDataString strMaker, dBid, nBidSize, dAsk, nAskSize, dTime, strActiveFlag, aNewData(i)
        If Len(strMaker) > 0 Then
            'bid data
            If m.Bid.SearchAsIndex(aIdxBid, 0, strMaker, j) Then
                SetTableRow m.Bid, strMaker, dBid, nBidSize, -1, dTime, strActiveFlag, aIdxBid(j), True
            Else
                m.Bid.AddRecord ""
                ZeroTableRow m.Bid, m.Bid.NumRecords - 1
                SetTableRow m.Bid, strMaker, dBid, nBidSize, 0, dTime, strActiveFlag, m.Bid.NumRecords - 1, True
                Set aIdxBid = m.Bid.CreateSortedIndex(0)
            End If
            'ask data
            If m.Ask.SearchAsIndex(aIdxAsk, 0, strMaker, j) Then
                SetTableRow m.Ask, strMaker, dAsk, nAskSize, -1, dTime, strActiveFlag, aIdxAsk(j), True
            Else
                m.Ask.AddRecord ""
                ZeroTableRow m.Ask, m.Ask.NumRecords - 1
                SetTableRow m.Ask, strMaker, dAsk, nAskSize, -1, dTime, strActiveFlag, m.Ask.NumRecords - 1, True
                Set aIdxAsk = m.Ask.CreateSortedIndex(0)
            End If
            If dTime > m.dNewestData Then m.dNewestData = dTime
        End If
    Next
    
    'save best 5 bids
    Dim aIdxBestBidAsk As cGdArray
    Set aIdxBestBidAsk = m.Bid.CreateSortedIndex(1, eGdSort_Descending)
    k = 0
    dPrev = kNullData
    For i = 0 To aIdxBestBidAsk.Size - 1
        If k >= 5 Then Exit For
        dBid = m.Bid.Num(1, aIdxBestBidAsk.Num(i))
        If dBid <> kNullData And dBid <> dPrev Then
            If m.Bid.Num(2, aIdxBestBidAsk.Num(i)) <> kNullData Then
                m.aBestBids.Num(k) = dBid
                dPrev = dBid
                k = k + 1
            Else
                'do nothing - YM,ZB is coming accross with @AMEX
                'StatusMsg "Bid: " & m.Bid(0, aIdxBestBidAsk(i))
            End If
        End If
    Next
        
    'save best 5 asks
    Set aIdxBestBidAsk = m.Ask.CreateSortedIndex(1)
    k = 0
    dPrev = kNullData
    For i = 0 To aIdxBestBidAsk.Size - 1
        If k >= 5 Then Exit For
        dAsk = m.Ask.Num(1, aIdxBestBidAsk.Num(i))
        If dAsk <> kNullData And dAsk <> dPrev Then
            If m.Ask.Num(2, aIdxBestBidAsk.Num(i)) <> kNullData Then
                m.aBestAsks.Num(k) = dAsk
                dPrev = dAsk
                k = k + 1
            Else
                'do nothing - YM,ZB is coming accross with @AMEX
                'StatusMsg "Ask: " & m.Ask(0, aIdxBestBidAsk(i))
            End If
        End If
    Next
        
    m.bDirtyBidAsk = True
    m.dLastNewData = gdTickCount
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.UpdateByMakers", eGDRaiseError_Raise

End Sub


Private Sub UpdateCombined(tbData As cGdTable, ByVal nFlag&, ByVal bSalmonTables As Boolean)
On Error GoTo ErrSection:

    Dim dPrevPrice#, strMaker$
    Dim dPrice#, nSize&, nNotUsed&, nTotal&
    
    Dim i&, k&
    
    Dim aIdx As cGdArray
    Dim tbBidAsk As cGdTable
    
    dPrevPrice = kNullData
    
    If tbData Is m.BidSalmon Or tbData Is m.AskSalmon Then
        Set tbBidAsk = m.BidAskSalmon
    ElseIf tbData Is m.Bid Or tbData Is m.Ask Then
        Set tbBidAsk = m.BidAsk
    Else
        Exit Sub
    End If
    
     'sort passed in data table by price
    Set aIdx = tbData.CreateSortedIndex(1)
   
   'nFlag: 0=bid table was passed in
    '       1=ask table was passed
    'clear out combined bid/ask price table
    For i = 0 To tbBidAsk.NumRecords - 1
        If nFlag = 0 Then
            tbBidAsk.Num(1, i) = kNullData
        Else
            tbBidAsk.Num(2, i) = kNullData
        End If
    Next
    
    'sum size by price then add to combined bid/ask table
    nTotal = 0
    For i = 0 To aIdx.Size - 1
        GetTableRow tbData, strMaker, dPrice, nSize, nNotUsed, aIdx(i), True
        If Len(strMaker) > 0 And dPrice <> kNullData Then    'this can happen on start-up when prices come through without
            If nSize < 0 Then nSize = 0             'sizes or vice-versa (tables are not initialized to zeros)
            If dPrice = dPrevPrice Then
                nTotal = nTotal + nSize
            Else
                If dPrevPrice <> kNullData Then
                    k = SearchForPrice(dPrevPrice, bSalmonTables)
                    If k >= 0 Then
                        If nFlag = 0 Then
                            SetTableRow tbBidAsk, "", dPrevPrice, nTotal, -1, 0, "", k, False
                        Else
                            SetTableRow tbBidAsk, "", dPrevPrice, -1, nTotal, 0, "", k, False
                        End If
                    Else
                        tbBidAsk.AddRecord ""
                        If nFlag = 0 Then
                            SetTableRow tbBidAsk, "", dPrevPrice, nTotal, 0, 0, "", tbBidAsk.NumRecords - 1, False
                        Else
                            SetTableRow tbBidAsk, "", dPrevPrice, 0, nTotal, 0, "", tbBidAsk.NumRecords - 1, False
                        End If
                    End If
                End If
                nTotal = nSize
            End If
            dPrevPrice = dPrice
        End If
    Next
    
    'save last price from passed in table to combined bid/ask table
    If dPrice <> kNullData Then
        k = SearchForPrice(dPrice, bSalmonTables)
        If k >= 0 Then
            If nFlag = 0 Then
                SetTableRow tbBidAsk, "", dPrice, nTotal, -1, 0, "", k, False
            Else
                SetTableRow tbBidAsk, "", dPrice, -1, nTotal, 0, "", k, False
            End If
        Else
            tbBidAsk.AddRecord ""
            If nFlag = 0 Then
                SetTableRow tbBidAsk, "", dPrice, nTotal, 0, 0, "", tbBidAsk.NumRecords - 1, False
            Else
                SetTableRow tbBidAsk, "", dPrice, 0, nTotal, 0, "", tbBidAsk.NumRecords - 1, False
            End If
        End If
    End If
    
    aIdx.Destroy
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBidAskDepth.UpdateCombined", eGDRaiseError_Raise

End Sub

Private Sub GetTableRow(DataTable As cGdTable, _
    strMaker$, dPrice#, nCount&, nCount2&, ByVal nRow&, _
    ByVal bMakerData As Boolean)
On Error GoTo ErrSection:
    
    If nRow < 0 Or nRow >= DataTable.NumRecords Then Exit Sub
    
    If bMakerData Then
        strMaker = DataTable(0, nRow)
        dPrice = DataTable.Num(1, nRow)
        nCount = DataTable.Num(2, nRow)
    Else
        dPrice = DataTable.Num(0, nRow)
        nCount = DataTable.Num(1, nRow)
        nCount2 = DataTable.Num(2, nRow)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.GetTableRow", eGDRaiseError_Raise

End Sub

Private Sub ZeroTableRow(DataTable As cGdTable, ByVal nRow&)

    DataTable(0, nRow) = ""
    DataTable.Num(1, nRow) = kNullData
    DataTable.Num(2, nRow) = 0
    DataTable.Num(3, nRow) = 0#
    DataTable(4, nRow) = ""
    DataTable(5, nRow) = ""

End Sub

Private Sub SetTableRow(DataTable As cGdTable, _
    ByVal strMaker$, ByVal dPrice#, ByVal nCount&, ByVal nCount2&, _
    ByVal dTime#, ByVal strActiveFlag$, ByVal nRow&, ByVal bMakerData As Boolean)
On Error GoTo ErrSection:

    Dim dPrevSize&

    If nRow < 0 Or nRow >= DataTable.NumRecords Then
        Exit Sub
    End If
    
    If bMakerData Then
        dPrevSize = DataTable.Num(2, nRow)
        DataTable(0, nRow) = strMaker       'Bid or Ask table was passed in
        If dPrice <> kNullData Then DataTable.Num(1, nRow) = dPrice
        If nCount <> 0 Then DataTable.Num(2, nRow) = nCount
        
        If DataTable(4, nRow) = "C" Then
            If Int(DataTable.Num(3, nRow)) <> m.nSessionDate Then
                DataTable(4, nRow) = strActiveFlag
                DataTable.Num(3, nRow) = dTime
            End If
        Else
            DataTable(4, nRow) = strActiveFlag
            DataTable.Num(3, nRow) = dTime
        End If
                
        If dPrevSize > 0 And nCount > 0 Then
            If nCount > dPrevSize Then
                DataTable(5, nRow) = "+"
            ElseIf nCount < dPrevSize Then
                DataTable(5, nRow) = "-"
            End If
        End If
    Else
        If dPrice <> kNullData Then DataTable.Num(0, nRow) = dPrice
        If nCount > 0 Then DataTable.Num(1, nRow) = nCount
        If nCount2 > 0 Then DataTable.Num(2, nRow) = nCount2
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.SetTableRow", eGDRaiseError_Raise

End Sub

Private Sub ParseDataString(strMaker$, dBid#, nBidSize&, dAsk#, nAskSize&, _
    dTime#, strFlag$, ByVal strData$)
On Error GoTo ErrSection:

    Dim aFields As New cGdArray
    
    'initialize return values
    strMaker = ""
    dBid = kNullData
    nBidSize = 0
    dAsk = kNullData
    nAskSize = 0
    dTime = 0#

    aFields.SplitFields (strData)
    
    'precautionary
    If aFields(0) <> m.strSymbol Then Exit Sub
    
    If aFields.Size >= 7 Then
        strMaker = aFields(1)
        dTime = Val(aFields(2))
        strFlag = aFields(8)
        dBid = RoundToMinMove(Val(aFields(3)), m.dMinMove)
        nBidSize = Val(aFields(4))
        dAsk = RoundToMinMove(Val(aFields(5)), m.dMinMove)
        nAskSize = Val(aFields(6))
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.ParseDataString", eGDRaiseError_Raise

End Sub

Private Function SearchForPrice(ByVal dPrice#, ByVal bSalmonTables As Boolean) As Long
On Error GoTo ErrSection:

    Dim dRoundedPrice#, i&, nPos&
    Dim tbBidAsk As cGdTable

    If bSalmonTables Then
        Set tbBidAsk = m.BidAskSalmon
    Else
        Set tbBidAsk = m.BidAsk
    End If

    If tbBidAsk.NumRecords <= 0 Then
        SearchForPrice = -1
        Exit Function
    End If

    dRoundedPrice = RoundToMinMove(dPrice, m.dMinMove)
    nPos = -1
    For i = 0 To tbBidAsk.NumRecords - 1
        If dRoundedPrice = tbBidAsk.Num(0, i) Then
            nPos = i
            Exit For
        End If
    Next

    SearchForPrice = nPos

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cBidAskDepth.SearchForPrice", eGDRaiseError_Raise

End Function

Public Property Get BidAskData() As cGdTable

    If m.BidSalmon.NumRecords > 0 Or m.AskSalmon.NumRecords > 0 Then
        If m.bDirtyBidAsk Then
            m.BidAskSalmon.NumRecords = 0
            UpdateCombined m.BidSalmon, 0, False
            UpdateCombined m.AskSalmon, 1, False
        End If
        Set BidAskData = m.BidAskSalmon
    ElseIf m.Bid.NumRecords > 0 Or m.Ask.NumRecords > 0 Then
        If m.bDirtyBidAsk Then
            UpdateCombined m.Bid, 0, False
            UpdateCombined m.Ask, 1, False
        End If
        Set BidAskData = m.BidAsk
    End If
        
    m.bDirtyBidAsk = False

End Property

Public Property Get AskData() As cGdTable

    If m.AskSalmon.NumRecords > 0 Then
        Set AskData = m.AskSalmon
    Else
        Set AskData = m.Ask
    End If

End Property

Public Property Get BidData() As cGdTable

    If m.BidSalmon.NumRecords > 0 Then
        Set BidData = m.BidSalmon
    Else
        Set BidData = m.Bid
    End If

End Property

Public Property Get BestBids() As cGdArray
    
    If m.BidSalmon.NumRecords > 0 Then
        Set BestBids = m.aBestBidSalmon
    Else
        Set BestBids = m.aBestBids
    End If

End Property

Public Property Get BestAsks() As cGdArray

    If m.AskSalmon.NumRecords > 0 Then
        Set BestAsks = m.aBestAskSalmon
    Else
        Set BestAsks = m.aBestAsks
    End If

End Property

Public Sub LargestBidAskSize(nLargestSize&, nLargestBid&, nLargestAsk&, dBidPrice#, dAskPrice#)
On Error GoTo ErrSection:
    
    Dim i&, nSize&
    
    Dim tbBidAsk As cGdTable
    
    nLargestSize = 0
    nLargestBid = 0
    nLargestAsk = 0
        
    nSize = 0
    
    If m.BidAskSalmon.NumRecords > 0 Then
        Set tbBidAsk = m.BidAskSalmon
    Else
        Set tbBidAsk = m.BidAsk
    End If
    
    For i = 0 To tbBidAsk.NumRecords - 1
        nSize = tbBidAsk.Num(1, i)
        If nSize > nLargestSize Then nLargestSize = nSize
        If nSize > nLargestBid Then
            nLargestBid = nSize
            dBidPrice = tbBidAsk.Num(0, i)
        End If
        
        nSize = tbBidAsk.Num(2, i)
        If nSize > nLargestSize Then nLargestSize = nSize
        If nSize > nLargestAsk Then
            nLargestAsk = nSize
            dAskPrice = tbBidAsk.Num(0, i)
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.LargestBidAskSize", eGDRaiseError_Raise

End Sub

'This property checks the individual bid or ask tables to
'see if all fields have valid data. Only changed data is
'sent so it is possible to get bid or ask sizes without
'corresponding prices or vice versa.
Public Property Get AllDataValid(ByVal nBidAskTable&, ByVal nRow&, _
    bIsActive As Boolean, strUpDownFlag$) As Boolean
On Error GoTo ErrSection:
    
    Dim bHasAll As Boolean
    Dim tblTable As cGdTable
        
    If nBidAskTable = 0 Then
        If m.BidSalmon.NumRecords > 0 Then
            Set tblTable = m.BidSalmon
        Else
            Set tblTable = m.Bid
        End If
    ElseIf nBidAskTable = 1 Then
        If m.AskSalmon.NumRecords > 0 Then
            Set tblTable = m.AskSalmon
        Else
            Set tblTable = m.Ask
        End If
    Else
        Exit Property
    End If
        
    bHasAll = True
    
    If tblTable(4, nRow) = "Y" Then
        bIsActive = True
    Else
        bIsActive = False
    End If
    
    If Len(tblTable(0, nRow)) < 1 Then bHasAll = False  'maker
    If tblTable.Num(1, nRow) = kNullData Then bHasAll = False      'price
    If tblTable.Num(2, nRow) <= 0 Then bHasAll = False      'size
    If tblTable.Num(3, nRow) <= 0 Then bHasAll = False      'time
    strUpDownFlag = tblTable(5, nRow)                   '+/-
    
    AllDataValid = bHasAll

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cBidAskDepth.AllDataValid.Get", eGDRaiseError_Raise

End Property

Public Sub CheckBestBidAsk(ByVal dBestBidPrice#, ByVal dBestAskPrice#)
On Error GoTo ErrSection:

    Dim i&, d#
    
    '01-17-2006: only do this for stocks
    If SecurityType(m.strSymbol) <> "S" Then
        Exit Sub
    End If
    
    If dBestBidPrice <> kNullData Then
        For i = m.Bid.NumRecords - 1 To 0 Step -1
            d = m.Bid.Num(1, i)
            If d <> kNullData Then
                If d > dBestBidPrice Then
                    m.Bid.RemoveRecords i
                End If
            End If
        Next
    End If

    If dBestAskPrice <> kNullData Then
        For i = m.Ask.NumRecords - 1 To 0 Step -1
            d = m.Ask.Num(1, i)
            If d <> kNullData Then
                If d < dBestAskPrice Then
                    m.Ask.RemoveRecords i
                End If
            End If
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.CheckBestBidAsk", eGDRaiseError_Raise

End Sub

Public Property Get UsageCounter() As Long
    UsageCounter = m.nUsageCounter
End Property

Public Property Let UsageCounter(ByVal nCount&)
    m.nUsageCounter = nCount
End Property

Public Sub NewData(ByVal strData$)
On Error GoTo ErrSection:

    Dim aStrings As New cGdArray
    Dim i&

    'save data to array if the necessary data not set
    If m.dMinMove = 0 Then
        m.aUnclaimedData.Add strData
        m.dLastNewData = gdTickCount
        Exit Sub
    End If

    If m.aUnclaimedData.Size > 0 Then
        For i = 0 To m.aUnclaimedData.Size - 1
            aStrings.Add m.aUnclaimedData(i)
        Next
        m.aUnclaimedData.Clear
        m.aUnclaimedData.Size = 0
    End If

    aStrings.Add strData

    UpdateByMakers aStrings

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cBidAskDepth.NewData", eGDRaiseError_Raise

End Sub

Public Function ClientInit(ByVal strSym$, ByVal dMinMove#, ByVal nSessionDate&) As Boolean
On Error GoTo ErrSection:

    If g.RealTime.ConvertContinuous Then
        strSym = RollSymbolForDate(strSym)
    End If
    If Len(m.strSymbol) = 0 Then m.strSymbol = strSym
    m.bIsSpreadSymbol = IsSpreadSymbol(m.strSymbol)
    If m.nSessionDate = 0 Then m.nSessionDate = nSessionDate
    If m.dMinMove = 0 Then m.dMinMove = dMinMove
    
    If strSym <> m.strSymbol Or m.nSessionDate <> nSessionDate Or m.dMinMove <> dMinMove Then
        ClientInit = False
    Else
        ClientInit = True
    End If
                    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBidAskDepth.ClienInit", eGDRaiseError_Raise

End Function

Public Property Get NewestData() As Double
On Error GoTo ErrSection:


    NewestData = m.dNewestData

    Exit Property
    
ErrSection:
    RaiseError "cBidAskDepth.NewestData.Get"
    
End Property

Private Sub UpdateBest(ByVal bSalmonTables As Boolean)
On Error GoTo ErrSection

    Dim i&, k&
    Dim dPrev#, dBid#, dAsk#
    
    Dim tbBid As cGdTable
    Dim tbAsk As cGdTable
    
    Dim arrayAsk As cGdArray
    Dim arrayBid As cGdArray
    Dim aIdxBestBidAsk As cGdArray
    
    If bSalmonTables Then
        If m.BidSalmon Is Nothing Or m.AskSalmon Is Nothing Or m.aBestAskSalmon Is Nothing Or m.aBestBidSalmon Is Nothing Then
            DebugLog m.strSymbol & " one of the data table is a null pointer."
            GoTo ErrExit
        Else
            'save 5 best bids from salmon table
            m.aBestBidSalmon.Size = 0
            Set aIdxBestBidAsk = m.BidSalmon.CreateSortedIndex(0)
            If Not aIdxBestBidAsk Is Nothing Then
                If aIdxBestBidAsk.Size > 5 Then
                    k = 4
                Else
                    k = aIdxBestBidAsk.Size - 1
                End If
                For i = 0 To 5
                    m.aBestBidSalmon(i) = m.BidSalmon(1, aIdxBestBidAsk(i))
                Next
            End If
        
            Set aIdxBestBidAsk = Nothing
            'save 5 best asks from salmon table
            m.aBestAskSalmon.Size = 0
            Set aIdxBestBidAsk = m.AskSalmon.CreateSortedIndex(0)
            If Not aIdxBestBidAsk Is Nothing Then
                If aIdxBestBidAsk.Size > 5 Then
                    k = 4
                Else
                    k = aIdxBestBidAsk.Size - 1
                End If
                For i = 0 To k
                    m.aBestAskSalmon(i) = m.AskSalmon(1, aIdxBestBidAsk(i))
                Next
            End If
        End If
    Else
        Set tbBid = m.Bid
        Set tbAsk = m.Ask
        Set arrayAsk = m.aBestAsks
        Set arrayBid = m.aBestBids
    
        'save best 5 bids
        Set aIdxBestBidAsk = tbBid.CreateSortedIndex(1, eGdSort_Descending)
        k = 0
        dPrev = kNullData
        For i = 0 To aIdxBestBidAsk.Size - 1
            If k >= 5 Then Exit For
            dBid = tbBid.Num(1, aIdxBestBidAsk(i))
            If dBid <> kNullData And dBid <> dPrev Then
                If tbBid.Num(2, aIdxBestBidAsk(i)) > 0 Then
                    arrayBid.Num(k) = dBid
                    dPrev = dBid
                    k = k + 1
                Else
                    'do nothing - YM,ZB is coming accross with @AMEX
                    'StatusMsg "Bid: " & m.Bid(0, aIdxBestBidAsk(i))
                End If
            End If
        Next
        
        'save best 5 asks
        Set aIdxBestBidAsk = tbAsk.CreateSortedIndex(1)
        k = 0
        dPrev = kNullData
        For i = 0 To aIdxBestBidAsk.Size - 1
            If k >= 5 Then Exit For
            dAsk = tbAsk.Num(1, aIdxBestBidAsk(i))
            If dAsk <> kNullData And dAsk <> dPrev Then
                If tbAsk.Num(2, aIdxBestBidAsk(i)) > 0 Then
                    arrayAsk.Num(k) = dAsk
                    dPrev = dAsk
                    k = k + 1
                Else
                    'do nothing - YM,ZB is coming accross with @AMEX
                    'StatusMsg "Ask: " & m.Ask(0, aIdxBestBidAsk(i))
                End If
            End If
        Next
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.UpdateBest"
    
End Sub

Private Sub DumpDepthData()
On Error GoTo ErrSection:
    
    Dim i&
    Dim aFile As New cGdArray
    
    Dim Bars As New cGdBars
    
    SetBarProperties Bars, m.strSymbol, True
    
    'dump bid table returned by salmon
    aFile.Add m.strSymbol & " Salmon bid table size = " & Str(m.BidSalmon.NumRecords)
    For i = 0 To m.BidSalmon.NumRecords - 1
        aFile.Add m.BidSalmon.GetRecord(i, vbTab) & vbTab & Bars.PriceDisplay(m.BidSalmon(1, i))
    Next
    
    'dump ask table returned by salmon
    aFile.Add vbCrLf & vbCrLf & m.strSymbol & " Salmon ask table size = " & Str(m.AskSalmon.NumRecords)
    For i = 0 To m.AskSalmon.NumRecords - 1
        aFile.Add m.AskSalmon.GetRecord(i, vbTab) & vbTab & Bars.PriceDisplay(m.AskSalmon(1, i))
    Next
    
    'dump best 5 bids as determined by BidAsk object
    aFile.Add vbCrLf & vbCrLf & "Best 5 Bids"
    For i = 0 To m.aBestBidSalmon.Size - 1
        aFile.Add m.aBestBidSalmon(i) & vbTab & Bars.PriceDisplay(m.aBestBidSalmon(i))
    Next
    
    'dump best 5 asks as determined by BidAsk object
    aFile.Add vbCrLf & vbCrLf & "Best 5 Asks"
    For i = 0 To m.aBestAskSalmon.Size - 1
        aFile.Add m.aBestAskSalmon(i) & vbTab & Bars.PriceDisplay(m.aBestAskSalmon(i))
    Next
    
    'dump combined BidAsk table
    aFile.Add vbCrLf & vbCrLf & "Combined BidAsk table'"
    For i = 0 To m.BidAskSalmon.NumRecords - 1
        aFile.Add m.BidAskSalmon.GetRecord(i, vbTab) & vbTab & Bars.PriceDisplay(m.BidAskSalmon(0, i))
    Next
    
    aFile.ToFile "SalmonDepthData_" & m.strSymbol & ".txt"
    
    
    m.bDumpFileSalmon = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBidAskDepth.DumpDepthData"
    
End Sub


Public Sub ToggleSalmonFile()
    m.bDumpFileSalmon = Not m.bDumpFileSalmon
End Sub
