VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimeSalesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVSFlexDataSource

Private Enum eDataFields
    eFld_Time = 0
    eFld_Price
    eFld_Volume
    eFld_Type
    eFld_Color
    eFld_UpDownFlag
    eFld_Flags
End Enum

Private Enum eCumulativeFieds
    eFldCumulative_Time = 0
    eFldCumulative_Price
    eFldCumulative_Trades
    eFldCumulative_Contracts
    eFldCumulative_AvgSize
    eFldCumulative_LargestSize
    eFldCumulative_BuyVol
    eFldCumulative_SellVol
    eFldCumulative_Color
    eFldCumulative_UpDownFlag
End Enum

Private Type mPrivate
    frmTS As frmTimeSales
    TickBars As cGdBars
    DailyBars As cGdBars
    
    tbMinute As New cGdTable        'table holding data for minute-by-minute view
    tbTick As New cGdTable          'table holding data for tick-by-tick view
    tbCumulative As New cGdTable
    
    tbCache As New cGdTable             'cache table for tick-by-tick view
    tbCacheCumulative As New cGdTable   'cache table fot cumulative view
    
    dLastTradePriceAtAsk As Double
    dLastTradePriceAtBid As Double
    dLastTradePrice As Double
    dLastBid As Double
    dLastAsk As Double
    dMinMove As Double
    
    nLastTradeColor As Long
    nLastBidSize As Long
    nLastAskSize As Long
    nLastBidColor As Long
    nLastAskColor As Long
    nLastBidUpDownFlag As Long
    nLastAskUpDownFlag As Long
    nColCount As Long
    
    nVolFilterMax As Long
    nVolFilterMin As Long
    
    nSumBuyVol As Long                  'for DTG start/stop button
    nSumSellVol As Long
        
    bHasBidAsk As Boolean
    bProfileUnload As Boolean
    bBuySellDebug As Boolean
    bSumBySell As Boolean
    
    strMsg As String
End Type
Private m As mPrivate

Private Sub Class_Initialize()

    'create table for minute-by-minute view
    m.tbMinute.CreateField eGDARRAY_Doubles, eFld_Time, "Time"
    m.tbMinute.CreateField eGDARRAY_Strings, eFld_Price, "Price"
    'create table for tick-by-tick view
    m.tbTick.CreateField eGDARRAY_Doubles, eFld_Time, "Time"
    m.tbTick.CreateField eGDARRAY_Doubles, eFld_Price, "Price"
    m.tbTick.CreateField eGDARRAY_Longs, eFld_Volume, "Volume"
    m.tbTick.CreateField eGDARRAY_Strings, eFld_Type, "Type"
    m.tbTick.CreateField eGDARRAY_Longs, eFld_Color, "Color"
    m.tbTick.CreateField eGDARRAY_Longs, eFld_UpDownFlag, "UpDownFlag"
    m.tbTick.CreateField eGDARRAY_Longs, eFld_Flags, "Flags"
    
    'create cache table for tick-by-tick view
    m.tbCache.CreateField eGDARRAY_Doubles, eFld_Time, "Time"
    m.tbCache.CreateField eGDARRAY_Doubles, eFld_Price, "Price"
    m.tbCache.CreateField eGDARRAY_Longs, eFld_Volume, "Volume"
    m.tbCache.CreateField eGDARRAY_Strings, eFld_Type, "Type"
    m.tbCache.CreateField eGDARRAY_Longs, eFld_Color, "Color"
    m.tbCache.CreateField eGDARRAY_Longs, eFld_UpDownFlag, "UpDownFlag"
    m.tbCache.CreateField eGDARRAY_Longs, eFld_Flags, "Flags"
    
    'create table for cumulative view
    m.tbCumulative.CreateField eGDARRAY_Doubles, eFldCumulative_Time, "Time"
    m.tbCumulative.CreateField eGDARRAY_Doubles, eFldCumulative_Price, "Price"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_Trades, "Trades"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_Contracts, "Contracts"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_AvgSize, "AvgTradeSize"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_LargestSize, "LargestTradeSize"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_BuyVol, "BuyVol"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_SellVol, "SellVol"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_Color, "Color"
    m.tbCumulative.CreateField eGDARRAY_Longs, eFldCumulative_UpDownFlag, "UpDownFlag"

    'create cache table for cumulative view
    m.tbCacheCumulative.CreateField eGDARRAY_Doubles, eFldCumulative_Time, "Time"
    m.tbCacheCumulative.CreateField eGDARRAY_Doubles, eFldCumulative_Price, "Price"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_Trades, "Trades"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_Contracts, "Contracts"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_AvgSize, "AvgTradeSize"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_LargestSize, "LargestTradeSize"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_BuyVol, "BuyVol"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_SellVol, "SellVol"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_Color, "Color"
    m.tbCacheCumulative.CreateField eGDARRAY_Longs, eFldCumulative_UpDownFlag, "UpDownFlag"

End Sub

Private Sub Class_Terminate()
    
    Set m.TickBars = Nothing
    Set m.tbMinute = Nothing
    Set m.tbTick = Nothing
    Set m.tbCache = Nothing
    Set m.tbCumulative = Nothing
    Set m.tbCacheCumulative = Nothing
    
End Sub

Private Function IVSFlexDataSource_GetData(ByVal field As Long, ByVal Record As Long) As String
On Error GoTo ErrSection:

    Dim eField As eDataFields
    Dim eFieldCumulative As eCumulativeFieds
    Dim nStyle&, nColor&, dPrice#
    
    If m.bProfileUnload Then Exit Function
    
    If m.frmTS Is Nothing Or m.TickBars Is Nothing Or m.nColCount = 0 Then
        'no data available
        IVSFlexDataSource_GetData = ""
        Exit Function
    ElseIf m.TickBars.Size = 0 Then
        With m.frmTS.fg
            If field = .Cols - 1 Then
                IVSFlexDataSource_GetData = m.strMsg
            Else
                IVSFlexDataSource_GetData = m.strMsg    '"N/A"
            End If
        End With
        Exit Function
    End If

    nStyle = m.frmTS.DisplayStyle
    
    eField = field
    eFieldCumulative = field
    If eField = eFld_Type And nStyle = 2 Then
        Exit Function
    ElseIf nStyle = 4 Then
        If eFieldCumulative < eFldCumulative_Time Or eFieldCumulative > eFldCumulative_SellVol Then
            'invalid field
            IVSFlexDataSource_GetData = ""
            Exit Function
        End If
    ElseIf eField < eFld_Time Or eField > eFld_Type Then
        'invalid field
        IVSFlexDataSource_GetData = ""
        Exit Function
    End If
                    
    If nStyle = 1 Or nStyle = 3 Then
        'tick by tick view
        IVSFlexDataSource_GetData = GetDataByTick(nColor, dPrice, eField, Record, nStyle)
        If m.frmTS.fg.Cell(flexcpForeColor, Record + 1, 0) <> nColor Then
            m.frmTS.fg.Cell(flexcpForeColor, Record + 1, eFld_Time, Record + 1, m.nColCount) = nColor
        End If
    ElseIf nStyle = 2 Then
        'minute by minute view
        IVSFlexDataSource_GetData = GetDataByMinute(eField, Record)
    ElseIf nStyle = 4 Then
        IVSFlexDataSource_GetData = GetDataByTick(nColor, dPrice, eFieldCumulative, Record, nStyle)
        
        If eFieldCumulative = eFldCumulative_Price Then
            If m.frmTS.fg.Cell(flexcpForeColor, Record + 1, 0) <> nColor Then
                m.frmTS.fg.Cell(flexcpForeColor, Record + 1, eFld_Time, Record + 1, m.nColCount) = nColor
            End If
        End If
    Else
        'invalid view
        IVSFlexDataSource_GetData = ""
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.IVSFlexDataSource.GetData", eGDRaiseError_Raise

End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
On Error GoTo ErrSection:

    Dim nStyle&
    
    If m.frmTS Is Nothing Then
        IVSFlexDataSource_GetFieldCount = 0
    Else
        nStyle = m.frmTS.DisplayStyle
        If nStyle = 1 Then              'TickByTick
            IVSFlexDataSource_GetFieldCount = 3
            m.nColCount = 2
        ElseIf nStyle = 2 Then          'MinByMin
            IVSFlexDataSource_GetFieldCount = 2
            m.nColCount = 1
        ElseIf nStyle = 3 Then          'TickByTick with best bid/ask
            IVSFlexDataSource_GetFieldCount = 4
            m.nColCount = 3
        ElseIf nStyle = 4 Then
            IVSFlexDataSource_GetFieldCount = 8
            m.nColCount = 7
        Else
            IVSFlexDataSource_GetFieldCount = 0
            m.nColCount = 0
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.IVSFlexDataSource.GetFieldCount", eGDRaiseError_Raise

End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal field As Long) As String
On Error GoTo ErrSection:

    Dim nStyle&, strSym$
    
    If m.frmTS Is Nothing Or m.TickBars Is Nothing Then
        IVSFlexDataSource_GetFieldName = ""
    Else
        nStyle = m.frmTS.DisplayStyle
        strSym = m.TickBars.Prop(eBARS_Symbol)
        Select Case nStyle
            Case 1, 3   '1=tick-by-tick, 3=tick-by-tick with best bid/ask
                Select Case field
                    Case 0:
                        IVSFlexDataSource_GetFieldName = "Time"
                    Case 1:
                        IVSFlexDataSource_GetFieldName = "Price"
                    Case 2:
                        'If InStr(strSym, "-") > 0 Then
                        '    IVSFlexDataSource_GetFieldName = "Trades"
                        'Else
                            IVSFlexDataSource_GetFieldName = "Size"
                        'End If
                    Case 3:
                        IVSFlexDataSource_GetFieldName = "Type"
                    Case Else:
                        IVSFlexDataSource_GetFieldName = ""
                End Select
            Case 2      '2 = minute by minute
                Select Case field
                    Case 0:
                        IVSFlexDataSource_GetFieldName = "Time"
                    Case 1:
                        IVSFlexDataSource_GetFieldName = "Price"
                    Case Else:
                        IVSFlexDataSource_GetFieldName = ""
                End Select
            Case 4      'cumulative
                Select Case field
                    Case eFldCumulative_Time
                        IVSFlexDataSource_GetFieldName = "Time"
                    Case eFldCumulative_Price
                        IVSFlexDataSource_GetFieldName = "Price"
                    Case eFldCumulative_Trades
                        IVSFlexDataSource_GetFieldName = "Trades"
                    Case eFldCumulative_Contracts
                        IVSFlexDataSource_GetFieldName = "Contracts"
                    Case eFldCumulative_AvgSize
                        IVSFlexDataSource_GetFieldName = "Avg Size"
                    Case eFldCumulative_LargestSize
                        IVSFlexDataSource_GetFieldName = "Largest Size"
                    Case eFldCumulative_BuyVol
                        IVSFlexDataSource_GetFieldName = "Buy Vol"
                    Case eFldCumulative_SellVol
                        IVSFlexDataSource_GetFieldName = "Sell Vol"
                End Select
            Case Else   'invalid style
                IVSFlexDataSource_GetFieldName = ""
        End Select
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.IVSFlexDataSource.GetFieldName", eGDRaiseError_Raise

End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
On Error GoTo ErrSection:

    Dim nStyle&
    
    If m.frmTS Is Nothing Or m.TickBars Is Nothing Then
        IVSFlexDataSource_GetRecordCount = 0
        Exit Function
    End If

    nStyle = m.frmTS.DisplayStyle
        
    If nStyle = 1 Or nStyle = 3 Then
        'tick by tick
        IVSFlexDataSource_GetRecordCount = m.tbTick.NumRecords          'm.TickBars.Size()
        If m.tbTick.NumRecords <> m.TickBars.Size Then
            nStyle = nStyle
        End If
    ElseIf nStyle = 2 Then
        'minute by minute
        IVSFlexDataSource_GetRecordCount = m.tbMinute.NumRecords
    ElseIf nStyle = 4 Then
        IVSFlexDataSource_GetRecordCount = m.tbCumulative.NumRecords
    Else
        'invalid style
        IVSFlexDataSource_GetRecordCount = 0
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.IVSFlexDataSource.GetRecordCount", eGDRaiseError_Raise

End Function

Private Sub IVSFlexDataSource_SetData(ByVal field As Long, ByVal Record As Long, ByVal NewData As String)
    'data is read only - do nothing
End Sub

Public Property Let FormTS(frm As frmTimeSales)
    Set m.frmTS = frm
End Property

Private Sub LoadDataTable()
On Error GoTo ErrSection:

    Dim strSym$, nStatus&
    
    m.bBuySellDebug = FileExist("TimeSalesDebug.flg")
    
    If m.TickBars.Size > 0 Then
        m.strMsg = "Loading data table for " & m.TickBars.Prop(eBARS_Symbol) & " ..."
        m.frmTS.UpdateMessage m.strMsg
        
        LoadMinuteTable
        LoadTickTable
        LoadCumulativeTable
        
        m.strMsg = ""
    ElseIf g.RealTime.SalmonIsRunning Then
        strSym = m.TickBars.Prop(eBARS_Symbol)
        nStatus = g.RealTime.SymbolInfo(strSym).GetDataRequestStatus(ePRD_EachTick)
        If nStatus = eSalmonPending Then
            m.strMsg = "Retrieving intraday data for " & strSym & " ..."
        ElseIf nStatus = eSalmonAvailable Then
            m.strMsg = ""
        Else
            m.strMsg = "No tick data. Salmon status = " & Str(nStatus)
        End If
    ElseIf Len(m.strMsg) = 0 Then
        m.strMsg = "Intraday data not available..."
    End If
    
    m.frmTS.UpdateMessage m.strMsg

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTimeSalesData.LoadDataTable"

End Sub

Public Function NewTradeProfileData(ByVal nSymID&, ByVal strSym$, Optional ByVal bSkipGetData As Boolean = False) As Boolean
On Error GoTo ErrSection:
   
    Dim dBid#, dAsk#
    
    m.tbMinute.NumRecords = 0
    m.tbTick.NumRecords = 0
    m.dLastAsk = 0
    m.dLastBid = 0
    m.bHasBidAsk = False
    
    m.nLastAskSize = 0
    m.nLastBidSize = 0
    
    If m.bProfileUnload Then Exit Function
    
    If bSkipGetData Then
        IVSFlexDataSource_GetData kNullData, 0      'to reset static variable in this subroutine
    Else
        Set m.TickBars = Nothing
        Set m.TickBars = New cGdBars
        Set m.DailyBars = Nothing
        GetBarData nSymID, strSym
    End If
    
    'build daily bar from ticks to get best/bid ask in real time
    If g.RealTime.Active Then
        Set m.DailyBars = New cGdBars
        SetBarProperties m.DailyBars, m.TickBars.Prop(eBARS_SymbolID)
        m.DailyBars.ArrayMask = eBARS_Eod Or eBARS_BidAsk
        m.DailyBars.BuildBars "D", m.TickBars.BarsHandle
    End If
    
    m.dMinMove = m.TickBars.MinMove(m.TickBars(eBARS_DateTime, 0))
    
    If Not CanHaveVolFilter Then
        m.nVolFilterMax = 0         '4466
        m.nVolFilterMin = 0
    End If
    
    LoadDataTable
    
ErrExit:
    NewTradeProfileData = Not m.bProfileUnload
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.NewData", eGDRaiseError_Raise

End Function

Private Sub GetBarData(ByVal nSymID&, ByVal strSym$)
On Error GoTo ErrSection:

    Dim rc As Byte
    Dim nDate&, i&
    Dim strMsg$
    Dim bFullTicksAvail As Boolean

    If Left(strSym, 1) = "$" And Not IsForex(strSym) Then
        InfBox "This feature does not support index symbols.", "I", , "Time & Sales"
        Exit Sub
    End If

    If m.frmTS.TS_SessionCurrent = True Then
        GetAvailTickData m.TickBars, nDate, strSym, nSymID, 0, -6
    Else
        GetAvailTickData m.TickBars, nDate, strSym, nSymID, m.frmTS.TS_SessionDate, 0
    End If

    If m.frmTS.TS_SessionCurrent = True Then      'current session
        If m.TickBars.Size = 0 Then
            If Not g.RealTime.Active Then
                ' ask if wish to refresh quote board and try again
                strMsg = "No intraday data available for " & strSym & ".  Would you like to refresh the data for the current session?"
                strMsg = InfBox(strMsg, "?", "+Refresh|-Not Now", "No Intraday Data")
            Else
                strMsg = "R"
            End If
            If strMsg = "R" Then
                If g.RealTime.RefreshSymbolList(Not g.RealTime.Active) Then
                    GetAvailTickData m.TickBars, nDate, strSym, nSymID, 0, 0
                End If
            End If
        End If
        m.frmTS.TS_SessionDate = nDate
    End If
            
    strMsg = ""
    If g.nReplaySession = 0 Then        'don't do check if streaming replay is on
        ' check if need to download full ticks (if they are currently minutized)
        If m.TickBars.Size > 0 And IsMinutized(m.TickBars, bFullTicksAvail) Then
            m.TickBars.Size = 0
            If Not ProcessIsBusy(True) Then
                If Not m.frmTS.TS_SessionCurrent Then
                    If bFullTicksAvail Then
                        'only prompt to download non-minutized data if user chose specific date and full tick is available per tick available flag file
                        strMsg = "Data for " & DateFormat(m.frmTS.TS_SessionDate) & " is currently compressed data. To view time & sales for this date, you will need to download all of the ticks.||Do you wish to download now?|"
                        If InfBox(strMsg, "?", "+Yes|-No", "Confirmation") = "Y" Then
                            If DownloadTicks(m.TickBars, m.frmTS.TS_SessionDate) Then
                                DM_GetBars m.TickBars, nSymID, ePRD_EachTick, m.frmTS.TS_SessionDate, m.frmTS.TS_SessionDate, , , , , False
                            End If
                        End If
                    Else
                        strMsg = "Full tick data for " & strSym & " on " & DateFormat(m.frmTS.TS_SessionDate) & " is not available for download at this time."
                        'InfBox strMsg, "I", , "Time and Sales"
                    End If
                End If
            End If
        End If
    End If

    ' if no data ...
    If m.TickBars.Size = 0 Then
        If InStr(strSym, "-") > 0 And Not HasModule("FT") Then
            strMsg = "You must be enabled for intraday data for Futures."
        ElseIf InStr(strSym, "-") = 0 And Not HasModule("ST") Then
            strMsg = "You must be enabled for intraday data for Stocks."
        ElseIf g.RealTime.SalmonIsRunning Then
            If g.RealTime.SymbolInfo(strSym).GetDataRequestStatus(ePRD_EachTick) = eSalmonPending Then
                strMsg = "Retrieving intraday data for " & strSym & " ..."
            Else
                strMsg = "No intraday data available for " & strSym & " on " & DateFormat(m.frmTS.TS_SessionDate)
            End If
        ElseIf Len(strMsg) = 0 Then
            strMsg = "No intraday data available for " & strSym & " on " & DateFormat(m.frmTS.TS_SessionDate)
        End If
    End If
    
    m.strMsg = strMsg
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTimeSalesData.GetBarData", eGDRaiseError_Raise

End Sub

Private Sub LoadMinuteTable()
On Error GoTo ErrSection:

    Dim nPointerSave&
    
    Dim i&, iStart&, nVol&, nSize&
    Dim dDate#, dDatePrev#
    Dim dPrice#, dPricePrev#
    
    Dim strPriceAtIdx$, strPrice$, strMsgSave$
    
    Dim hBarClose&, hBarTime&, hBarVolume&
    Dim hTime&, hPrice&
    
    Dim bAddRecoord As Boolean
    
    nPointerSave = Screen.MousePointer
    
    m.tbMinute.NumRecords = 0
    
    If m.TickBars.Size = 0 Then
        m.tbMinute.AddRecord m.strMsg
        GoTo ErrExit
    End If
    
    hBarTime = m.TickBars.ArrayHandle(eBARS_DateTime)
    hBarClose = m.TickBars.ArrayHandle(eBARS_Close)
    hBarVolume = m.TickBars.ArrayHandle(eBARS_Vol)
    
    hTime = m.tbMinute.FieldArrayHandle(eFld_Time)
    hPrice = m.tbMinute.FieldArrayHandle(eFld_Price)
               
    strMsgSave = "Loading minute table for " & m.TickBars.Prop(eBARS_Symbol) & " ..."
    'find first trade that meets volume filter criteria
    For i = 0 To m.TickBars.Size - 1
        'If IncludeTrade(m.TickBars(eBARS_Vol, i)) Then
        If IncludeTrade(gdGetNum(hBarVolume, i)) Then
            'strPrice = m.TickBars.PriceDisplay(m.TickBars(eBARS_Close, i)) & " "
            'dDate = m.TickBars(eBARS_DateTime, i)
            
            dPrice = gdGetNum(hBarClose, i)
            strPrice = m.TickBars.PriceDisplay(dPrice) & " "
            dDate = gdGetNum(hBarTime, i)
            
            dDatePrev = dDate
            iStart = i + 1
            Exit For
        End If
    Next
    
    nSize = m.TickBars.Size
    For i = iStart To nSize - 2
        bAddRecoord = False
        nVol = gdGetNum(hBarVolume, i)
        dPrice = gdGetNum(hBarClose, i)
        
        If IncludeTrade(nVol) Then
            'dDate = m.TickBars(eBARS_DateTime, i)
            dDate = gdGetNum(hBarTime, i)
            If dDate = dDatePrev Then
                'strPrice = strPrice & m.TickBars.PriceDisplay(m.TickBars(eBARS_Close, i)) & " "
                If dPrice <> dPricePrev Then
                    strPriceAtIdx = m.TickBars.PriceDisplay(dPrice) & " "
                    dPricePrev = dPrice
                End If
                strPrice = strPrice & strPriceAtIdx
                
            ElseIf Int(dDate) = Int(dDatePrev) Then
                If Minute(dDate) = Minute(dDatePrev) Then
                    'strPrice = strPrice & m.TickBars.PriceDisplay(m.TickBars(eBARS_Close, i)) & " "
                    If dPrice <> dPricePrev Then
                        strPriceAtIdx = m.TickBars.PriceDisplay(dPrice) & " "
                        dPricePrev = dPrice
                    End If
                    strPrice = strPrice & strPriceAtIdx
                Else
                    bAddRecoord = True
                End If
            Else
                bAddRecoord = True
            End If
            If bAddRecoord Then
                m.tbMinute.AddRecord ""
                
                gdSetNum hTime, m.tbMinute.NumRecords - 1, dDatePrev
                gdSetStr hPrice, m.tbMinute.NumRecords - 1, strPrice
                
                dDatePrev = dDate
                dPricePrev = dPrice
                strPrice = m.TickBars.PriceDisplay(m.TickBars(eBARS_Close, i)) & " "
                strPriceAtIdx = strPrice
            End If
        End If
        
        If nSize > 100000 Then
            If i Mod 1000 = 0 Then
                m.strMsg = strMsgSave & Str(i)
                m.frmTS.UpdateMessage m.strMsg
                DoEvents
                Screen.MousePointer = vbHourglass
                If m.bProfileUnload Then Exit For
            End If
        End If
    Next
    
    'add last time & price
    m.tbMinute.AddRecord ""
    gdSetNum hTime, m.tbMinute.NumRecords - 1, dDatePrev
    gdSetStr hPrice, m.tbMinute.NumRecords - 1, strPrice
        
ErrExit:
    Screen.MousePointer = nPointerSave
    Exit Sub

ErrSection:
    Screen.MousePointer = nPointerSave
    RaiseError "cTimeSalesData.LoadMinuteTable", eGDRaiseError_Raise

End Sub

Private Function GetDataByTick(nColor&, dPrice#, ByVal eField As eDataFields, _
    ByVal nRecord&, ByVal nStyle&) As String
On Error GoTo ErrSection:

    Dim i&, nSort&, strData$
    Dim dDateTime#
    
    If m.bProfileUnload Then Exit Function
    
    If m.tbTick.NumRecords = 0 Then
        LoadDataTable
    Else
        m.strMsg = ""
    End If
        
    nSort = m.frmTS.SortOrder
    
    If nSort = flexSortGenericDescending Then
        If nStyle = 4 Then
            nSort = m.tbCumulative.NumRecords
            i = nSort - (nRecord + 1)
        Else
            nSort = m.tbTick.NumRecords
            i = nSort - (nRecord + 1)
        End If
            
        If i >= nSort Then
            i = nSort - 1
        ElseIf i < 0 Then
            i = 0
        End If
    Else
        i = nRecord
        If i < 0 Then
            GoTo ErrExit
        ElseIf nStyle = 4 Then
            If i >= m.tbCumulative.NumRecords Then
                GoTo ErrExit
            End If
        ElseIf i >= m.tbTick.NumRecords Then
            GoTo ErrExit
        End If
    End If
                    
    If nStyle = 4 Then
        nColor = m.tbCumulative(eFldCumulative_Color, i)
        Select Case eField
            Case eFldCumulative_Time:
                dDateTime = m.tbCumulative(eFldCumulative_Time, i)
                If g.bShowInLocalTimeZone Then
                    dDateTime = ConvertTimeZone(dDateTime, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
                End If
                strData = DateFormat(RoundToSecond(dDateTime), NO_DATE, HH_MM_SS)
            
            Case eFldCumulative_Price:
                strData = m.TickBars.PriceDisplay(m.tbCumulative(eFldCumulative_Price, i))
                dPrice = m.tbCumulative(eFld_Price, i)
            
            Case eFldCumulative_Trades, eFldCumulative_Contracts, eFldCumulative_AvgSize, eFldCumulative_LargestSize
                If m.tbCumulative(eField, i) > 0 Then
                    strData = Format(m.tbCumulative(eField, i), "#,##0")
                Else
                    strData = ""
                End If
            
            Case eFldCumulative_AvgSize
                strData = Format(m.tbCumulative(eFldCumulative_AvgSize, i), "#,##0")
            
            Case eFldCumulative_BuyVol, eFldCumulative_SellVol
                If m.tbCumulative(eField, i) > 0 Then
                    strData = Format(m.tbCumulative(eField, i), "#,##0")
                Else
                    strData = ""
                End If
            
            Case Else:
                strData = ""
        End Select
    Else
        nColor = m.tbTick(eFld_Color, i)
        Select Case eField
            Case eFld_Time:
                dDateTime = m.tbTick(eFld_Time, i)
                If g.bShowInLocalTimeZone Then
                    dDateTime = ConvertTimeZone(dDateTime, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
                End If
                strData = DateFormat(RoundToSecond(dDateTime), NO_DATE, HH_MM_SS)
            Case eFld_Price:
                strData = m.TickBars.PriceDisplay(m.tbTick(eFld_Price, i))
                dPrice = m.tbTick(eFld_Price, i)
            Case eFld_Volume:
                If m.tbTick(eFld_Volume, i) > 0 Then
                    strData = Format(m.tbTick(eFld_Volume, i), "#,##0")
                Else
                    strData = ""
                End If
            Case eFld_Type:
                strData = m.tbTick(eFld_Type, i)
                If m.bBuySellDebug Then
                    If UCase(m.tbTick(eFld_Type, i)) = "TRADE" Then
                        'eTICK_AtBid = 1     ' i.e. SELL Volume
                        'eTICK_AtAsk = 2     ' i.e. BUY Volume
                        'If gdGetNum(hFlags, i) = eTICK_AtBid Then
                        If m.tbTick(eFld_Flags, i) = eTICK_AtBid Then
                            strData = strData & " (SELL)"
                        ElseIf m.tbTick(eFld_Flags, i) = eTICK_AtAsk Then
                            strData = strData & " (BUY)"
                        Else
                            strData = strData & " (Unknown:" & Str(m.tbTick(eFld_Flags, i)) & ")"
                        End If
                    End If
                End If
            Case Else:
                strData = ""
        End Select
    End If
    
    GetDataByTick = strData

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.GetDataByTick", eGDRaiseError_Raise

End Function

Private Function GetDataByMinute(ByVal eField As eDataFields, ByVal nRecord&) As String
On Error GoTo ErrSection:

    Dim strData$, nSort&, i&
    Dim dDateTime As Double
    
    If m.bProfileUnload Then Exit Function
    
    If m.tbMinute.NumRecords = 0 Then
        LoadDataTable
    Else
        m.strMsg = ""
    End If
    
    nSort = m.frmTS.SortOrder
    
    If nSort = flexSortGenericDescending Then
        i = m.tbMinute.NumRecords - (nRecord + 1)
        If i >= m.tbMinute.NumRecords Then
            i = m.tbMinute.NumRecords - 1
        ElseIf i < 0 Then
            i = 0
        End If
    Else
        i = nRecord
    End If
    
    Select Case eField
        Case eFld_Time:
            dDateTime = m.tbMinute(eFld_Time, i)
            If g.bShowInLocalTimeZone Then
                dDateTime = ConvertTimeZone(dDateTime, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
            strData = DateFormat(dDateTime, NO_DATE, HH_MM)
        Case eFld_Price:
            strData = m.tbMinute(eFld_Price, i)
        Case Else:
            strData = ""
    End Select

    GetDataByMinute = strData

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.GetDataByMinute", eGDRaiseError_Raise

End Function

Public Property Get DataSize() As Long
On Error GoTo ErrSection:

    Dim nStyle&
    
    If m.frmTS Is Nothing Or m.TickBars Is Nothing Then
        DataSize = 0
        Exit Property
    End If
    
    nStyle = m.frmTS.DisplayStyle
    
    If nStyle = 1 Then
        DataSize = m.TickBars.Size      'tick by tick
    ElseIf nStyle = 2 Or nStyle = 3 Then
        DataSize = m.tbMinute.NumRecords
    Else
        DataSize = 0
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.Get.DataSize", eGDRaiseError_Raise

End Property

Public Function UpdateDataRT(ByVal bCacheData As Boolean) As Boolean
On Error GoTo ErrSection:
    
    Dim nOldSize&, i&
    Dim nSymID&, strSym$
    Dim dBid#, dAsk#, dTime#
    Dim nBidSize&, nAskSize&
    Dim bNewDay As Boolean, bNewTrade As Boolean, bNewBidAsk As Boolean
    
    Dim tbToUse As cGdTable
    
    If Len(m.strMsg) > 0 Then
        LoadDataTable
        Exit Function
    End If
    
    If bCacheData Then
        Set tbToUse = m.tbCache
    Else
        If m.tbCache.NumRecords > 0 Then
            m.tbTick.AppendFromTable m.tbCache
            m.tbCache.NumRecords = 0
        End If
        Set tbToUse = m.tbTick
    End If

    'get updated tick bars
    nOldSize = m.TickBars.Size
    If g.RealTime.UpdateBars(m.TickBars, bNewDay) Then
        If bNewDay Then
            nSymID = m.TickBars.Prop(eBARS_SymbolID)
            strSym = m.TickBars.Prop(eBARS_Symbol)
            NewTradeProfileData nSymID, strSym
        ElseIf m.TickBars.Size > nOldSize Then
            bNewTrade = True
        End If
    End If

    If Not bNewDay Then
        If Not m.DailyBars Is Nothing And m.frmTS.TS_DisplayStyle = 3 Then
            bNewBidAsk = g.RealTime.UpdateBidAsk(m.DailyBars, True)
            If bNewBidAsk Then
                dBid = m.DailyBars(eBARS_Bid, m.DailyBars.Size - 1)
                dAsk = m.DailyBars(eBARS_Ask, m.DailyBars.Size - 1)
                dTime = m.DailyBars.Prop(eBARS_LastTickTime)
                nBidSize = m.DailyBars(eBARS_BidSize, m.DailyBars.Size - 1)
                nAskSize = m.DailyBars(eBARS_AskSize, m.DailyBars.Size - 1)
                
                'update table for tick-by-tick view
                If dBid > 0 And (m.dLastBid <> dBid Or m.nLastBidSize <> nBidSize) Then
                    tbToUse.AddRecord ""
                    'set time,price,size,type
                    tbToUse(eFld_Time, tbToUse.NumRecords - 1) = dTime / 1440# + m.DailyBars(eBARS_DateTime, m.DailyBars.Size - 1)
                    tbToUse(eFld_Price, tbToUse.NumRecords - 1) = dBid
                    tbToUse(eFld_Volume, tbToUse.NumRecords - 1) = nBidSize
                    tbToUse(eFld_Type, tbToUse.NumRecords - 1) = "Bid"
                    'set color
                    If m.dLastBid = 0 Then
                        m.nLastBidColor = m.frmTS.TS_UpColorBid
                        m.nLastBidUpDownFlag = 1
                    Else
                        dBid = RoundToMinMove(dBid, m.dMinMove)
                        If dBid > m.dLastBid Then
                            m.nLastBidColor = m.frmTS.TS_UpColorBid
                            m.nLastBidUpDownFlag = 1
                        ElseIf dBid < m.dLastBid Then
                            m.nLastBidColor = m.frmTS.TS_DownColorBid
                            m.nLastBidUpDownFlag = 0
                        End If
                    End If
                    m.dLastBid = RoundToMinMove(dBid, m.dMinMove)
                    m.nLastBidSize = nBidSize
                    tbToUse(eFld_Color, tbToUse.NumRecords - 1) = m.nLastBidColor
                    tbToUse(eFld_UpDownFlag, tbToUse.NumRecords - 1) = m.nLastBidUpDownFlag
                    m.bHasBidAsk = True
                End If
                If dAsk > 0 And (dAsk <> m.dLastAsk Or nAskSize <> m.nLastAskSize) Then
                    tbToUse.AddRecord ""
                    'set time,price,size,type
                    tbToUse(eFld_Time, tbToUse.NumRecords - 1) = dTime / 1440# + m.DailyBars(eBARS_DateTime, m.DailyBars.Size - 1)
                    tbToUse(eFld_Price, tbToUse.NumRecords - 1) = dAsk
                    tbToUse(eFld_Volume, tbToUse.NumRecords - 1) = nAskSize
                    tbToUse(eFld_Type, tbToUse.NumRecords - 1) = "Ask"
                    'set color
                    If m.dLastAsk = 0 Then
                        m.nLastAskColor = m.frmTS.TS_UpColorAsk
                        m.nLastAskUpDownFlag = 1
                    Else
                        dAsk = RoundToMinMove(dAsk, m.dMinMove)
                        If dAsk > m.dLastAsk Then
                            m.nLastAskColor = m.frmTS.TS_UpColorAsk
                            m.nLastAskUpDownFlag = 1
                        ElseIf dAsk < m.dLastAsk Then
                            m.nLastAskColor = m.frmTS.TS_DownColorAsk
                            m.nLastAskUpDownFlag = 0
                        End If
                    End If
                    m.dLastAsk = RoundToMinMove(dAsk, m.dMinMove)
                    m.nLastAskSize = nAskSize
                    tbToUse(eFld_Color, tbToUse.NumRecords - 1) = m.nLastAskColor
                    tbToUse(eFld_UpDownFlag, tbToUse.NumRecords - 1) = m.nLastAskUpDownFlag
                    m.bHasBidAsk = True
                End If
            End If
        End If
        If bNewTrade Then
            UpdateTableRT nOldSize, bCacheData
        End If
    End If
    
    If bNewDay Or bNewTrade Or bNewBidAsk Then
        UpdateDataRT = True
    Else
        UpdateDataRT = False
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.UpdateDataRT", eGDRaiseError_Raise

End Function

Private Sub UpdateTableRT(ByVal nOldSize&, ByVal bCacheData As Boolean)
On Error GoTo ErrSection:

    Dim iPos&, i&, iRec&, strPrice$
    Dim dPrice#, dDate#, dDatePrev#
    Dim nVol&, nColor&, nUpDownFlag&
    
    Dim bAddRecord As Boolean
    
    Dim tbToUse As cGdTable
    
    If nOldSize = m.TickBars.Size Then
        Exit Sub
    End If
    
    If bCacheData Then
        Set tbToUse = m.tbCache
    Else
        Set tbToUse = m.tbTick
    End If
        
    UpdateCumulativeTable nOldSize, m.TickBars.Size - 1, bCacheData
        
    For i = nOldSize To m.TickBars.Size - 1
        nVol = m.TickBars(eBARS_Vol, i)
        If IncludeTrade(nVol) Then
'update table for minute-by-minute view
            dPrice = RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)
            dDate = m.TickBars(eBARS_DateTime, i)
            
            iPos = m.tbMinute.NumRecords - 1
            dDatePrev = m.tbMinute(eFld_Time, iPos)
            
            If dDate = dDatePrev Then
                strPrice = m.tbMinute(eFld_Price, iPos) & " " & m.TickBars.PriceDisplay(dPrice)
            ElseIf Int(dDate) = Int(dDatePrev) Then
                If Minute(dDate) = Minute(dDatePrev) Then
                    strPrice = m.tbMinute(eFld_Price, iPos) & " " & m.TickBars.PriceDisplay(dPrice)
                Else
                    bAddRecord = True
                End If
            Else
                bAddRecord = True
            End If
            
            If bAddRecord Then
                strPrice = m.TickBars.PriceDisplay(dPrice)
                m.tbMinute.AddRecord ""
                m.tbMinute(eFld_Time, m.tbMinute.NumRecords - 1) = dDate
                m.tbMinute(eFld_Price, m.tbMinute.NumRecords - 1) = strPrice
            Else
                m.tbMinute(eFld_Price, iPos) = strPrice
            End If
            
'update table for tick-by-tick view
            For iRec = tbToUse.NumRecords To 1 Step -1
                If dDate >= tbToUse(eFld_Time, iRec - 1) Then
                    Exit For
                End If
            Next
            
            tbToUse.AddRecord "", iRec
            tbToUse(eFld_Time, iRec) = m.TickBars(eBARS_DateTime, i)
            tbToUse(eFld_Price, iRec) = m.TickBars(eBARS_Close, i)
            tbToUse(eFld_Volume, iRec) = m.TickBars(eBARS_Vol, i)
            tbToUse(eFld_Type, iRec) = "Trade"
            tbToUse(eFld_Flags, iRec) = m.TickBars(eBARS_Flags, i)
            
'            If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Or m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
'                'do nothing
'            Else
'                StatusMsg m.TickBars.Prop(eBARS_Symbol) & " " _
'                    & DateFormat(m.TickBars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM_SS) _
'                    & " Flag = " & Str(m.TickBars(eBARS_Flags, i))
'            End If
            
            If dPrice > m.dLastTradePrice Then
                nColor = m.frmTS.TS_UpColor
                nUpDownFlag = 1
            ElseIf dPrice < m.dLastTradePrice Then
                nColor = m.frmTS.TS_DownColor
                nUpDownFlag = 0
            Else
                nColor = m.nLastTradeColor
            End If
            tbToUse(eFld_Color, iRec) = nColor
            tbToUse(eFld_UpDownFlag, iRec) = nUpDownFlag
        End If
    Next
    
    If dPrice > 0 Then
        m.dLastTradePrice = dPrice
        m.nLastTradeColor = nColor
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTimeSalesData.UpdateTableRT", eGDRaiseError_Raise

End Sub

Private Sub LoadTickTable()
On Error GoTo ErrSection:

    Dim nPointerSave&
    
    Dim i&, j&
    Dim nColor&, nUpColor&, nDownColor&, nUpDownFlag&
    Dim nVol&, nSize&
    Dim dPrice#, dPricePrev#, strMsgSave$
    
    Dim hTime&, hPrice&, hVolume&, hType&, hColor&, hUpdownFlag&, hFlags&
    Dim hBarTime&, hBarClose&, hBarVolume&, hBarFlags&
        
    If m.TickBars Is Nothing Then Exit Sub
    
    nPointerSave = Screen.MousePointer
    
    m.tbTick.NumRecords = 0
    
    If m.TickBars.Size = 0 Then
        m.tbTick.AddRecord m.strMsg
        GoTo ErrExit
    End If
        
    nUpColor = m.frmTS.TS_UpColor
    nDownColor = m.frmTS.TS_DownColor
    
    hTime = m.tbTick.FieldArrayHandle(eFld_Time)
    hPrice = m.tbTick.FieldArrayHandle(eFld_Price)
    hVolume = m.tbTick.FieldArrayHandle(eFld_Volume)
    hType = m.tbTick.FieldArrayHandle(eFld_Type)
    hColor = m.tbTick.FieldArrayHandle(eFld_Color)
    hUpdownFlag = m.tbTick.FieldArrayHandle(eFld_UpDownFlag)
    hFlags = m.tbTick.FieldArrayHandle(eFld_Flags)
    
    If hTime = 0 Or hPrice = 0 Or hVolume = 0 Or hColor = 0 Or hUpdownFlag = 0 Then
        GoTo ErrExit
    End If
    
    hBarTime = m.TickBars.ArrayHandle(eBARS_DateTime)
    hBarClose = m.TickBars.ArrayHandle(eBARS_Close)
    hBarVolume = m.TickBars.ArrayHandle(eBARS_Vol)
    hBarFlags = m.TickBars.ArrayHandle(eBARS_Flags)
    
    If hBarTime = 0 Or hBarClose = 0 Or hBarVolume = 0 Then
        GoTo ErrExit
    End If
        
    strMsgSave = "Loading tick table for " & m.TickBars.Prop(eBARS_Symbol) & " ..."
    nSize = m.TickBars.Size
    
    For i = 0 To nSize - 1
        nVol = gdGetNum(hBarVolume, i)
        If IncludeTrade(nVol) Then
            j = m.tbTick.NumRecords
            m.tbTick.AddRecord ""
            
            gdSetNum hTime, j, gdGetNum(hBarTime, i)
            gdSetNum hPrice, j, gdGetNum(hBarClose, i)
            gdSetNum hVolume, j, nVol       'gdGetNum(hBarVolume, i)
            gdSetStr hType, j, "Trade"
            
            dPrice = RoundToMinMove(gdGetNum(hBarClose, i), m.dMinMove)
            If i = 0 Then
                nColor = 0
                nUpDownFlag = 0
            ElseIf dPrice > dPricePrev Then
                nColor = nUpColor
                nUpDownFlag = 1
            ElseIf dPrice < dPricePrev Then
                nColor = nDownColor
                nUpDownFlag = 0
            End If
            
            gdSetNum hColor, j, nColor
            gdSetNum hUpdownFlag, j, nUpDownFlag
            gdSetNum hFlags, j, gdGetNum(hBarFlags, i)
            
            dPricePrev = dPrice
        End If
        
        If nSize > 100000 Then
            If i Mod 1000 = 0 Then
                m.strMsg = strMsgSave & Str(i)
                m.frmTS.UpdateMessage m.strMsg
                DoEvents
                Screen.MousePointer = vbHourglass
                If m.bProfileUnload Then Exit For
            End If
        End If
    Next
    
    m.dLastTradePrice = dPrice
    m.nLastTradeColor = nColor
            
ErrExit:
    Screen.MousePointer = nPointerSave
    Exit Sub
    
ErrSection:
    Screen.MousePointer = nPointerSave
    RaiseError "cTimeSalesData.LoadTickTable", eGDRaiseError_Raise

End Sub

Public Sub RemoveBidAskRecords()
On Error GoTo ErrSection:

    Dim i&
    
    If Not m.bHasBidAsk Then Exit Sub
    
    For i = m.tbTick.NumRecords - 1 To 0 Step -1
        If m.tbTick(eFld_Type, i) = "Ask" Or m.tbTick(eFld_Type, i) = "Bid" Then
            m.tbTick.RemoveRecords i
        End If
    Next
    
    m.bHasBidAsk = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTimeSalesData.RemoveBidAskRecords", eGDRaiseError_Raise

End Sub

Public Sub ChangeColor(ByVal strType$, ByVal nNewColor&, ByVal nUpDownFlag&)
On Error GoTo ErrSection:

    Dim i&
    
    For i = 0 To m.tbTick.NumRecords - 1
        If m.tbTick(eFld_Type, i) = strType And m.tbTick(eFld_UpDownFlag, i) = nUpDownFlag Then
            m.tbTick(eFld_Color, i) = nNewColor
        End If
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTimeSalesData.RemoveBidAskRecords", eGDRaiseError_Raise

End Sub

Private Function IncludeTrade(ByVal iVol&) As Boolean
On Error GoTo ErrSection:

    Dim bInclude As Boolean

    If m.nVolFilterMin > 0 Then
        If iVol >= m.nVolFilterMin Then
            If m.nVolFilterMax > 0 Then
                If iVol <= m.nVolFilterMax Then bInclude = True
            Else
                bInclude = True
            End If
        End If
    ElseIf m.nVolFilterMax > 0 Then
        If iVol < m.nVolFilterMax Then bInclude = True
    Else
        bInclude = True
    End If
    
    IncludeTrade = bInclude

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cTimeSalesData.IncludeTrade"

End Function

Public Property Get VolFilterMax() As Long
On Error GoTo ErrSection:

    VolFilterMax = m.nVolFilterMax

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.VolFilterMax.Get"

End Property

Public Property Let VolFilterMax(ByVal nVolMax&)
On Error GoTo ErrSection:
    
    If nVolMax >= 0 Then m.nVolFilterMax = nVolMax

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.VolFilterMax.Let"

End Property

Public Property Get VolFilterMin() As Long
On Error GoTo ErrSection:

    VolFilterMin = m.nVolFilterMin

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.nVolFilterMin.Get"

End Property

Public Property Let VolFilterMin(ByVal nVolMin&)
On Error GoTo ErrSection:

    If nVolMin >= 0 Then m.nVolFilterMin = nVolMin

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.nVolFilterMin.Let"

End Property

Public Property Get CanHaveVolFilter() As Boolean
On Error GoTo ErrSection:

    If Not m.TickBars Is Nothing Then
        If m.TickBars.Size > 0 Then
            If m.TickBars(eBARS_Vol, m.TickBars.Size - 1) > 0 Then CanHaveVolFilter = True
        End If
    End If
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.CanHaveVolFilter.Get"

End Property

Public Property Get TradeProfileUnload() As Boolean
    TradeProfileUnload = m.bProfileUnload
End Property

Public Property Let TradeProfileUnload(ByVal bUnload As Boolean)
    m.bProfileUnload = bUnload
End Property

Private Sub LoadCumulativeTable()
On Error GoTo ErrExit:
    
    If m.TickBars Is Nothing Then Exit Sub
    
    m.tbCumulative.NumRecords = 0
    
    m.nSumBuyVol = 0
    m.nSumSellVol = 0
    m.bSumBySell = False
    
    If m.TickBars.Size = 0 Then
        m.tbCumulative.AddRecord m.strMsg
        Exit Sub
    End If
        
    UpdateCumulativeTable 0, m.TickBars.Size - 1, False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTimeSalesData.LoadCumulativeTable"

End Sub

Private Sub UpdateCumulativeTable(ByVal nStart&, ByVal nEnd&, ByVal bCacheData As Boolean)
On Error GoTo ErrExit:

    Dim i&, j&, k&, dPrice#, dPricePrev#
    Dim nVol&, nLargestVol&, nVolInTbl&
    Dim nColor&, nUpDownFlag&, nSize&
    
    Dim hTime&, hPrice&, hTrades&, hContracts&, hAvgSize&, hLargestSize&
    Dim hBuyVol&, hSellVol&, hColor&, hUpdownFlag&
    Dim hBarTime&, hBarClose&, hBarVolume&, hBarContracts&, hFlags&
    
    Dim eFlag As eTickFlags
    Dim tbToUse As cGdTable
    
    If bCacheData Then
        Set tbToUse = m.tbCacheCumulative
        If m.tbCacheCumulative.NumRecords > 0 Then
            i = m.tbCacheCumulative.NumRecords - 1
            dPricePrev = m.tbCacheCumulative(eFldCumulative_Price, i)
            nColor = m.tbCacheCumulative(eFldCumulative_Color, i)
        ElseIf m.tbCumulative.NumRecords > 0 Then
            i = m.tbCumulative.NumRecords - 1
            dPricePrev = m.tbCumulative(eFldCumulative_Price, i)
            nColor = m.tbCumulative(eFldCumulative_Color, i)
        End If
    Else
        If m.tbCacheCumulative.NumRecords > 0 Then
            m.tbCumulative.AppendFromTable m.tbCacheCumulative
            m.tbCacheCumulative.NumRecords = 0
        End If
        
        Set tbToUse = m.tbCumulative
        If m.tbCumulative.NumRecords > 0 Then
            i = m.tbCumulative.NumRecords - 1
            dPricePrev = m.tbCumulative(eFldCumulative_Price, i)
            nColor = m.tbCumulative(eFldCumulative_Color, i)
        End If
    End If
    
    'get handles to arrays in cumulative table
    hTime = tbToUse.FieldArrayHandle(eFldCumulative_Time)
    hPrice = tbToUse.FieldArrayHandle(eFldCumulative_Price)
    hTrades = tbToUse.FieldArrayHandle(eFldCumulative_Trades)
    hContracts = tbToUse.FieldArrayHandle(eFldCumulative_Contracts)
    hAvgSize = tbToUse.FieldArrayHandle(eFldCumulative_AvgSize)
    hLargestSize = tbToUse.FieldArrayHandle(eFldCumulative_LargestSize)
    hBuyVol = tbToUse.FieldArrayHandle(eFldCumulative_BuyVol)
    hSellVol = tbToUse.FieldArrayHandle(eFldCumulative_SellVol)
    hColor = tbToUse.FieldArrayHandle(eFldCumulative_Color)
    hUpdownFlag = tbToUse.FieldArrayHandle(eFldCumulative_UpDownFlag)
    
    If hTime = 0 Or hPrice = 0 Or hTrades = 0 Or hContracts = 0 Or hAvgSize = 0 Or _
       hLargestSize = 0 Or hBuyVol = 0 Or hSellVol = 0 Or hColor = 0 Or hUpdownFlag = 0 Then
        GoTo ErrExit
    End If
    
    'get handles to arrays in tick bars
    hBarTime = m.TickBars.ArrayHandle(eBARS_DateTime)
    hBarClose = m.TickBars.ArrayHandle(eBARS_Close)
    hBarVolume = m.TickBars.ArrayHandle(eBARS_Vol)
    hFlags = m.TickBars.ArrayHandle(eBARS_Flags)
    
    If hBarTime = 0 Or hBarClose = 0 Or hBarVolume = 0 Or hFlags = 0 Then
        GoTo ErrExit
    End If

'Pete 's email - 04/23/2012
'Columns that can be displayed:
'Time - of the first trade on that line
'Price (existing with uptick/down tick color change)
'Trades - total number of trades at that price since line was created until price moves away.
'Contracts – total number of contracts at this price line since it was created until price moves away.
'Avg Trade Size – number of contracts at this price divided by number of trades.
'Largest Trade – the largest trade by contract size at this price since it was created until price moves away.
'Bid – Counts how many contracts traded at the bid.  Like the buy / sell volume on the ladder.
'Ask – counts how many contracts traded at the ask.  Like the buy / sell volume on the ladder.

    For i = nStart To nEnd
        nVol = gdGetNum(hBarVolume, i)
        
        If nVol <> kNullData Then
            j = tbToUse.NumRecords
            If IncludeTrade(nVol) Then
                dPrice = RoundToMinMove(gdGetNum(hBarClose, i), m.dMinMove)
                
                If dPrice = dPricePrev Then
                    k = j - 1
                    If nVol > nLargestVol Then nLargestVol = nVol
                    If gdGetNum(hLargestSize, k) > nLargestVol Then nLargestVol = gdGetNum(hLargestSize, k)
                    
                    gdSetNum hTrades, k, gdGetNum(hTrades, k) + 1
                    gdSetNum hContracts, k, nVol + gdGetNum(hContracts, k)
                Else
                    k = j
                    nLargestVol = nVol
                    tbToUse.AddRecord ""
                    
                    If i = 0 Then
                        nColor = 0
                        nUpDownFlag = 0
                    ElseIf dPrice > dPricePrev Then
                        nColor = m.frmTS.TS_UpColor
                        nUpDownFlag = 1
                    Else
                        nColor = m.frmTS.TS_DownColor
                        nUpDownFlag = 0
                    End If
                    
                    gdSetNum hTrades, k, 1
                    gdSetNum hContracts, k, nVol
                    gdSetNum hBuyVol, k, 0          'new row in table
                    gdSetNum hSellVol, k, 0
                    gdSetNum hTime, k, gdGetNum(hBarTime, i)
                    
                    dPricePrev = dPrice
                End If
                
                
                gdSetNum hPrice, k, gdGetNum(hBarClose, i)
                gdSetNum hAvgSize, k, gdGetNum(hContracts, k) / gdGetNum(hTrades, k)
                gdSetNum hLargestSize, k, nLargestVol

                'eTICK_AtBid = 1     ' i.e. SELL Volume
                'eTICK_AtAsk = 2     ' i.e. BUY Volume
                
                eFlag = gdGetNum(hFlags, i)
                
                If eFlag = eTICK_AtBid Then
                    nVolInTbl = gdGetNum(hSellVol, k)
                    gdSetNum hSellVol, k, nVol + nVolInTbl
                    m.dLastTradePriceAtBid = dPrice
                    If m.bSumBySell Then m.nSumSellVol = m.nSumSellVol + nVol
                ElseIf eFlag = eTICK_AtAsk Then
                    nVolInTbl = gdGetNum(hBuyVol, k)
                    gdSetNum hBuyVol, k, nVol + nVolInTbl
                    m.dLastTradePriceAtAsk = dPrice
                    If m.bSumBySell Then m.nSumBuyVol = m.nSumBuyVol + nVol
                End If
                
                gdSetNum hColor, k, nColor
                gdSetNum hUpdownFlag, k, nUpDownFlag
            End If
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTimeSalesData.UpdateCumulativeTable"

End Sub

Public Sub ClearData()
On Error GoTo ErrSection:

    m.tbCache.NumRecords = 0
    m.tbCacheCumulative.NumRecords = 0
    
    m.tbCumulative.NumRecords = 0
    m.tbMinute.NumRecords = 0
    m.tbTick.NumRecords = 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTimeSalesData.ClearData"

End Sub

Public Property Let SumBuySell(ByVal bStart As Boolean)
On Error GoTo ErrSection:

    m.bSumBySell = bStart
    m.nSumBuyVol = 0
    m.nSumSellVol = 0

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cTimeSalesData.SumBuySell.Let"

End Property

Public Property Get SumBuyVol() As Long
    SumBuyVol = m.nSumBuyVol
End Property

Public Property Get SumSellVol() As Long
    SumSellVol = m.nSumSellVol
End Property
