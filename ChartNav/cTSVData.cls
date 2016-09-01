VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTSVData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVSFlexDataSource

Private Const kVolAtPriceFile = "\Provided\mpVolAtPrice.txt"

Public Enum eTSV_Table_Type
    eTSVTb_TbType_Invalid = -1
    eTSVTb_TbType_VolAtPrice
    eTSVTb_TbType_Reversal
End Enum

Public Enum eTSV_TbFields
    eTSVTb_Price = 0
    eTSVTb_VolBid_A
    eTSVTb_VolAsk_A
    eTSVTb_VolBid_B
    eTSVTb_VolAsk_B
    eTSVTb_VolBid_C
    eTSVTb_VolAsk_C
    eTSVTb_VolBid_D
    eTSVTb_VolAsk_D
    eTSVTb_VolBid_E
    eTSVTb_VolAsk_E
    eTSVTb_VolBid_F
    eTSVTb_VolAsk_F
    eTSVTb_VolBid_G
    eTSVTb_VolAsk_G
    eTSVTb_VolBid_H
    eTSVTb_VolAsk_H
    eTSVTb_VolBid_I
    eTSVTb_VolAsk_I
    eTSVTb_VolBid_J
    eTSVTb_VolAsk_J
    eTSVTb_VolBid_K
    eTSVTb_VolAsk_K
    eTSVTb_VolBid_L
    eTSVTb_VolAsk_L
    eTSVTb_VolBid_M
    eTSVTb_VolAsk_M
    eTSVTb_VolBid_N
    eTSVTb_VolAsk_N
    eTSVTb_VolRow_Total
    eTSVTb_VolOther_A
    eTSVTb_VolOther_B
    eTSVTb_VolOther_C
    eTSVTb_VolOther_D
    eTSVTb_VolOther_E
    eTSVTb_VolOther_F
    eTSVTb_VolOther_G
    eTSVTb_VolOther_H
    eTSVTb_VolOther_I
    eTSVTb_VolOther_J
    eTSVTb_VolOther_K
    eTSVTb_VolOther_L
    eTSVTb_VolOther_M
    eTSVTb_VolOther_N
End Enum

Public Enum eTSV_TbStatsFields
    eTSVTb_Col_A
    eTSVTb_Col_B
    eTSVTb_Col_C
    eTSVTb_Col_D
    eTSVTb_Col_E
    eTSVTb_Col_F
    eTSVTb_Col_G
    eTSVTb_Col_H
    eTSVTb_Col_I
    eTSVTb_Col_J
    eTSVTb_Col_K
    eTSVTb_Col_L
    eTSVTb_Col_M
    eTSVTb_Col_N
    eTSVTb_Col_Total
End Enum

Public Enum eTSV_TbStatsRows
    eTSVTb_Row_Bid
    eTSVTb_Row_Ask
    eTSVTb_Row_Delta
    eTSVTb_Row_DeltaCum
    eTSVTb_Row_Percent
    eTSVTb_Row_VAM
    eTSVTb_Row_Mean
    eTSVTb_Row_VBM
    eTSVTb_Row_VolTotal
    eTSVTb_Row_AvgVol
    eTSVTb_Row_SHVol
    eTSVTb_Row_Range
    eTSVTb_Row_AvgRange
    eTSVTb_Row_SHRange
    eTSVTb_Row_OtherVol
End Enum

Public Enum eTSV_TbAuctionIdx
    eTSVTb_Idx_Hi
    eTSVTb_Idx_Low
    eTSVTb_Idx_Close
    eTSVTb_Idx_Above
    eTSVTb_Idx_Mean
    eTSVTb_Idx_Below
    eTSVTb_Idx_HiestVol
End Enum

Public Enum eTSV_Groups
    eTSV_Group_AB
    eTSV_Group_AE
    eTSV_Group_AI
    eTSV_Group_AN
    eTSV_Group_Current
End Enum

'structure to hold index of which row in price vol table to
'color green/yellow for each of the A-B, A-E, A-I, A-N group
'the color bars actually go from A-B, C-E, F-I, J-N
'but IOAMT document refer to them A-B,A-E,A-I,A-N to
'emphasize that the calculations are cumulative from col A
Private Type eTSV_ColorRow
    iMeanCurrent As Long
    iMeanAB As Long
    iMeanAE As Long
    iMeanAI As Long
    iMeanAN As Long
    iAboveCurrent As Long
    iAboveAB As Long
    iAboveAE As Long
    iAboveAI As Long
    iAboveAN As Long
    iBelowCurrent As Long
    iBelowAB As Long
    iBelowAE As Long
    iBelowAI As Long
    iBelowAN As Long
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PriceVol table (tbPriceVol) usage: volume at price & bid/ask directional analysis
'   volume at price:
'       - fixed fields(columns)  - 1 for price, 3x14 for bid/ask/other vol, 1 for totals = 44
'       - variable records(rows) - row 0 holds bar up/down flag
'                                  row 1 holds timestamp
'                                  rows >= 2 hold every price minmove apart from high to low
'   bid/ask directional analysis
'       - variable fields(columns) - 1 for price, 3 per bar for bid/ask/other vol
'       - variable records(rows)   - holds every price minmove apart from high to low
'
'Stats table (tbStats) usage: volume at price & bid/ask directional analysis
'   volume at price:
'       - fixed fields(columns) - 2x14 for A-N, 1 for totals = 15
'       - fixed records(rows)   - BidVol,AskVol,[+ -],[+ -]Cum,Percent,VAM,Mean,VBM,
'                                 Totals,AvgVol,SHVol,Range,AvgRange,SHRange = 14
'   bid/ask directional analysis
'       - variable fields(columns) - 1 for name, 3 per bar for bid/ask/other vol
'       - fixed records(rows)      - Totals,Differential,Total %,Big Lots,Big Lots %,
'                                    Cum Big,Cum Big Diff,Cum Big %,Cum Big D% of Tot,
'                                    Cum,Cum Diff,Cum %,Cum Diff% of Tot = 13
'
'PullBack tables (tbPBVol, tbPBStats) usage: bid/ask directional analysis
'   tbPBVol:
'       - fixed fields(columns)  - 3 per bar for bid/ask/other vol
'       - variable records(rows) - row 0 & 1 are blank placeholder to match PriceVol table
'                                  rows >= 2 hold every price minmove apart from high to low
'
'   tbPBStats:
'       - fixed fields(columns)  - 1 each for bid/ask/other vol = 3
'       - fixed records(rows)    - Totals,Differential,Big Lots = 3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type mPrivate
    TickBars As cGdBars         'full ticks
    
    BidAskVolBars As cGdBars    'bid/ask volume grid (minute, vol, tick or range bars - TimeSalesAnalyzer)
    TradeSpeedBars As cGdBars   'trade speed grid (minute, vol, or tick bars - TimeSalesAnalyzer)
    aTimeRangeBars As cGdArray  'array of bars for custom time ranges (TimeSalesAnalyzer)

    tbPriceVol As cGdTable
    tbStats As cGdTable
    tbPBVol As cGdTable
    tbPBStats As cGdTable
    tbCumVolAtPrice As cGdTable 'holds cumulative volume at specific price
    
    aTbIdx As cGdArray          'index for PriceVol table
    aLineTool As cGdArray       'holds color for rows highlighted with [Draw] tool
    aBoxTool As cGdArray        'holds info for boxes drawn with [Draw] tool
    aBoxToolOther As cGdArray   'holds info for boxes drawn with [Draw] tool with tick reversal <> current tick reversal
    aIconTool As cGdArray
    
    frmAnalyzer As frmTimeSalesAnalyzer   'for coloring virtual grid (TimeSalesAnalyzer)
    
    'holds index into price-vol table identifying hi,low,close,mean,value area etc. on a per-column basis
    tbAuction As cGdTable
    'holds index of rows to color green/yellow for group (VolAtPrice)
    tpColorRows As eTSV_ColorRow
    'holds index of previous color rows
    tpPrevColorRows As eTSV_ColorRow
    'holds statistically high values for drawing blue box
    aStatHighVal As cGdArray
    
    'for updating data real time (VolAtPrice)
    dTbColTimeStart As Double
    dTbColTimeEnd As Double
    dPxVx As Double                 'Px*Vx up to last complete col
    dVx As Double                   'dVx up to last complete col
    ePxVxCol As eTSV_TbFields       'col that PxVx & Vx was summed up to
    nTbLastCol As Long
    nTbStatsLastCol As Long
    iTotalColVolHi As Long          'index of row in total column that has highest volume
    
    'configurable settings for Bid/Ask Directional Analysis
    iReverseTick As Long            'number of ticks before reversal
    iBigLot As Long                 'number for highlighting big lot
    
    iFirstBar As Long               'for building bid/ask directional table
    iBlankRows As Long              'extra rows above/below price high/low
    dHighTradeAtBid As Double
    dLowTradeAtAsk As Double
    dPriceCumVAM As Double          'for displaying in price col of bid/ask directional grid
    dPriceCumVBM As Double          'for displaying in price col of bid/ask directional grid
    
    dTradeSpeed As Double           'number for highlighting trades per minute
    dBidTrades As Double            'running count of trades at bid
    dAskTrades As Double            'running ocunt of trades at ask
    aSpeedCells As cGdArray         'holds "row,col" for highlighting cells with tradeSpeed > dTradeSpeed
    
    nSessionDate As Long
    nStartTime As Long              'requested start time as minutes since midnight
    nTimeSalesVolTotal As Long
    dMinMove As Double
    
    nTradesAtBid As Long            'running total used for bid ask directional form
    nTradesAtAsk As Long            'running total used for bid ask directional form
    dPriceOfLastSwitch As Double    'price where differential goes from pos->neg or vice versa
    
    strFile As String               'file name for draw tools

    eTbType As eTSV_Table_Type
End Type

Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.TickBars = New cGdBars
    Set m.BidAskVolBars = New cGdBars
    Set m.TradeSpeedBars = New cGdBars
    Set m.aStatHighVal = New cGdArray
    Set m.aSpeedCells = New cGdArray
    
    m.dMinMove = -1
    m.nStartTime = -1
    m.eTbType = eTSVTb_TbType_Invalid
    
    m.iBlankRows = -1
    m.iReverseTick = 4
    m.iBigLot = 500
    m.dTradeSpeed = 10000
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.Class_Initialize"

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    SaveDrawTools

    Set m.TickBars = Nothing
    Set m.BidAskVolBars = Nothing
    Set m.TradeSpeedBars = Nothing
    Set m.aStatHighVal = Nothing
    Set m.aSpeedCells = Nothing
    
    Set m.tbPriceVol = Nothing
    Set m.tbStats = Nothing
    Set m.tbAuction = Nothing
    Set m.tbPBVol = Nothing
    Set m.tbPBStats = Nothing
    Set m.tbCumVolAtPrice = Nothing
    
    Set m.aTbIdx = Nothing
    Set m.aLineTool = Nothing
    Set m.aBoxTool = Nothing
    Set m.aBoxToolOther = Nothing
    Set m.aIconTool = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.Class_Terminate"

End Sub

Private Function IVSFlexDataSource_GetData(ByVal field As Long, ByVal Record As Long) As String

    Dim strData$, dDateTime#, i&
                    
    strData = ""
    
    i = m.TickBars.Size - (Record + 1)
    If i >= 0 And i < m.TickBars.Size Then
        If field = 0 Then
            dDateTime = m.TickBars(eBARS_DateTime, i)
            If g.bShowInLocalTimeZone Then
                dDateTime = ConvertTimeZone(dDateTime, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
            strData = DateFormat(dDateTime, NO_DATE, HH_MM_SS)
        ElseIf field = 1 Then
            strData = m.TickBars.PriceDisplay(m.TickBars(eBARS_Close, i))
        ElseIf field = 2 Then
            If m.TickBars(eBARS_Vol, i) > 0 Then
                strData = Format(m.TickBars(eBARS_Vol, i), "#,##0")
                m.nTimeSalesVolTotal = m.nTimeSalesVolTotal + m.TickBars(eBARS_Vol, i)
            Else
                strData = ""
            End If
        ElseIf field = 3 Then
            strData = m.TickBars(eBARS_Flags, i)
        End If
    End If
    
    IVSFlexDataSource_GetData = strData
    
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = 4
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal field As Long) As String
    
    If field = 0 Then
        IVSFlexDataSource_GetFieldName = "Time"
    ElseIf field = 1 Then
        IVSFlexDataSource_GetFieldName = "Price"
    ElseIf field = 2 Then
        IVSFlexDataSource_GetFieldName = "Volume"
    ElseIf field = 3 Then
        IVSFlexDataSource_GetFieldName = "Flag"
        If Not m.frmAnalyzer Is Nothing Then
            m.frmAnalyzer.fgTimeSales.ColHidden(3) = True
        End If
    End If

End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = m.TickBars.Size
End Function

Private Sub IVSFlexDataSource_SetData(ByVal field As Long, ByVal Record As Long, ByVal NewData As String)
    'not editing
End Sub

Public Sub TimeRangeBarData(Bars As cGdBars, ByVal dDateStart#, ByVal dDateEnd#, _
    ByVal iStart&, ByVal iEnd&, ByVal iBarIdx&)
    
    Dim dVol#, dBidVol#, dAskVol#
    Dim dTradeAtBid&, dTradeAtAsk&
    Dim dBarDate#, i&
        
    'go through bars and sum up total trade vol, trade at bid vol & trade at ask vol
    For i = iStart To iEnd
        If i < m.TickBars.Size Then
            dBarDate = m.TickBars(eBARS_DateTime, i)
            If dBarDate <> kNullData And dDateStart <> kNullData And dDateEnd <> kNullData Then
                'disregard seconds in tick bars date time
                'check that time of bar is between start & end time passed in
                If Hour(dBarDate) > Hour(dDateStart) Or (Hour(dBarDate) = Hour(dDateStart) And Minute(dBarDate) >= Minute(dDateStart)) Then
                    If Hour(dBarDate) < Hour(dDateEnd) Or (Hour(dBarDate) = Hour(dDateEnd) And Minute(dBarDate) < Minute(dDateEnd)) Then
                        dVol = dVol + m.TickBars(eBARS_Vol, i)
                        If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                            dBidVol = dBidVol + m.TickBars(eBARS_Vol, i)
                            dTradeAtBid = dTradeAtBid + 1
                        ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                            dAskVol = dAskVol + m.TickBars(eBARS_Vol, i)
                            dTradeAtAsk = dTradeAtAsk + 1
                        End If
                    End If
                End If
            End If
        End If
    Next
            
    If iBarIdx >= 0 And iBarIdx < Bars.Size Then          'precautionary
        If dDateEnd > 0 Then
            Bars(eBARS_DateTime, iBarIdx) = dDateEnd
            If dVol > 0 Then
                Bars(eBARS_Vol, iBarIdx) = Bars(eBARS_Vol, iBarIdx) + dVol
            End If
            If dBidVol > 0 Then
                Bars(eBARS_BidVol, iBarIdx) = Bars(eBARS_BidVol, iBarIdx) + dBidVol
                Bars(eBARS_Bid, iBarIdx) = Bars(eBARS_Bid, iBarIdx) + dTradeAtBid
                Bars(eBARS_UpTicks, iBarIdx) = Bars(eBARS_UpTicks, iBarIdx) + dTradeAtBid
            End If
            If dAskVol > 0 Then
                Bars(eBARS_AskVol, iBarIdx) = Bars(eBARS_AskVol, iBarIdx) + dAskVol
                Bars(eBARS_Ask, iBarIdx) = Bars(eBARS_Ask, iBarIdx) + dTradeAtAsk
                Bars(eBARS_DownTicks, iBarIdx) = Bars(eBARS_DownTicks, iBarIdx) + dTradeAtAsk
            End If
        End If
    End If

End Sub

Public Sub ResetTimeRangeBars(fg As VSFlexGrid, ByVal iGridStart&, ByVal iGridEnd&)
On Error GoTo ErrSection:

    Dim CustomRange As cGdBars
    
    Dim j&, i&, iStart&, iEnd&
    Dim dDate1#, dDate2#
    
    Dim bOkay As Boolean
    Dim bToday As Boolean       'flag to indicate whether start time is for today's date
    
        
    Set m.aTimeRangeBars = New cGdArray
    
    With fg
        For j = iGridStart To iGridEnd
            bOkay = True
            bToday = True
            If Not .MergeRow(j) And Len(.TextMatrix(j, 0)) > 0 And Len(.TextMatrix(j, 1)) > 0 Then
                'locate bars with start time
                dDate1 = DateValue(Now) + TimeValue(.TextMatrix(j, 0))
                iStart = m.TickBars.FindDateTime(dDate1)
                If iStart >= m.TickBars.Size Then
                    'go back a day for symbols that trade overnight
                    dDate1 = DateValue(Now - 1) + TimeValue(.TextMatrix(j, 0))
                    iStart = m.TickBars.FindDateTime(dDate1)
                    If iStart = 0 Then
                        If m.TickBars(eBARS_DateTime, 0) = dDate1 Then
                            bToday = False
                        Else
                            'data not found
                            Set CustomRange = New cGdBars
                            m.aTimeRangeBars.Add CustomRange
                            bOkay = False
                        End If
                    ElseIf iStart < 0 Then
                        'data not found
                        Set CustomRange = New cGdBars
                        m.aTimeRangeBars.Add CustomRange
                        bOkay = False
                    ElseIf iStart >= m.TickBars.Size Then
                        'data not yet available
                        Set CustomRange = New cGdBars
                        m.aTimeRangeBars.Add CustomRange
                        bOkay = False
                    Else
                        bToday = False
                    End If
                End If
            Else
                'empty strings in grid
                Set CustomRange = New cGdBars
                m.aTimeRangeBars.Add CustomRange
                bOkay = False
            End If
            If bOkay Then
                'locate bars with end time
                If bToday Then
                    dDate2 = DateValue(Now) + TimeValue(.TextMatrix(j, 1))
                Else
                    dDate2 = DateValue(Now - 1) + TimeValue(.TextMatrix(j, 1))
                End If
                
                'do this if endtime is inclusive
                'add 1 minute to end time on grid to get all bars matching end time
                'iEnd = m.TickBars.FindDateTime(dDate2 + 60 / 1440 / 60)
                'If iEnd > 0 Then iEnd = iEnd - 1    'back up one to not include ending minute
                    
                iEnd = m.TickBars.FindDateTime(dDate2)
                If iEnd >= m.TickBars.Size Then
                    If m.TickBars(eBARS_DateTime, m.TickBars.Size - 1) < dDate2 Then
                        iEnd = m.TickBars.Size - 1
                    Else
                        Set CustomRange = New cGdBars
                        m.aTimeRangeBars.Add CustomRange
                        bOkay = False
                    End If
                End If

            End If
            
            'create new bar and set data for that bar
            If bOkay Then
                Set CustomRange = New cGdBars
                CustomRange.Size = 1
                CustomRange(eBARS_Vol, 0) = 0
                CustomRange(eBARS_AskVol, 0) = 0
                CustomRange(eBARS_BidVol, 0) = 0
                CustomRange(eBARS_Ask, 0) = 0
                CustomRange(eBARS_Bid, 0) = 0
                TimeRangeBarData CustomRange, dDate1, dDate2, iStart, iEnd, 0
                m.aTimeRangeBars.Add CustomRange
            End If
        Next
    End With
    
    Set CustomRange = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.ResetTimeRangeBars"

End Sub

Private Sub ResetCustomTickBars(Bars As cGdBars, ByVal strPeriod)
On Error GoTo ErrSection:

    Dim dDate1#, dDate2#
    Dim dBidVol#, dAskVol#, dVol#, dTicks#
    
    Dim iTicks&, i&, j&

    iTicks = Val(Left(strPeriod, Len(strPeriod) - 1))
    
    Set Bars = New cGdBars
    Bars.BuildBars strPeriod, m.TickBars.BarsHandle
    j = 0
    
    Bars(eBARS_DateTime, j) = m.TickBars(eBARS_DateTime, 0)
    Bars(eBARS_Vol, j) = 0
    Bars(eBARS_BidVol, j) = 0
    Bars(eBARS_AskVol, j) = 0
    Bars(eBARS_Bid, j) = 0
    Bars(eBARS_Ask, j) = 0
    Bars(eBARS_UpTicks, j) = 0      'see note in ResetMinuteBars for this
    Bars(eBARS_DownTicks, j) = 0
    
    For i = 0 To m.TickBars.Size - 1
        If dTicks < iTicks Then
            dTicks = dTicks + 1
            dVol = m.TickBars(eBARS_Vol, i)
            Bars(eBARS_Vol, j) = Bars(eBARS_Vol, j) + dVol
            Bars(eBARS_Open, j) = dTicks
            If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                Bars(eBARS_BidVol, j) = Bars(eBARS_BidVol, j) + dVol
                Bars(eBARS_Bid, j) = Bars(eBARS_Bid, j) + 1
                Bars(eBARS_UpTicks, j) = Bars(eBARS_UpTicks, j) + 1
            ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                Bars(eBARS_AskVol, j) = Bars(eBARS_AskVol, j) + dVol
                Bars(eBARS_Ask, j) = Bars(eBARS_Ask, j) + 1
                Bars(eBARS_DownTicks, j) = Bars(eBARS_DownTicks, j) + 1
            End If
        Else
            dTicks = 0
            If j + 1 < Bars.Size Then
                j = j + 1
            
                Bars(eBARS_Vol, j) = 0
                Bars(eBARS_BidVol, j) = 0
                Bars(eBARS_AskVol, j) = 0
                Bars(eBARS_Bid, j) = 0
                Bars(eBARS_Ask, j) = 0
                Bars(eBARS_UpTicks, j) = 0
                Bars(eBARS_DownTicks, j) = 0
                
                Bars(eBARS_DateTime, j) = m.TickBars(eBARS_DateTime, i)
                i = i - 1
            End If
        End If
    Next
    
    If j = Bars.Size - 1 Then
        'store index of last full ticks bar that was processed
        Bars(eBARS_Close, Bars.Size - 1) = m.TickBars.Size - 1
        'store total number of trades included in this bar
        Bars(eBARS_Open, Bars.Size - 1) = dTicks
    Else
        Set Bars = Nothing
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.ResetCustomTickBars"

End Sub

Private Sub ResetVolumeBars(Bars As cGdBars, ByVal strPeriod)
On Error GoTo ErrSection:

    Dim dVol#, dVolSum#
    Dim iVol&, i&, j&

    iVol = Val(Left(strPeriod, Len(strPeriod) - 1))
    
    Set Bars = New cGdBars
    Bars.BuildBars strPeriod, m.TickBars.BarsHandle
    j = 0
    
    Bars(eBARS_DateTime, j) = m.TickBars(eBARS_DateTime, 0)
    Bars(eBARS_Vol, j) = 0
    Bars(eBARS_BidVol, j) = 0
    Bars(eBARS_AskVol, j) = 0
    Bars(eBARS_Bid, j) = 0
    Bars(eBARS_Ask, j) = 0
    Bars(eBARS_UpTicks, j) = 0              'see note in ResetMinuteBars for this
    Bars(eBARS_DownTicks, j) = 0
    
    For i = 0 To m.TickBars.Size - 1
        If dVolSum < iVol Then
            dVol = m.TickBars(eBARS_Vol, i)
            dVolSum = dVolSum + dVol
            Bars(eBARS_Vol, j) = Bars(eBARS_Vol, j) + dVol
            If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                Bars(eBARS_BidVol, j) = Bars(eBARS_BidVol, j) + dVol
                Bars(eBARS_Bid, j) = Bars(eBARS_Bid, j) + 1
                Bars(eBARS_UpTicks, j) = Bars(eBARS_UpTicks, j) + 1
            ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                Bars(eBARS_AskVol, j) = Bars(eBARS_AskVol, j) + dVol
                Bars(eBARS_Ask, j) = Bars(eBARS_Ask, j) + 1
                Bars(eBARS_DownTicks, j) = Bars(eBARS_DownTicks, j) + 1
            End If
        Else
            dVolSum = 0
            If j + 1 < Bars.Size Then
                j = j + 1
            
                Bars(eBARS_Vol, j) = 0
                Bars(eBARS_BidVol, j) = 0
                Bars(eBARS_AskVol, j) = 0
                Bars(eBARS_Bid, j) = 0
                Bars(eBARS_Ask, j) = 0
                Bars(eBARS_UpTicks, j) = 0
                Bars(eBARS_DownTicks, j) = 0
                
                Bars(eBARS_DateTime, j) = m.TickBars(eBARS_DateTime, i)
                i = i - 1
            End If
        End If
    Next
    
    If j = Bars.Size - 1 Then
        'store index of last full ticks bar that was processed
        Bars(eBARS_Close, Bars.Size - 1) = m.TickBars.Size - 1
    Else
        Set Bars = Nothing          'shouldn't get here
    End If
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.ResetVolumeBars"

End Sub

Private Sub ResetMinuteBars(Bars As cGdBars, ByVal strPeriod)
On Error GoTo ErrSection:

    Dim TempBars As New cGdBars
    Dim dDate1#, dDate2#
    Dim iStart&, iEnd&
    Dim iMinute&, i&
        
    TempBars.ArrayMask = eBARS_Intraday
    TempBars.BuildBars strPeriod, m.TickBars.BarsHandle
    
    Set Bars = New cGdBars
    Bars.CopyFromHandle (TempBars.BarsHandle)
    
    iMinute = Val(Left(strPeriod, Len(strPeriod) - 1))
        
    For i = 0 To TempBars.Size - 1
        dDate2 = TempBars(eBARS_DateTime, i)
        If dDate1 = 0 Then dDate1 = dDate2 - (60 / 1440 / 60) * iMinute
                
        iStart = m.TickBars.FindDateTime(dDate1)
        iEnd = m.TickBars.FindDateTime(dDate2)
        If iEnd >= m.TickBars.Size Then iEnd = m.TickBars.Size - 1
    
        Bars(eBARS_Vol, i) = 0
        Bars(eBARS_BidVol, i) = 0
        Bars(eBARS_AskVol, i) = 0
        Bars(eBARS_Bid, i) = 0
        Bars(eBARS_Ask, i) = 0
        
        '08-07-2006: eBARS_Bid & eBARS_Ask arrays not available for intraday bars
        'setting the bars type to eBARS_Intraday OR eBARS_BidAsk does not help
        'using UpTics to hold number of trades at bid and DownTicks for number of trades at ask
        Bars(eBARS_UpTicks, i) = 0
        Bars(eBARS_DownTicks, i) = 0
        
        TimeRangeBarData Bars, dDate1, dDate2, iStart, iEnd, i
        Bars(eBARS_DateTime, i) = dDate2
        dDate1 = dDate2
    Next

    'store index of last full ticks bar that was processed
    Bars(eBARS_Close, Bars.Size - 1) = iEnd
        
    Set TempBars = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.ResetMinuteBars"

End Sub

Public Sub ResetTradeBars(ByVal strPeriod$)
On Error GoTo ErrSection:

    If InStr(strPeriod, "m") <> 0 Then
        ResetMinuteBars m.TradeSpeedBars, strPeriod
    ElseIf InStr(strPeriod, "v") <> 0 Then
        ResetVolumeBars m.TradeSpeedBars, strPeriod
    ElseIf InStr(strPeriod, "t") <> 0 Then
        ResetCustomTickBars m.TradeSpeedBars, strPeriod
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.ResetTradeBars"

End Sub

Public Sub ResetVolBars(ByVal strPeriod$)
On Error GoTo ErrSection:
            
    If InStr(strPeriod, "m") <> 0 Then
        ResetMinuteBars m.BidAskVolBars, strPeriod
    ElseIf InStr(strPeriod, "v") <> 0 Then
        ResetVolumeBars m.BidAskVolBars, strPeriod
    ElseIf InStr(strPeriod, "t") <> 0 Then
        ResetCustomTickBars m.BidAskVolBars, strPeriod
    End If
    
    m.BidAskVolBars.Prop(eBARS_StartTime) = m.TickBars.Prop(eBARS_StartTime)
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.ResetVolBars"

End Sub

Public Sub ResetTickBars(ByVal strSym$, ByVal nSymID&, frm As frmTimeSalesAnalyzer, _
    Optional ByVal nTimeStart& = 0)         'this is minutes from midnight
On Error GoTo ErrSection:
    
    Dim iBarStart&, i&
    Dim dTimeStart#
    Dim bHasVolume As Boolean
    
    m.TickBars.Size = 0
    GetAvailTickData m.TickBars, m.nSessionDate, strSym, nSymID, 0, 0
    
    'try one session date back so data for ES1 can be viewed after market close
    If m.TickBars.Size = 0 Then
        m.nSessionDate = m.nSessionDate - 1
        If Not IsWeekday(m.nSessionDate) Then m.nSessionDate = m.nSessionDate - 1
        If Not IsWeekday(m.nSessionDate) Then m.nSessionDate = m.nSessionDate - 1
        GetAvailTickData m.TickBars, m.nSessionDate, strSym, nSymID, m.nSessionDate, 0
    End If
    
    If m.TickBars.Size = 0 Then Exit Sub
    
    If m.TickBars.Size > 0 Then
        If nTimeStart > 0 Then
            dTimeStart = m.nSessionDate + nTimeStart / 1440     '570 minutes from midnight = 09:30AM
            gdBinarySearch m.TickBars.ArrayHandle(eBARS_DateTime), dTimeStart, iBarStart, eGdSort_Default, 0, m.TickBars.Size
            m.TickBars.DeleteFirstBars (iBarStart)
            If DateFormat(m.TickBars(eBARS_DateTime, 0), NO_DATE, HH_MM) = DateFormat(dTimeStart, NO_DATE, HH_MM) Then
                m.nStartTime = nTimeStart
            Else
                m.TickBars.Size = 0         'precautionary check because binary search is not based on exact match
                m.nStartTime = -1
            End If
        End If
    End If

    m.dMinMove = m.TickBars.MinMove(m.nSessionDate)
    Set m.frmAnalyzer = frm
    
    'aardvark 4089: make sure bars has volume data
    For i = 0 To m.TickBars.Size - 1
        If m.TickBars(eBARS_Vol, i) <> kNullData Then
            bHasVolume = True
            Exit For
        End If
    Next
    
    If Not bHasVolume Then m.TickBars.Size = 0
    
    'test code (begin)
'    Dim aTest As New cGdTable
'    Dim aTestArray As New cGdArray
'    Dim aIndex As cGdArray
'    Dim j As Long
'
'    aTest.CreateField eGDARRAY_Doubles
'    aTest.CreateField eGDARRAY_Doubles
'    aTest.CreateField eGDARRAY_Longs
'    aTest.CreateField eGDARRAY_Longs
'    For i = 0 To m.TickBars.Size - 1
'        aTest.AddRecord ""
'        aTest(0, i) = m.TickBars(eBARS_DateTime, i)
'        aTest(1, i) = RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)
'        aTest(2, i) = m.TickBars(eBARS_Vol, i)
'        aTest(3, i) = m.TickBars(eBARS_Flags, i)
''        strTest = DateFormat(m.TickBars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM_SS) & vbTab & _
''        Str(RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)) & vbTab & Str(m.TickBars(eBARS_Vol, i)) _
''        & vbTab & Str(m.TickBars(eBARS_Flags, i))
''        aTest.Add strTest
'    Next
'    Set aIndex = aTest.CreateIndex
'    aTest.SortIndex aIndex, 3, eGdSort_Stable
'    If Not aIndex Is Nothing Then
'        For i = 0 To aIndex.Size - 1
'            j = aIndex(i)
'            aTestArray.Add DateFormat(aTest(0, j), MM_DD_YYYY, HH_MM_SS) & vbTab & _
'            aTest(1, j) & vbTab & aTest(2, j) & vbTab & aTest(3, j)
'        Next
'        aTestArray.ToFile "YM_089.txt"
'    End If
'
'    Set aTest = Nothing
'    Set aTestArray = Nothing
'    Set aIndex = Nothing
    'test code (end)
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.ResetTickBars"

End Sub

Private Sub CalcHistoryStats(ByVal nDays&)
On Error GoTo ErrSection:
    
    Dim i&, j&, iSize&
    Dim nFirstDate&, nLastDate&, nSymID&
    
    Dim Stats As gdArrayStatistics
    Dim Bars As cGdBars
    
    Dim tbRange As New cGdTable
    Dim tbVol As New cGdTable
    
    Dim aRangeEOD As New cGdArray
    Dim aVolEOD As New cGdArray
    
    Dim strFile$           'using hard-coded text file from Rasa for now
    Dim strBaseSym$, strTemp$
                    
    For i = 0 To 13
        tbRange.CreateField eGDARRAY_Doubles, , , 0
        tbVol.CreateField eGDARRAY_Doubles, , , 0
    Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'range for current session
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Bars = New cGdBars
    DM_GetBars Bars, m.TickBars.Prop(eBARS_SymbolID), "30m", m.nSessionDate, m.nSessionDate
    For i = 0 To Bars.Size - 1
        m.tbStats(i, eTSVTb_Row_Range) = RoundNum(Bars(eBARS_High, i) - Bars(eBARS_Low, i), 2)
    Next

    If m.TickBars.SecurityType = "F" Then
        nSymID = GetSymbolID(m.TickBars.Prop(eBARS_BaseSymbol) & "-067")
    Else
        nSymID = m.TickBars.Prop(eBARS_SymbolID)
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'range & vol for daily bars (shown in total column)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    nLastDate = m.nSessionDate
    Set Bars = New cGdBars
    DM_GetBars Bars, nSymID, "Daily", nLastDate - nDays, nLastDate - 1
    
    If Bars.Size > 200 Then Bars.DeleteFirstBars (Bars.Size - 200)
    
    For i = 0 To Bars.Size - 1
        aVolEOD.Add Bars(eBARS_Vol, i)
        aRangeEOD.Add Bars(eBARS_High, i) - Bars(eBARS_Low, i)
    Next
    
    If gdCalcStatistics(aRangeEOD.ArrayHandle, Stats, False, 0, -1) Then
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_AvgRange) = RoundNum(Stats.Average, 2)
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_SHRange) = RoundNum(Stats.Average + Stats.StdDev, 2)
    End If
    If gdCalcStatistics(aVolEOD.ArrayHandle, Stats, False, 0, -1) Then
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_AvgVol) = RoundNum(Stats.Average, 0)
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_SHVol) = RoundNum(Stats.Average + Stats.StdDev, 0)
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'range & vol for each 30min colum A-N
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Bars = New cGdBars
    
    nFirstDate = nLastDate - 1
    
    'make sure 30min bar aligns with requested start time (eg 09:30)
    Dim dTimeStart, iBarStart&
    
    For i = 1 To nDays
        Bars.Size = 0
        iBarStart = -1
        
        nFirstDate = nLastDate - 1
        DM_GetBars Bars, nSymID, "30m", nFirstDate, nFirstDate
        
        If Bars.Size > 14 Then      'have more than A thru N
            If m.nStartTime > 0 Then
                dTimeStart = nFirstDate + m.nStartTime / 1440
            Else
                dTimeStart = nFirstDate + 570 / 1440            '570 minutes from midnight = 09:30AM
            End If
            For j = 0 To Bars.Size - 1
                If Bars(eBARS_DateTime, j) = dTimeStart Then
                    If j + 13 <= Bars.Size Then iBarStart = j
                    Exit For
                End If
            Next
        ElseIf Bars.Size = 14 Then
            iBarStart = 0
        End If
        
        If iBarStart <> -1 Then   'we have all bars from A thru N
            tbRange.AddRecord ""
            tbVol.AddRecord ""
            iSize = tbRange.NumRecords - 1
            For j = 0 To 13
                tbRange(j, iSize) = Bars(eBARS_High, j + iBarStart) - Bars(eBARS_Low, j + iBarStart)
                tbVol(j, iSize) = Bars(eBARS_Vol, j + iBarStart)
            Next
        End If
        nLastDate = nFirstDate
    Next
    
    For i = 0 To 13
        If gdCalcStatistics(tbRange.FieldArrayHandle(i), Stats, False, 0, -1) Then
            m.tbStats(i, eTSVTb_Row_AvgRange) = RoundNum(Stats.Average, 2)
            m.tbStats(i, eTSVTb_Row_SHRange) = RoundNum(Stats.Average + Stats.StdDev, 2)
        End If
        If gdCalcStatistics(tbVol.FieldArrayHandle(i), Stats, False, 0, -1) Then
            m.tbStats(i, eTSVTb_Row_AvgVol) = RoundNum(Stats.Average, 0)
            m.tbStats(i, eTSVTb_Row_SHVol) = RoundNum(Stats.Average + Stats.StdDev, 0)
        End If
    Next
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'read in statistically high vol at price for text file for coloring blue box
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    m.aStatHighVal.Size = 0
    strFile = g.strAppPath & kVolAtPriceFile
    If FileExist(strFile) Then
        Dim aTemp As New cGdArray
        aTemp.FromFile strFile
        strBaseSym = m.TickBars.Prop(eBARS_BaseSymbol)
        
        strTemp = ConvertSynthetic(strBaseSym, False)
        If Len(strTemp) > 0 Then strBaseSym = strTemp
        
        For i = aTemp.Size - 1 To 0 Step -1
            If Parse(aTemp(i), vbTab, 1) = strBaseSym Then
                m.aStatHighVal.Add aTemp(i)
                Exit For
            End If
        Next
    End If
    
    Set Bars = Nothing
    Set tbRange = Nothing
    Set tbVol = Nothing
    Set aRangeEOD = Nothing
    Set aVolEOD = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.CalcHistoryStats"

End Sub

Private Sub UpdateCustomTickBarsRT(Bars As cGdBars)
On Error GoTo ErrSection:

    Dim iStart&, iTicks&
    Dim dTicks#, dVol#
    Dim i&, j&
    
    iStart = Bars(eBARS_Close, Bars.Size - 1) + 1   'index of last full ticks bar processed
    If iStart = kNullData Then Exit Sub             'something fubar here, just exit

    iTicks = Bars.Prop(eBARS_PeriodsPerBar)
    dTicks = Bars(eBARS_Open, Bars.Size - 1)

    j = Bars.Size - 1

    For i = iStart To m.TickBars.Size - 1
        If dTicks < iTicks Then
            dTicks = dTicks + 1
            dVol = m.TickBars(eBARS_Vol, i)
            Bars(eBARS_Vol, j) = Bars(eBARS_Vol, j) + dVol
            Bars(eBARS_UpTicks, j) = dTicks
            If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                Bars(eBARS_BidVol, j) = Bars(eBARS_BidVol, j) + dVol
                Bars(eBARS_Bid, j) = Bars(eBARS_Bid, j) + 1
                Bars(eBARS_UpTicks, j) = Bars(eBARS_UpTicks, j) + 1     'see note in ResetMinuteBars for this
            ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                Bars(eBARS_AskVol, j) = Bars(eBARS_AskVol, j) + dVol
                Bars(eBARS_Ask, j) = Bars(eBARS_Ask, j) + 1
                Bars(eBARS_DownTicks, j) = Bars(eBARS_DownTicks, j) + 1
            End If
        Else
            dTicks = 0
            Bars.Size = Bars.Size + 1
            j = Bars.Size - 1
        
            Bars(eBARS_Vol, j) = 0
            Bars(eBARS_BidVol, j) = 0
            Bars(eBARS_AskVol, j) = 0
            Bars(eBARS_Bid, j) = 0
            Bars(eBARS_Ask, j) = 0
            Bars(eBARS_UpTicks, j) = 0
            Bars(eBARS_DownTicks, j) = 0
            
            Bars(eBARS_DateTime, j) = m.TickBars(eBARS_DateTime, i)
            i = i - 1
        End If
    Next
    
    'store index of last full ticks bar that was processed
    Bars(eBARS_Close, Bars.Size - 1) = m.TickBars.Size - 1
    'store total number of trades included in this bar
    Bars(eBARS_Open, Bars.Size - 1) = dTicks
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.UpdateCustomTickBarsRT"

End Sub

Private Sub UpdateVolumeBarsRT(Bars As cGdBars)
On Error GoTo ErrSection:
    
    Dim dVol#, dVolSum#
    Dim iStart&, iVol&
    Dim i&, j&

    iStart = Bars(eBARS_Close, Bars.Size - 1) + 1   'index of last full ticks bar processed
    If iStart = kNullData Then Exit Sub             'something fubar here, just exit

    iVol = Bars.Prop(eBARS_PeriodsPerBar)

    j = Bars.Size - 1
    dVolSum = Bars(eBARS_Vol, j)
    
    For i = iStart To m.TickBars.Size - 1
        If dVolSum < iVol Then
            dVol = m.TickBars(eBARS_Vol, i)
            dVolSum = dVolSum + dVol
            Bars(eBARS_Vol, j) = Bars(eBARS_Vol, j) + dVol
            If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                Bars(eBARS_BidVol, j) = Bars(eBARS_BidVol, j) + dVol
                Bars(eBARS_Bid, j) = Bars(eBARS_Bid, j) + 1
                Bars(eBARS_UpTicks, j) = Bars(eBARS_UpTicks, j) + 1     'see note in ResetMinuteBars for this
            ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                Bars(eBARS_AskVol, j) = Bars(eBARS_AskVol, j) + dVol
                Bars(eBARS_Ask, j) = Bars(eBARS_Ask, j) + 1
                Bars(eBARS_DownTicks, j) = Bars(eBARS_DownTicks, j) + 1
            End If
        Else
            dVolSum = 0
            Bars.Size = Bars.Size + 1
            j = Bars.Size - 1
        
            Bars(eBARS_Vol, j) = 0
            Bars(eBARS_BidVol, j) = 0
            Bars(eBARS_AskVol, j) = 0
            Bars(eBARS_Bid, j) = 0
            Bars(eBARS_Ask, j) = 0
            Bars(eBARS_UpTicks, j) = 0
            Bars(eBARS_DownTicks, j) = 0
            
            Bars(eBARS_DateTime, j) = m.TickBars(eBARS_DateTime, i)
            i = i - 1
        End If
    Next
    
    'store index of last full ticks bar that was processed
    Bars(eBARS_Close, Bars.Size - 1) = m.TickBars.Size - 1
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.UpdateVolumeBarsRT"

End Sub

Private Sub UpdateMinuteBarsRT(Bars As cGdBars)
On Error GoTo ErrSection:
    
    Dim dDate1#, dDate2#
    Dim iStart&, iEnd&
    Dim iMinute&, i&, j&
            
    iStart = Bars(eBARS_Close, Bars.Size - 1) + 1   'index of last full ticks bar processed
    If iStart = kNullData Then Exit Sub             'something fubar here, just exit
    iEnd = m.TickBars.Size - 1
    
    'these 2 dates should be start/end times shown in last row of time analyzer grids
    dDate1 = Bars(eBARS_DateTime, Bars.Size - 2)
    dDate2 = Bars(eBARS_DateTime, Bars.Size - 1)
    j = Bars.Size       'save current size of bars to update
    
    For i = iStart To iEnd
        If m.TickBars(eBARS_DateTime, i) > Bars(eBARS_DateTime, Bars.Size - 1) Then
            Bars.Size = Bars.Size + 1
            Bars(eBARS_Vol, Bars.Size - 1) = 0
            Bars(eBARS_BidVol, Bars.Size - 1) = 0
            Bars(eBARS_AskVol, Bars.Size - 1) = 0
            Bars(eBARS_Bid, Bars.Size - 1) = 0
            Bars(eBARS_Ask, Bars.Size - 1) = 0
            Bars(eBARS_UpTicks, Bars.Size - 1) = 0          'see note in ResetMinuteBars for this
            Bars(eBARS_DownTicks, Bars.Size - 1) = 0
            Bars(eBARS_DateTime, Bars.Size - 1) = Bars(eBARS_DateTime, Bars.Size - 1) + 60 / 1440 / 60 * Bars.Prop(eBARS_PeriodsPerBar)
        End If
    Next
        
    For i = j To Bars.Size
        dDate1 = Bars(eBARS_DateTime, j - 2)
        dDate2 = Bars(eBARS_DateTime, j - 1)
        
        TimeRangeBarData Bars, dDate1, dDate2, iStart, iEnd, j - 1
        'Bars(eBARS_DateTime, j - 1) = dDate2
    Next

    'store index of last full ticks bar that was processed
    Bars(eBARS_Close, Bars.Size - 1) = m.TickBars.Size - 1
    
    Exit Sub

ErrSection:
    RaiseError "cTSVData.UpdateMinuteBarsRT"

End Sub

Public Function UpdateDataRT(bNewBarVol As Boolean, bNewBarTrade As Boolean, _
    bTableRebuild As Boolean) As Long
On Error GoTo ErrSection:

    Dim nOldSize&, nSymID&, strSym$
    Dim bNewDay As Boolean
    
    If nOldSize = 0 And m.TickBars.Size = 0 Then
        UpdateDataRT = -1
    Else
        'get updated tick bars
        nOldSize = m.TickBars.Size
        If g.RealTime.UpdateBars(m.TickBars, bNewDay) Then
            If bNewDay Then
                nSymID = m.TickBars.Prop(eBARS_SymbolID)
                strSym = m.TickBars.Prop(eBARS_Symbol)
                UpdateDataRT = -1           'things are faster with Salmon (partially related to issue 5465)
            ElseIf m.TickBars.Size > nOldSize Then
                If m.eTbType = eTSVTb_TbType_VolAtPrice Then
                    bTableRebuild = UpdatePriceVolTbRT(nOldSize)
                ElseIf m.eTbType = eTSVTb_TbType_Reversal Then
                    bTableRebuild = UpdateReverseTbRT(nOldSize)
                End If
                UpdateDataRT = m.TickBars.Size - nOldSize
            End If
            If Not m.BidAskVolBars Is Nothing Then
                If InStr(m.BidAskVolBars.Prop(eBARS_PeriodicityStr), "minute") <> 0 Then
                    UpdateMinuteBarsRT m.BidAskVolBars
                ElseIf InStr(m.BidAskVolBars.Prop(eBARS_PeriodicityStr), "vol") <> 0 Then
                    UpdateVolumeBarsRT m.BidAskVolBars
                ElseIf InStr(m.BidAskVolBars.Prop(eBARS_PeriodicityStr), "ticks") <> 0 Then
                    UpdateCustomTickBarsRT m.BidAskVolBars
                End If
            End If
            If Not m.TradeSpeedBars Is Nothing Then
                If InStr(m.TradeSpeedBars.Prop(eBARS_PeriodicityStr), "minute") <> 0 Then
                    UpdateMinuteBarsRT m.TradeSpeedBars
                ElseIf InStr(m.TradeSpeedBars.Prop(eBARS_PeriodicityStr), "vol") <> 0 Then
                    UpdateVolumeBarsRT m.TradeSpeedBars
                ElseIf InStr(m.TradeSpeedBars.Prop(eBARS_PeriodicityStr), "ticks") <> 0 Then
                    UpdateCustomTickBarsRT m.TradeSpeedBars
                End If
            End If
        End If
    End If
    
    Exit Function

ErrSection:
    RaiseError "cTSVData.UpdateDataRT"

End Function

Public Property Get TbColStats(ByVal eCol As eTSV_TbFields) As eTSV_TbStatsFields

    TbColStats = -1

    Select Case eCol
        Case eTSVTb_VolAsk_A, eTSVTb_VolBid_A
            TbColStats = eTSVTb_Col_A
        Case eTSVTb_VolAsk_B, eTSVTb_VolBid_B
            TbColStats = eTSVTb_Col_B
        Case eTSVTb_VolAsk_C, eTSVTb_VolBid_C
            TbColStats = eTSVTb_Col_C
        Case eTSVTb_VolAsk_D, eTSVTb_VolBid_D
            TbColStats = eTSVTb_Col_D
        Case eTSVTb_VolAsk_E, eTSVTb_VolBid_E
            TbColStats = eTSVTb_Col_E
        Case eTSVTb_VolAsk_F, eTSVTb_VolBid_F
            TbColStats = eTSVTb_Col_F
        Case eTSVTb_VolAsk_G, eTSVTb_VolBid_G
            TbColStats = eTSVTb_Col_G
        Case eTSVTb_VolAsk_H, eTSVTb_VolBid_H
            TbColStats = eTSVTb_Col_H
        Case eTSVTb_VolAsk_I, eTSVTb_VolBid_I
            TbColStats = eTSVTb_Col_I
        Case eTSVTb_VolAsk_J, eTSVTb_VolBid_J
            TbColStats = eTSVTb_Col_J
        Case eTSVTb_VolAsk_K, eTSVTb_VolBid_K
            TbColStats = eTSVTb_Col_K
        Case eTSVTb_VolAsk_L, eTSVTb_VolBid_L
            TbColStats = eTSVTb_Col_L
        Case eTSVTb_VolAsk_M, eTSVTb_VolBid_M
            TbColStats = eTSVTb_Col_M
        Case eTSVTb_VolAsk_N, eTSVTb_VolBid_N
            TbColStats = eTSVTb_Col_N
        Case eTSVTb_VolRow_Total
            TbColStats = eTSVTb_Col_Total
        Case Else
            TbColStats = -1
    End Select

End Property

Public Property Get TbColVolBid(ByVal eCol As eTSV_TbStatsFields) As eTSV_TbFields

    Select Case eCol
        Case eTSVTb_Col_A
            TbColVolBid = eTSVTb_VolBid_A
        Case eTSVTb_Col_B
            TbColVolBid = eTSVTb_VolBid_B
        Case eTSVTb_Col_C
            TbColVolBid = eTSVTb_VolBid_C
        Case eTSVTb_Col_D
            TbColVolBid = eTSVTb_VolBid_D
        Case eTSVTb_Col_E
            TbColVolBid = eTSVTb_VolBid_E
        Case eTSVTb_Col_F
            TbColVolBid = eTSVTb_VolBid_F
        Case eTSVTb_Col_G
            TbColVolBid = eTSVTb_VolBid_G
        Case eTSVTb_Col_H
            TbColVolBid = eTSVTb_VolBid_H
        Case eTSVTb_Col_I
            TbColVolBid = eTSVTb_VolBid_I
        Case eTSVTb_Col_J
            TbColVolBid = eTSVTb_VolBid_J
        Case eTSVTb_Col_K
            TbColVolBid = eTSVTb_VolBid_K
        Case eTSVTb_Col_L
            TbColVolBid = eTSVTb_VolBid_L
        Case eTSVTb_Col_M
            TbColVolBid = eTSVTb_VolBid_M
        Case eTSVTb_Col_N
            TbColVolBid = eTSVTb_VolBid_N
        Case eTSVTb_Col_Total
            TbColVolBid = eTSVTb_VolRow_Total
        Case Else
            TbColVolBid = -1
    End Select

End Property

Public Property Get TbColVolOther(Optional ByVal iCol& = -1) As eTSV_TbFields

    Dim eColBid As eTSV_TbFields

    If iCol >= 0 Then
        eColBid = iCol
    Else
        eColBid = m.nTbLastCol
    End If
        
    Select Case eColBid
        Case eTSVTb_VolBid_A
            TbColVolOther = eTSVTb_VolOther_A
        Case eTSVTb_VolBid_B
            TbColVolOther = eTSVTb_VolOther_B
        Case eTSVTb_VolBid_C
            TbColVolOther = eTSVTb_VolOther_C
        Case eTSVTb_VolBid_D
            TbColVolOther = eTSVTb_VolOther_D
        Case eTSVTb_VolBid_E
            TbColVolOther = eTSVTb_VolOther_E
        Case eTSVTb_VolBid_F
            TbColVolOther = eTSVTb_VolOther_F
        Case eTSVTb_VolBid_G
            TbColVolOther = eTSVTb_VolOther_G
        Case eTSVTb_VolBid_H
            TbColVolOther = eTSVTb_VolOther_H
        Case eTSVTb_VolBid_I
            TbColVolOther = eTSVTb_VolOther_I
        Case eTSVTb_VolBid_J
            TbColVolOther = eTSVTb_VolOther_J
        Case eTSVTb_VolBid_K
            TbColVolOther = eTSVTb_VolOther_K
        Case eTSVTb_VolBid_L
            TbColVolOther = eTSVTb_VolOther_L
        Case eTSVTb_VolBid_M
            TbColVolOther = eTSVTb_VolOther_M
        Case eTSVTb_VolBid_N
            TbColVolOther = eTSVTb_VolOther_N
        Case Else
            TbColVolOther = -1
    End Select

End Property

'returns true if table was rebuilt
Private Function UpdateReverseTbRT(ByVal iOldSize&) As Boolean

    Dim dMinPrice#, dMaxPrice#, dReverseAmt#, i&
        
    dMinPrice = gdMinValue(m.TickBars.ArrayHandle(eBARS_Close), iOldSize, m.TickBars.Size)
    dMaxPrice = gdMaxValue(m.TickBars.ArrayHandle(eBARS_Close), iOldSize, m.TickBars.Size)
    
    If m.tbPriceVol.NumRecords = 0 Or _
       dMinPrice < m.tbPriceVol(eTSVTb_Price, m.tbPriceVol.NumRecords - 1) Or _
       dMaxPrice > m.tbPriceVol(eTSVTb_Price, 2) Then
       BuildTables eTSVTb_TbType_Reversal
       UpdateReverseTbRT = True
       Exit Function
    End If
    
    dReverseAmt = m.iReverseTick * m.dMinMove
        
    For i = iOldSize To m.TickBars.Size - 1
        SetReverseTbData i, dReverseAmt, True
    Next
    
    CalcPullbackStats
    m.aSpeedCells.Sort eGdSort_Default Or eGdSort_DeleteDuplicates
    
End Function

'returns true if table was rebuilt
Private Function UpdatePriceVolTbRT(ByVal iOldSize&) As Boolean

    Dim dMinPrice#, dMaxPrice#, dPrice#
    Dim dTimeStart#, dTimeEnd#, dBarTime#, dLastBarTime#
    Dim dVol#, dVolPrev#, dColTotalHiVol#
    Dim i&, j&, iCol&, iPrevStatsCol&
    
    Dim eOtherVolCol As eTSV_TbFields
    Dim bAdd As Boolean
    
    If m.tbPriceVol Is Nothing Then Exit Function

    dMinPrice = gdMinValue(m.TickBars.ArrayHandle(eBARS_Close), iOldSize, m.TickBars.Size)
    dMaxPrice = gdMaxValue(m.TickBars.ArrayHandle(eBARS_Close), iOldSize, m.TickBars.Size)
    
    If m.tbPriceVol.NumRecords = 0 Or _
       dMinPrice < m.tbPriceVol(eTSVTb_Price, m.tbPriceVol.NumRecords - 1) Or _
       dMaxPrice > m.tbPriceVol(eTSVTb_Price, 0) Then
       BuildTables eTSVTb_TbType_VolAtPrice
       UpdatePriceVolTbRT = True
       Exit Function
    End If
    
    If m.aTbIdx Is Nothing Then             'precautionary - should never happen
        Set m.aTbIdx = m.tbPriceVol.CreateSortedIndex(eTSVTb_Price)
    End If
    
    dTimeStart = m.dTbColTimeStart
    dTimeEnd = m.dTbColTimeEnd
    dLastBarTime = m.TickBars(eBARS_DateTime, m.TickBars.Size - 1)
    
    iCol = m.nTbLastCol
    iPrevStatsCol = m.nTbStatsLastCol
    eOtherVolCol = TbColVolOther
    dColTotalHiVol = m.tbPriceVol(eTSVTb_VolRow_Total, m.iTotalColVolHi)
    
    For i = iOldSize To m.TickBars.Size - 1
        bAdd = False
        dBarTime = m.TickBars(eBARS_DateTime, i)
        If Hour(dBarTime) < Hour(dTimeEnd) Then
            bAdd = True
        ElseIf Hour(dBarTime) = Hour(dTimeEnd) Then
            If Minute(dBarTime) < Minute(dTimeEnd) Then
                bAdd = True
            End If
        ElseIf Int(dBarTime) < Int(dTimeEnd) Then
            'data crossing midnight into next day
            If Hour(dBarTime) = 23 And Hour(dTimeEnd) = 0 Then bAdd = True
        End If
                   
        If bAdd Then
            dPrice = RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)
            dVol = m.TickBars(eBARS_Vol, i)
            If (m.tbPriceVol.SearchAsIndex(m.aTbIdx, eTSVTb_Price, dPrice, j)) Then
                If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                    dVolPrev = m.tbPriceVol(iCol, m.aTbIdx(j))
                    m.tbPriceVol(iCol, m.aTbIdx(j)) = m.tbPriceVol(iCol, m.aTbIdx(j)) + dVol
                    m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_Bid) = m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_Bid) + dVol
                ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                    dVolPrev = m.tbPriceVol(iCol + 1, m.aTbIdx(j))
                    m.tbPriceVol(iCol + 1, m.aTbIdx(j)) = m.tbPriceVol(iCol + 1, m.aTbIdx(j)) + dVol
                    m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_Ask) = m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_Ask) + dVol
                Else
                    dVolPrev = m.tbPriceVol(eOtherVolCol, m.aTbIdx(j))
                    m.tbPriceVol(eOtherVolCol, m.aTbIdx(j)) = m.tbPriceVol(eOtherVolCol, m.aTbIdx(j)) + dVol
                    m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_OtherVol) = m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_OtherVol) + dVol
                End If
                
                m.tbPriceVol(eTSVTb_VolRow_Total, m.aTbIdx(j)) = m.tbPriceVol(eTSVTb_VolRow_Total, m.aTbIdx(j)) + dVol
                m.nTimeSalesVolTotal = m.nTimeSalesVolTotal + dVol
                
                'for auction bar view
                dVol = m.tbPriceVol(eTSVTb_VolRow_Total, m.aTbIdx(j))
                If dVol > dColTotalHiVol Then
                    m.iTotalColVolHi = m.aTbIdx(j)
                End If
                
                m.nTbLastCol = iCol
                m.dTbColTimeStart = dTimeStart
                m.dTbColTimeEnd = dTimeEnd
            End If
        ElseIf Int(dTimeEnd) > Int(dLastBarTime) Then
            Exit For            'something is wrong here
        Else
            iCol = iCol + 2     'set to vol1 of next letter, i.e. Vol1_B, Vol1_C, Vol1_D etc.
            dTimeStart = dTimeEnd
            dTimeEnd = dTimeEnd + 30 / 1440
            i = i - 1
        End If
    Next
    
    Select Case m.nTbLastCol
        Case eTSVTb_VolBid_A
            m.nTbStatsLastCol = eTSVTb_Col_A
        Case eTSVTb_VolBid_B
            m.nTbStatsLastCol = eTSVTb_Col_B
        Case eTSVTb_VolBid_C
            m.nTbStatsLastCol = eTSVTb_Col_C
        Case eTSVTb_VolBid_D
            m.nTbStatsLastCol = eTSVTb_Col_D
        Case eTSVTb_VolBid_E
            m.nTbStatsLastCol = eTSVTb_Col_E
        Case eTSVTb_VolBid_F
            m.nTbStatsLastCol = eTSVTb_Col_F
        Case eTSVTb_VolBid_G
            m.nTbStatsLastCol = eTSVTb_Col_G
        Case eTSVTb_VolBid_H
            m.nTbStatsLastCol = eTSVTb_Col_H
        Case eTSVTb_VolBid_I
            m.nTbStatsLastCol = eTSVTb_Col_I
        Case eTSVTb_VolBid_J
            m.nTbStatsLastCol = eTSVTb_Col_J
        Case eTSVTb_VolBid_K
            m.nTbStatsLastCol = eTSVTb_Col_K
        Case eTSVTb_VolBid_L
            m.nTbStatsLastCol = eTSVTb_Col_L
        Case eTSVTb_VolBid_M
            m.nTbStatsLastCol = eTSVTb_Col_M
        Case eTSVTb_VolBid_N
            m.nTbStatsLastCol = eTSVTb_Col_N
    End Select
    
    CalcStatsTable eTSVTb_Col_A, m.nTbStatsLastCol
    CalcStatsRT iPrevStatsCol, m.aTbIdx(j)

End Function

Private Sub CalcStatsRT(ByVal iPrevStatsCol&, ByVal iClose&)
    
    Dim dPxVx#, dVx#, dVol#, dVolHi#
    Dim dPrice#, dHigh#, dLow#
    Dim dVAM#, dVBM#
    
    Dim i&, j&, k&, iMeanIdx&
    Dim iHigh&, iLow&, iHighestVol&            'for per column auction bar view
    
    Dim aPrices As New cGdArray
    
    Dim eOtherVolCol As eTSV_TbFields
           
    dHigh = -999999
    dLow = 999999
    
    eOtherVolCol = TbColVolOther
    
    If iPrevStatsCol = m.nTbStatsLastCol Then
        'okay to continue
    ElseIf m.nTbStatsLastCol > iPrevStatsCol Then
         j = -1
         If iPrevStatsCol = eTSVTb_Col_A Then
             j = eTSVTb_Col_B
             m.ePxVxCol = eTSVTb_VolBid_A
         ElseIf iPrevStatsCol = eTSVTb_Col_B Then
             j = eTSVTb_Col_C
             m.ePxVxCol = eTSVTb_VolBid_B
         ElseIf iPrevStatsCol = eTSVTb_Col_C Then
             j = eTSVTb_Col_D
             m.ePxVxCol = eTSVTb_VolBid_C
         ElseIf iPrevStatsCol = eTSVTb_Col_D Then
             j = eTSVTb_Col_E
             m.ePxVxCol = eTSVTb_VolBid_D
         ElseIf iPrevStatsCol = eTSVTb_Col_E Then
             j = eTSVTb_Col_F
             m.ePxVxCol = eTSVTb_VolBid_E
         ElseIf iPrevStatsCol = eTSVTb_Col_F Then
             j = eTSVTb_Col_G
             m.ePxVxCol = eTSVTb_VolBid_F
         ElseIf iPrevStatsCol = eTSVTb_Col_G Then
             j = eTSVTb_Col_H
             m.ePxVxCol = eTSVTb_VolBid_G
         ElseIf iPrevStatsCol = eTSVTb_Col_H Then
             j = eTSVTb_Col_I
             m.ePxVxCol = eTSVTb_VolBid_H
         ElseIf iPrevStatsCol = eTSVTb_Col_I Then
             j = eTSVTb_Col_J
             m.ePxVxCol = eTSVTb_VolBid_I
         ElseIf iPrevStatsCol = eTSVTb_Col_J Then
             j = eTSVTb_Col_K
             m.ePxVxCol = eTSVTb_VolBid_J
         ElseIf iPrevStatsCol = eTSVTb_Col_K Then
             j = eTSVTb_Col_L
             m.ePxVxCol = eTSVTb_VolBid_K
         ElseIf iPrevStatsCol = eTSVTb_Col_L Then
             j = eTSVTb_Col_M
             m.ePxVxCol = eTSVTb_VolBid_L
         ElseIf iPrevStatsCol = eTSVTb_Col_M Then
             j = eTSVTb_Col_N
             m.ePxVxCol = eTSVTb_VolBid_M
         ElseIf iPrevStatsCol = eTSVTb_Col_N Then
             j = -1             'theoretically should not get here
         End If
                 
        If j = m.nTbStatsLastCol Then
            eOtherVolCol = TbColVolOther(m.ePxVxCol)
            For i = 0 To m.tbPriceVol.NumRecords - 1
                If m.tbPriceVol(m.ePxVxCol, i) > 0 Or m.tbPriceVol(m.ePxVxCol + 1, i) > 0 Or m.tbPriceVol(eOtherVolCol, i) > 0 Then
                    dVol = m.tbPriceVol(m.ePxVxCol, i) + m.tbPriceVol(m.ePxVxCol + 1, i) + m.tbPriceVol(eOtherVolCol, i)
                    dPrice = m.tbPriceVol(0, i)
                    
                    dPxVx = dPxVx + dPrice * dVol
                    dVx = dVx + dVol
                End If
            Next
            m.dVx = m.dVx + dVx
            m.dPxVx = m.dPxVx + dPxVx
        Else
            Exit Sub            'theoretically should not get here
        End If
    Else
        Exit Sub               'theoretically should never get here)
    End If
    
    eOtherVolCol = TbColVolOther(m.nTbLastCol)
    For i = 0 To m.tbPriceVol.NumRecords - 1
        If m.tbPriceVol(m.nTbLastCol, i) > 0 Or m.tbPriceVol(m.nTbLastCol + 1, i) > 0 Or m.tbPriceVol(eOtherVolCol, i) > 0 Then
           dVol = m.tbPriceVol(m.nTbLastCol, i) + m.tbPriceVol(m.nTbLastCol + 1, i) + m.tbPriceVol(eOtherVolCol, i)
           dPrice = m.tbPriceVol(0, i)
           
           dPxVx = dPxVx + dPrice * dVol
           dVx = dVx + dVol
           If dVol > dVolHi Then
               iHighestVol = i
               dVolHi = dVol
           End If
           
           aPrices.Add i
           
           If dPrice > dHigh Then
               dHigh = dPrice
               iHigh = i
           End If
           If dPrice < dLow Then
               dLow = dPrice
               iLow = i
           End If
        End If
    Next
    'per column data for this column for auction bar view
    SetPerColIdx m.nTbLastCol, m.nTbStatsLastCol, dPxVx, dVx, iHigh, iLow, iClose, iHighestVol
    
    'range for current column
    If dHigh > 0 And dLow > 0 Then
        m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_Range) = RoundNum(dHigh - dLow, 2)
    End If
    If dVx > 0 Then
        dPxVx = dPxVx + m.dPxVx
        dVx = dVx + m.dVx
        dPrice = RoundToMinMove(dPxVx / dVx, m.dMinMove)
        m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_Mean) = dPrice
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Mean) = dPrice
        'calculate cumulative VAM & VBM for column
        If (m.tbPriceVol.SearchAsIndex(m.aTbIdx, eTSVTb_Price, dPrice, j)) Then
           iMeanIdx = m.aTbIdx(j)
           aPrices.Sort eGdSort_DeleteDuplicates
           For k = 0 To aPrices.Size - 1
           If aPrices(k) > iMeanIdx Then
               dVBM = dVBM + m.tbPriceVol(m.nTbLastCol, aPrices(k))
           ElseIf aPrices(k) < iMeanIdx Then
               dVAM = dVAM + m.tbPriceVol(m.nTbLastCol, aPrices(k))
           End If
           Next
           m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_VAM) = m.tbStats(m.nTbStatsLastCol - 1, eTSVTb_Row_VAM) + dVAM
           m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_VBM) = m.tbStats(m.nTbStatsLastCol - 1, eTSVTb_Row_VBM) + dVBM
           SetAboveBelowIdx eTSV_Group_Current, iMeanIdx, dVx
           'VAM, VMB & Mean of 'totals' column is same
           m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_VAM) = m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_VAM)
           m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_VBM) = m.tbStats(m.nTbStatsLastCol, eTSVTb_Row_VBM)
           If m.nTbStatsLastCol = eTSVTb_Col_B Then
               SetAboveBelowIdx eTSV_Group_AB, iMeanIdx, dVx
           ElseIf m.nTbStatsLastCol = eTSVTb_Col_E Then
               SetAboveBelowIdx eTSV_Group_AE, iMeanIdx, dVx
           ElseIf m.nTbStatsLastCol = eTSVTb_Col_I Then
               SetAboveBelowIdx eTSV_Group_AI, iMeanIdx, dVx
           ElseIf m.nTbStatsLastCol = eTSVTb_Col_N Then
               SetAboveBelowIdx eTSV_Group_AN, iMeanIdx, dVx
           End If
           'per column mean & value area for total column
           m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Hi) = 0
           m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Low) = m.tbPriceVol.NumRecords - 1
           m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Above) = m.tpColorRows.iAboveCurrent
           m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Mean) = m.tpColorRows.iMeanCurrent
           m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Below) = m.tpColorRows.iBelowCurrent
           m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_HiestVol) = m.iTotalColVolHi
        End If
    End If
    
End Sub

Private Sub CalcPullbackStats()

    Dim i&
    Dim dTotalBid#, dTotalAsk#, dOther#
    
    dTotalBid = 0
    dTotalAsk = 0
    dOther = 0
    For i = 0 To m.tbPBVol.NumRecords
        dTotalBid = dTotalBid + m.tbPBVol(0, i)
        dTotalAsk = dTotalAsk + m.tbPBVol(1, i)
        dOther = dOther + m.tbPBVol(2, i)
    Next
        
    m.tbPBStats(0, 0) = dTotalBid
    m.tbPBStats(1, 0) = dTotalAsk
    m.tbPBStats(2, 0) = dOther
        
    dOther = dTotalBid - dTotalAsk
    If dOther > 0 Then
        m.tbPBStats(0, 1) = dOther
        m.tbPBStats(1, 1) = 0
    ElseIf dOther < 0 Then
        m.tbPBStats(1, 1) = Abs(dOther)
        m.tbPBStats(0, 1) = 0
    End If

End Sub

Private Sub CalcRevCumulative(ByVal iCol&)

    Dim dTemp#
    
'cumulative
'[Cum Big] [Cum Big Diff] [Cum Big%] [Cum Big Diff% of Total] (rows 7 thru 10)
'[Cum Totals] [Cum Diff] [Cum%] [Cum Diff% of Total] (rows 11 thru 14)
    If iCol > 3 Then
        '[Cum Big]
        m.tbStats(iCol, 7) = m.tbStats(iCol - 3, 7) + m.tbStats(iCol, 4)
        m.tbStats(iCol + 1, 7) = m.tbStats(iCol - 2, 7) + m.tbStats(iCol + 1, 4)
        m.tbStats(iCol + 2, 7) = m.tbStats(iCol - 1, 7) + m.tbStats(iCol + 2, 4)
        '[Cum Totals]
        m.tbStats(iCol, 11) = m.tbStats(iCol - 3, 11) + m.tbStats(iCol, 0)
        m.tbStats(iCol + 1, 11) = m.tbStats(iCol - 2, 11) + m.tbStats(iCol + 1, 0)
        m.tbStats(iCol + 2, 11) = m.tbStats(iCol - 1, 11) + m.tbStats(iCol + 2, 0)
    ElseIf iCol = 1 Then
        '[Cum Big]
        m.tbStats(iCol, 7) = m.tbStats(iCol, 4)
        m.tbStats(iCol + 1, 7) = m.tbStats(iCol + 1, 4)
        m.tbStats(iCol + 2, 7) = m.tbStats(iCol + 2, 4)
        '[Cum Totals]
        m.tbStats(iCol, 11) = m.tbStats(iCol, 0)
        m.tbStats(iCol + 1, 11) = m.tbStats(iCol + 1, 0)
        m.tbStats(iCol + 2, 11) = m.tbStats(iCol + 2, 0)
    Else
        Exit Sub        'theoretically should not get here
    End If

    '[Cum Big Diff]
    dTemp = m.tbStats(iCol, 7) - m.tbStats(iCol + 1, 7)
    If dTemp > 0 Then
        m.tbStats(iCol, 8) = dTemp
        m.tbStats(iCol + 1, 8) = 0
    ElseIf dTemp < 0 Then
        m.tbStats(iCol + 1, 8) = Abs(dTemp)
        m.tbStats(iCol, 8) = 0
    End If
    '[Cum Big%]
    dTemp = m.tbStats(iCol, 7) + m.tbStats(iCol + 1, 7) + m.tbStats(iCol + 2, 7)
    If dTemp > 0 Then
        m.tbStats(iCol, 9) = RoundNum(m.tbStats(iCol, 7) / dTemp * 100)
        m.tbStats(iCol + 1, 9) = RoundNum(m.tbStats(iCol + 1, 7) / dTemp * 100)
    End If
    '[Cum Diff]
    dTemp = m.tbStats(iCol, 11) - m.tbStats(iCol + 1, 11)
    If dTemp > 0 Then
        m.tbStats(iCol, 12) = dTemp
        m.tbStats(iCol + 1, 12) = 0
    ElseIf dTemp < 0 Then
        m.tbStats(iCol + 1, 12) = Abs(dTemp)
        m.tbStats(iCol, 12) = 0
    End If
    
    'percentages of total
    dTemp = m.tbStats(iCol, 11) + m.tbStats(iCol + 1, 11)       '+ m.tbStats(iCol + 2, 11)
    If dTemp > 0 Then
        '[Cum Big Diff% of Total]
        m.tbStats(iCol, 10) = Abs(RoundNum(m.tbStats(iCol, 8) / dTemp * 100))
        m.tbStats(iCol + 1, 10) = Abs(RoundNum(m.tbStats(iCol + 1, 8) / dTemp * 100))
        '[Cum %]
        m.tbStats(iCol, 13) = RoundNum(m.tbStats(iCol, 11) / dTemp * 100)
        m.tbStats(iCol + 1, 13) = RoundNum(m.tbStats(iCol + 1, 11) / dTemp * 100)
        '[Cum Diff% of Total]
        m.tbStats(iCol, 14) = Abs(RoundNum(m.tbStats(iCol, 12) / dTemp * 100))
        m.tbStats(iCol + 1, 14) = Abs(RoundNum(m.tbStats(iCol + 1, 12) / dTemp * 100))
    End If

End Sub

Private Sub CalcRevRT(ByVal iCol&)

    Dim dColVolBid#, dColVolAsk#, dColVolOther#
    Dim dColBigLotBid#, dColBigLotAsk#, dColBigLotOther#
    
    Dim dTemp#, i&
           
    dColVolBid = m.tbStats(iCol, 0)
    dColVolAsk = m.tbStats(iCol + 1, 0)
    dColVolOther = m.tbStats(iCol + 2, 0)
    
    dColBigLotBid = m.tbStats(iCol, 4)
    dColBigLotAsk = m.tbStats(iCol + 1, 4)
    dColBigLotOther = m.tbStats(iCol + 2, 4)

    '[Differential]    bid-ask (positive goes in bid col, neg goes into ask col)
    dTemp = dColVolBid - dColVolAsk
    If dTemp > 0 Then
        m.tbStats(iCol, 1) = dTemp
        m.tbStats(iCol + 1, 1) = 0
    ElseIf dTemp < 0 Then
        m.tbStats(iCol + 1, 1) = Abs(dTemp)    'From Rasa: don't show negative numbers in grid (08-24-2006)
        m.tbStats(iCol, 1) = 0
    End If
    '[Total%]  bid/ask vol as % of total vol
    dTemp = dColVolBid + dColVolAsk
    If dTemp > 0 Then
        m.tbStats(iCol, 2) = RoundNum(dColVolBid / dTemp * 100)
        m.tbStats(iCol + 1, 2) = RoundNum(dColVolAsk / dTemp * 100)
    End If

    '[Big Lots %]
    dTemp = dColBigLotBid + dColBigLotAsk
    If dTemp > 0 Then
        m.tbStats(iCol, 5) = RoundNum(dColBigLotBid / dTemp * 100)
        m.tbStats(iCol + 1, 5) = RoundNum(dColBigLotAsk / dTemp * 100)
        m.tbStats(iCol + 2, 5) = RoundNum(dColBigLotOther / dTemp * 100)
    End If
    
    CalcRevCumulative iCol
    CalcRevTbMean
    
End Sub

Private Sub CalcStatsTable(ByVal eColStart As eTSV_TbStatsFields, ByVal eColEnd As eTSV_TbStatsFields)

    Dim i&
    Dim dDelta#, dDeltaCum#
    Dim dVolCum#, dPercent#
    Dim dVolBid#, dVolAsk#
    
    'precautionary
    If eColStart < eTSVTb_Col_A Then
        eColStart = eTSVTb_Col_A
    ElseIf eColStart > eTSVTb_Col_N Then
        eColStart = eTSVTb_Col_N
    End If
    
    If eColEnd < eTSVTb_Col_A Then
        eColEnd = eTSVTb_Col_A
    ElseIf eColEnd > eTSVTb_Col_N Then
        eColEnd = eTSVTb_Col_N
    End If
        
    For i = eColStart To eColEnd            'A thru N
        dDelta = m.tbStats(i, eTSVTb_Row_Ask) - m.tbStats(i, eTSVTb_Row_Bid)
        dVolCum = dVolCum + m.tbStats(i, eTSVTb_Row_Bid) + m.tbStats(i, eTSVTb_Row_Ask) + m.tbStats(i, eTSVTb_Row_OtherVol)
        
        If i > eTSVTb_Col_A Then
            dDeltaCum = m.tbStats(i - 1, eTSVTb_Row_DeltaCum) + dDelta
        Else
            dDeltaCum = dDelta
        End If
                
        'bid/ask vol totals
        dVolBid = dVolBid + m.tbStats(i, eTSVTb_Row_Bid)
        dVolAsk = dVolAsk + m.tbStats(i, eTSVTb_Row_Ask)
        '[+ -]
        m.tbStats(i, eTSVTb_Row_Delta) = dDelta
        
        '[+ -] Cum, Percent
         m.tbStats(i, eTSVTb_Row_DeltaCum) = dDeltaCum
        
        If dVolCum > 0 Then
            dPercent = Abs(RoundNum((dDeltaCum / dVolCum * 100), 0))
            m.tbStats(i, eTSVTb_Row_Percent) = dPercent
        End If
                
        'Total volume
        m.tbStats(i, eTSVTb_Row_VolTotal) = m.tbStats(i, eTSVTb_Row_Bid) + m.tbStats(i, eTSVTb_Row_Ask) + m.tbStats(i, eTSVTb_Row_OtherVol)
        
        'handle right-most total column
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Bid) = dVolBid
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Ask) = dVolAsk
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Delta) = dVolAsk - dVolBid
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_DeltaCum) = dDeltaCum
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Percent) = dPercent
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_VolTotal) = dVolCum
    Next
    
    'entire range for current session
    m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Range) = RoundNum(m.tbPriceVol(eTSVTb_Price, 0) - m.tbPriceVol(eTSVTb_Price, m.tbPriceVol.NumRecords - 1), 2)

    m.nTbStatsLastCol = eColEnd
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Identify weighted mean (green line) & value area (yellow lines) per column,
' then set the row number of the record where these lines are.
' Mean: (P1*V1 + P1*V2 + ... Pn*Vn) / Sum(V1,Vn)
' Value Area: sum volume on either side of mean until sum is at least 68% of total
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPerColIdx(ByVal iColBid&, ByVal iColStats&, ByVal dPxVx#, ByVal dVx#, _
    ByVal iHiIdx&, ByVal iLowIdx&, ByVal iCloseIdx&, ByVal iHiestVolIdx&)

    Dim dPrice#, dVolPercent#, dVol#
    Dim iAbove&, iMean&, iBelow&
    Dim iRec&, iColOther&, i&

    iMean = -1
    iAbove = -1
    iBelow = -1
    
    If dVx <= 0 Then Exit Sub
    
    dPrice = RoundToMinMove(dPxVx / dVx, m.dMinMove)    'mean for this column
    
    'search price table for the mean
    m.tbPriceVol.SearchAsIndex m.aTbIdx, eTSVTb_Price, dPrice, iMean
    
    dVol = 0
    dVolPercent = dVx * 0.68
    
    iMean = m.aTbIdx(iMean)
    iRec = m.tbPriceVol.NumRecords
    iColOther = TbColVolOther(iColBid)
    
    If iMean >= 0 And iMean < iRec Then
        'save row numbers for high, low, close & hiest volume
        m.tbAuction(iColStats, eTSVTb_Idx_Hi) = iHiIdx
        m.tbAuction(iColStats, eTSVTb_Idx_Low) = iLowIdx
        m.tbAuction(iColStats, eTSVTb_Idx_Close) = iCloseIdx
        m.tbAuction(iColStats, eTSVTb_Idx_HiestVol) = iHiestVolIdx
        'save row number where mean price is at
        m.tbAuction(iColStats, eTSVTb_Idx_Mean) = iMean
        'sum up volume on either side of mean until sum is at least 68% of total volume
        iAbove = iMean
        iBelow = iMean
        Do While dVol < dVolPercent
            iAbove = iAbove - 1
            If iAbove >= 0 And iAbove < iRec Then
                dVol = dVol + m.tbPriceVol(iColBid, iAbove) + m.tbPriceVol(iColBid + 1, iAbove) + m.tbPriceVol(iColOther, iAbove)
            End If
            If dVol < dVolPercent Then
                iBelow = iBelow + 1
                If iBelow >= 0 And iBelow < iRec Then
                    dVol = dVol + m.tbPriceVol(iColBid, iBelow) + m.tbPriceVol(iColBid + 1, iBelow) + m.tbPriceVol(iColOther, iBelow)
                End If
            End If
            If iAbove < 0 And iBelow >= iRec Then
                Exit Do
            End If
        Loop
        m.tbAuction(iColStats, eTSVTb_Idx_Above) = iAbove
        m.tbAuction(iColStats, eTSVTb_Idx_Below) = iBelow
    End If
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Identify value area (yellow lines) cumulative by group,
' then set the row number of the record where yellow lines are.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAboveBelowIdx(ByVal eGroup As eTSV_Groups, ByVal iMeanRow&, ByVal dVxVol#)

    Dim dVolPercent#, dVolSum#
    Dim iAbove&, iBelow&, iRec&

    dVolPercent = dVxVol * 0.68
    dVolSum = m.tbPriceVol(eTSVTb_VolRow_Total, iMeanRow)

    iAbove = iMeanRow
    iBelow = iMeanRow
    iRec = m.tbPriceVol.NumRecords

    Do While dVolSum < dVolPercent
        iAbove = iAbove - 1
        If iAbove >= 0 And iAbove < iRec Then
            dVolSum = dVolSum + m.tbPriceVol(eTSVTb_VolRow_Total, iAbove)
        End If
        If dVolSum < dVolPercent Then
            iBelow = iBelow + 1
            If iBelow >= 0 And iBelow < iRec Then
                dVolSum = dVolSum + m.tbPriceVol(eTSVTb_VolRow_Total, iBelow)
            End If
        End If
        If iAbove < 0 And iBelow >= iRec Then
            Exit Do
        End If
    Loop
    
    If eGroup = eTSV_Group_AB Then
    
        m.tpPrevColorRows.iMeanAB = m.tpColorRows.iMeanAB
        m.tpPrevColorRows.iAboveAB = m.tpColorRows.iAboveAB
        m.tpPrevColorRows.iBelowAB = m.tpColorRows.iBelowAB

        m.tpColorRows.iMeanAB = iMeanRow
        m.tpColorRows.iAboveAB = iAbove
        m.tpColorRows.iBelowAB = iBelow
        
    ElseIf eGroup = eTSV_Group_AE Then
    
        m.tpPrevColorRows.iMeanAE = m.tpColorRows.iMeanAE
        m.tpPrevColorRows.iAboveAE = m.tpColorRows.iAboveAE
        m.tpPrevColorRows.iBelowAE = m.tpColorRows.iBelowAE

        m.tpColorRows.iMeanAE = iMeanRow
        m.tpColorRows.iAboveAE = iAbove
        m.tpColorRows.iBelowAE = iBelow
        
    ElseIf eGroup = eTSV_Group_AI Then
    
        m.tpPrevColorRows.iMeanAI = m.tpColorRows.iMeanAI
        m.tpPrevColorRows.iAboveAI = m.tpColorRows.iAboveAI
        m.tpPrevColorRows.iBelowAI = m.tpColorRows.iBelowAI

        m.tpColorRows.iMeanAI = iMeanRow
        m.tpColorRows.iAboveAI = iAbove
        m.tpColorRows.iBelowAI = iBelow
    
    ElseIf eGroup = eTSV_Group_AN Then
        
        m.tpPrevColorRows.iMeanAN = m.tpColorRows.iMeanAN
        m.tpPrevColorRows.iAboveAN = m.tpColorRows.iAboveAN
        m.tpPrevColorRows.iBelowAN = m.tpColorRows.iBelowAN

        m.tpColorRows.iMeanAN = iMeanRow
        m.tpColorRows.iAboveAN = iAbove
        m.tpColorRows.iBelowAN = iBelow
    
    ElseIf eGroup = eTSV_Group_Current Then
        
        m.tpPrevColorRows.iMeanCurrent = m.tpColorRows.iMeanCurrent
        m.tpPrevColorRows.iAboveCurrent = m.tpColorRows.iAboveCurrent
        m.tpPrevColorRows.iBelowCurrent = m.tpColorRows.iBelowCurrent

        m.tpColorRows.iMeanCurrent = iMeanRow
        m.tpColorRows.iAboveCurrent = iAbove
        m.tpColorRows.iBelowCurrent = iBelow
    
    End If

End Sub

Private Sub SaveDrawTools()

    Dim i&, strText$
    Dim dPrice#, dPrice2#
    
    Dim aFile As New cGdArray
    Dim aFields As New cGdArray
    
    If Not m.aLineTool Is Nothing Then
        For i = 0 To m.aLineTool.Size - 1
            If m.aLineTool(i) > 0 Then
                dPrice = m.tbPriceVol(eTSVTb_Price, i)
                aFile.Add "LINE|" & Str(dPrice) & "|" & Str(m.aLineTool(i))
            End If
        Next
                
'loaded format: id|row|column|row|column|color|pixel|view|fill
'save to format: BOX|session|price|column|price|column|color|pixel|view OR revTick|fill
        For i = 0 To m.aBoxTool.Size - 1
            aFields.SplitFields m.aBoxTool(i), "|"
            
            If aFields(0) > 0 Then
                dPrice = m.tbPriceVol(eTSVTb_Price, ValOfText(aFields(1)))
                dPrice2 = m.tbPriceVol(eTSVTb_Price, ValOfText(aFields(3)))
                
                strText = "BOX|" & Str(m.nSessionDate)
                strText = strText & "|" & Str(dPrice) & "|" & aFields(2)
                strText = strText & "|" & Str(dPrice2) & "|" & aFields(4)
                strText = strText & "|" & aFields(5)
                strText = strText & "|" & aFields(6)
                strText = strText & "|" & aFields(7)
                strText = strText & "|" & aFields(8)
                
                aFile.Add strText
            End If
        Next
        
'format: ICON|key|icon|price|column|color|alignment (key = icon type concat with number)
        If Not m.aIconTool Is Nothing Then
            For i = 0 To m.aIconTool.Size - 1
                aFile.Add m.aIconTool(i)
            Next
        End If
        
        For i = 0 To m.aBoxToolOther.Size - 1
            aFile.Add m.aBoxToolOther(i)
        Next
        
        If aFile.Size > 0 Then
            aFile.ToFile m.strFile
        ElseIf FileExist(m.strFile) Then
            KillFile m.strFile
        End If
    End If
            
End Sub

Private Sub LoadDrawTools()
    
    Dim i&, j&
    Dim Top&, Bottom&
    Dim strText$, strFile$, strKey$
    Dim nSymID&, dPrice#, dPrice2#
    
    Dim aFile As New cGdArray
    Dim aFields As New cGdArray

    If m.TickBars Is Nothing Then Exit Sub
    If m.TickBars.Size = 0 Then Exit Sub

    If m.aBoxTool Is Nothing Then Set m.aBoxTool = New cGdArray
    m.aBoxTool.Size = 0
    
    If m.aBoxToolOther Is Nothing Then Set m.aBoxToolOther = New cGdArray
    m.aBoxToolOther.Size = 0
    
    If m.aLineTool Is Nothing Then Set m.aLineTool = New cGdArray
    m.aLineTool.Size = 0
    
    If m.aIconTool Is Nothing Then Set m.aIconTool = New cGdArray
    m.aIconTool.Size = 0

    If m.TickBars.SecurityType = "F" Then
        nSymID = GetSymbolID(m.TickBars.Prop(eBARS_BaseSymbol) & "-067")
    Else
        nSymID = m.TickBars.Prop(eBARS_SymbolID)
    End If
    
    If m.eTbType = eTSVTb_TbType_VolAtPrice Then
        strText = "^" & nSymID & "VP"
    Else
        strText = "^" & nSymID & "BD"
    End If
    
    m.strFile = g.ChartGlobals.strCPCRoot & "\Charts\" & strText & ".IOA"
    
    If Not m.tbPriceVol Is Nothing Then
        For i = 0 To m.tbPriceVol.NumRecords - 1
            m.aLineTool.Add 0
        Next
    End If
    
    If Not FileExist(m.strFile) Then Exit Sub

    aFile.FromFile m.strFile
    If aFile.Size = 0 Then Exit Sub
    
    For i = 0 To aFile.Size - 1
        strText = UCase(Parse(aFile(i), "|", 1))
        If strText = "LINE" Then
            dPrice = RoundToMinMove(ValOfText(Parse(aFile(i), "|", 2)), m.dMinMove)
            If m.tbPriceVol.SearchAsIndex(m.aTbIdx, 0, dPrice, j) Then
                m.aLineTool(m.aTbIdx(j)) = ValOfText(Parse(aFile(i), "|", 3))
            End If
        ElseIf strText = "BOX" Then
'saved format: BOX|session|price|column|price|column|color|pixel|view OR revTick
'load to format: id|row|column|row|column|color|pixel|view OR revTick
            aFields.SplitFields aFile(i), "|"
            If aFields(1) = m.nSessionDate Then
                'VolAtPrice: always load box if session date matches.
                'Bid/Ask Dir: load box only if session date AND reverse tick matches
                If m.eTbType = eTSVTb_TbType_VolAtPrice Or ValOfText(aFields(8)) = m.iReverseTick Then
                    dPrice = RoundToMinMove(aFields(2), m.dMinMove)
                    dPrice2 = RoundToMinMove(aFields(4), m.dMinMove)
                    If m.tbPriceVol.SearchAsIndex(m.aTbIdx, 0, dPrice, Top) Then
                        If m.tbPriceVol.SearchAsIndex(m.aTbIdx, 0, dPrice2, Bottom) Then
                            strText = m.aBoxTool.Size + 1       'run-time ID
                            strText = strText & "|" & m.aTbIdx(Top) & "|" & aFields(3)      'row|col
                            strText = strText & "|" & m.aTbIdx(Bottom) & "|" & aFields(5)   'row|col
                            strText = strText & "|" & aFields(6)    'color
                            strText = strText & "|" & aFields(7)    'pixel
                            strText = strText & "|" & aFields(8)    'view
                            strText = strText & "|" & aFields(9)    'fill
                            m.aBoxTool.Add strText
                        End If
                    End If
                Else
                    m.aBoxToolOther.Add aFile(i)    'boxes drawn on bid/ask directional form for other reverse tick values
                End If
            End If
        ElseIf strText = "ICON" Then
'format: ICON|key|icon|price|column|color|alignment|session (key = icon type concat with number)
            aFields.SplitFields aFile(i), "|"
            aFields(1) = aFields(2) & m.aIconTool.Size + 1
            If aFields(7) = m.nSessionDate Then
                m.aIconTool.Add aFields.JoinFields("|")
            End If
        End If
    Next
    
End Sub

Public Sub BuildTables(ByVal eTbType As eTSV_Table_Type)

    If eTbType = eTSVTb_TbType_VolAtPrice Then
        SaveDrawTools
        m.eTbType = eTSVTb_TbType_VolAtPrice
        BuildPriceVolTable
        LoadDrawTools
    ElseIf eTbType = eTSVTb_TbType_Reversal Then
        SaveDrawTools
        m.eTbType = eTSVTb_TbType_Reversal
        BuildReverseTable
        LoadDrawTools
    Else
        m.eTbType = eTSVTb_TbType_Invalid
    End If
                
End Sub

Private Sub ClearTbPBVol()

    Dim i&

    'clear pullback table
    For i = 0 To m.tbPBVol.NumRecords - 1
        m.tbPBVol.SetRecord "", i
    Next
    
    For i = 0 To m.tbPBStats.NumRecords - 1
        m.tbPBStats.SetRecord "", i
    Next

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calculate weighted mean (green line) & value areas (yellow lines) for reverse
' table then save prices at mean & value areas to last row in stats table
' This is different from PriceAtVol: columns are not fixed, there are no groupings
' Mean: (P1*V1 + P1*V2 + ... Pn*Vn) / Sum(V1,Vn)
' Value Area: sum volume on either side of mean until sum is at least 68% of total
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcRevTbMean()

    Dim dPriceAtMean#, dPriceAbove#, dPriceBelow#
    Dim dVol#, dVolPercent#
    Dim iAbove&, iMean&, iBelow&
    Dim iRec&, iCol&, i&
    
    'search price table for cumulative mean
    dPriceAtMean = RoundToMinMove(m.dPxVx / m.dVx, m.dMinMove)
    m.tbPriceVol.SearchAsIndex m.aTbIdx, eTSVTb_Price, dPriceAtMean, iMean
        
    iMean = m.aTbIdx(iMean)
    iRec = m.tbPriceVol.NumRecords
    iCol = m.tbPriceVol.NumFields
    
    'sum up volume on either side of mean until sum is at least 68% of total volume
    If iMean >= 0 And iMean < iRec Then
        dVol = 0
        'dVolPercent = (m.tbStats(iCol - 1, 0) + m.tbStats(iCol - 2, 0) + m.tbStats(iCol - 3, 0)) * 0.68
        dVolPercent = m.dVx * 0.68
        'yellow bands for individual column
        iAbove = iMean
        iBelow = iMean
        Do While dVol < dVolPercent
            iAbove = iAbove - 1
            If iAbove >= 0 And iAbove < iRec Then
                dVol = dVol + m.tbCumVolAtPrice(1, iAbove)
            End If
            If dVol < dVolPercent Then
                iBelow = iBelow + 1
                If iBelow >= 0 And iBelow < iRec Then
                    dVol = dVol + m.tbCumVolAtPrice(1, iBelow)
                End If
            End If
            If iAbove < 0 And iBelow >= iRec Then
                Exit Do
            End If
            dPriceAbove = m.tbCumVolAtPrice(0, iAbove)
            dPriceBelow = m.tbCumVolAtPrice(0, iBelow)
        Loop

        m.dPriceCumVAM = m.tbCumVolAtPrice(0, iAbove)
        m.dPriceCumVBM = m.tbCumVolAtPrice(0, iBelow)

'original code: leave awhile then remove 04-12-2007
'        Do While dVol < dVolPercent
'            iAbove = iAbove - 1
'            If iAbove >= 0 And iAbove < iRec Then
'                dVol = dVol + m.tbPriceVol(iCol - 1, iAbove) + m.tbPriceVol(iCol - 2, iAbove) + m.tbPriceVol(iCol - 3, iAbove)
'            End If
'            If dVol < dVolPercent Then
'                iBelow = iBelow + 1
'                If iBelow >= 0 And iBelow < iRec Then
'                    dVol = dVol + m.tbPriceVol(iCol - 1, iBelow) + m.tbPriceVol(iCol - 2, iBelow) + m.tbPriceVol(iCol - 3, iBelow)
'                End If
'            End If
'            If iAbove < 0 And iBelow >= iRec Then
'                Exit Do
'            End If
'            dPriceAbove = m.tbPriceVol(0, iAbove)
'            dPriceBelow = m.tbPriceVol(0, iBelow)
'        Loop
    End If

'NOTE: need to re-visit this for prices with ^ - such as TQ
'    If InStr(m.TickBars.PriceDisplay(dPriceAbove), "^") = 0 Then
'        m.tbStats(iCol - 1, 15) = m.TickBars.PriceDisplay(dPriceAbove)
'        m.tbStats(iCol - 2, 15) = m.TickBars.PriceDisplay(dPriceAtMean)
'        m.tbStats(iCol - 3, 15) = m.TickBars.PriceDisplay(dPriceBelow)
'    Else
        m.tbStats(iCol - 1, 15) = dPriceAbove
        m.tbStats(iCol - 2, 15) = dPriceAtMean
        m.tbStats(iCol - 3, 15) = dPriceBelow
'    End If

End Sub

Private Sub SetReverseTbData(ByVal iBar&, ByVal dReverseAmt#, ByVal bRT As Boolean)
    
    Dim dVol#, dPrice#, iColUpDownFLag&, j&
    Dim dMinutes#, dSpeed#
    Dim nPrevTradesAtBid&, nPrevTradesAtAsk&
    
    Dim aBid As cGdArray, aAsk As cGdArray, aOther As cGdArray
    
    Dim bStartNew As Boolean

    If iBar < 0 Or iBar >= m.TickBars.Size Then
        Exit Sub            'precautionary
    End If

    dPrice = RoundToMinMove(m.TickBars(eBARS_Close, iBar), m.dMinMove)
    iColUpDownFLag = m.tbPriceVol(m.nTbLastCol, 0)
    
    m.dVx = m.dVx + m.TickBars(eBARS_Vol, iBar)
    m.dPxVx = m.dPxVx + (dPrice * m.TickBars(eBARS_Vol, iBar))
    
    nPrevTradesAtBid = m.nTradesAtBid
    nPrevTradesAtAsk = m.nTradesAtAsk
            
    If m.TickBars(eBARS_Flags, iBar) = eTICK_AtBid Then
        If m.dHighTradeAtBid = 0 Then
            m.dHighTradeAtBid = dPrice
        ElseIf dPrice > m.dHighTradeAtBid Then
            m.dHighTradeAtBid = dPrice
            If iColUpDownFLag = 1 Then ClearTbPBVol
        End If
        m.nTradesAtBid = m.nTradesAtBid + m.TickBars(eBARS_Vol, iBar)
        If (nPrevTradesAtBid = 0 And nPrevTradesAtAsk = 0) Or (nPrevTradesAtBid < nPrevTradesAtAsk And m.nTradesAtBid > m.nTradesAtAsk) Then
            m.dPriceOfLastSwitch = RoundToMinMove(m.TickBars(eBARS_Close, iBar), m.dMinMove)
        End If
    ElseIf m.TickBars(eBARS_Flags, iBar) = eTICK_AtAsk Then
        If m.dLowTradeAtAsk = 0 Then
            m.dLowTradeAtAsk = dPrice
        ElseIf dPrice < m.dLowTradeAtAsk Then
            m.dLowTradeAtAsk = dPrice
            If iColUpDownFLag = -1 Then ClearTbPBVol
        End If
        m.nTradesAtAsk = m.nTradesAtAsk + m.TickBars(eBARS_Vol, iBar)
        If (nPrevTradesAtAsk = 0 And nPrevTradesAtBid = 0) Or (nPrevTradesAtAsk < nPrevTradesAtBid And m.nTradesAtAsk > m.nTradesAtBid) Then
            m.dPriceOfLastSwitch = RoundToMinMove(m.TickBars(eBARS_Close, iBar), m.dMinMove)
        End If
    End If
    
   If iBar > m.iFirstBar Then
        dMinutes = (m.TickBars(eBARS_DateTime, iBar) - m.TickBars(eBARS_DateTime, m.iFirstBar)) * 1440#
    End If
    
    If iColUpDownFLag = 0 Then
        If m.dHighTradeAtBid > 0 Then
            If m.dHighTradeAtBid - dPrice >= dReverseAmt Then
                iColUpDownFLag = 1
                m.tbPriceVol(m.nTbLastCol, 0) = iColUpDownFLag
                bStartNew = True
            End If
        ElseIf m.dLowTradeAtAsk > 0 Then
            If dPrice - m.dLowTradeAtAsk >= dReverseAmt Then
                iColUpDownFLag = -1
                m.tbPriceVol(m.nTbLastCol, 0) = iColUpDownFLag
                bStartNew = True
            End If
        End If
    ElseIf iColUpDownFLag = 1 Then
        If m.dHighTradeAtBid > 0 Then
            If m.dHighTradeAtBid - dPrice >= dReverseAmt Then
                bStartNew = True
            End If
        End If
    ElseIf iColUpDownFLag = -1 Then
        If m.dLowTradeAtAsk > 0 Then
            If dPrice - m.dLowTradeAtAsk >= dReverseAmt Then
                bStartNew = True
            End If
        End If
    Else
        'theoretically should never get here (precautionary)
    End If
            
    'Per Rasa: vol for tick that starts new bar should be added to old bar (08-21-2006)
    If (m.tbPriceVol.SearchAsIndex(m.aTbIdx, eTSVTb_Price, dPrice, j)) Then
        dVol = m.TickBars(eBARS_Vol, iBar)
        m.tbPriceVol(m.nTbLastCol, 1) = m.TickBars(eBARS_DateTime, iBar)
        If m.TickBars(eBARS_Flags, iBar) = eTICK_AtBid Then
            m.tbPriceVol(m.nTbLastCol, m.aTbIdx(j)) = m.tbPriceVol(m.nTbLastCol, m.aTbIdx(j)) + dVol
            m.tbStats(m.nTbLastCol, 0) = m.tbStats(m.nTbLastCol, 0) + dVol
            m.tbPBVol(0, m.aTbIdx(j)) = m.tbPBVol(0, m.aTbIdx(j)) + dVol
            m.tbCumVolAtPrice(1, m.aTbIdx(j)) = m.tbCumVolAtPrice(1, m.aTbIdx(j)) + dVol
            m.dBidTrades = m.dBidTrades + 1
            
            If dVol >= m.iBigLot Then
                m.tbStats(m.nTbLastCol, 4) = m.tbStats(m.nTbLastCol, 4) + dVol
                m.tbPBStats(0, 2) = m.tbPBStats(0, 2) + dVol
            End If
            If dMinutes > 0 Then
                dSpeed = m.dBidTrades / dMinutes
                If dSpeed >= m.dTradeSpeed Then
                    m.aSpeedCells.Add Str(m.aTbIdx(j)) & "," & Str(m.nTbLastCol)
                End If
            End If
        
        ElseIf m.TickBars(eBARS_Flags, iBar) = eTICK_AtAsk Then
            m.tbPriceVol(m.nTbLastCol + 1, m.aTbIdx(j)) = m.tbPriceVol(m.nTbLastCol + 1, m.aTbIdx(j)) + dVol
            m.tbStats(m.nTbLastCol + 1, 0) = m.tbStats(m.nTbLastCol + 1, 0) + dVol
            m.tbPBVol(1, m.aTbIdx(j)) = m.tbPBVol(1, m.aTbIdx(j)) + dVol
            m.tbCumVolAtPrice(1, m.aTbIdx(j)) = m.tbCumVolAtPrice(1, m.aTbIdx(j)) + dVol
            m.dAskTrades = m.dAskTrades + 1
            
            If dVol >= m.iBigLot Then
                m.tbPBStats(1, 2) = m.tbPBStats(1, 2) + dVol
                m.tbStats(m.nTbLastCol + 1, 4) = m.tbStats(m.nTbLastCol + 1, 4) + dVol
            End If
            If dMinutes > 0 Then
                dSpeed = m.dAskTrades / dMinutes
                If dSpeed >= m.dTradeSpeed Then
                    m.aSpeedCells.Add Str(m.aTbIdx(j)) & "," & Str(m.nTbLastCol + 1)
                End If
            End If
        
        Else
            m.tbPriceVol(m.nTbLastCol + 2, m.aTbIdx(j)) = m.tbPriceVol(m.nTbLastCol + 2, m.aTbIdx(j)) + dVol
            m.tbStats(m.nTbLastCol + 2, 0) = m.tbStats(m.nTbLastCol + 2, 0) + dVol
            m.tbPBVol(2, m.aTbIdx(j)) = m.tbPBVol(2, m.aTbIdx(j)) + dVol
            m.tbCumVolAtPrice(1, m.aTbIdx(j)) = m.tbCumVolAtPrice(1, m.aTbIdx(j)) + dVol

            If dVol >= m.iBigLot Then
                m.tbStats(m.nTbLastCol + 2, 4) = m.tbStats(m.nTbLastCol + 2, 4) + dVol
            End If
        End If
    End If
    
    If bRT Then CalcRevRT m.nTbLastCol
    
    If bStartNew Then
        If Not bRT Then CalcRevRT m.nTbLastCol
                
        ClearTbPBVol            'pullback tables
                
        'price vol table
        Set aBid = New cGdArray
        aBid.Create eGDARRAY_Doubles, m.tbPriceVol.NumRecords, 0
        Set aAsk = New cGdArray
        aAsk.Create eGDARRAY_Doubles, m.tbPriceVol.NumRecords, 0
        Set aOther = New cGdArray
        aOther.Create eGDARRAY_Doubles, m.tbPriceVol.NumRecords, 0
        'start new column & flag it as a down column
        m.tbPriceVol.AttachField aBid, m.tbPriceVol.NumFields
        m.tbPriceVol.AttachField aAsk, m.tbPriceVol.NumFields
        m.tbPriceVol.AttachField aOther, m.tbPriceVol.NumFields
        m.nTbLastCol = m.nTbLastCol + 3
        m.tbPriceVol(m.nTbLastCol, 0) = -1 * iColUpDownFLag
        
        'stats table
        Set aBid = New cGdArray
        aBid.Create eGDARRAY_Doubles, m.tbStats.NumRecords, 0
        Set aAsk = New cGdArray
        aAsk.Create eGDARRAY_Doubles, m.tbStats.NumRecords, 0
        Set aOther = New cGdArray
        aOther.Create eGDARRAY_Doubles, m.tbStats.NumRecords, 0
        'start new column
        m.tbStats.AttachField aBid, m.tbStats.NumFields
        m.tbStats.AttachField aAsk, m.tbStats.NumFields
        m.tbStats.AttachField aOther, m.tbStats.NumFields
        
        m.dHighTradeAtBid = 0
        m.dLowTradeAtAsk = 0
        m.dAskTrades = 0
        m.dBidTrades = 0
        
        bStartNew = False
        m.iFirstBar = iBar + 1
    End If

End Sub

Private Sub BuildReverseTable()

    Dim dPrice#, dPrice1#, dReverseAmt#
    Dim dMinutes#, dSpeed#, i&
            
    If m.TickBars Is Nothing Then
        Exit Sub
    ElseIf m.TickBars.Size = 0 Then
        Exit Sub
    End If
    
    m.aSpeedCells.Size = 0
    m.dBidTrades = 0
    m.dAskTrades = 0

'create table that holds per price data (upper portion of grid)
    Set m.tbPriceVol = New cGdTable
    Set m.tbPBVol = New cGdTable
        
    m.tbPriceVol.CreateField eGDARRAY_Doubles, 0, , 0       'price
    m.tbPriceVol.CreateField eGDARRAY_Doubles, 1, , 0       'trades at bid (also holds datetime)
    m.tbPriceVol.CreateField eGDARRAY_Doubles, 2, , 0       'trades at ask
    m.tbPriceVol.CreateField eGDARRAY_Doubles, 3, , 0       'trades at other
    
    m.tbPBVol.CreateField eGDARRAY_Doubles, 0, , 0           'trades at bid (also holds datetime)
    m.tbPBVol.CreateField eGDARRAY_Doubles, 1, , 0           'trades at ask
    m.tbPBVol.CreateField eGDARRAY_Doubles, 2, , 0           'trades at other
    
    m.nTbLastCol = 1
    
    m.tbPriceVol.AddRecord ""       'first row for up/down flag (1=up, 0=unknown, -1=down)
    m.tbPriceVol.AddRecord ""       'second row for timestamp
    
    m.tbPBVol.AddRecord ""          'first 2 rows blank placeholders
    m.tbPBVol.AddRecord ""
        
'create table that holds cumulative volume at a particular price for calculating cumulative VAM & VBM
    Set m.tbCumVolAtPrice = New cGdTable
    
    m.tbCumVolAtPrice.CreateField eGDARRAY_Doubles, 0, , 0       'price
    m.tbCumVolAtPrice.CreateField eGDARRAY_Doubles, 1, , 0       'cumulative volume at price
    
    m.tbCumVolAtPrice.AddRecord ""  'first 2 rows blank placeholders
    m.tbCumVolAtPrice.AddRecord ""
    
'create table that holds summary/cumulative totals (bottom portion of grid)
    Set m.tbStats = New cGdTable
    Set m.tbPBStats = New cGdTable
        
    m.tbStats.CreateField eGDARRAY_Strings, 0, , 0           'field name
    m.tbStats.CreateField eGDARRAY_Doubles, 1, , 0           'trades at bid
    m.tbStats.CreateField eGDARRAY_Doubles, 2, , 0           'trades at ask
    m.tbStats.CreateField eGDARRAY_Longs, 3, , 0             'trades at other
        
    m.tbPBStats.CreateField eGDARRAY_Doubles, 0, , 0           'trades at bid
    m.tbPBStats.CreateField eGDARRAY_Doubles, 1, , 0           'trades at ask
    m.tbPBStats.CreateField eGDARRAY_Longs, 2, , 0             'trades at other
    
    'fixed number of rows to hold:
    '   [Totals] [Diff] [Total%] (rows 0 thru 2)
    '   separator (row 3)
    '   [Big Lots] [Big Lots%] (rows 4 thru 5)
    '   separator (row 6)
    '   [Cum Big] [Cum Big Diff] [Cum Big%] [Cum Big Diff% of Total] (rows 7 thru 10)
    '   [Cum Totals] [Cum Diff] [Cum%] [Cum Diff% of Total] (rows 11 thru 14)
    For i = 0 To 15
        m.tbStats.AddRecord ""
    Next
    m.tbStats(0, 0) = "Totals"
    m.tbStats(0, 1) = "Differential"
    m.tbStats(0, 2) = "Total %"
    
    m.tbStats(0, 4) = "Big Lots"
    m.tbStats(0, 5) = "Big Lots %"

    m.tbStats(0, 7) = "Cumulative Big"
    m.tbStats(0, 8) = "Cum Big Diff"
    m.tbStats(0, 9) = "Cumulative Big %"
    m.tbStats(0, 10) = "Cum Big D% of Tot"
    
    m.tbStats(0, 11) = "Cumulative"
    m.tbStats(0, 12) = "Cumulative Diff"
    m.tbStats(0, 13) = "Cumulative %"
    m.tbStats(0, 14) = "Cum Diff% of Tot"
    m.tbStats(0, 15) = "Mean,VAM,VBM"               'hidden record containing price at MEAN, VAM and VBM
    
    'fixed number of rows to hold: [Totals] [Diff] [Big Lots]
    m.tbPBStats.AddRecord ""
    m.tbPBStats.AddRecord ""
    m.tbPBStats.AddRecord ""

'fill table with all possible values that are dMinMove apart from dMaxPrice to dMinPrice
    dPrice = RoundToMinMove(gdMaxValue(m.TickBars.ArrayHandle(eBARS_Close), 0, m.TickBars.Size), m.dMinMove)
    dPrice1 = RoundToMinMove(gdMinValue(m.TickBars.ArrayHandle(eBARS_Close), 0, m.TickBars.Size), m.dMinMove)
    
    
    If m.iBlankRows > 0 Then
        dPrice = dPrice + m.dMinMove * m.iBlankRows
        dPrice1 = dPrice1 - m.dMinMove * m.iBlankRows
        dPrice = RoundToMinMove(dPrice, m.dMinMove)
        dPrice1 = RoundToMinMove(dPrice1, m.dMinMove)
    End If
    
    Do While dPrice >= dPrice1
        m.tbPBVol.AddRecord ""
        
        m.tbPriceVol.AddRecord ""
        m.tbPriceVol(eTSVTb_Price, m.tbPriceVol.NumRecords - 1) = RoundToMinMove(dPrice, m.dMinMove)
        
        m.tbCumVolAtPrice.AddRecord ""
        m.tbCumVolAtPrice(eTSVTb_Price, m.tbCumVolAtPrice.NumRecords - 1) = m.tbPriceVol(eTSVTb_Price, m.tbPriceVol.NumRecords - 1)
        
        dPrice = RoundToMinMove(dPrice - m.dMinMove, m.dMinMove)
    Loop
    
    Set m.aTbIdx = m.tbPriceVol.CreateSortedIndex(eTSVTb_Price)
               
    If m.iBigLot <= 0 Then m.iBigLot = 500
    If m.iReverseTick <= 2 Then m.iReverseTick = 2
    If m.dTradeSpeed <= 0 Then m.dTradeSpeed = 10000
    
    dReverseAmt = m.iReverseTick * m.dMinMove
    m.iFirstBar = 0
    m.dHighTradeAtBid = 0
    m.dLowTradeAtAsk = 0
    
    m.nTradesAtBid = 0
    m.nTradesAtAsk = 0
    m.dPriceOfLastSwitch = 0#
    
    m.dVx = 0
    m.dPxVx = 0
    
    For i = 0 To m.TickBars.Size - 1
        SetReverseTbData i, dReverseAmt, False
    Next
    
    CalcRevRT m.nTbLastCol
    CalcPullbackStats
        
    m.aSpeedCells.Sort eGdSort_Default Or eGdSort_DeleteDuplicates
    
End Sub

Private Sub BuildPriceVolTable()
    
    Static strPrevSymID&
    
    Dim dMinPrice#, dMaxPrice#, dPrice#
    Dim dTimeStart#, dTimeEnd#, dBarTime#, dLastBarTime#, dVol#
    Dim bAdd As Boolean, bCalcHistory As Boolean
    
    Dim iCol&, iColOtherVol&, iColStats&, iBarStart&
    Dim i&, j&, k&, iMeanIdx&
        
    Dim aPxVxRows As New cGdArray       'rows to sum Px*Vx
    Dim dPxVx#, dVx#                    'Mean: Sum(Px*Vx)/Sum(Vx)
    Dim dVAM#, dVBM#                    'these are cumulative sums
        
    Dim iHigh&, iLow&, iClose&
    Dim dColPxVx#, dColVx#
    Dim dColVolHi#, dTotalColVolHi#, iColVolHi&
        
    If m.tbPriceVol Is Nothing Then
        Set m.tbPriceVol = New cGdTable
        m.tbPriceVol.CreateField eGDARRAY_Doubles, eTSVTb_Price, , 0
        '3 fields each for volume from A-N (14 x 2)
        '1 fields for volume for totals
        For i = 0 To 42
            m.tbPriceVol.CreateField eGDARRAY_Doubles, , , 0
        Next
    Else
        m.tbPriceVol.NumRecords = 0
    End If
        
    If m.tbStats Is Nothing Then
        Set m.tbStats = New cGdTable
        bCalcHistory = True
                
        '1 field each for volume from A-N
        '1 field for data in [total] colum
        For i = 0 To 14
            m.tbStats.CreateField eGDARRAY_Doubles, , , 0
        Next
        'add fixed number of records
        'BidVol,AskVol,[+ -],[+ -]Cum,Percent,VAM,Mean,VBM,Totals,AvgVol,SHVol,Range,AvgRange,SHRange,Other
        For i = 0 To 14
            m.tbStats.AddRecord ""
        Next
    End If
    
    If m.tbAuction Is Nothing Then
        Set m.tbAuction = New cGdTable
        '1 field each for A-N and one field for totals
        For i = 0 To 14
            m.tbAuction.CreateField eGDARRAY_Longs, , , -1
        Next
        'fixed number of records
        'high,low,close,above,mean,below,hiestVol
        For i = 0 To 6
            m.tbAuction.AddRecord ""
        Next
    Else
        m.tbAuction.NumRecords = 0
    End If
    
    ResetColorRowIdx m.tpColorRows
    ResetColorRowIdx m.tpPrevColorRows
    
    If m.TickBars Is Nothing Then
        Exit Sub        'precautionary
    ElseIf m.TickBars.Size = 0 Then
        Exit Sub        'no data for current session
    End If
    
    If m.TickBars.Prop(eBARS_SymbolID) <> strPrevSymID Then
        bCalcHistory = True
    Else
        m.tbStats.SetRecord "", eTSVTb_Row_Bid
        m.tbStats.SetRecord "", eTSVTb_Row_Ask
        m.tbStats.SetRecord "", eTSVTb_Row_Delta
        m.tbStats.SetRecord "", eTSVTb_Row_DeltaCum
        m.tbStats.SetRecord "", eTSVTb_Row_Percent
        m.tbStats.SetRecord "", eTSVTb_Row_VAM
        m.tbStats.SetRecord "", eTSVTb_Row_Mean
        m.tbStats.SetRecord "", eTSVTb_Row_VBM
        m.tbStats.SetRecord "", eTSVTb_Row_VolTotal
        m.tbStats.SetRecord "", eTSVTb_Row_OtherVol
    End If
    
    If bCalcHistory Then
        InfBox "Loading data.  Please wait...", , , "Volume at Price", True
    End If
    
    dTimeStart = m.TickBars(eBARS_DateTime, 0)
    dTimeEnd = dTimeStart + 30 / 1440
    dLastBarTime = m.TickBars(eBARS_DateTime, m.TickBars.Size - 1)
       
    dMaxPrice = RoundToMinMove(gdMaxValue(m.TickBars.ArrayHandle(eBARS_Close), iBarStart, m.TickBars.Size), m.dMinMove)
    dMinPrice = RoundToMinMove(gdMinValue(m.TickBars.ArrayHandle(eBARS_Close), iBarStart, m.TickBars.Size), m.dMinMove)
    
    If m.iBlankRows > 0 Then
        dMaxPrice = dMaxPrice + m.dMinMove * m.iBlankRows
        dMinPrice = dMinPrice - m.dMinMove * m.iBlankRows
        dMaxPrice = RoundToMinMove(dMaxPrice, m.dMinMove)
        dMinPrice = RoundToMinMove(dMinPrice, m.dMinMove)
    End If
    
    'fill table with all possible values that are dMinMove apart from dMaxPrice to dMinPrice
    Do While dMaxPrice >= dMinPrice
        m.tbPriceVol.AddRecord ""
        m.tbPriceVol(eTSVTb_Price, m.tbPriceVol.NumRecords - 1) = RoundToMinMove(dMaxPrice, m.dMinMove)
        dMaxPrice = RoundToMinMove(dMaxPrice - m.dMinMove, m.dMinMove)
    Loop
    
    Set m.aTbIdx = m.tbPriceVol.CreateSortedIndex(eTSVTb_Price)

    'fill tables with volume data
    iCol = eTSVTb_VolBid_A
    iColOtherVol = eTSVTb_VolOther_A
    iColStats = eTSVTb_Col_A
    m.nTimeSalesVolTotal = 0
    'per column values for auction bar
    dColPxVx = 0
    dColVx = 0
    iHigh = -1
    iLow = -1
    iClose = -1
    iColVolHi = -1
    
    If m.aTbIdx.Size = m.tbPriceVol.NumRecords Then
        For i = 0 To m.TickBars.Size - 1
            bAdd = False
            dBarTime = m.TickBars(eBARS_DateTime, i)
            If Hour(dBarTime) < Hour(dTimeEnd) Then
                bAdd = True
            ElseIf Hour(dBarTime) = Hour(dTimeEnd) Then
                If Minute(dBarTime) < Minute(dTimeEnd) Then
                    bAdd = True
                End If
            ElseIf Int(dBarTime) < Int(dTimeEnd) Then
                'data crossing midnight into next day
                If Hour(dBarTime) = 23 And Hour(dTimeEnd) = 0 Then bAdd = True
            End If
            
            If iCol > eTSVTb_VolBid_N Then
                Exit For            'need to find out why this is happening for YM1
            End If
                       
            If bAdd Then
                dPrice = RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)
                dVol = m.TickBars(eBARS_Vol, i)
                If (m.tbPriceVol.SearchAsIndex(m.aTbIdx, eTSVTb_Price, dPrice, j)) Then
                    If m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                        m.tbPriceVol(iCol, m.aTbIdx(j)) = m.tbPriceVol(iCol, m.aTbIdx(j)) + dVol
                        m.tbStats(iColStats, eTSVTb_Row_Bid) = m.tbStats(iColStats, eTSVTb_Row_Bid) + dVol
                    ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                        m.tbPriceVol(iCol + 1, m.aTbIdx(j)) = m.tbPriceVol(iCol + 1, m.aTbIdx(j)) + dVol
                        m.tbStats(iColStats, eTSVTb_Row_Ask) = m.tbStats(iColStats, eTSVTb_Row_Ask) + dVol
                    Else
                        m.tbPriceVol(iColOtherVol, m.aTbIdx(j)) = m.tbPriceVol(iColOtherVol, m.aTbIdx(j)) + dVol
                        m.tbStats(iColStats, eTSVTb_Row_OtherVol) = m.tbStats(iColStats, eTSVTb_Row_OtherVol) + dVol
                    End If
                    
                    m.tbPriceVol(eTSVTb_VolRow_Total, m.aTbIdx(j)) = m.tbPriceVol(eTSVTb_VolRow_Total, m.aTbIdx(j)) + dVol
                    m.nTimeSalesVolTotal = m.nTimeSalesVolTotal + dVol
                    
                    'cumulative values
                    dVx = dVx + dVol
                    aPxVxRows.Add m.aTbIdx(j)
                    'per column values
                    dColVx = dColVx + dVol
                    dColPxVx = dColPxVx + (dPrice * dVol)
                    dVol = m.tbPriceVol(iCol, m.aTbIdx(j)) + m.tbPriceVol(iCol + 1, m.aTbIdx(j)) + m.tbPriceVol(iColOtherVol, m.aTbIdx(j))
                    If dVol > dColVolHi Then
                        iColVolHi = m.aTbIdx(j)
                        dColVolHi = dVol
                    End If
                    dVol = m.tbPriceVol(eTSVTb_VolRow_Total, m.aTbIdx(j))
                    If dVol > dTotalColVolHi Then
                        m.iTotalColVolHi = m.aTbIdx(j)
                        dTotalColVolHi = dVol
                    End If
                    iClose = m.aTbIdx(j)
                    
                    m.nTbLastCol = iCol
                    m.dTbColTimeStart = dTimeStart
                    m.dTbColTimeEnd = dTimeEnd
                End If
            ElseIf Int(dTimeEnd) > Int(dLastBarTime) Then
                Exit For            'something is wrong here
            Else
                'calculate cumulative Mean for column
                aPxVxRows.Sort eGdSort_DeleteDuplicates
                For k = 0 To aPxVxRows.Size - 1
                    dVol = m.tbPriceVol(iCol, aPxVxRows(k)) + m.tbPriceVol(iCol + 1, aPxVxRows(k)) + m.tbPriceVol(iColOtherVol, aPxVxRows(k))
                    dPxVx = dPxVx + (dVol * m.tbPriceVol(eTSVTb_Price, aPxVxRows(k)))
                Next
                If dVx > 0 Then
                    dPrice = RoundToMinMove(dPxVx / dVx, m.dMinMove)
                    m.tbStats(iColStats, eTSVTb_Row_Mean) = dPrice
                End If
                'calculate cumulative VAM & VBM for column
                If (m.tbPriceVol.SearchAsIndex(m.aTbIdx, eTSVTb_Price, dPrice, j)) Then
                    iMeanIdx = m.aTbIdx(j)
                    For k = 0 To aPxVxRows.Size - 1
                        If aPxVxRows(k) > iMeanIdx Then
                            dVBM = dVBM + m.tbPriceVol(iCol, aPxVxRows(k)) + m.tbPriceVol(iCol + 1, aPxVxRows(k)) + m.tbPriceVol(iColOtherVol, aPxVxRows(k))
                        ElseIf aPxVxRows(k) < iMeanIdx Then
                            dVAM = dVAM + m.tbPriceVol(iCol, aPxVxRows(k)) + m.tbPriceVol(iCol + 1, aPxVxRows(k)) + m.tbPriceVol(iColOtherVol, aPxVxRows(k))
                        End If
                    Next
                    m.tbStats(iColStats, eTSVTb_Row_VAM) = dVAM
                    m.tbStats(iColStats, eTSVTb_Row_VBM) = dVBM
                    If iColStats = eTSVTb_Col_B Then
                        SetAboveBelowIdx eTSV_Group_AB, iMeanIdx, dVx
                    ElseIf iColStats = eTSVTb_Col_E Then
                        SetAboveBelowIdx eTSV_Group_AE, iMeanIdx, dVx
                    ElseIf iColStats = eTSVTb_Col_I Then
                        SetAboveBelowIdx eTSV_Group_AI, iMeanIdx, dVx
                    ElseIf iColStats = eTSVTb_Col_N Then
                        SetAboveBelowIdx eTSV_Group_AN, iMeanIdx, dVx
                    End If
                End If
                
                'calculate per column mean & value area
                SetPerColIdx iCol, iColStats, dColPxVx, dColVx, aPxVxRows(0), aPxVxRows(aPxVxRows.Size - 1), iClose, iColVolHi
                
                'reset variables
                aPxVxRows.Size = 0
                dColPxVx = 0
                dColVx = 0
                iHigh = -1
                iLow = -1
                iClose = -1
                iColVolHi = -1
                dColVolHi = 0
                
                'save cumulative values
                m.dPxVx = dPxVx
                m.dVx = dVx
                m.ePxVxCol = iCol
                
                'set iCol to vol bid of next letter, i.e. VolBid_B, VolBid_C, VolBid_D etc.
                iCol = iCol + 2
                iColOtherVol = iColOtherVol + 1
                iColStats = iColStats + 1
                dTimeStart = dTimeEnd
                dTimeEnd = dTimeEnd + 30 / 1440
                i = i - 1
            End If
        Next
    End If

    'calculate cumulative Mean for last column
    If iCol <= eTSVTb_VolBid_N Then
        aPxVxRows.Sort eGdSort_DeleteDuplicates
        For k = 0 To aPxVxRows.Size - 1
            dVol = m.tbPriceVol(iCol, aPxVxRows(k)) + m.tbPriceVol(iCol + 1, aPxVxRows(k)) + m.tbPriceVol(iColOtherVol, aPxVxRows(k))
            dPxVx = dPxVx + (dVol * m.tbPriceVol(0, aPxVxRows(k)))
        Next
        If dVx > 0 Then
            dPrice = RoundToMinMove(dPxVx / dVx, m.dMinMove)
            m.tbStats(iColStats, eTSVTb_Row_Mean) = dPrice
            m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_Mean) = dPrice
        End If
    End If
    'calculate cumulative VAM & VBM for last column
    If (m.tbPriceVol.SearchAsIndex(m.aTbIdx, eTSVTb_Price, dPrice, j)) Then
        iMeanIdx = m.aTbIdx(j)
        For k = 0 To aPxVxRows.Size - 1
            If aPxVxRows(k) > iMeanIdx Then
                dVBM = dVBM + m.tbPriceVol(iCol, aPxVxRows(k)) + m.tbPriceVol(iCol + 1, aPxVxRows(k)) + m.tbPriceVol(iColOtherVol, aPxVxRows(k))
            ElseIf aPxVxRows(k) < iMeanIdx Then
                dVAM = dVAM + m.tbPriceVol(iCol, aPxVxRows(k)) + m.tbPriceVol(iCol + 1, aPxVxRows(k)) + m.tbPriceVol(iColOtherVol, aPxVxRows(k))
            End If
        Next
        m.tbStats(iColStats, eTSVTb_Row_VAM) = dVAM
        m.tbStats(iColStats, eTSVTb_Row_VBM) = dVBM
        SetAboveBelowIdx eTSV_Group_Current, iMeanIdx, dVx
        'VAM & VBM of 'totals' column is same
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_VAM) = dVAM
        m.tbStats(eTSVTb_Col_Total, eTSVTb_Row_VBM) = dVBM
        'set indexes for green/yellow bars
        If iColStats = eTSVTb_Col_B Then
            SetAboveBelowIdx eTSV_Group_AB, iMeanIdx, dVx
        ElseIf iColStats = eTSVTb_Col_E Then
            SetAboveBelowIdx eTSV_Group_AE, iMeanIdx, dVx
        ElseIf iColStats = eTSVTb_Col_I Then
            SetAboveBelowIdx eTSV_Group_AI, iMeanIdx, dVx
        ElseIf iColStats = eTSVTb_Col_N Then
            SetAboveBelowIdx eTSV_Group_AN, iMeanIdx, dVx
        End If
    End If
    'per column mean & value area for last column
    SetPerColIdx iCol, iColStats, dColPxVx, dColVx, aPxVxRows(0), aPxVxRows(aPxVxRows.Size - 1), iClose, iColVolHi
    
    'per column mean & value area for total column
    m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Hi) = 0
    m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Low) = m.tbPriceVol.NumRecords - 1
    m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Above) = m.tpColorRows.iAboveCurrent
    m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Mean) = m.tpColorRows.iMeanCurrent
    m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_Below) = m.tpColorRows.iBelowCurrent
    m.tbAuction(eTSVTb_Col_Total, eTSVTb_Idx_HiestVol) = m.iTotalColVolHi
    
    aPxVxRows.Size = 0
    m.nTbStatsLastCol = iColStats
    
    If bCalcHistory Then
        CalcHistoryStats 290        'want 200 trading days request more for holidays & weekends
    End If
    
    CalcStatsTable eTSVTb_Col_A, iColStats
    
    strPrevSymID = m.TickBars.Prop(eBARS_SymbolID)
        
    InfBox ""

End Sub

Public Sub DumpTable()

    Dim aTableTest As New cGdArray
    Dim i&, strText$
    
'    If Not m.tbAuction Is Nothing Then
'        For i = 0 To m.tbAuction.NumRecords - 1
'            strText = m.tbAuction.GetRecord(i, vbTab)
'            aTableTest.Add strText
'        Next
'        aTableTest.ToFile "AuctionTable.txt"
'    End If
'
'    If Not m.tbStats Is Nothing Then
'        For i = 0 To m.tbStats.NumRecords - 1
'            strText = m.tbStats.GetRecord(i, vbTab)
'            aTableTest.Add strText
'        Next
'        aTableTest.ToFile "StatsTable.txt"
'    End If
'
    If Not m.tbPriceVol Is Nothing Then
        aTableTest.Size = 0
        aTableTest.Add m.tbPriceVol.GetRecord(0, vbTab)
        For i = 0 To m.tbPriceVol.NumFields
            If m.tbPriceVol(i, 1) > 0 Then
                strText = strText & vbTab & DateFormat(m.tbPriceVol(i, 1), NO_DATE, HH_MM_SS)
            Else
                strText = strText & vbTab
            End If
        Next
        aTableTest.Add strText
        For i = 2 To m.tbPriceVol.NumRecords - 1
            strText = m.tbPriceVol.GetRecord(i, vbTab)
            aTableTest.Add strText
        Next
        aTableTest.ToFile "PriceVolTable.txt"
    End If
'
'    If Not m.TradeSpeedBars Is Nothing Then
'        For i = 0 To m.TradeSpeedBars.Size - 1
'            strText = DateFormat(m.TradeSpeedBars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM_SS) & vbTab
'            strText = strText & m.TradeSpeedBars(eBARS_Vol, i) & vbTab
'            strText = strText & m.TradeSpeedBars(eBARS_Bid, i) & vbTab
'            strText = strText & m.TradeSpeedBars(eBARS_BidVol, i) & vbTab
'            strText = strText & m.TradeSpeedBars(eBARS_Ask, i) & vbTab
'            strText = strText & m.TradeSpeedBars(eBARS_AskVol, i)
'            aTableTest.Add strText
'        Next
'        aTableTest.ToFile "TradeSpeed.txt"
'    End If

'    If Not m.tbPBVol Is Nothing Then
'        aTableTest.Size = 0
'        For i = 0 To m.tbPBVol.NumRecords
'            strText = m.tbPriceVol(0, i)
'            strText = strText & vbTab & m.tbPBVol.GetRecord(i, vbTab)
'            aTableTest.Add strText
'        Next
'        aTableTest.ToFile "PullBack.txt"
'    End If
        
    Set aTableTest = Nothing

End Sub

Public Property Get SessionDate() As Long
    SessionDate = m.nSessionDate
End Property

Public Property Get BidAskVolBars() As cGdBars
    Set BidAskVolBars = m.BidAskVolBars
End Property

Public Property Get TradeSpeedBars() As cGdBars
    Set TradeSpeedBars = m.TradeSpeedBars
End Property

Public Property Get TimeRangeBars() As cGdArray
    Set TimeRangeBars = m.aTimeRangeBars
End Property

Public Property Get TickBars() As cGdBars
    Set TickBars = m.TickBars
End Property

Public Property Get TimeSalesVolTotal() As Long
    TimeSalesVolTotal = m.nTimeSalesVolTotal
End Property

Public Property Get AuctionTable() As cGdTable
    Set AuctionTable = m.tbAuction
End Property

Public Property Get StatsTable() As cGdTable
    Set StatsTable = m.tbStats
End Property

Public Property Get PriceVolTable() As cGdTable
    Set PriceVolTable = m.tbPriceVol
End Property

Public Property Get PriceVolPBTable() As cGdTable
    Set PriceVolPBTable = m.tbPBVol
End Property

Public Property Get StatsPBTable() As cGdTable
    Set StatsPBTable = m.tbPBStats
End Property

Public Property Get CumVolPriceTable() As cGdTable
    Set CumVolPriceTable = m.tbCumVolAtPrice
End Property

Public Property Get PriceVolTableIdx() As cGdArray
    Set PriceVolTableIdx = m.aTbIdx
End Property

Public Property Get PriceVolLastDataCol() As Long
    PriceVolLastDataCol = m.nTbLastCol
End Property

Public Property Get SpeedCells() As cGdArray
    Set SpeedCells = m.aSpeedCells
End Property

Public Property Get StatsLastDataCol() As Long
    StatsLastDataCol = m.nTbStatsLastCol
End Property

Public Property Get StatHighVolAtPrice(ByVal eCol As eTSV_TbFields) As Long

    Select Case eCol
        Case eTSVTb_VolBid_A
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 2))
        Case eTSVTb_VolBid_B
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 3))
        Case eTSVTb_VolBid_C
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 4))
        Case eTSVTb_VolBid_D
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 5))
        Case eTSVTb_VolBid_E
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 6))
        Case eTSVTb_VolBid_F
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 7))
        Case eTSVTb_VolBid_G
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 8))
        Case eTSVTb_VolBid_H
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 9))
        Case eTSVTb_VolBid_I
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 10))
        Case eTSVTb_VolBid_J
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 11))
        Case eTSVTb_VolBid_K
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 12))
        Case eTSVTb_VolBid_L
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 13))
        Case eTSVTb_VolBid_M
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 14))
        Case eTSVTb_VolBid_N
            StatHighVolAtPrice = Val(Parse(m.aStatHighVal(0), vbTab, 15))
    End Select

End Property

Private Sub ResetColorRowIdx(tpRowIdx As eTSV_ColorRow)
    
    tpRowIdx.iMeanCurrent = 0
    tpRowIdx.iMeanAB = 0
    tpRowIdx.iMeanAE = 0
    tpRowIdx.iMeanAI = 0
    tpRowIdx.iMeanAN = 0
    
    tpRowIdx.iAboveCurrent = 0
    tpRowIdx.iAboveAB = 0
    tpRowIdx.iAboveAE = 0
    tpRowIdx.iAboveAI = 0
    tpRowIdx.iAboveAN = 0
    
    tpRowIdx.iBelowCurrent = 0
    tpRowIdx.iBelowAB = 0
    tpRowIdx.iBelowAE = 0
    tpRowIdx.iBelowAI = 0
    tpRowIdx.iBelowAN = 0
    
End Sub

Public Sub ColorRowIdx(ByVal eGroup As eTSV_Groups, ByVal bPrevious As Boolean, _
    iMean&, iAbove&, iBelow&)
        
    Dim tpRows As eTSV_ColorRow
    
    If bPrevious Then
        tpRows = m.tpPrevColorRows
    Else
        tpRows = m.tpColorRows
    End If
        
    If eGroup = eTSV_Group_AB Then
        iMean = tpRows.iMeanAB
        iAbove = tpRows.iAboveAB
        iBelow = tpRows.iBelowAB
    ElseIf eGroup = eTSV_Group_AE Then
        iMean = tpRows.iMeanAE
        iAbove = tpRows.iAboveAE
        iBelow = tpRows.iBelowAE
    ElseIf eGroup = eTSV_Group_AI Then
        iMean = tpRows.iMeanAI
        iAbove = tpRows.iAboveAI
        iBelow = tpRows.iBelowAI
    ElseIf eGroup = eTSV_Group_AN Then
        iMean = tpRows.iMeanAN
        iAbove = tpRows.iAboveAN
        iBelow = tpRows.iBelowAN
    ElseIf eGroup = eTSV_Group_Current Then
        iMean = tpRows.iMeanCurrent
        iAbove = tpRows.iAboveCurrent
        iBelow = tpRows.iBelowCurrent
    End If
    
End Sub

Public Property Get ReverseTick() As Long
    ReverseTick = m.iReverseTick
End Property

Public Property Get BigLot() As Long
    BigLot = m.iBigLot
End Property

Public Property Get HighlightTrades() As Long
    HighlightTrades = m.dTradeSpeed
End Property

Public Property Let ReverseTick(ByVal iTick&)
    m.iReverseTick = iTick
End Property

Public Property Let BigLot(ByVal iLot&)
    m.iBigLot = iLot
End Property

Public Property Let HighlightTrades(ByVal iTrades&)
    m.dTradeSpeed = iTrades
End Property

Public Property Let LineToolColor(ByVal idx&, ByVal iColor&)
    
    Dim i&
    
    If idx >= 0 Then
        If idx < m.aLineTool.Size Then
            m.aLineTool(idx) = iColor
        Else
            For i = m.aLineTool.Size To idx
                m.aLineTool(i) = 0
            Next
            m.aLineTool(idx) = iColor
        End If
    End If
    
End Property

Public Property Get LineToolColor(ByVal idx&) As Long

    If idx >= 0 And idx < m.aLineTool.Size Then
        LineToolColor = ValOfText(m.aLineTool(idx))
    Else
        LineToolColor = 0
    End If

End Property

Public Property Get BoxColor(ByVal iBoxId&) As Long

    Dim i&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            BoxColor = aFields(5)
            Exit For
        End If
    Next

End Property

Public Property Let BoxColor(ByVal iBoxId&, ByVal iColor&)

    Dim i&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            aFields(5) = iColor
            m.aBoxTool(i) = aFields.JoinFields("|")
            Exit For
        End If
    Next

End Property

Public Property Get BoxThickness(ByVal iBoxId&) As Long

    Dim i&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            BoxThickness = aFields(6)
            Exit For
        End If
    Next

End Property

Public Property Let BoxThickness(ByVal iBoxId&, ByVal iPix&)

    Dim i&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            aFields(6) = iPix
            m.aBoxTool(i) = aFields.JoinFields("|")
            Exit For
        End If
    Next

End Property

Public Property Get BoxFill(ByVal iBoxId&) As Long

    Dim i&
    Dim aFields As New cGdArray

'format: id|price|column|price|column|color|pixel|view|fill

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            BoxFill = ValOfText(aFields(8))
            Exit For
        End If
    Next

End Property

Public Property Let BoxFill(ByVal iBoxId&, ByVal iFill&)

    Dim i&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view|fill

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            aFields(8) = iFill
            m.aBoxTool(i) = aFields.JoinFields("|")
            Exit For
        End If
    Next

End Property

Public Property Get BoxView(ByVal iBoxId&) As Long

    Dim i&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            BoxView = aFields(7)
            Exit For
        End If
    Next

End Property

Public Sub BoxRect(ByVal iBoxId&, iTop&, iLeft&, iBottom&, iRight&)

    Dim i&
    Dim aFields As New cGdArray
    
    iTop = -1
    iLeft = -1
    iBottom = -1
    iRight = -1
    
'format: id|price|column|price|column|color|pixel|view

    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        If aFields(0) = iBoxId Then
            iTop = aFields(1)
            iLeft = aFields(2)
            iBottom = aFields(3)
            iRight = aFields(4)
            Exit For
        End If
    Next
        
End Sub

Public Sub AddBox(ByVal iTop&, ByVal iLeft&, ByVal iBottom&, ByVal iRight&, _
    ByVal iColor&, ByVal iPix&, ByVal iViewOrTickR&, ByVal iBoxFill&)

    Dim strBox$
    
'format: id|price|column|price|column|color|pixel|view OR revTick
    
    strBox = Str(m.aBoxTool.Size + 1) & "|"
    strBox = strBox & Str(iTop) & "|" & Str(iLeft) & "|"
    strBox = strBox & Str(iBottom) & "|" & Str(iRight) & "|"
    strBox = strBox & Str(iColor) & "|" & Str(iPix) & "|"
    
    If m.eTbType = eTSVTb_TbType_VolAtPrice Then
        strBox = strBox & Str(iViewOrTickR)
    Else
        strBox = strBox & Str(m.iReverseTick)
    End If
    
    strBox = strBox & "|" & Str(iBoxFill)
    
    m.aBoxTool.Add strBox
    
End Sub

Public Sub RemoveBox(ByVal iBoxId&)

    Dim i&, ID&
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view
    
    For i = 0 To m.aBoxTool.Size - 1
        aFields.SplitFields m.aBoxTool(i), "|"
        ID = ValOfText(aFields(0))
        If ID = iBoxId Then
            aFields(0) = -1
            m.aBoxTool(i) = aFields.JoinFields("|")
        End If
    Next

End Sub

Public Property Get BoxCount() As Long
    BoxCount = m.aBoxTool.Size
End Property

Public Function BoxId(ByVal iRow&, ByVal iCol&) As Long

    Dim iTop&, iLeft&, iBottom&, iRight&, i&
    Dim bInRange As Boolean
    Dim aFields As New cGdArray
    
'format: id|price|column|price|column|color|pixel|view
    
    For i = 0 To m.aBoxTool.Size - 1
        bInRange = False
        
        aFields.SplitFields m.aBoxTool(i), "|"
        
        If aFields(0) > 0 Then
            iTop = aFields(1)
            iLeft = aFields(2)
            iBottom = aFields(3)
            iRight = aFields(4)
        
            If iBottom > iTop Then
                If iRow <= iBottom And iRow >= iTop Then bInRange = True
            Else
                If iRow >= iBottom And iRow <= iTop Then bInRange = True
            End If
        
            If bInRange Then
                bInRange = False
                If iRight > iLeft Then
                    If iCol <= iRight And iCol >= iLeft Then bInRange = True
                Else
                    If iCol >= iRight And iCol <= iLeft Then bInRange = True
                End If
            End If
        
            If bInRange Then
                BoxId = aFields(0)
                Exit For
            End If
        End If
    Next
    
End Function

Public Sub HighLow(dHigh#, dLow#)
    
    If m.TickBars Is Nothing Then
        dHigh = kNullData
        dLow = kNullData
    Else
        dHigh = RoundToMinMove(gdMaxValue(m.TickBars.ArrayHandle(eBARS_Close), 0, m.TickBars.Size), m.dMinMove)
        dLow = RoundToMinMove(gdMinValue(m.TickBars.ArrayHandle(eBARS_Close), 0, m.TickBars.Size), m.dMinMove)
    End If

End Sub

Public Property Get BlankRows() As Long
    BlankRows = m.iBlankRows
End Property

Public Property Let BlankRows(ByVal iRows&)

    If m.iBlankRows = -1 Then
        m.iBlankRows = iRows
    ElseIf iRows <> m.iBlankRows And m.iBlankRows >= 0 Then
        m.iBlankRows = iRows
        BuildTables m.eTbType
    End If

End Property

Public Property Get TotalTradesAtBid() As Long
    TotalTradesAtBid = m.nTradesAtBid
End Property

Public Property Get TotalTradesAtAsk() As Long
    TotalTradesAtAsk = m.nTradesAtAsk
End Property

Public Property Get PriceAtSwitch() As Double
    PriceAtSwitch = m.dPriceOfLastSwitch
End Property

Public Property Get TbRowForPrice(ByVal dPrice#)
On Error GoTo ErrSection:

    Dim iIndex&, iSize&, aHandle&
    Dim dSearchVal#
    
    aHandle = m.tbPriceVol.FieldArrayHandle(0)
    iSize = gdGetSize(aHandle)
    dSearchVal = RoundToMinMove(dPrice, m.dMinMove)
    
    gdBinarySearch aHandle, dSearchVal, iIndex, eGdSort_Descending, 0, iSize
    
    TbRowForPrice = iIndex
    
    Exit Property
    
ErrSection:
    RaiseError "cTSVData.TbRowForPrice"
End Property

Public Property Get IconCount() As Long
On Error GoTo ErrSection:
    
    If Not m.aIconTool Is Nothing Then
        IconCount = m.aIconTool.Size
    End If
    
    Exit Property
    
ErrSection:
    RaiseError "cTSVData.IconCount"

End Property

Public Property Get IconString(ByVal iIndex&)
On Error GoTo ErrSection:

    If iIndex >= 0 And iIndex < m.aIconTool.Size Then
        IconString = m.aIconTool(iIndex)
    End If
    
    Exit Property

ErrSection:
    RaiseError "cTSVData.IconString"

End Property


Public Sub AddIcon(ByVal strIconInfo$)
On Error GoTo ErrSection:

    'format: ICON|key|icon|price|column|color|alignment|session (key = icon type concat with number)
    m.aIconTool.Add strIconInfo & "|" & Str(m.nSessionDate)
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.AddIcon"

End Sub

Public Sub RemoveIcon(ByVal strKey$)
On Error GoTo ErrSection:
    
    Dim i&
    
    'format: ICON|key|icon|price|column|color|alignment (key = icon type concat with number)
    For i = 0 To m.aIconTool.Size - 1
        If strKey = Parse(m.aIconTool(i), "|", 2) Then
            m.aIconTool.Remove i
            Exit For
        End If
    Next
    
    Exit Sub
    
ErrSection:
    RaiseError "cTSVData.RemoveIcon"

End Sub

Public Property Get PriceCumVAM() As Double
    PriceCumVAM = m.dPriceCumVAM            'used only for bid/ask directional form
End Property

Public Property Get PriceCumVBM() As Double
    PriceCumVBM = m.dPriceCumVBM            'used only for bid/ask directional form
End Property
