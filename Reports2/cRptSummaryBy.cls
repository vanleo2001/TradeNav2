VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptSummaryBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    dMinProfit As Currency
    lSubHeadRow As Long
    strReportName As String
    strStatToChart As String
    strShowCents As String
    strSubClass As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary
    Results As cGdTable
    lRecord As Long
    
    strDefaultFields As String
    strFields As String
    lAppID As Long
    bAutoSize As Boolean
    
    hRuleNames As Long
    hSignalType As Long
    hTradeDate As Long
    hSkip As Long
    hSkipRpt As Long
    hRuleID As Long
    hSignalIndex As Long
    hSysNbr As Long
    hEntryExitPtr As Long
    hAccountBalance As Long
    hSorted As Long
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_SysNbr = 0
    eGDCol_EntrySignal
    eGDCol_ExitSignal
    eGDCol_CustomText
    eGDCol_TotTrades
    eGDCol_WinPct
    eGDCol_WinAvg
    eGDCol_LossAvg
    eGDCol_AvgRunUp
    eGDCol_AvgRunDown
    eGDCol_PLRatio
    eGDCol_PF
    eGDCol_AvgTrade
    eGDCol_LossConsec
    eGDCol_LossLargest
    eGDCol_Profit
    eGDCol_NetProfit
    eGDCol_AvgNumBars
    eGDCol_ROI
    eGDCol_NumCols
End Enum

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

Public Sub BeforeScrollTip(ByVal Row As Long)
End Sub
Public Sub AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Resize
'' Description: Upon resizing the form, extend the custom column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Resize()
On Error Resume Next

    'm.RptForm.ExtendCustomColumn m.vsGrid

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayValues
'' Description: When the user moves the mouse, display the appropriate value
'' Inputs:      X and Y location of the mouse, Label to display values in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayValues(ByVal X As Double, ByVal Y As Double, lbl As Object)
On Error GoTo ErrSection:

    With lbl
        .Caption = ""
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptSummaryBy.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Run
'' Description: Initialize and run the report
'' Inputs:      Form to show report on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Run(RptForm As Object)
On Error GoTo ErrSection:
    
    Set m.RptForm = RptForm
    Set m.Chart = RptForm!BarChart
    Set m.vsGrid = RptForm!vsReport
    
    m.strSubClass = m.RptForm.OptionValue("SubClass")
    m.strDefaultFields = m.RptForm.OptionValue("DefaultFields")
    m.lAppID = m.RptForm.ReportID
    m.hRuleNames = m.RptForm.Trades.RuleNameHandle
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hSignalType = .FieldHandle(entd_SignalType)
        Select Case UCase(m.strSubClass)
            Case "TDOW"
                m.hTradeDate = .FieldHandle(entd_TradeDayOfWeek)
                'm.RptForm.CustomColumn = GDCol(eGDCol_CustomText)
            Case "TDOM"
                m.hTradeDate = .FieldHandle(entd_TradeDayOfMonth)
                'm.RptForm.CustomColumn = GDCol(eGDCol_CustomText)
            Case "TDOY"
                m.hTradeDate = .FieldHandle(entd_TradeDayOfYear)
                'm.RptForm.CustomColumn = GDCol(eGDCol_CustomText)
            Case "ENTRY SIGNAL", "ENTRY/EXIT SIGNAL"
                'm.RptForm.CustomColumn = GDCol(eGDCol_EntrySignal)
            Case "EXIT SIGNAL"
                'm.RptForm.CustomColumn = GDCol(eGDCol_ExitSignal)
            Case Else
                m.hTradeDate = .FieldHandle(entd_TradeDate)
                'm.RptForm.CustomColumn = GDCol(eGDCol_CustomText)
        End Select
        m.hSkip = .FieldHandle(entd_Skip)
        m.hSkipRpt = .FieldHandle(entd_SkipRpt)
        m.hRuleID = .FieldHandle(entd_RuleID)
        m.hSignalIndex = .FieldHandle(entd_SignalIndex)
        m.hSysNbr = .FieldHandle(entd_SysNbr)
        m.hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hSorted = .SortHandle
    End With
    
    m.RptForm.SetToSingleUnit False
    m.RptForm.FilterTrades
    
    'Report options
    m.dMinProfit = m.RptForm.OptionValue("ProfitGT")
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strStatToChart = m.RptForm.OptionValue("StatToChart")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    Select Case UCase(m.strSubClass)
        Case "DAY"
            CalculateStatsByDay
        Case "WEEK"
            If FileExist(AddSlash(g.strAppPath) & "CWeek.FLG") Then
                CalculateStatsByCalendarWeek
            Else
                CalculateStatsByTradingWeek
            End If
        Case "WEEK DAY"
            CalculateStatsByWeekDay
        Case "MONTH"
            CalculateStatsByMonth
        Case "MONTHLY"
            'CalculateStatsByMonthYear
            CalculateStatsOverHistory "M"
        Case "WEEKLY"
            CalculateStatsOverHistory "W"
        Case "YEAR"
            'CalculateStatsByYear
            CalculateStatsOverHistory "Y"
        Case "TDOW"
            CalculateStatsByTDOW
        Case "TDOM"
            CalculateStatsByTDOM
        Case "TDOY"
            CalculateStatsByTDOY
        Case "ENTRY SIGNAL"
            CalculateStatsByEntry
        Case "EXIT SIGNAL"
            CalculateStatsByExit
        Case "ENTRY/EXIT SIGNAL"
            CalculateStatsByEntryExit
    End Select
    ShowReport
    BuildChart
    m.vsGrid.Redraw = flexRDBuffered
           
    If UCase(m.strSubClass) = "MONTHLY" Or UCase(m.strSubClass) = "WEEKLY" Then
        If m.vsGrid.Rows > m.vsGrid.FixedRows Then
            m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.Run", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildChart
'' Description: Build the chart
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildChart()
On Error GoTo ErrSection:
    
    Dim lItemsToChart As Long           ' Number of items to chart
    Dim lTrade As Long                  ' Index into a for loop
    Dim YData() As Single               ' Array for the Y-Axis of the chart
    Dim lPoints As Long                 ' Number of points on the chart
    Dim dBarWidth As Double             ' Number of points per bar
    Dim strSignal As String             ' Signal name
    
    ' Calculate Items to chart (number of rows in grid)
    lItemsToChart = m.vsGrid.Rows - 1
    ShowChart lItemsToChart, m.Chart, m.RptForm.lblNoTrades, m.RptForm.lblNoMMTrades
    If lItemsToChart = 0 Then Exit Sub
    
    Select Case UCase(m.strSubClass)
        Case "DAY"
            lPoints = 31
        Case "WEEK"
            lPoints = 4
        Case "MONTH"
            lPoints = 12
        Case "MONTHLY", "WEEKLY"
            lPoints = lItemsToChart
        Case "YEAR"
            lPoints = lItemsToChart
        Case "TDOW", "WEEK DAY"
            lPoints = 5
        Case "TDOM"
            lPoints = 23
        Case "TDOY"
            lPoints = 255
        Case "ENTRY SIGNAL", "EXIT SIGNAL", "ENTRY/EXIT SIGNAL"
            lPoints = lItemsToChart
    End Select
    
    '** Set Subsets, Points, and fill with YData **'
    With m.Chart
        .PEactions = 20     'Empty chart and reinitialize
        
        .MainTitle = m.strReportName
        .SubTitle = "by " & m.strStatToChart
        .YAxisLabel = m.strStatToChart
        .AxisFormatY = "|,|"
        .XAxisLabel = ""
        
        Select Case UCase(m.strSubClass)
            Case "ENTRY SIGNAL", "EXIT SIGNAL", "ENTRY/EXIT SIGNAL"
                SetChartSettings m.Chart, GPM_HORIZONTALBAR
            Case Else
                SetChartSettings m.Chart, GPM_BAR
        End Select
            
        .Subsets = 1
        .Points = lPoints
        If UCase(m.strSubClass) = "TDOY" Then
            dBarWidth = .BarWidth
            .BarWidth = 1
            .AltFrequencies(0) = 25
        ElseIf UCase(m.strSubClass) = "MONTHLY" Then
            .FirstPtLabelOffset = 1
            .AltFrequencies(0) = 12
        Else
            .AltFrequencies(0) = 1
        End If
        
        ReDim YData(lPoints) As Single
        
        Select Case m.strStatToChart
            Case "Trades"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_TotTrades), lTrade + 1)
                Next lTrade
            Case "Win Pct"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_WinPct), lTrade + 1)
                Next lTrade
            Case "PL Ratio"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_PLRatio), lTrade + 1)
                Next lTrade
            Case "Profit Factor"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_PF), lTrade + 1)
                Next lTrade
            Case "Profit"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_Profit), lTrade + 1)
                Next lTrade
            Case "Win Avg"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_WinAvg), lTrade + 1)
                Next lTrade
            Case "Loss Avg"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_LossAvg), lTrade + 1)
                Next lTrade
            Case "Avg Trade"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_AvgTrade), lTrade + 1)
                Next lTrade
            Case "Avg Run Up"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_AvgRunUp), lTrade + 1)
                Next lTrade
            Case "Avg Run Down"
                For lTrade = 0 To lPoints - 1
                    YData(lTrade) = m.Results.Num(GDCol(eGDCol_AvgRunDown), lTrade + 1)
                Next lTrade
        End Select
        
        PEvset .hObject, PEP_faYDATA, YData(0), lPoints
        
        Select Case UCase(m.strSubClass)
            Case "WEEK"
                .PointLabels(0) = "Week 1"
                .PointLabels(1) = "Week 2"
                .PointLabels(2) = "Week 3"
                .PointLabels(3) = "Week 4"
            Case "MONTH"
                .PointLabels(0) = "Jan"
                .PointLabels(1) = "Feb"
                .PointLabels(2) = "Mar"
                .PointLabels(3) = "Apr"
                .PointLabels(4) = "May"
                .PointLabels(5) = "Jun"
                .PointLabels(6) = "Jul"
                .PointLabels(7) = "Aug"
                .PointLabels(8) = "Sep"
                .PointLabels(9) = "Oct"
                .PointLabels(10) = "Nov"
                .PointLabels(11) = "Dec"
            Case "MONTHLY"
                For lTrade = 1 To m.Results.NumRecords - 1
                    .PointLabels(lTrade - 1) = Str(Year((m.Results.Item(GDCol(eGDCol_CustomText), lTrade))))
                Next lTrade
            Case "YEAR"
                For lTrade = 1 To m.Results.NumRecords - 1
                    .PointLabels(lTrade - 1) = m.Results.Item(GDCol(eGDCol_CustomText), lTrade)
                Next lTrade
            Case "WEEK DAY"
                .PointLabels(0) = "Monday"
                .PointLabels(1) = "Tuesday"
                .PointLabels(2) = "Wednesday"
                .PointLabels(3) = "Thursday"
                .PointLabels(4) = "Friday"
            Case "ENTRY SIGNAL"
                For lTrade = 0 To lItemsToChart - 1
                    If UCase(m.RptForm.AppName) = "SYSTEM NAVIGATOR" Then
                        strSignal = gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_EntrySignal), lTrade + 1))
                        If Len(strSignal) > 48 Then
                            .PointLabels(lTrade) = Left(strSignal, 48)
                        Else
                            .PointLabels(lTrade) = strSignal
                        End If
                    Else
                        .PointLabels(lTrade) = Str(m.Results.Num(GDCol(eGDCol_SysNbr), lTrade + 1)) & " - " & _
                            gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_EntrySignal), lTrade + 1))
                    End If
                Next lTrade
                .PointsToGraph = 15
            Case "EXIT SIGNAL"
                For lTrade = 0 To lItemsToChart - 1
                    If UCase(m.RptForm.AppName) = "SYSTEM NAVIGATOR" Then
                        strSignal = gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_ExitSignal), lTrade + 1))
                        If Len(strSignal) > 48 Then
                            .PointLabels(lTrade) = Left(strSignal, 48)
                        Else
                            .PointLabels(lTrade) = strSignal
                        End If
                    Else
                        .PointLabels(lTrade) = Str(m.Results.Num(GDCol(eGDCol_SysNbr), lTrade + 1)) & " - " & _
                            gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_ExitSignal), lTrade + 1))
                    End If
                Next lTrade
                .PointsToGraph = 15
            Case "ENTRY/EXIT SIGNAL"
                For lTrade = 0 To lItemsToChart - 1
                    strSignal = Left(gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_EntrySignal), lTrade + 1)), 24) & " / " & Left(gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_ExitSignal), lTrade + 1)), 24)
                    If Len(strSignal) > 48 Then
                        .PointLabels(lTrade) = Left(strSignal, 48)
                    Else
                        .PointLabels(lTrade) = strSignal
                    End If
                Next lTrade
                .PointsToGraph = 15
        End Select
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
        
        If UCase(m.strSubClass) = "TDOY" Then .BarWidth = dBarWidth
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the Grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitGrid()
On Error GoTo ErrSection:
    
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    'Format grid columns and headings...
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        ClearGrid m.vsGrid
        
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = g.nAltGridRowColor '&HC0FFFF
        .AutoSizeMode = flexAutoSizeColWidth
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        Select Case UCase(m.strSubClass)
            Case "ENTRY SIGNAL", "EXIT SIGNAL"
                .FrozenCols = 2
            Case "ENTRY/EXIT SIGNAL"
                .FrozenCols = 3
            Case Else
                .FrozenCols = 1
        End Select
        m.lSubHeadRow = 0
        
        '1st column can be customized
        .ColWidth(GDCol(eGDCol_SysNbr)) = 500
        .ColAlignment(GDCol(eGDCol_SysNbr)) = flexAlignLeftBottom
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SysNbr)) = "Sys#"
        .ColHidden(GDCol(eGDCol_SysNbr)) = (UCase(m.RptForm.AppName) = "SYSTEM NAVIGATOR")
        
        .ColWidth(GDCol(eGDCol_EntrySignal)) = 1500
        .ColAlignment(GDCol(eGDCol_EntrySignal)) = flexAlignLeftBottom
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntrySignal)) = "Entry Signal"
        .ColHidden(GDCol(eGDCol_EntrySignal)) = (InStr(UCase(m.strSubClass), "ENTRY") = 0)
        
        .ColWidth(GDCol(eGDCol_ExitSignal)) = 1500
        .ColAlignment(GDCol(eGDCol_ExitSignal)) = flexAlignLeftBottom
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ExitSignal)) = "Exit Signal"
        .ColHidden(GDCol(eGDCol_ExitSignal)) = (InStr(UCase(m.strSubClass), "EXIT") = 0)
        
        .ColWidth(GDCol(eGDCol_CustomText)) = 1000
        .ColAlignment(GDCol(eGDCol_CustomText)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_CustomText)) = m.strSubClass
        .ColHidden(GDCol(eGDCol_CustomText)) = InStr(UCase(m.strSubClass), "SIGNAL")
        
        .ColAlignment(GDCol(eGDCol_TotTrades)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotTrades)) = "Trades"
        
        .ColAlignment(GDCol(eGDCol_WinPct)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WinPct)) = "Win Pct"
        
        .ColAlignment(GDCol(eGDCol_WinAvg)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WinAvg)) = "Win Avg"
        
        .ColAlignment(GDCol(eGDCol_LossAvg)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LossAvg)) = "Loss Avg"
        
        .ColAlignment(GDCol(eGDCol_PLRatio)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_PLRatio)) = "PayoutR" '"P/L"
        
        .ColAlignment(GDCol(eGDCol_PF)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_PF)) = "ProfitF" '"P/F"
        
        .ColAlignment(GDCol(eGDCol_AvgTrade)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AvgTrade)) = "AvgTrade"
        
        .ColAlignment(GDCol(eGDCol_AvgRunUp)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AvgRunUp)) = "Run Up"
        
        .ColAlignment(GDCol(eGDCol_AvgRunDown)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AvgRunDown)) = "Run Dn"
                
        .ColAlignment(GDCol(eGDCol_Profit)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Profit)) = "Profit"
        
        .ColAlignment(GDCol(eGDCol_NetProfit)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfit)) = "Net Profit"
        
        .ColAlignment(GDCol(eGDCol_LossConsec)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LossConsec)) = "C/L"
        
        .ColAlignment(GDCol(eGDCol_LossLargest)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LossLargest)) = "Max Loss"
        
        .ColAlignment(GDCol(eGDCol_AvgNumBars)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AvgNumBars)) = "Avg Bars"
        
        .ColAlignment(GDCol(eGDCol_ROI)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ROI)) = "ROI%"
                
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        'm.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        'm.RptForm.ExtendCustomColumn m.vsGrid
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByDay
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByDay()
On Error GoTo ErrSection:
    
    Dim lDay As Long                    ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        For lDay = 1 To 31
            
            'Filter Trades by entry signal...
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If Day(CDate(gdGetNum(m.hTradeDate, lIndex))) = lDay Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults "Day " & Format(lDay, "00")
        Next lDay
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByDay", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByCalendarWeek
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByCalendarWeek()
On Error GoTo ErrSection:
    
    Dim lWeek As Long                   ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lEntryDay As Long               ' Day of the entry signal
    Dim lFirstDay As Long               ' First day of the week
    Dim lLastDay As Long                ' Last day of the week
    
    ' Week1 = 1st of month through 7th of month
    ' Week2 = 8th of month through 14th of month
    ' Week3 = 15th of month through 21st of month
    ' Week4 = 22nd of month through 31st of month
    
    ' Set Properties for loading trades to analyze
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        For lWeek = 1 To 4
            lLastDay = lWeek * 7
            lFirstDay = lLastDay - 6
            If lLastDay = 28 Then lLastDay = 31
            
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    lEntryDay = Day(gdGetNum(m.hTradeDate, lIndex))
                    If lEntryDay >= lFirstDay And lEntryDay <= lLastDay Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults "Week " & Str(lWeek)
        Next lWeek
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByCalendarWeek", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByTradingWeek
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByTradingWeek()
On Error GoTo ErrSection:
    
    Dim lWeek As Long                   ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lEntryDay As Long               ' Day of the entry signal
    Dim lFirstDay As Long               ' First day of the week
    Dim lLastDay As Long                ' Last day of the week
    Dim hTDOM As Long                   ' Handle to the TDOM array
    
    ' Week1 = 1st TDOM through 5th TDOM
    ' Week2 = 6th TDOM through 10th TDOM
    ' Week3 = 11th TDOM through 15th TDOM
    ' Week4 = 16th TDOM through 23rd TDOM
    
    ' Set Properties for loading trades to analyze
    With m.Stats
        hTDOM = m.RptForm.Trades.FieldHandle(entd_TradeDayOfMonth)
        
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        For lWeek = 1 To 4
            lLastDay = lWeek * 5
            lFirstDay = lLastDay - 4
            If lLastDay = 20 Then lLastDay = 23
            
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    lEntryDay = gdGetNum(hTDOM, lIndex)
                    If lEntryDay >= lFirstDay And lEntryDay <= lLastDay Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults "Week " & Str(lWeek)
        Next lWeek
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByTradingWeek", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByWeekDay
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByWeekDay()
On Error GoTo ErrSection:
    
    Dim lDay As Long                    ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lSize As Long                   ' Number of records in the trades
    
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        lSize = m.RptForm.Trades.NumRecords
        For lDay = 2 To 6
            For lIndex2 = 1 To lSize
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If Weekday(CDate(gdGetNum(m.hTradeDate, lIndex))) = lDay Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults Str(DateOf(20030600 + lDay))
        Next lDay
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByWeekDay", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByMonth
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByMonth()
On Error GoTo ErrSection:
    
    Dim lMonth As Long                  ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop

    ' Set properties for loading trades to analyze
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        For lMonth = 1 To 12
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If Month(gdGetNum(m.hTradeDate, lIndex)) = lMonth Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults Str(DateOf("1980" & Format(lMonth, "00") & "01"))
        Next lMonth
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByMonth", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByYear
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByYear()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dBeginDate As Date
    Dim dEndDate As Date
    Dim dFromDate As Date
    Dim dFromDate2 As Date
    Dim dToDate As Date
    Dim dToDate2 As Date
    Dim bFirst As Boolean
    Dim dNetProfit As Double
    Dim dROI As Double
    Dim dAccountReq As Double
    Dim lExit As Long
    
    ' Figure out the beginning and ending dates...
    For lIndex2 = 1 To gdGetSize(m.hTradeDate) - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
    
        'If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            If gdGetNum(m.hSkip, lIndex) = 0 Then
                If dBeginDate = 0 Then dBeginDate = gdGetNum(m.hTradeDate, lIndex)
                dEndDate = gdGetNum(m.hTradeDate, lIndex)
            End If
        'End If
    Next lIndex2
    
    ' Set properties for loading trades to analyze
    bFirst = True
    dNetProfit = 0#
    
    ' Calculate stats with all of the non-ignored trades so that we can get the
    ' account size required to use as a divisor...
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        .Calculate False, False, True, False
        dAccountReq = .AccountSizeReq
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With

    ' Run statistics for each year in the system
    'dBeginDate = gdGetNum(m.hTradeDate, 1)
    'dEndDate = gdGetNum(m.hTradeDate, m.RptForm.Trades.NumRecords - 1)
    dFromDate = DateValue("1/1/" & Str(Year(dBeginDate)))
    dToDate = DateValue("12/31/" & Str(Year(dBeginDate)))
    Do While dFromDate < dEndDate
    
        ' Override for exact beginning/ending dates from prompt
        dFromDate2 = dFromDate
        dToDate2 = dToDate
        If dFromDate < dBeginDate Then
            dFromDate2 = dBeginDate
        End If
        If dToDate > dEndDate Then
            dToDate2 = dEndDate
        End If
        If InStr(1, dToDate2, ".") <= 0 Then
            dToDate2 = dToDate + 0.9999
        End If
        
        'Process all trades for current year
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lIndex = gdGetNum(m.hSorted, lIndex2)
            lExit = gdGetNum(m.hEntryExitPtr, lIndex)
            
            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                If CDate(gdGetNum(m.hTradeDate, lExit)) >= dFromDate2 And _
                   CDate(gdGetNum(m.hTradeDate, lExit)) <= dToDate2 Then
                    gdSetNum m.hSkipRpt, lIndex, 0
                Else
                    gdSetNum m.hSkipRpt, lIndex, 1
                End If
            End If
        Next lIndex2
        
        If Not bFirst Then
            gdSetNum m.hAccountBalance, 0, m.Stats.EndingBalance
        End If
        bFirst = False
        m.Stats.SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        m.Stats.SetSortHandle m.RptForm.Trades.SortHandle
        
        ' Now do a calculate using the skip report filter and the equity filter...
        m.Stats.Calculate False, False, False, True
        
        dNetProfit = dNetProfit + m.Stats.TotalNetProfit
        If dAccountReq = 0 Then
            dROI = 99999
        Else
            dROI = m.Stats.TotalNetProfit / dAccountReq
        End If
        SaveResults Year(dFromDate), , dNetProfit, dROI
        
        dFromDate = DateAdd("yyyy", 1, dFromDate)     'add one year to from date
        dToDate = DateValue("12/31/" & Str(Year(dFromDate)))
    Loop

    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByYear", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByMonthYear
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByMonthYear()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dBeginDate As Date
    Dim dEndDate As Date
    Dim dFromDate As Date
    Dim dFromDate2 As Date
    Dim dToDate As Date
    Dim dToDate2 As Date
    Dim bFirst As Boolean
    Dim lMonth As Long
    Dim dNetProfit As Double
    Dim dROI As Double
    Dim dAccountReq As Double
    Dim lExit As Long
    
    ' Figure out the beginning and ending dates...
    For lIndex2 = 1 To gdGetSize(m.hTradeDate) - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        lExit = gdGetNum(m.hEntryExitPtr, lIndex)
        
        If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            If gdGetNum(m.hSkip, lIndex) = 0 Then
                If dBeginDate = 0 Then dBeginDate = gdGetNum(m.hTradeDate, lExit)
                dEndDate = gdGetNum(m.hTradeDate, lExit)
            End If
        End If
    Next lIndex2
    
    ' Set properties for loading trades to analyze
    bFirst = True
    dNetProfit = 0#
    
    ' Calculate stats with all of the non-ignored trades so that we can get the
    ' account size required to use as a divisor...
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        .Calculate False, False, True, False
        dAccountReq = .AccountSizeReq
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With

    ' Run statistics for each year in the system
    'dBeginDate = gdGetNum(m.hTradeDate, 1)
    'dEndDate = gdGetNum(m.hTradeDate, m.RptForm.Trades.NumRecords - 1)
    dFromDate = DateValue("1/1/" & Str(Year(dBeginDate)))
    dToDate = DateValue("12/31/" & Str(Year(dBeginDate)))
    Do While dFromDate < dEndDate
    
        ' Override for exact beginning/ending dates from prompt
        dFromDate2 = dFromDate
        dToDate2 = dToDate
        If dFromDate < dBeginDate Then
            dFromDate2 = dBeginDate
        End If
        If dToDate > dEndDate Then
            dToDate2 = dEndDate
        End If
        If InStr(1, dToDate2, ".") <= 0 Then
            dToDate2 = dToDate + 0.9999
        End If
    
        For lMonth = 1 To 12
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                lIndex = gdGetNum(m.hSorted, lIndex2)
                lExit = gdGetNum(m.hEntryExitPtr, lIndex)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If CDate(gdGetNum(m.hTradeDate, lExit)) >= dFromDate2 And _
                       CDate(gdGetNum(m.hTradeDate, lExit)) <= dToDate2 And _
                       Month(gdGetNum(m.hTradeDate, lExit)) = lMonth Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
        
            If Not bFirst Then
                gdSetNum m.hAccountBalance, 0, m.Stats.EndingBalance
            End If
            bFirst = False
            m.Stats.SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            m.Stats.SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            m.Stats.Calculate False, False, False, True
            
            dNetProfit = dNetProfit + m.Stats.TotalNetProfit
            If dAccountReq = 0 Then
                dROI = 99999
            Else
                dROI = m.Stats.TotalNetProfit / dAccountReq
            End If
            SaveResults Str(DateOf((Year(dFromDate) * 10000) + (lMonth * 100) + 1)), , dNetProfit, dROI
        Next lMonth
        
        dFromDate = DateAdd("yyyy", 1, dFromDate)     'add one year to from date
        dToDate = DateValue("12/31/" & Str(Year(dFromDate)))
    Loop

    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByMonthYear", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsOverHistory
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsOverHistory(ByVal strMode$)
On Error GoTo ErrSection:
    
    Dim i&
    Dim lIndex As Long
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dBeginDate As Date
    Dim dEndDate As Date
    Dim dFromDate As Date
    Dim dFromDate2 As Date
    Dim dToDate As Date
    Dim dToDate2 As Date
    Dim bFirst As Boolean
    Dim lMonth As Long
    Dim dNetProfit As Double
    Dim dROI As Double
    Dim dAccountReq As Double
    Dim lExit As Long
    
    ' Mode = Yearly, Monthly, or Weekly
    strMode = UCase(Left(strMode, 1)) ' Y/M/W
    
    ' Figure out the beginning and ending dates...
    For lIndex2 = 1 To gdGetSize(m.hTradeDate) - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        lExit = gdGetNum(m.hEntryExitPtr, lIndex)
        
        If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            If gdGetNum(m.hSkip, lIndex) = 0 Then
                If dBeginDate = 0 Then dBeginDate = gdGetNum(m.hTradeDate, lExit)
                dEndDate = gdGetNum(m.hTradeDate, lExit)
            End If
        End If
    Next lIndex2
    
    ' Set properties for loading trades to analyze
    bFirst = True
    dNetProfit = 0#
    
    ' Calculate stats with all of the non-ignored trades so that we can get the
    ' account size required to use as a divisor...
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        .Calculate False, False, True, False
        dAccountReq = .AccountSizeReq
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    dFromDate = dBeginDate
    Do While dFromDate <= dEndDate
        ' bump dToDate up until last day of period
        dToDate = Int(dFromDate)
        Select Case strMode
        Case "Y" ' Yearly
            i = Year(dFromDate)
            Do While Year(dToDate + 1) = i
                dToDate = dToDate + 1
            Loop
        Case "M" ' Monthly
            i = Month(dFromDate)
            Do While Month(dToDate + 1) = i
                dToDate = dToDate + 1
            Loop
        Case "W" ' Weekly
            i = WkNum(dFromDate)
            Do While WkNum(dToDate + 1) = i
                dToDate = dToDate + 1
            Loop
        End Select
        
        ' filter for just the dates within this period
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lIndex = gdGetNum(m.hSorted, lIndex2)
            lExit = gdGetNum(m.hEntryExitPtr, lIndex)
            
            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                If CDate(gdGetNum(m.hTradeDate, lExit)) >= dFromDate And _
                   CDate(gdGetNum(m.hTradeDate, lExit)) < dToDate + 0.999999 Then
                    gdSetNum m.hSkipRpt, lIndex, 0
                Else
                    gdSetNum m.hSkipRpt, lIndex, 1
                End If
            End If
        Next lIndex2
    
        If Not bFirst Then
            gdSetNum m.hAccountBalance, 0, m.Stats.EndingBalance
        End If
        bFirst = False
        m.Stats.SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        m.Stats.SetSortHandle m.RptForm.Trades.SortHandle
        
        ' Now do a calculate using the skip report filter and the equity filter...
        m.Stats.Calculate False, False, False, True
        
        dNetProfit = dNetProfit + m.Stats.TotalNetProfit
        If dAccountReq = 0 Then
            dROI = 99999
        Else
            dROI = m.Stats.TotalNetProfit / dAccountReq
        End If
        If strMode = "Y" Then
            SaveResults Year(dFromDate), , dNetProfit, dROI
        Else
            SaveResults Str(dToDate), , dNetProfit, dROI
        End If
        
        ' increment to the next period
        dFromDate = dToDate + 1
    Loop
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsOverHistory", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByTDOW
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByTDOW()
On Error GoTo ErrSection:
    
    Dim lTDOW As Long                   ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lSize As Long                   ' Number of records in the trades
    
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        For lTDOW = 1 To 5
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If gdGetNum(m.hTradeDate, lIndex) = lTDOW Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults "Day " & Format(lTDOW, "00")
        Next lTDOW
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByTDOW", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByTDOM
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByTDOM()
On Error GoTo ErrSection:
    
    Dim lTDOM As Long                   ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lSize As Long                   ' Number of records in the trades
    
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        For lTDOM = 1 To 23
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If gdGetNum(m.hTradeDate, lIndex) = lTDOM Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults "Day " & Format(lTDOM, "00")
        Next lTDOM
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByTDOM", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByTDOY
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByTDOY()
On Error GoTo ErrSection:
    
    Dim lTDOY As Long                   ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lSize As Long                   ' Number of records in the trades
    
    With m.Stats
        .MinProfit = m.dMinProfit
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
        
        lSize = m.RptForm.Trades.NumRecords
        For lTDOY = 1 To 260
            For lIndex2 = 1 To lSize
                lIndex = gdGetNum(m.hSorted, lIndex2)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    If gdGetNum(m.hTradeDate, lIndex) = lTDOY Then
                        gdSetNum m.hSkipRpt, lIndex, 0
                    Else
                        gdSetNum m.hSkipRpt, lIndex, 1
                    End If
                End If
            Next lIndex2
            
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
            
            SaveResults "Day " & Format(lTDOY, "000")
        Next lTDOY
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummaryBy.CalculateStatsByTDOY", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByEntry
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByEntry()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim hInd As Long                    ' Handle to a temporary array
    Dim lSystemNum As Long              ' Current System Number
    
    ' Create a unique, sorted index of Entry Name...
    hInd = gdCreateArray(eGDARRAY_Longs, m.RptForm.Trades.NumRecords - 1, USE_DEFAULT_NULL)
    gdSortAsIndex hInd, m.hSignalIndex, 1, eGdSort_DeleteDuplicates, 0, gdGetSize(m.hSignalIndex)

    With m.Stats
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
        
    For lIndex = 1 To gdGetSize(hInd) - 1
        ' Skip all but Entries with the given Entry Signal...
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords
            If gdGetNum(m.hSignalType, lIndex2) = gEntrySignal And gdGetNum(m.hSignalIndex, lIndex2) = gdGetNum(m.hSignalIndex, gdGetNum(hInd, lIndex)) Then
                gdSetNum m.hSkipRpt, lIndex2, 0
            Else
                gdSetNum m.hSkipRpt, lIndex2, 1
            End If
        Next lIndex2
        
        ' Calculate statistics...
        With m.Stats
            .MinProfit = m.dMinProfit
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
        End With
                
        ' Save the results for the given Entry Signal...
        If m.Stats.TotTrades > 0 Then
            lSystemNum = m.RptForm.SystemArrayInd(gdGetNum(m.hSysNbr, gdGetNum(hInd, lIndex)))
            SaveResults gdGetNum(m.hSignalIndex, gdGetNum(hInd, lIndex)), lSystemNum
        End If
    Next lIndex
    
ErrExit:
    gdDestroyArray hInd
    Exit Sub

ErrSection:
    gdDestroyArray hInd
    RaiseError "cRptSummaryBy.CalculateStatsByEntry", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByExit
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByExit()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim hInd As Long                    ' Handle to a temporary array
    Dim lSystemNum As Long              ' Current System Number
    
    ' Create a unique, sorted index of Entry Name...
    hInd = gdCreateArray(eGDARRAY_Longs, m.RptForm.Trades.NumRecords, USE_DEFAULT_NULL)
    gdSortAsIndex hInd, m.hSignalIndex, 1, eGdSort_DeleteDuplicates, 0, gdGetSize(m.hSignalIndex)
    
    With m.Stats
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
        
    For lIndex = 1 To gdGetSize(hInd) - 1
        ' Skip all but Exits with the given Exit Signal and their corresponding Entries...
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords
            If gdGetNum(m.hSignalType, lIndex2) = gEntrySignal And gdGetNum(m.hSignalIndex, gdGetNum(m.hEntryExitPtr, lIndex2)) = gdGetNum(m.hSignalIndex, gdGetNum(hInd, lIndex)) Then
                gdSetNum m.hSkipRpt, lIndex2, 0
            Else
                gdSetNum m.hSkipRpt, lIndex2, 1
            End If
            
            'If gdGetNum(m.hSignalType, lIndex2) = gExitSignal And gdGetNum(m.hSignalIndex, lIndex2) = gdGetNum(m.hSignalIndex, gdGetNum(hInd, lIndex)) Then
            '    gdSetNum m.hSkipRpt, lIndex2, 0
            'ElseIf gdGetNum(m.hSignalType, lIndex2) = gEntrySignal And gdGetNum(m.hSignalIndex, gdGetNum(m.hEntryExitPtr, lIndex2)) = gdGetNum(m.hSignalIndex, gdGetNum(hInd, lIndex)) Then
            '    gdSetNum m.hSkipRpt, lIndex2, 0
            'Else
            '    gdSetNum m.hSkipRpt, lIndex2, 1
            'End If
        Next lIndex2
        
        ' Calculate statistics...
        With m.Stats
            .MinProfit = m.dMinProfit
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
        
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
        End With
                
        ' Save the results for the given Entry Signal...
        If m.Stats.TotTrades > 0 Then
            lSystemNum = m.RptForm.SystemArrayInd(gdGetNum(m.hSysNbr, gdGetNum(hInd, lIndex)))
            SaveResults Str(gdGetNum(m.hSignalIndex, gdGetNum(hInd, lIndex))), lSystemNum
        End If
    Next lIndex
    
ErrExit:
    gdDestroyArray hInd
    Exit Sub
    
ErrSection:
    gdDestroyArray hInd
    RaiseError "cRptSummaryBy.CalculateStatsByExit", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStatsByEntryExit
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStatsByEntryExit()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim hKey As Long                    ' Handle to a temporary array
    Dim hInd As Long                    ' Handle to a temporary array
    Dim strSignal As String             ' Current signal combination
    Dim lSystemNum As Long              ' Current System Number
    
    'Create a sorted index of Entry/Exit combination
    hKey = gdCreateArray(eGDARRAY_Strings, m.RptForm.Trades.NumRecords, USE_DEFAULT_NULL)
    With m.RptForm.Trades
        For lIndex = 1 To .NumRecords
            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                gdSetStr hKey, lIndex, gdGetNum(m.hSignalIndex, lIndex) & vbTab & gdGetNum(m.hSignalIndex, gdGetNum(m.hEntryExitPtr, lIndex))
            End If
        Next lIndex
    End With
    hInd = gdCreateArray(eGDARRAY_Longs, m.RptForm.Trades.NumRecords, USE_DEFAULT_NULL)
    gdSortAsIndex hInd, hKey, 1, eGdSort_DeleteDuplicates, 0, 999999999
    
    With m.Stats
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        ' First, need to do a calculate so that the equity filter can take out
        ' trades as appropriate...
        .Calculate False, False, True, False
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
        
    For lIndex = 1 To gdGetSize(hInd) - 1
        ' Don't Skip trades with the current combination...
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords
            strSignal = Str(gdGetNum(m.hSignalIndex, lIndex2)) & vbTab & Str(gdGetNum(m.hSignalIndex, gdGetNum(m.hEntryExitPtr, lIndex2)))
            If gdGetNum(m.hSignalType, lIndex2) = gEntrySignal And gdGetStr(hKey, gdGetNum(hInd, lIndex)) = strSignal Then
                gdSetNum m.hSkipRpt, lIndex2, 0
            Else
                gdSetNum m.hSkipRpt, lIndex2, 1
            End If
        Next lIndex2
        
        ' Calculate statistics...
        With m.Stats
            .MinProfit = m.dMinProfit
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
        
            ' Now do a calculate using the skip report filter and the equity filter...
            .Calculate False, False, False, True
        End With
        
        If m.Stats.TotTrades > 0 Then
            lSystemNum = m.RptForm.SystemArrayInd(gdGetNum(m.hSysNbr, gdGetNum(hInd, lIndex)))
            SaveResults gdGetStr(hKey, gdGetNum(hInd, lIndex)), lSystemNum
        End If
    Next lIndex
    
ErrExit:
    gdDestroyArray hKey
    gdDestroyArray hInd
    Exit Sub

ErrSection:
    gdDestroyArray hKey
    gdDestroyArray hInd
    RaiseError "cRptSummaryBy.CalculateStatsByEntryExit", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveResults
'' Description: Save the results from the calculation in the table
'' Inputs:      Custom Text
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveResults(ByVal strCustomText As String, Optional ByVal lSysNbr As Long = 0, Optional ByVal dNetProfit As Double, Optional ByVal dROI As Double)
On Error GoTo ErrSection:

    With m.Stats
        m.lRecord = m.lRecord + 1
        Select Case UCase(m.strSubClass)
            Case "ENTRY SIGNAL"
                m.Results.Num(GDCol(eGDCol_SysNbr), m.lRecord) = lSysNbr
                m.Results.Num(GDCol(eGDCol_EntrySignal), m.lRecord) = CLng(strCustomText)
            Case "EXIT SIGNAL"
                m.Results.Num(GDCol(eGDCol_SysNbr), m.lRecord) = lSysNbr
                m.Results.Num(GDCol(eGDCol_ExitSignal), m.lRecord) = CLng(strCustomText)
            Case "ENTRY/EXIT SIGNAL"
                m.Results.Num(GDCol(eGDCol_SysNbr), m.lRecord) = lSysNbr
                m.Results.Num(GDCol(eGDCol_EntrySignal), m.lRecord) = CLng(Parse(strCustomText, vbTab, 1))
                m.Results.Num(GDCol(eGDCol_ExitSignal), m.lRecord) = CLng(Parse(strCustomText, vbTab, 2))
            Case Else
                m.Results.Item(GDCol(eGDCol_CustomText), m.lRecord) = strCustomText
        End Select
        m.Results.Num(GDCol(eGDCol_TotTrades), m.lRecord) = .TotTrades
        m.Results.Num(GDCol(eGDCol_WinPct), m.lRecord) = .WinPct
        m.Results.Num(GDCol(eGDCol_WinAvg), m.lRecord) = .WinAvg
        m.Results.Num(GDCol(eGDCol_LossAvg), m.lRecord) = .LossAvg
        m.Results.Num(GDCol(eGDCol_AvgRunUp), m.lRecord) = .WinRunUpAvg
        m.Results.Num(GDCol(eGDCol_AvgRunDown), m.lRecord) = .LossRunDownAvg
        m.Results.Num(GDCol(eGDCol_PLRatio), m.lRecord) = .PLRatio
        m.Results.Num(GDCol(eGDCol_PF), m.lRecord) = .ProfitFactor
        m.Results.Num(GDCol(eGDCol_AvgTrade), m.lRecord) = .AvgTrade
        m.Results.Num(GDCol(eGDCol_LossConsec), m.lRecord) = .LossConsec
        m.Results.Num(GDCol(eGDCol_LossLargest), m.lRecord) = .LossLargest
        m.Results.Num(GDCol(eGDCol_Profit), m.lRecord) = .TotalNetProfit
        If Not IsMissing(dNetProfit) Then
            m.Results.Num(GDCol(eGDCol_NetProfit), m.lRecord) = dNetProfit
        End If
        m.Results.Num(GDCol(eGDCol_AvgNumBars), m.lRecord) = .AvgBarsInTrade
        If Not IsMissing(dROI) Then
            m.Results.Num(GDCol(eGDCol_ROI), m.lRecord) = dROI
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptSummaryBy.SaveResults", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReport
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowReport()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = .Rows + m.Results.NumRecords - 1
        
        Select Case UCase(m.strSubClass)
            Case "WEEKLY"
                .ColFormat(GDCol(eGDCol_CustomText)) = DateFormat("Format", MM_DD_YYYY)
            Case "MONTHLY"
                .ColFormat(GDCol(eGDCol_CustomText)) = "mmm YYYY"
            Case "MONTH"
                .ColFormat(GDCol(eGDCol_CustomText)) = "mmmm"
            Case "WEEK DAY"
                .ColFormat(GDCol(eGDCol_CustomText)) = "dddd"
            Case Else
                .ColFormat(GDCol(eGDCol_CustomText)) = ""
        End Select
        
        For lIndex = 1 To m.Results.NumRecords - 1
            lRow = m.lSubHeadRow + lIndex
            Select Case UCase(m.strSubClass)
                Case "ENTRY SIGNAL"
                    .TextMatrix(lRow, GDCol(eGDCol_SysNbr)) = Str(m.Results.Num(GDCol(eGDCol_SysNbr), lIndex))
                    .TextMatrix(lRow, GDCol(eGDCol_EntrySignal)) = gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_EntrySignal), lIndex))
                Case "EXIT SIGNAL"
                    .TextMatrix(lRow, GDCol(eGDCol_SysNbr)) = Str(m.Results.Num(GDCol(eGDCol_SysNbr), lIndex))
                    .TextMatrix(lRow, GDCol(eGDCol_ExitSignal)) = gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_ExitSignal), lIndex))
                Case "ENTRY/EXIT SIGNAL"
                    .TextMatrix(lRow, GDCol(eGDCol_SysNbr)) = Str(m.Results.Num(GDCol(eGDCol_SysNbr), lIndex))
                    .TextMatrix(lRow, GDCol(eGDCol_EntrySignal)) = gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_EntrySignal), lIndex))
                    .TextMatrix(lRow, GDCol(eGDCol_ExitSignal)) = gdGetStr(m.hRuleNames, m.Results.Num(GDCol(eGDCol_ExitSignal), lIndex))
                Case "MONTHLY", "MONTH", "WEEK DAY", "WEEKLY"
                    .Cell(flexcpText, lRow, GDCol(eGDCol_CustomText)) = Val(m.Results.Item(GDCol(eGDCol_CustomText), lIndex))
                Case Else
                    .TextMatrix(lRow, GDCol(eGDCol_CustomText)) = m.Results.Item(GDCol(eGDCol_CustomText), lIndex)
            End Select
            .TextMatrix(lRow, GDCol(eGDCol_TotTrades)) = Format(m.Results.Num(GDCol(eGDCol_TotTrades), lIndex), "#,##0")
            .TextMatrix(lRow, GDCol(eGDCol_WinPct)) = Format(m.Results.Num(GDCol(eGDCol_WinPct), lIndex), "##0.00%")
            .TextMatrix(lRow, GDCol(eGDCol_WinAvg)) = Format(m.Results.Num(GDCol(eGDCol_WinAvg), lIndex), FormatDollar(m.strShowCents))
            .TextMatrix(lRow, GDCol(eGDCol_AvgRunUp)) = Format(m.Results.Num(GDCol(eGDCol_AvgRunUp), lIndex), FormatDollar(m.strShowCents))
            .TextMatrix(lRow, GDCol(eGDCol_AvgRunDown)) = Format(m.Results.Num(GDCol(eGDCol_AvgRunDown), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_AvgRunDown), lIndex), lRow, GDCol(eGDCol_AvgRunDown)
            .TextMatrix(lRow, GDCol(eGDCol_LossAvg)) = Format(m.Results.Num(GDCol(eGDCol_LossAvg), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_LossAvg), lIndex), lRow, GDCol(eGDCol_LossAvg)
            .TextMatrix(lRow, GDCol(eGDCol_PLRatio)) = Format(m.Results.Num(GDCol(eGDCol_PLRatio), lIndex), "##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_PF)) = Format(m.Results.Num(GDCol(eGDCol_PF), lIndex), "##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_AvgTrade)) = Format(m.Results.Num(GDCol(eGDCol_AvgTrade), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_AvgTrade), lIndex), lRow, GDCol(eGDCol_AvgTrade)
            .TextMatrix(lRow, GDCol(eGDCol_Profit)) = Format(m.Results.Num(GDCol(eGDCol_Profit), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_Profit), lIndex), lRow, GDCol(eGDCol_Profit)
            .TextMatrix(lRow, GDCol(eGDCol_NetProfit)) = Format(m.Results.Num(GDCol(eGDCol_NetProfit), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_NetProfit), lIndex), lRow, GDCol(eGDCol_NetProfit)
            .TextMatrix(lRow, GDCol(eGDCol_LossConsec)) = Format(m.Results.Num(GDCol(eGDCol_LossConsec), lIndex), "#,##0")
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_LossConsec), lIndex), lRow, GDCol(eGDCol_LossConsec)
            .TextMatrix(lRow, GDCol(eGDCol_LossLargest)) = Format(m.Results.Num(GDCol(eGDCol_LossLargest), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_LossLargest), lIndex), lRow, GDCol(eGDCol_LossLargest)
            .TextMatrix(lRow, GDCol(eGDCol_AvgNumBars)) = Format(m.Results.Num(GDCol(eGDCol_AvgNumBars), lIndex), "#,##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_ROI)) = Format(m.Results.Num(GDCol(eGDCol_ROI), lIndex), "##0.00%")
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_ROI), lIndex), lRow, GDCol(eGDCol_ROI)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        'm.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        'm.RptForm.ExtendCustomColumn m.vsGrid
        'If UCase(m.strSubClass) = "MONTHLY" Then
            SetUpColumns
        'End If
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptSummaryBy.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary
    Set m.Results = New cGdTable
    With m.Results
        .CreateField eGDARRAY_Longs, GDCol(eGDCol_SysNbr), "SystemNumber"
        .CreateField eGDARRAY_Longs, GDCol(eGDCol_EntrySignal), "SystemNumber"
        .CreateField eGDARRAY_Longs, GDCol(eGDCol_ExitSignal), "SystemNumber"
        .CreateField eGDARRAY_Strings, GDCol(eGDCol_CustomText), "CustomText"
        .CreateField eGDARRAY_Longs, GDCol(eGDCol_TotTrades), "TotTrades"
        .CreateField eGDARRAY_Floats, GDCol(eGDCol_WinPct), "WinPct"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_WinAvg), "WinAvg"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_LossAvg), "LossAvg"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_AvgRunUp), "AvgRunUp"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_AvgRunDown), "AvgRunDown"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_PLRatio), "PLRatio"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_PF), "PF"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_AvgTrade), "AvgTrade"
        .CreateField eGDARRAY_Shorts, GDCol(eGDCol_LossConsec), "LossConsec"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_LossLargest), "LossLargest"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_Profit), "Profit"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_NetProfit), "NetProfit"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_AvgNumBars), "AvgNumBars"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_ROI), "ROI"
    End With

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing
    Set m.Results = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeFields
'' Description: Allow the user to change the order of the fields
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeFields()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of field information
    Dim astrUsed As New cGdArray        ' Array of fields to send to add/remove form
    Dim astrAvailable As New cGdArray   ' Array of available fields
    Dim astrDefault As New cGdArray     ' Array of default fields
    Dim astrNew As New cGdArray         ' Final array of column information
    Dim astrHidden As New cGdArray      ' Array of always hidden fields
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strTemp As String               ' Temporary string
    Dim strIniFile As String            ' File name and path of the ini file

    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    
    ' Compile list of default fields...
    astrFields.SplitFields m.strDefaultFields, ","
    For lIndex = 0 To astrFields.Size - 1
        If Parse(astrFields(lIndex), ";", 2) = "0" Then
            astrDefault.Add Parse(astrFields(lIndex), ";", 1)
        End If
    Next lIndex
    
    ' Compile lists of used and available fields...
    astrFields.SplitFields m.strFields, ","
    For lIndex = 0 To astrFields.Size - 1
        Select Case Parse(astrFields(lIndex), ";", 2)
            Case "1"
                astrNew.Add astrFields(lIndex)
            Case "0"
                astrUsed.Add Parse(astrFields(lIndex), ";", 1)
            Case "-1"
                astrAvailable.Add Parse(astrFields(lIndex), ";", 1)
            Case "-2"
                astrHidden.Add Parse(astrFields(lIndex), ";", 1)
        End Select
    Next lIndex
    
    ' Show the add/remove form...
    If frmAddRemove.ShowMe(astrAvailable, astrUsed, eOrderMode_Ordered, astrDefault) Then
        ' If the user hit OK, recompile the fields list and setup the columns...
        For lIndex = 0 To astrUsed.Size - 1
            For lIndex2 = 0 To astrFields.Size - 1
                If Parse(astrFields(lIndex2), ";", 1) = astrUsed(lIndex) Then
                    If Len(Parse(astrFields(lIndex2), ";", 3)) > 0 Then
                        astrNew.Add astrUsed(lIndex) & ";0;" & Parse(astrFields(lIndex2), ";", 3)
                    Else
                        astrNew.Add astrUsed(lIndex) & ";0"
                    End If
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        
        For lIndex = 0 To astrAvailable.Size - 1
            For lIndex2 = 0 To astrFields.Size - 1
                If Parse(astrFields(lIndex2), ";", 1) = astrAvailable(lIndex) Then
                    If Len(Parse(astrFields(lIndex2), ";", 3)) > 0 Then
                        astrNew.Add astrAvailable(lIndex) & ";-1;" & Parse(astrFields(lIndex2), ";", 3)
                    Else
                        astrNew.Add astrAvailable(lIndex) & ";-1"
                    End If
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        
        For lIndex = 0 To astrHidden.Size - 1
            astrNew.Add astrHidden(lIndex) & ";-2"
        Next lIndex
        
        m.strFields = astrNew.JoinFields(",")
        SetIniFileProperty "Fields", m.strFields, Str(m.lAppID), strIniFile
        SetUpColumns
    End If
    
ErrExit:
    Set astrFields = Nothing
    Set astrUsed = Nothing
    Set astrAvailable = Nothing
    Set astrDefault = Nothing
    Set astrNew = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    Set astrUsed = Nothing
    Set astrAvailable = Nothing
    Set astrDefault = Nothing
    Set astrNew = Nothing
    RaiseError "cRptSummaryBy.ChangeFields", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpColumns
'' Description: Set up the column order/width/visibility according to spec
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpColumns()
On Error GoTo ErrSection:

    Dim strFields As String             ' Fields from the ini file
    Dim strIniFile As String            ' File path and name of the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Array of field information
    Dim strColName As String            ' Column Name
    Dim strHidden As String             ' Is Column Hidden?
    Dim lColWidth As Long               ' Width of the column
    Dim lTotWidth As Long               ' Total width from the fields string
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColPos As Long                 ' Column position
    
    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    strFields = GetIniFileProperty("Fields", m.strDefaultFields, Str(m.lAppID), strIniFile)
    
    If InStr(strFields, ",Profit;") = 0 Then
        Select Case UCase(m.strSubClass)
            Case "MONTHLY", "YEAR", "WEEKLY"
                If strFields = m.strDefaultFields Then
                    strFields = Replace(strFields, ",Net Profit;", ",Profit;0,Net Profit;")
                Else
                    strFields = Replace(strFields, ",Net Profit;", ",Profit;0;915,Net Profit;")
                    SetIniFileProperty "Fields", strFields, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
                End If
                
            Case Else
                strFields = Replace(strFields, ",Net Profit;", ",Profit;") & ",Net Profit;-2"
                
        End Select
    End If
    If InStr(strFields, ",ROI%;") = 0 Then
        Select Case UCase(m.strSubClass)
            Case "MONTHLY", "YEAR", "WEEKLY"
                If strFields = m.strDefaultFields Then
                    strFields = strFields = strFields & ",ROI%;0"
                Else
                    strFields = strFields = strFields & ",ROI%;0;915"
                    SetIniFileProperty "Fields", strFields, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
                End If
                
            Case Else
                strFields = strFields & ",ROI%;-2"
                
        End Select
    End If
    m.strFields = strFields
    
    lTotWidth = 0&
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        astrFields.SplitFields strFields, ","
        lColPos = 0&
        For lIndex = 0 To astrFields.Size - 1
            strColName = Parse(astrFields(lIndex), ";", 1)
            strHidden = Parse(astrFields(lIndex), ";", 2)
            lColWidth = CLng(ValOfText(Parse(astrFields(lIndex), ";", 3)))
            
            If strHidden <> "1" Then lColPos = lColPos + 1
            
            For lCol = 0 To .Cols - 1
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    If strHidden <> "1" Then
                        .ColPosition(lCol) = GDCol(eGDCol_CustomText) + lColPos
                        lTotWidth = lTotWidth + lColWidth
                        If strHidden = "-1" Or strHidden = "-2" Then
                            .ColHidden(GDCol(eGDCol_CustomText) + lColPos) = True
                        Else
                            .ColHidden(GDCol(eGDCol_CustomText) + lColPos) = False
                        End If
                        .ColWidth(GDCol(eGDCol_CustomText) + lColPos) = lColWidth
                    Else
                        .ColWidth(lCol) = lColWidth
                    End If
                    Exit For
                End If
            Next lCol
        Next lIndex
        
        If lTotWidth = 0& Then
            .AutoSize 0, .Cols - 1, False, 75
            'm.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
            'm.RptForm.ExtendCustomColumn m.vsGrid
            m.bAutoSize = True
        Else
            'm.RptForm.ExtendCustomColumn m.vsGrid
            m.bAutoSize = False
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Set astrFields = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "cRptSummaryBy.SetUpColumns", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveCols
'' Description: Save the column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveCols()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of field information
    Dim lIndex As Long                  ' Index into a for loop
    
    astrFields.Create eGDARRAY_Strings
    For lIndex = 0 To m.vsGrid.Cols - 1
        If lIndex <= GDCol(eGDCol_CustomText) Then
            If m.vsGrid.ColHidden(lIndex) = False Then
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";1;" & Str(m.vsGrid.ColWidth(lIndex))
            End If
        Else
            If m.vsGrid.ColHidden(lIndex) = False Then
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";0;" & Str(m.vsGrid.ColWidth(lIndex))
            ElseIf InStr(m.strFields, "," & m.vsGrid.TextMatrix(0, lIndex) & ";-2") > 0 Then
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";-2;" & Str(m.vsGrid.ColWidth(lIndex))
            Else
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";-1;" & Str(m.vsGrid.ColWidth(lIndex))
            End If
        End If
    Next lIndex

    m.strFields = astrFields.JoinFields(",")
    SetIniFileProperty "Fields", m.strFields, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
    m.bAutoSize = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptSummaryBy.SaveCols", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportToCSV
'' Description: Export the grid to a comma separated format
'' Inputs:      Filename to Export to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExportToCSV(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array to dump to the file
    Dim astrLine As New cGdArray        ' Array for column headers
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim s$, strFmt$
    
    astrFile.Create eGDARRAY_Strings
    astrLine.Create eGDARRAY_Strings
    
    With m.vsGrid
        For lRow = 0 To .Rows - 1
            astrLine.Size = 0
            For lCol = 0 To .Cols - 1
                If .ColHidden(lCol) = False Then
                    s = .TextMatrix(lRow, lCol)
                    If lRow > 0 And astrLine.Size = 0 Then
                        strFmt = .ColFormat(GDCol(eGDCol_CustomText))
                        If Len(strFmt) > 0 Then
                            s = Format(Val(s), strFmt)
                        End If
                    End If
                    astrLine.Add s
                End If
            Next lCol
            astrFile.Add astrLine.JoinFields(vbTab)
        Next
    End With
    
    astrFile.ToFile strFileName

ErrExit:
    Set astrFile = Nothing
    Exit Sub
    
ErrSection:
    Set astrFile = Nothing
    RaiseError "cRptSummaryBy.ExportToCSV", eGDRaiseError_Raise, g.strAppPath
    
End Sub

