VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMSummary2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    lSubHeadRow As Long
    strReportName As String
    strShowCents As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary
    hAccountBalance As Long
End Type
Private m As mPrivate

Public Sub BeforeScrollTip(ByVal Row As Long)
End Sub
Public Sub AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
End Sub
Public Sub Resize()
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
    RaiseError "cRptMMSummary2.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    'Get Trades handles
    With m.RptForm.Trades
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
    End With
    
    'Report options
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    
    'Calculate Statistics
    With m.Stats
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        .StartDate = m.RptForm.FromDate
        .EndDate = m.RptForm.ToDate
        .MinProfit = 0
        .Calculate True
    End With
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    ShowReport
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMSummary2.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Dim lIndex As Long                  ' Index into a for loop
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        ClearGrid m.vsGrid
        
        .Editable = True
        .ExplorerBar = flexExSort
        .GridLines = flexGridNone
        .ExtendLastCol = False
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightNever
        .AutoSizeMode = flexAutoSizeRowHeight
        .AllowUserResizing = flexResizeColumns
        .ScrollTips = True
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        
        .FixedRows = 0
        .FixedCols = 0
        
        .Rows = 23 '30
        .Cols = 5 '6
        .FrozenCols = 0
        
        .ColFormat(0) = ""
        .ColFormat(1) = ""
        .ColFormat(2) = ""
        .ColFormat(3) = ""
        .ColFormat(4) = ""
        
        .ColWidth(0) = 2400
        .ColWidth(1) = 1400
        .ColWidth(2) = 250
        .ColWidth(3) = 2400
        .ColWidth(4) = 1400
        
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignLeftTop
        
        For lIndex = .FixedRows To .Rows - 1
            .MergeRow(lIndex) = False
        Next lIndex
        .MergeCells = flexMergeNever
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMSummary2.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim Results As gdArrayStatistics    ' Results array
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .BackColorAlternate = vbDefault
        .WordWrap = False
        
        'OVERALL section
        m.lSubHeadRow = 0
        OutLookTitle "Overall", m.lSubHeadRow, 0, 1, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        
        'Begin Balance/Peak to Valley dip pct
        .TextMatrix(m.lSubHeadRow + 1, 0) = "Begin Balance:"
        If gdGetNum(m.hAccountBalance, 0) > 0 Then
            .TextMatrix(m.lSubHeadRow + 1, 1) = Format(gdGetNum(m.hAccountBalance, 0), FormatDollar(m.strShowCents))
        Else
            .TextMatrix(m.lSubHeadRow + 1, 1) = "N/A"
        End If
        .TextMatrix(m.lSubHeadRow + 1, 3) = "Peak to Valley dip Pct:"
        If gdGetNum(m.hAccountBalance, 0) > 0 Then
            .TextMatrix(m.lSubHeadRow + 1, 4) = Format(m.Stats.PeakToValleyDrawdown, "##0.0%")
        Else
            .TextMatrix(m.lSubHeadRow + 1, 4) = "N/A"
        End If
        .Cell(flexcpForeColor, m.lSubHeadRow + 1, 4) = vbRed
        
        'Equity Peak/Drawdown intra day
        .TextMatrix(m.lSubHeadRow + 2, 0) = "Equity Peak:"
        If gdGetNum(m.hAccountBalance, 0) > 0 Then
            .TextMatrix(m.lSubHeadRow + 2, 1) = Format(m.Stats.EquityPeak, FormatDollar(m.strShowCents))
        Else
            .TextMatrix(m.lSubHeadRow + 2, 1) = "N/A"
        End If
        ColorNegValue m.vsGrid, m.Stats.EquityPeak, m.lSubHeadRow + 2, 1
        .TextMatrix(m.lSubHeadRow + 2, 3) = "Max drawdown Intraday:"
        .TextMatrix(m.lSubHeadRow + 2, 4) = Format(m.Stats.MaxDrawdownIntra, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.MaxDrawdownIntra, m.lSubHeadRow + 2, 4
        
        'Ending Balance/Max drawdown closed
        .TextMatrix(m.lSubHeadRow + 3, 0) = "Ending Balance:"
        If gdGetNum(m.hAccountBalance, 0) > 0 Then
            .TextMatrix(m.lSubHeadRow + 3, 1) = Format(m.Stats.EndingBalance, FormatDollar(m.strShowCents))
        Else
            .TextMatrix(m.lSubHeadRow + 3, 1) = "N/A"
        End If
        ColorNegValue m.vsGrid, m.Stats.EndingBalance, m.lSubHeadRow + 3, 1
        .TextMatrix(m.lSubHeadRow + 3, 3) = "Max drawdown closed:"
        .TextMatrix(m.lSubHeadRow + 3, 4) = Format(m.Stats.MaxDrawdownClosed, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.MaxDrawdownClosed, m.lSubHeadRow + 3, 4
        
        .Select m.lSubHeadRow + 1, 0, m.lSubHeadRow + 3, 4
        .CellFontBold = False
        .CellFontSize = 8
        .Cell(flexcpFontBold, m.lSubHeadRow + 1, 4) = True    'PV dip%
        
        
        'Return Summary
        '===============
        OutLookTitle "Return Summary", m.lSubHeadRow + 5, 0, 1, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        
        'Calculate standard statistics on Monthly returns array
        Dim MthRets   As cGdTable
        Dim MthRetResults   As gdArrayStatistics
        Const MR_ReturnPct = 3
        Set MthRets = New cGdTable
        MthRets.CopyFromHandle m.Stats.MthReturnValuesHandle
        gdCalcStatistics MthRets.FieldArrayHandle(MR_ReturnPct), _
            MthRetResults, 0, 0, -1
        
        'Cumulative Return%
        .TextMatrix(m.lSubHeadRow + 6, 0) = "Cumulative return%:"
        .TextMatrix(m.lSubHeadRow + 6, 1) = Format(m.Stats.ReturnPctMM, "#,##0.0%")
        
        'Average Annual Return%
        .TextMatrix(m.lSubHeadRow + 7, 0) = "Avg annual return%:"
        .TextMatrix(m.lSubHeadRow + 7, 1) = Format(m.Stats.AvgYearlyReturn, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 7, 3) = "Compounded Avg Annual (ROR):"
        .TextMatrix(m.lSubHeadRow + 7, 4) = Format(m.Stats.CompoundedAvgAnnualROR, "#,##0.0%")
        
        'Average Monthly Return%/Std of Monthly Return
        .TextMatrix(m.lSubHeadRow + 8, 0) = "Avg monthly return%:"
        .TextMatrix(m.lSubHeadRow + 8, 1) = Format(m.Stats.MthReturnAvg, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 8, 3) = "STD of monthly return%:"
        .TextMatrix(m.lSubHeadRow + 8, 4) = Format(MthRetResults.StdDev, "#,##0.0%")
    
        'Avg Winning Month/Avg Losing Month
        .TextMatrix(m.lSubHeadRow + 9, 0) = "Avg winning return%:"
        .TextMatrix(m.lSubHeadRow + 9, 1) = Format(m.Stats.MthReturnWinAvg, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 9, 3) = "Avg losing month:"
        .TextMatrix(m.lSubHeadRow + 9, 4) = Format(m.Stats.MthReturnLossAvg, "#,##0.0%")
        ColorNegValue m.vsGrid, m.Stats.MthReturnLossAvg, m.lSubHeadRow + 9, 4

        'Largest Monthly Gain/Largest monthly loss
        .TextMatrix(m.lSubHeadRow + 10, 0) = "Largest monthly gain%:"
        .TextMatrix(m.lSubHeadRow + 10, 1) = Format(m.Stats.MthReturnMax, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 10, 3) = "Largest monthly loss%:"
        .TextMatrix(m.lSubHeadRow + 10, 4) = Format(m.Stats.MthReturnMin, "#,##0.0%")
        ColorNegValue m.vsGrid, m.Stats.MthReturnMax, m.lSubHeadRow + 10, 1
        ColorNegValue m.vsGrid, m.Stats.MthReturnMin, m.lSubHeadRow + 10, 4
        
        'Pct of Winning Months/Pct of Losing Months
        .TextMatrix(m.lSubHeadRow + 11, 0) = "Pct of gain months:"
        .TextMatrix(m.lSubHeadRow + 11, 1) = Format(m.Stats.MthReturnPctOfWinMths, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 11, 3) = "Pct of loss months:"
        .TextMatrix(m.lSubHeadRow + 11, 4) = Format(1 - m.Stats.MthReturnPctOfWinMths, "#,##0.0%")
        ColorNegValue m.vsGrid, 1 - m.Stats.MthReturnPctOfWinMths, m.lSubHeadRow + 11, 4
        
        .Select m.lSubHeadRow + 6, 0, m.lSubHeadRow + 11, 4
        .CellFontBold = False
        .CellFontSize = 8
        .Cell(flexcpFontBold, m.lSubHeadRow + 6, 1) = True
        .Cell(flexcpFontBold, m.lSubHeadRow + 7, 1) = True
        
        
        'Standard Measures...
        OutLookTitle "Standard Measures", m.lSubHeadRow + 13, 0, 1, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        
        'SHARPE ratio/Curr 3-yr Sharpe ratio
        .TextMatrix(m.lSubHeadRow + 14, 0) = "Sharpe ratio overall:"
        .TextMatrix(m.lSubHeadRow + 14, 1) = Format(m.Stats.SharpeRatio, "#,##0.00")
        .TextMatrix(m.lSubHeadRow + 14, 3) = "Sterling Ratio"
        .TextMatrix(m.lSubHeadRow + 14, 4) = Format(m.Stats.SterlingRatioCust, "#,##0.00")
        
        'Sterling Ratio/Effiency Index
        .TextMatrix(m.lSubHeadRow + 15, 0) = "Sharpe ratio (3-year):"
        .TextMatrix(m.lSubHeadRow + 15, 1) = Format(m.Stats.SharpeRatio(3), "#,##0.00")
        .TextMatrix(m.lSubHeadRow + 15, 3) = "Sterling Ratio (3-year):"
        .TextMatrix(m.lSubHeadRow + 15, 4) = Format(m.Stats.SterlingRatioCust(3), "#,##0.00")
        
        .Cell(flexcpFontBold, m.lSubHeadRow + 14, 1) = True
        .Cell(flexcpFontBold, m.lSubHeadRow + 14, 4) = True
        
        'Optimial f/Kelly ratio
        .TextMatrix(m.lSubHeadRow + 16, 0) = "Optimal f:"
        .TextMatrix(m.lSubHeadRow + 16, 1) = Format(m.Stats.OptimalF, "##0.0%")
        
        'Z-Score
        ' >=  2 Winners tend to follow losers, losers follow winners
        ' <= -2 Winners tend to follow winners, losers follow losers
        .TextMatrix(m.lSubHeadRow + 16, 3) = "Z Score (W/L predictability):"
        .TextMatrix(m.lSubHeadRow + 16, 4) = Format(m.Stats.ZScore, "##0.00")
        
        'Note: Can't get the description to spill into next column
        '.MergeCells = flexMergeSpill
        'If m.Stats.ZScore <= -2 Then
        '    .TextMatrix(m.lSubHeadRow + 16, 1) = _
        '        Format(m.Stats.ZScore, "##0.0") & " (W follow wins, losses follow losses)"
        'Else
        '    If m.Stats.ZScore >= 2 Then
        '        .TextMatrix(m.lSubHeadRow + 16, 1) = _
        '            Format(m.Stats.ZScore, "##0.0") & " (wins tend to follow losses, losses follow wins)"
        '    Else
        '        .TextMatrix(m.lSubHeadRow + 16, 1) = _
        '            Format(m.Stats.ZScore, "##0.0") & " (no predictability)"
        '    End If
        'End If
        
        .Select m.lSubHeadRow + 14, 0, m.lSubHeadRow + 16, 4
        .CellFontBold = False
        .CellFontSize = 8
        
        
        'TRADE MANAGEMENT...
        OutLookTitle "Trade Management", m.lSubHeadRow + 18, 0, 1, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        
        'Calculate stats...
        Dim PctConflicts        As Single
        Dim AvgAlloc            As Single
        Dim AvgOpen             As Single
        Dim PctSkipped          As Single
        Dim MaxAlloc            As Single
        Dim maxOpen             As Single
        m.Stats.CalcTradeMgt PctConflicts, AvgAlloc, AvgOpen, PctSkipped, _
            MaxAlloc, maxOpen
            
        'Pct of conflict signals/Pct of signals skipped
        .TextMatrix(m.lSubHeadRow + 19, 0) = "Pct of conflict signals:"
        .TextMatrix(m.lSubHeadRow + 19, 1) = Format(PctConflicts, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 19, 3) = "Pct of signals skipped:"
        .TextMatrix(m.lSubHeadRow + 19, 4) = Format(PctSkipped, "#,##0.0%")

        'Avg equity allocated/trade/Max equity allocated/trade
        .TextMatrix(m.lSubHeadRow + 20, 0) = "Avg equity allocated/trade:"
        .TextMatrix(m.lSubHeadRow + 20, 1) = Format(AvgAlloc, "#,##0.0%")
        .TextMatrix(m.lSubHeadRow + 20, 3) = "Max equity allocated/trade:"
        .TextMatrix(m.lSubHeadRow + 20, 4) = Format(MaxAlloc, "#,##0.0%")
        
        'Avg open positions/Max open positions
        .TextMatrix(m.lSubHeadRow + 21, 0) = "Avg open positions:"
        .TextMatrix(m.lSubHeadRow + 21, 1) = Format(AvgOpen, "#,##0.0")
        .TextMatrix(m.lSubHeadRow + 21, 3) = "Max open positions:"
        .TextMatrix(m.lSubHeadRow + 21, 4) = Format(maxOpen, "#,##0.0")
        
        .Select m.lSubHeadRow + 19, 0, m.lSubHeadRow + 21, 4
        .CellFontBold = False
        .CellFontSize = 8
        
        .AutoSize 0, .Cols - 1
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMSummary2.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing

End Sub
