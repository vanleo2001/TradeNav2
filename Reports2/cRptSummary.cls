VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptSummary"
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
    strShowCents As String
    strSubClass As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary

    aIndex As cGdArray
    strChartBy As String
    strShowMaxIntra As String
    lBottomRow As Long
    alGraphToTrades As cGdArray
    
    bShowMovingAverage As Boolean       ' Does the user want to show the moving average on the chart?
    bShowFiltered As Boolean            ' Does the user want to show the filtered equity on the chart?
End Type
Private m As mPrivate

Public Sub BeforeScrollTip(ByVal Row As Long)
End Sub
Public Sub AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
End Sub
Public Sub Resize()
End Sub

Private Sub Class_Initialize()
    Set m.Stats = New cStatsSummary
End Sub

Private Sub Class_Terminate()
    Set m.Stats = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayValues
'' Description: Display the Transaction Number, Date, and Equity at the current
''              mouse location
'' Inputs:      Current Mouse Location, Label to display values in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayValues(ByVal X As Double, ByVal Y As Double, lbl As Object)
On Error GoTo ErrSection:

    Dim lTradeNum As Long               ' Transaction Number
    Dim dProfit As Double               ' Profit for this transaction
    Dim dDate As Double                 ' Date of this transaction
    Dim dEquity As Double               ' Equity at this point

    lTradeNum = m.RptForm.Trades.Item(m.alGraphToTrades(X - 1), entd_TradeNbr)
    dProfit = m.RptForm.Trades.Item(m.alGraphToTrades(X - 1), entd_TotalProfit)
    dDate = m.RptForm.Trades.Item(m.alGraphToTrades(X - 1), entd_TradeDate)
    dEquity = m.Chart.YData(0, X - 1)

    With lbl
        .Caption = "Transaction: " & lTradeNum & _
                ", Date: " & DateFormat(dDate) & _
                ", Equity: " & Format(dEquity, "$#,##0.00")
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptSummary.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Run
'' Description: Run and show the report
'' Inputs:      Form to show report on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Run(RptForm As Object)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim hPosition As Long               ' Array Handle to the Position array
    Dim hSkipRpt As Long                ' Array Handle to the SkipRpt array
    
    Set m.RptForm = RptForm
    Set m.Chart = RptForm!BarChart
    Set m.vsGrid = RptForm!vsReport

    With m.RptForm
        ' Save Array Handles to some of the Trades class arrays...
        hPosition = .Trades.FieldHandle(entd_Position)
        hSkipRpt = .Trades.FieldHandle(entd_SkipRpt)
        
        .SetToSingleUnit False
        .FilterTrades
        
        ' Save the Report options...
        m.dMinProfit = .OptionValue("ProfitGT")
        m.strReportName = .OptionValue("ReportName")
        m.strShowCents = .OptionValue("ShowCents")
        m.strChartBy = .OptionValue("ChartBy")
        m.strShowMaxIntra = .OptionValue("ShowMaxIntra")
        m.strSubClass = .OptionValue("SubClass")
        
        m.bShowMovingAverage = True '.OptionValue("ShowMovAvg")
        m.bShowFiltered = True '.OptionValue("ShowFiltered")
    End With
    
    ' Skip trades if necessary...
    Select Case UCase(m.strSubClass)
        Case "LONG"
            For lIndex = 1 To m.RptForm.Trades.NumRecords - 1
                If gdGetNum(hPosition, lIndex) = 0 Then gdSetNum hSkipRpt, lIndex, 1
            Next lIndex
        Case "SHORT"
            For lIndex = 1 To m.RptForm.Trades.NumRecords - 1
                If gdGetNum(hPosition, lIndex) = 1 Then gdSetNum hSkipRpt, lIndex, 1
            Next lIndex
    End Select
    
    ' Calculate Statistics...
    With m.Stats
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        .MinProfit = m.dMinProfit
        .StartDate = m.RptForm.OptionValue("FromDate")
        .EndDate = m.RptForm.OptionValue("ToDate")
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
        
        .Calculate True
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    ' Fill in the grid and build the chart...
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    ShowReport
    BuildChart
    m.vsGrid.Redraw = flexRDBuffered
    
    m.RptForm.AvgTradesPerYear = Str(Round(m.Stats.AvgTradesPerYear))
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummary.Run", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
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
        .AutoSizeMode = flexAutoSizeColWidth ' = flexAutoSizeRowHeight
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
        .ColWidth(3) = 2200
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
    RaiseError "cRptSummary.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Dim lRow As Long, d As Double, s As String
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .BackColorAlternate = vbDefault
        .WordWrap = False
        .Rows = 29 '28 '24
        
        'OVERALL section
        m.lSubHeadRow = 0
        lRow = 0
        'OutLookTitle "Overall", lRow, 0, 0, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        OutLookTitle "Overall", lRow, 0, 0, m.vsGrid, _
            vbWhite, vbBlue, m.vsGrid.Font.Name, m.vsGrid.Font.Size, True, 1.3
        
        'Total net profit / Profit Factor
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Total Net Profit:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinTotal + m.Stats.LossTotal, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinTotal + m.Stats.LossTotal, lRow, 1
        .TextMatrix(lRow, 3) = "Profit Factor ($Wins/$Losses):"
        .TextMatrix(lRow, 4) = Format(m.Stats.ProfitFactor, "#,##0.00")
        .Cell(flexcpFontBold, lRow, 4) = True
        
        'Total Trades / Winning Pct
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Total Trades:"
        .TextMatrix(lRow, 1) = Format(m.Stats.TotTrades, "#,##0")
        .TextMatrix(lRow, 3) = "Winning Percentage:"
        .TextMatrix(lRow, 4) = Format(m.Stats.WinPct, "#,##0.0%")
        '.Cell(flexcpFontBold, lRow, 4) = True
        
        'Average Trade / Payout Ratio (avgwin/avgloss)
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Average Trade:"
        .TextMatrix(lRow, 1) = Format(m.Stats.AvgTrade, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.AvgTrade, lRow, 1
        .TextMatrix(lRow, 3) = "Payout Ratio (AvgWin/AvgLoss):"
        .TextMatrix(lRow, 4) = Format(m.Stats.PLRatio, "#,##0.00")
        '.Cell(flexcpFontBold, lRow, 4) = True
        
        'Max closed-out drawdown / CPC Index
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Max Closed-out Drawdown:"
        .TextMatrix(lRow, 1) = Format(m.Stats.MaxDrawdownClosed, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.MaxDrawdownClosed, lRow, 1
        .TextMatrix(lRow, 3) = "CPC Index (PF x Win% x PR):"
        .TextMatrix(lRow, 4) = Format(m.Stats.CPCIndex, "#,##0.00")
        
        'Max intraday drawdown / Expectancy%
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Max Intra-trade Drawdown:"
        If m.Stats.MaxDrawdownIntra = 0 Then
            .TextMatrix(lRow, 1) = "N/A"
        Else
            .TextMatrix(lRow, 1) = Format(m.Stats.MaxDrawdownIntra, FormatDollar(m.strShowCents))
        End If
        ColorNegValue m.vsGrid, m.Stats.MaxDrawdownIntra, lRow, 1
        .TextMatrix(lRow, 3) = "Expectancy (AvgTrade/AvgLoss):"
        If m.Stats.Expectancy = 0 Then
            .TextMatrix(lRow, 4) = "N/A"
        Else
            .TextMatrix(lRow, 4) = Format(m.Stats.Expectancy, "#,##0.00%")
            '.Cell(flexcpFontBold, lRow, 4) = True
        End If
        ColorNegValue m.vsGrid, m.Stats.Expectancy, lRow, 4
        
        'Acct size reqd / Return %
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Account Size Required:"
        .TextMatrix(lRow, 1) = Format(m.Stats.AccountSizeReq, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Return Pct:"
        .TextMatrix(lRow, 4) = Format(m.Stats.ReturnPct, "#,##0.0%")
        .Cell(flexcpFontBold, lRow, 4) = True
        
        'Open equity / Kelly%
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Open Equity:"
        .TextMatrix(lRow, 1) = Format(m.Stats.OpenEquity, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.OpenEquity, lRow, 1
        .TextMatrix(lRow, 3) = "Kelly Pct (AvgTrade/AvgWin):"
        .TextMatrix(lRow, 4) = Format(m.Stats.KellyRatio, "#0.00%") '"0.0000")
        '.Cell(flexcpFontBold, lRow, 4) = True
        
        '% in market / Optimal f
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Percent in the Market:"
        If m.Stats.PctInMarket = 0 Then
            .TextMatrix(lRow, 1) = "N/A"
        Else
            .TextMatrix(lRow, 1) = Format(m.Stats.PctInMarket, "#,##0.0%")
        End If
        .TextMatrix(lRow, 3) = "Optimal f:"
        .TextMatrix(lRow, 4) = Format(m.Stats.OptimalF, "0.00")
        
        'Avg # bars in trade / Z-score
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg # of Bars in Trade:"
        If m.Stats.AvgBarsInTrade = 0 Then
            .TextMatrix(lRow, 1) = "N/A"
        Else
            .TextMatrix(lRow, 1) = Format(m.Stats.AvgBarsInTrade, "#,##0.00")
        End If
        .TextMatrix(lRow, 3) = "Z-Score (W/L Predictability):"
        .TextMatrix(lRow, 4) = Format(m.Stats.ZScore, "#,##0.0")
        
        'Avg # trades per year / current streak
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg # of Trades per Year:"
        .TextMatrix(lRow, 1) = Format(m.Stats.AvgTradesPerYear, "#,##0.0")
        .TextMatrix(lRow, 3) = "Current Streak:"
        If m.Stats.CurrentRun >= 0 Then
            .TextMatrix(lRow, 4) = Format(m.Stats.CurrentRun, "#,##0") & " Wins"
        Else
            .TextMatrix(lRow, 4) = Format(Abs(m.Stats.CurrentRun), "#,##0") & " Losses"
        End If
        
        ' calc and display the SharpeRatio stuff
        m.Stats.CalcSharpeRatio
        lRow = lRow + 2
        OutLookTitle "Monthly Profit Analysis", lRow, 0, 0, m.vsGrid, _
            vbWhite, vbBlue, m.vsGrid.Font.Name, m.vsGrid.Font.Size, True, 1.3
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Average Monthly Profit:"
        .TextMatrix(lRow + 1, 0) = "Std Dev of Monthly Profits:"
        If m.Stats.MonthlyProfitStdDev = 0 Then
            .TextMatrix(lRow, 1) = "N/A"
            .TextMatrix(lRow + 1, 1) = "N/A"
        Else
            .TextMatrix(lRow, 1) = Format(m.Stats.MonthlyProfitAvg, FormatDollar(m.strShowCents))
            .TextMatrix(lRow + 1, 1) = Format(m.Stats.MonthlyProfitStdDev, FormatDollar(m.strShowCents))
        End If
        .TextMatrix(lRow, 3) = "Monthly Sharpe Ratio:"
        .TextMatrix(lRow + 1, 3) = "Annualized Sharpe Ratio:"
        If m.Stats.SharpeRatioMonthly = 0 Then
            .TextMatrix(lRow, 4) = "N/A"
            .TextMatrix(lRow + 1, 4) = "N/A"
        Else
            .TextMatrix(lRow, 4) = Format(m.Stats.SharpeRatioMonthly, "#0.00")
            .TextMatrix(lRow + 1, 4) = Format(m.Stats.SharpeRatioAnnualized, "#0.00")
        End If
        lRow = lRow + 2
        .TextMatrix(lRow, 3) = "Calmar Ratio:"
        If m.Stats.CalmarRatio = 0 Then
            .TextMatrix(lRow, 4) = "N/A"
        Else
            .TextMatrix(lRow, 4) = Format(m.Stats.CalmarRatio, "#0.00")
        End If
        
        
        'WINNING/LOSING SECTION
        '======================
        lRow = lRow + 2
        'Format Winning/Losing Statistics in four column table
        'OutLookTitle "Winning Trades", lRow, 0, 0, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        OutLookTitle "Winning Trades", lRow, 0, 0, m.vsGrid, _
            vbWhite, vbBlue, m.vsGrid.Font.Name, m.vsGrid.Font.Size, True, 1.3
        'OutLookTitle "Losing Trades", lRow, 3, 3, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        OutLookTitle "Losing Trades", lRow, 3, 3, m.vsGrid, _
            vbWhite, vbBlue, m.vsGrid.Font.Name, m.vsGrid.Font.Size, True, 1.3
        
        'Total Winners/Total Losers
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Total Winners:"
        .TextMatrix(lRow, 1) = Format(m.Stats.Wins, "#,##0")
        .TextMatrix(lRow, 3) = "Total Losers:"
        .TextMatrix(lRow, 4) = Format(m.Stats.Losses, "#,##0")

        'Total Gross profit/loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Gross Profit:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinTotal, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Gross Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossTotal, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinTotal, lRow, 1
        ColorNegValue m.vsGrid, m.Stats.LossTotal, lRow, 4
        
        'Average Win/Loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Average Win:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinAvg, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Average Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossAvg, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.LossAvg, lRow, 4
        
        'Largest Win/Loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Largest Win:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinLargest, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Largest Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossLargest, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.LossLargest, lRow, 4
        
        'Largest Drawdown in win/Peak in Loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Largest Drawdown in Win:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinMaxDrawdown, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Largest Peak in Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossMaxPeak, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinMaxDrawdown, lRow, 1
        ColorNegValue m.vsGrid, m.Stats.LossMaxPeak, lRow, 4
        
        'Avg Drawdown in win/Peak in Loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg Drawdown in Win:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinAvgDrawdown, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Avg Peak in Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossAvgPeak, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinAvgDrawdown, lRow, 1
        ColorNegValue m.vsGrid, m.Stats.LossAvgPeak, lRow, 4
        
        'Avg Run up in win/loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg Run Up in Win:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinRunUpAvg, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Avg Run Up in Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossRunUpAvg, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinRunUpAvg, lRow, 1
        ColorNegValue m.vsGrid, m.Stats.LossRunUpAvg, lRow, 4
        
        'Avg Run down in win/loss
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg Run Down in Win:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinRunDownAvg, FormatDollar(m.strShowCents))
        .TextMatrix(lRow, 3) = "Avg Run Down in Loss:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossRunDownAvg, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinRunDownAvg, lRow, 1
        ColorNegValue m.vsGrid, m.Stats.LossRunDownAvg, lRow, 4
        
        'Most Consec Wins/Losses
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Most Consec Wins:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinConsec, "#,##0")
        .TextMatrix(lRow, 3) = "Most Consec Losses:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossConsec, "#,##0")
        
        'Average Number of Consec Wins/Losses
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg # of Consec Wins:"
        If m.Stats.WinRuns > 0 Then
            .TextMatrix(lRow, 1) = Format((m.Stats.Wins / m.Stats.WinRuns), "#,##0.00")
        Else
            .TextMatrix(lRow, 1) = "0"
        End If
        .TextMatrix(lRow, 3) = "Avg # of Consec Losses:"
        If m.Stats.LossRuns > 0 Then
            .TextMatrix(lRow, 4) = Format((m.Stats.Losses / m.Stats.LossRuns), "#,##0.00")
        Else
            .TextMatrix(lRow, 4) = "0"
        End If
        
        'Average Number of Consec Wins/Losses
        lRow = lRow + 1
        .TextMatrix(lRow, 0) = "Avg # of Bars in Wins:"
        .TextMatrix(lRow, 1) = Format(m.Stats.WinAvgBars, "#,###.00")
        .TextMatrix(lRow, 3) = "Avg # of Bars in Losses:"
        .TextMatrix(lRow, 4) = Format(m.Stats.LossAvgBars, "#,###.00")
        
        '.Select m.lSubHeadRow + 12, 0, m.lSubHeadRow + 22, 4
        '.CellFontBold = False
        '.CellFontSize = 8
        
        .AutoSize 0, .Cols - 1
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummary.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Public Sub ShowShuffledTrades(ByVal iShowRandomSequences)

    BuildChart iShowRandomSequences

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildChart
'' Description: Intialize and load the chart
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildChart(Optional ByVal iShowRandomSequences As Long = 0)
    
#If 1 Then
On Error GoTo ErrSection:
    
    Dim lItemsToChart As Long           ' Number of items to chart
    Dim lTrade As Long                  ' Index into a for loop
    Dim dEquity As Double               ' Running total of equity
    Dim dEquityMA As Double             ' Equity Moving Average value
    Dim dFiltered As Double             ' Filtered equity value
    Dim dMaxIntra As Double             ' Max Intra-Trade drawdown
    Dim dEntryDate As Double            ' Date of entry
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    
    Dim lLocationOfPeak As Long         ' Location of the last peak
    Dim dPeak As Double                 ' Equity at the last peak
    Dim lPeakPos As Long                ' Position in Trades array of the Peak
    Dim lLocationOfMax As Long          ' Location of max drawdown
    Dim dDrawdown As Double             ' Value of the last drawdown
    Dim dMaxDrawdown As Double          ' Value of the max drawdown
    Dim dDateOfMax As Double            ' Date of the max drawdown
    Dim lLocOfFlat As Long              ' Location of the end of flat period
    Dim lNumDaysFlat As Long            ' Number of days flat
    Dim lLongestFlat As Long            ' Longest number of days flat
    Dim dDateOfFlat As Double           ' Date of the longest flat period end
    
    Dim dStartDate As Double            ' Starting date for the report
    Dim dEndDate As Double              ' Ending date for the report
    Dim dPrevDate As Double             ' Previous date handled
    Dim lDate As Long                   ' Index into a for loop
    Dim lAnnot As Long                  ' Annotation index
    Dim lPrev As Long                   ' Previous trade handled
    Dim YData() As Single               ' Array to assign to the Y-Axis
    Dim XData() As Single               ' Array to assign to the X-Axis
    
    Dim lMI As Long                     ' Subset index for max intraday
    Dim lFiltered As Long               ' Subset index for filtered equity
    
    Dim hSignalType As Long
    Dim hSkip As Long
    Dim hSkipRpt As Long
    Dim hEntryExitPtr As Long
    Dim hTradeDate As Long
    Dim hMaxLoss As Long
    Dim hTotalProfit As Long
    Dim hTradeNbr As Long
    Dim hUnfilteredEquity As Long
    Dim hEquityMA As Long
    Dim hFilteredEquity As Long
    Dim hSorted As Long
    
    Dim bChartByDate As Boolean, bDoMonthLabels As Boolean
    Dim iShuffle As Long
    Dim aProfits As New cGdArray
    Dim dAvgDrawdown#, dLargestDrawdown#, dSmallestDrawdown#
    Dim Chart As Pegoa
    
    If iShowRandomSequences > 0 Then
        Set Chart = m.RptForm.BarChart2
    Else
        Set Chart = m.Chart
    End If
    
    If iShowRandomSequences > 0 Or UCase(m.strChartBy) <> "DATE" Then
        bChartByDate = False
    Else
        bChartByDate = True
    End If
    
    With m.RptForm.Trades
        hSignalType = .FieldHandle(entd_SignalType)
        hSkip = .FieldHandle(entd_Skip)
        hSkipRpt = .FieldHandle(entd_SkipRpt)
        hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        hTradeDate = .FieldHandle(entd_TradeDate)
        hMaxLoss = .FieldHandle(entd_MaxLoss)
        hTotalProfit = .FieldHandle(entd_TotalProfit)
        hTradeNbr = .FieldHandle(entd_TradeNbr)
        hUnfilteredEquity = .FieldHandle(entd_UnfilteredEquity)
        hEquityMA = .FieldHandle(entd_EquityMA)
        hFilteredEquity = .FieldHandle(entd_FilteredEquity)
        hSorted = .SortHandle
    End With
    
    aProfits.Create eGDARRAY_Doubles, m.RptForm.Trades.NumRecords / 2, 0
    
    'Calculate Items to chart (number of rows in grid)
    lIndex = 0
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lTrade = gdGetNum(hSorted, lIndex2)
        
        If (gdGetNum(hSkip, lTrade) = 0) And (gdGetNum(hSkipRpt, lTrade) = 0) Then
            If dStartDate = 0 Then dStartDate = gdGetNum(hTradeDate, lTrade)
            dEndDate = gdGetNum(hTradeDate, lTrade)
            If gdGetNum(hSignalType, lTrade) = 1 Then
                aProfits.Num(lIndex) = gdGetNum(hTotalProfit, lTrade)
                lIndex = lIndex + 1
            End If
        End If
    Next lIndex2
    aProfits.Size = lIndex
        
    If bChartByDate Then
        lItemsToChart = Int(dEndDate) - Int(dStartDate) + 1
        ' TLB 4/30/2015: if charting less than about 2 years, then show month labels
        If dEndDate - dStartDate < 720 Then
            bDoMonthLabels = True
        End If
    Else
        lItemsToChart = lIndex
    End If
    
    ShowChart lItemsToChart, Chart, m.RptForm.lblNoTrades, m.RptForm.lblNoMMTrades
    If lItemsToChart = 0 Then Exit Sub
    
    Set m.alGraphToTrades = New cGdArray
    m.alGraphToTrades.Create eGDARRAY_Longs, lItemsToChart
    
    '** Set Subsets, Points, and fill with YData **'
    With Chart
        .PEactions = 20     'Empty chart and reinitialize
        .Subsets = 1
        If iShowRandomSequences > 0 Then
            .Subsets = .Subsets + iShowRandomSequences
        Else
            If m.bShowMovingAverage Then .Subsets = .Subsets + 1
            If m.bShowFiltered Then .Subsets = .Subsets + 1
            lFiltered = .Subsets - 1
            If m.strShowMaxIntra = "Yes" Then .Subsets = .Subsets + 1
            lMI = .Subsets - 1
        End If
        .Points = lItemsToChart
        
        ReDim YData(lItemsToChart * .Subsets) As Single
        ReDim XData(lItemsToChart) As Single

        Select Case UCase(m.strSubClass)
            Case "LONG"
                .MainTitle = "Equity Curve - Long Trades Only"
            Case "SHORT"
                .MainTitle = "Equity Curve - Short Trades Only"
            Case Else
                If bChartByDate Then
                    .MainTitle = "Equity Curve (by Date)"
                Else
                    .MainTitle = "Equity Curve (by Trade)"
                End If
        End Select
        .SubTitle = ""
        .YAxisLabel = "" ' = "Equity"
        .XAxisLabel = "" ' = "Trades"
        
        SetChartSettings Chart, GPM_LINE
        If iShowRandomSequences > 0 Then
            For iShuffle = 1 To .Subsets - 1
                .SubsetLineTypes(iShuffle) = PELT_THINSOLID
            Next
        Else
            If m.bShowMovingAverage Then .SubsetLineTypes(1) = PELT_THINSOLID
            If m.bShowFiltered Then .SubsetLineTypes(lFiltered) = PELT_THINSOLID
            If m.strShowMaxIntra = "Yes" Then .SubsetLineTypes(lMI) = PELT_DOT
        End If
        
        If bChartByDate Then
            '.DateTimeMode = PEDTM_VB
            '.StartTime = dStartDate
            .TargetPointsToTable = 0
            .FirstPtLabelOffset = 0
        Else
            .DateTimeMode = PEDTM_NONE
            .AltFrequencies(0) = 10
            .AltFrequencies(1) = 50
            .AltFrequencies(2) = 100
            .AltFrequencies(3) = 1000
            .FirstPtLabelOffset = 1
        End If

        lIndex = 0
        dEquity = 0
        dPeak = 0
        dMaxDrawdown = 0
        dMaxIntra = 0
        dPrevDate = dStartDate
        lLongestFlat = 0&
        lNumDaysFlat = 0&
        lPeakPos = 1&
        
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lTrade = gdGetNum(hSorted, lIndex2)
            If (gdGetNum(hSkip, lTrade) = 0) And (gdGetNum(hSkipRpt, lTrade) = 0) Then
                If bChartByDate Then
                    For lDate = Int(dPrevDate) To Int(gdGetNum(hTradeDate, lTrade)) - 1
                        lIndex = lIndex + 1
                        YData(lIndex - 1) = dEquity
                        If dEquityMA <> -999999# And m.bShowMovingAverage Then
                            YData((lIndex - 1) + lItemsToChart) = dEquityMA
                        End If
                        If dFiltered <> -999999# And m.bShowFiltered Then
                            YData((lIndex - 1) + lItemsToChart * lFiltered) = dFiltered
                        End If
                        If m.strShowMaxIntra = "Yes" Then
                            YData((lIndex - 1) + lItemsToChart * lMI) = dMaxIntra
                        End If
                        XData(lIndex - 1) = lDate
                        gdSetNum m.alGraphToTrades.ArrayHandle, lIndex - 1, lPrev

                        If Day(lDate) = 1 Then
                            If Month(lDate) = 1 Then
                                .VertLineAnnotationType(lAnnot) = PELAT_GRIDLINE
                                .VertLineAnnotation(lAnnot) = lIndex - 1
                                .VertLineAnnotationText(lAnnot) = "|B" & Format(lDate, "yyyy")
                                .VertLineAnnotationColor(lAnnot) = 1
                                lAnnot = lAnnot + 1
                            ElseIf bDoMonthLabels Then
                                .VertLineAnnotationType(lAnnot) = PELAT_GRIDLINE
                                .VertLineAnnotation(lAnnot) = lIndex - 1
                                .VertLineAnnotationText(lAnnot) = "|B" & Format(lDate, "mmm")
                                .VertLineAnnotationColor(lAnnot) = RGB(192, 192, 192)
                                lAnnot = lAnnot + 1
                            End If
                        End If
                    Next lDate
                End If
                
                lPrev = lTrade
                
                If gdGetNum(hSignalType, lTrade) = gEntrySignal Then
                    dMaxIntra = dMaxIntra + gdGetNum(hMaxLoss, lTrade)
                Else
                    If (Not bChartByDate) And (m.strShowMaxIntra = "Yes") Then
                        If lIndex > 0 Then
                            YData((lIndex - 1) + lItemsToChart * lMI) = dMaxIntra
                        End If
                    End If
                    dEquity = gdGetNum(hUnfilteredEquity, lTrade)
                    dEquityMA = gdGetNum(hEquityMA, lTrade)
                    dFiltered = gdGetNum(hFilteredEquity, lTrade)
                    dMaxIntra = dMaxIntra + gdGetNum(hTotalProfit, lTrade) - gdGetNum(hMaxLoss, lTrade)
                    
                    If Not bChartByDate Then
                        lIndex = lIndex + 1
                    End If
                    If lIndex > 0 Then
                        YData(lIndex - 1) = dEquity
                        If dEquityMA <> -999999# And m.bShowMovingAverage Then
                            YData((lIndex - 1) + lItemsToChart) = dEquityMA
                        End If
                        If dFiltered <> -999999# And m.bShowFiltered Then
                            YData((lIndex - 1) + lItemsToChart * lFiltered) = dFiltered
                        End If
                    
                        XData(lIndex - 1) = lIndex - 1
                        If bChartByDate Then
                            gdSetNum m.alGraphToTrades.ArrayHandle, lIndex - 1, lPrev
                        Else
                            .PointLabels(lIndex - 1) = lIndex - 1
                            gdSetNum m.alGraphToTrades.ArrayHandle, lIndex - 1, lTrade
                        End If
                    End If
                
                    ' Figure drawdown
                    If dEquity > dPeak Then
                        dPeak = dEquity
                        lLocationOfPeak = lIndex
                        If Not bChartByDate Then
                            lLocationOfPeak = lLocationOfPeak - 1
                        End If
                        lPeakPos = lTrade
                        
                        If lNumDaysFlat > lLongestFlat Then
                            lLongestFlat = lNumDaysFlat
                            lLocOfFlat = lLocationOfPeak
                            dDateOfFlat = gdGetNum(hTradeDate, lTrade)
                        End If
                        
                        lNumDaysFlat = 0&
                    Else
                        lNumDaysFlat = gdGetNum(hTradeDate, lTrade) - gdGetNum(hTradeDate, lPeakPos)
                        dDrawdown = dEquity - dPeak
                        If dDrawdown < dMaxDrawdown Then
                            dMaxDrawdown = dDrawdown
                            lLocationOfMax = lIndex
                            If Not bChartByDate Then
                                lLocationOfMax = lLocationOfMax - 1
                            End If
                            dDateOfMax = gdGetNum(hTradeDate, lTrade)
                        End If
                    End If
                End If
                
                dPrevDate = gdGetNum(hTradeDate, lTrade)
            End If
        Next lIndex2
        
        If iShowRandomSequences > 0 Then
            dAvgDrawdown = dMaxDrawdown
            dLargestDrawdown = dMaxDrawdown
            dSmallestDrawdown = dMaxDrawdown
            For iShuffle = 1 To iShowRandomSequences
                gdShuffle aProfits.ArrayHandle, 0, aProfits.Size - 1
                dEquity = 0
                dPeak = 0
                dMaxDrawdown = 0
                For lIndex = 0 To aProfits.Size - 1
                    dEquity = dEquity + aProfits.Num(lIndex)
                    YData(lIndex + lItemsToChart * iShuffle) = dEquity
                    ' Figure drawdown
                    If dEquity > dPeak Then
                        dPeak = dEquity
                    Else
                        dDrawdown = dEquity - dPeak
                        If dDrawdown < dMaxDrawdown Then
                            dMaxDrawdown = dDrawdown
                        End If
                    End If
                Next
                dAvgDrawdown = dAvgDrawdown + dMaxDrawdown
                If dMaxDrawdown < dLargestDrawdown Then
                    dLargestDrawdown = dMaxDrawdown
                End If
                If dMaxDrawdown > dSmallestDrawdown Then
                    dSmallestDrawdown = dMaxDrawdown
                End If
            Next
            dAvgDrawdown = dAvgDrawdown / (iShowRandomSequences + 1)
        End If
        
        ' Assign the XData and YData points to the chart all at once
        PEvset .hObject, PEP_faYDATA, YData(0), .Points * .Subsets
        PEvset .hObject, PEP_faXDATA, XData(0), .Points

        If lNumDaysFlat > lLongestFlat Then
            lLongestFlat = lNumDaysFlat
            lLocOfFlat = lTrade ' lIndex
            dDateOfFlat = gdGetNum(hTradeDate, lTrade) ' lIndex)
            If Not bChartByDate Then
                lLocOfFlat = lLocOfFlat - 1
            End If
        End If
            
            
        ' Draw a dotted zero line
        .GraphAnnotationType(0) = PEGAT_DOTLINE
        .GraphAnnotationX(0) = 0
        .GraphAnnotationY(0) = 0
        .GraphAnnotationColor(0) = vbBlack
        .GraphAnnotationType(1) = PEGAT_LINECONTINUE
        .GraphAnnotationX(1) = lItemsToChart
        .GraphAnnotationY(1) = 0
        .GraphAnnotationColor(1) = vbBlack
        
        .FontSizeTitleCntl = 0.9
            
        '** Set SubsetLabels property array for 4 subsets **'
        .SubsetLabels(0) = "Unfiltered Equity"
        .SubsetColors(0) = vbRed
        .SubsetLineTypes(0) = PELT_MEDIUMSOLID
        
        If iShowRandomSequences > 0 Then
            For iShuffle = 1 To iShowRandomSequences
                Select Case iShuffle Mod 10
                Case 1
                    .SubsetColors(iShuffle) = RGB(0, 0, 255)
                Case 2
                    .SubsetColors(iShuffle) = RGB(0, 0, 192)
                Case 3
                    .SubsetColors(iShuffle) = RGB(0, 255, 0)
                Case 4
                    .SubsetColors(iShuffle) = RGB(0, 192, 0)
                Case 5
                    .SubsetColors(iShuffle) = RGB(255, 0, 255)
                Case 6
                    .SubsetColors(iShuffle) = RGB(0, 0, 0)
                Case 7
                    .SubsetColors(iShuffle) = RGB(192, 192, 192)
                Case 8
                    .SubsetColors(iShuffle) = RGB(0, 255, 255)
                Case 9
                    .SubsetColors(iShuffle) = RGB(0, 192, 192)
                Case Else
                    .SubsetColors(iShuffle) = RGB(192, 192, 0)
                End Select
            Next
        
            .SubTitleFont = "Arial"
            .SubTitle = "Max Drawdowns:  Largest = " & Format(dLargestDrawdown, "$#,##0") & " , Average = " & Format(dAvgDrawdown, "$#,##0") & " , Smallest = " & Format(dSmallestDrawdown, "$#,##0")
            'm.RptForm.fraLegend.Visible = False
        Else
            'm.RptForm.fraLegend.Visible = True
            If m.bShowMovingAverage Then
                .SubsetLabels(1) = "Equity Moving Average"
                .SubsetColors(1) = vbBlue
            End If
            If m.bShowFiltered Then
                .SubsetLabels(lFiltered) = "Filtered Equity"
                .SubsetColors(lFiltered) = QBColor(2)
            End If
            If m.strShowMaxIntra = "Yes" Then
                .SubsetLabels(lMI) = "Max Intra-Trade Drawdown"
                .SubsetColors(lMI) = vbBlack
            End If
        
            ' Add the Drawdown Annotation
            .GraphAnnotationType(2) = PEGAT_SMALLUPTRIANGLESOLID
            .GraphAnnotationX(2) = lLocationOfMax + 1
            If m.bShowFiltered Then
                .GraphAnnotationY(2) = .YData(lFiltered, lLocationOfMax)
            Else
                .GraphAnnotationY(2) = .YData(0, lLocationOfMax)
            End If
            .GraphAnnotationColor(2) = vbBlack
            .LegendAnnotationType(0) = PEGAT_SMALLUPTRIANGLESOLID
            .LegendAnnotationColor(0) = vbBlack
            .LegendAnnotationText(0) = "Max Drawdown: " & Format(dMaxDrawdown, "$#,##0.00") & " on " & DateFormat(dDateOfMax)
            
            ' Add the Longest Flat Period to the Graph
            .GraphAnnotationType(3) = PEGAT_SMALLDOWNTRIANGLESOLID
            .GraphAnnotationX(3) = lLocOfFlat + 1
            If m.bShowFiltered Then
                .GraphAnnotationY(3) = .YData(lFiltered, lLocOfFlat)
            Else
                .GraphAnnotationY(3) = .YData(0, lLocOfFlat)
            End If
            .GraphAnnotationColor(3) = vbBlack
            .LegendAnnotationType(1) = PEGAT_SMALLDOWNTRIANGLESOLID
            .LegendAnnotationColor(1) = vbBlack
            .LegendAnnotationText(1) = "Longest Flat: " & Format(lLongestFlat, "#,##0") & " days ending " & DateFormat(dDateOfFlat)
        End If
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummary.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath


#Else '==================================================================================


On Error GoTo ErrSection:
    
    Dim lItemsToChart As Long           ' Number of items to chart
    Dim lTrade As Long                  ' Index into a for loop
    Dim dEquity As Double               ' Running total of equity
    Dim dEquityMA As Double             ' Equity Moving Average value
    Dim dFiltered As Double             ' Filtered equity value
    Dim dMaxIntra As Double             ' Max Intra-Trade drawdown
    Dim dEntryDate As Double            ' Date of entry
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    
    Dim lLocationOfPeak As Long         ' Location of the last peak
    Dim dPeak As Double                 ' Equity at the last peak
    Dim lPeakPos As Long                ' Position in Trades array of the Peak
    Dim lLocationOfMax As Long          ' Location of max drawdown
    Dim dDrawdown As Double             ' Value of the last drawdown
    Dim dMaxDrawdown As Double          ' Value of the max drawdown
    Dim dDateOfMax As Double            ' Date of the max drawdown
    Dim lLocOfFlat As Long              ' Location of the end of flat period
    Dim lNumDaysFlat As Long            ' Number of days flat
    Dim lLongestFlat As Long            ' Longest number of days flat
    Dim dDateOfFlat As Double           ' Date of the longest flat period end
    
    Dim dStartDate As Double            ' Starting date for the report
    Dim dEndDate As Double              ' Ending date for the report
    Dim dPrevDate As Double             ' Previous date handled
    Dim lDate As Long                   ' Index into a for loop
    Dim lAnnot As Long                  ' Annotation index
    Dim lPrev As Long                   ' Previous trade handled
    Dim YData() As Single               ' Array to assign to the Y-Axis
    Dim XData() As Single               ' Array to assign to the X-Axis
    
    Dim lMI As Long                     ' Subset index for max intraday
    Dim lFiltered As Long               ' Subset index for filtered equity
    
    Dim hSignalType As Long
    Dim hSkip As Long
    Dim hSkipRpt As Long
    Dim hEntryExitPtr As Long
    Dim hTradeDate As Long
    Dim hMaxLoss As Long
    Dim hTotalProfit As Long
    Dim hTradeNbr As Long
    Dim hUnfilteredEquity As Long
    Dim hEquityMA As Long
    Dim hFilteredEquity As Long
    Dim hSorted As Long
    
    With m.RptForm.Trades
        hSignalType = .FieldHandle(entd_SignalType)
        hSkip = .FieldHandle(entd_Skip)
        hSkipRpt = .FieldHandle(entd_SkipRpt)
        hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        hTradeDate = .FieldHandle(entd_TradeDate)
        hMaxLoss = .FieldHandle(entd_MaxLoss)
        hTotalProfit = .FieldHandle(entd_TotalProfit)
        hTradeNbr = .FieldHandle(entd_TradeNbr)
        hUnfilteredEquity = .FieldHandle(entd_UnfilteredEquity)
        hEquityMA = .FieldHandle(entd_EquityMA)
        hFilteredEquity = .FieldHandle(entd_FilteredEquity)
        hSorted = .SortHandle
    End With
    
    'Calculate Items to chart (number of rows in grid)
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lTrade = gdGetNum(hSorted, lIndex2)
        
        If (gdGetNum(hSkip, lTrade) = 0) And (gdGetNum(hSkipRpt, lTrade) = 0) Then
            If dStartDate = 0 Then dStartDate = gdGetNum(hTradeDate, lTrade)
            dEndDate = gdGetNum(hTradeDate, lTrade)
            If gdGetNum(hSignalType, lTrade) = 1 Then lIndex = lIndex + 1
        End If
    Next lIndex2
    
    If m.strChartBy = "Date" Then
        lItemsToChart = Int(dEndDate) - Int(dStartDate) + 1
    Else
        lItemsToChart = lIndex
    End If
    
    ShowChart lItemsToChart, m.Chart, m.RptForm.lblNoTrades, m.RptForm.lblNoMMTrades
    If lItemsToChart = 0 Then Exit Sub
    
    Set m.alGraphToTrades = New cGdArray
    m.alGraphToTrades.Create eGDARRAY_Longs, lItemsToChart
    
    '** Set Subsets, Points, and fill with YData **'
    With m.Chart
        .PEactions = 20     'Empty chart and reinitialize
        .Subsets = 1
        If m.bShowMovingAverage Then .Subsets = .Subsets + 1
        If m.bShowFiltered Then .Subsets = .Subsets + 1
        lFiltered = .Subsets - 1
        If m.strShowMaxIntra = "Yes" Then .Subsets = .Subsets + 1
        lMI = .Subsets - 1
        .Points = lItemsToChart
        
        ReDim YData(lItemsToChart * .Subsets) As Single
        ReDim XData(lItemsToChart) As Single

        Select Case UCase(m.strSubClass)
            Case "LONG"
                .MainTitle = "Equity Curve - Long Trades Only"
            Case "SHORT"
                .MainTitle = "Equity Curve - Short Trades Only"
            Case Else
                .MainTitle = "Equity Curve"
        End Select
        .SubTitle = ""
        .YAxisLabel = "" ' = "Equity"
        .XAxisLabel = "" ' = "Trades"
        
        SetChartSettings m.Chart, GPM_LINE
        If m.bShowMovingAverage Then .SubsetLineTypes(1) = PELT_THINSOLID
        If m.bShowFiltered Then .SubsetLineTypes(lFiltered) = PELT_THINSOLID
        If m.strShowMaxIntra = "Yes" Then .SubsetLineTypes(lMI) = PELT_DOT
        
        If m.strChartBy = "Date" Then
            '.DateTimeMode = PEDTM_VB
            '.StartTime = dStartDate
            .TargetPointsToTable = 0
            .FirstPtLabelOffset = 0
        Else
            .DateTimeMode = PEDTM_NONE
            .AltFrequencies(0) = 10
            .AltFrequencies(1) = 50
            .AltFrequencies(2) = 100
            .AltFrequencies(3) = 1000
            .FirstPtLabelOffset = 1
        End If

        lIndex = 0
        dEquity = 0
        dPeak = 0
        dMaxIntra = 0
        dPrevDate = dStartDate
        lLongestFlat = 0&
        lNumDaysFlat = 0&
        lPeakPos = 1&
        
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lTrade = gdGetNum(hSorted, lIndex2)
            If (gdGetNum(hSkip, lTrade) = 0) And (gdGetNum(hSkipRpt, lTrade) = 0) Then
                If m.strChartBy = "Date" Then
                    For lDate = Int(dPrevDate) To Int(gdGetNum(hTradeDate, lTrade)) - 1
                        lIndex = lIndex + 1
                        YData(lIndex - 1) = dEquity
                        If dEquityMA <> -999999# And m.bShowMovingAverage Then
                            YData((lIndex - 1) + lItemsToChart) = dEquityMA
                        End If
                        If dFiltered <> -999999# And m.bShowFiltered Then
                            YData((lIndex - 1) + lItemsToChart * lFiltered) = dFiltered
                        End If
                        If m.strShowMaxIntra = "Yes" Then
                            YData((lIndex - 1) + lItemsToChart * lMI) = dMaxIntra
                        End If
                        XData(lIndex - 1) = lDate
                        gdSetNum m.alGraphToTrades.ArrayHandle, lIndex - 1, lPrev

                        If Month(lDate) = 1 And Day(lDate) = 1 Then
                            .VertLineAnnotationType(lAnnot) = PELAT_GRIDLINE
                            .VertLineAnnotation(lAnnot) = lIndex - 1
                            .VertLineAnnotationText(lAnnot) = "|B" & Format(lDate, "yyyy")
                            .VertLineAnnotationColor(lAnnot) = 1
                            lAnnot = lAnnot + 1
                        End If
                    Next lDate
                End If
                
                lPrev = lTrade
                
                If gdGetNum(hSignalType, lTrade) = gEntrySignal Then
                    dMaxIntra = dMaxIntra + gdGetNum(hMaxLoss, lTrade)
                Else
                    If (m.strChartBy <> "Date") And (m.strShowMaxIntra = "Yes") Then
                        If lIndex > 0 Then YData((lIndex - 1) + lItemsToChart * lMI) = dMaxIntra
                    End If
                    dEquity = gdGetNum(hUnfilteredEquity, lTrade)
                    dEquityMA = gdGetNum(hEquityMA, lTrade)
                    dFiltered = gdGetNum(hFilteredEquity, lTrade)
                    dMaxIntra = dMaxIntra + gdGetNum(hTotalProfit, lTrade) - gdGetNum(hMaxLoss, lTrade)
                    
                    If m.strChartBy <> "Date" Then
                        lIndex = lIndex + 1
                    End If
                    If lIndex > 0 Then YData(lIndex - 1) = dEquity
                    If dEquityMA <> -999999# And m.bShowMovingAverage Then
                        If lIndex > 0 Then YData((lIndex - 1) + lItemsToChart) = dEquityMA
                    End If
                    If dFiltered <> -999999# And m.bShowFiltered Then
                        If lIndex > 0 Then YData((lIndex - 1) + lItemsToChart * lFiltered) = dFiltered
                    End If
                    
                    If lIndex > 0 Then XData(lIndex - 1) = lIndex - 1
                    If m.strChartBy = "Date" Then
                        If lIndex > 0 Then gdSetNum m.alGraphToTrades.ArrayHandle, lIndex - 1, lPrev
                    Else
                        .PointLabels(lIndex - 1) = lIndex - 1
                        If lIndex > 0 Then gdSetNum m.alGraphToTrades.ArrayHandle, lIndex - 1, lTrade
                    End If
                
                    ' Figure drawdown
                    If dEquity > dPeak Then
                        dPeak = dEquity
                        lLocationOfPeak = lIndex
                        If m.strChartBy = "Trade" Then lLocationOfPeak = lLocationOfPeak - 1
                        lPeakPos = lTrade
                        
                        If lNumDaysFlat > lLongestFlat Then
                            lLongestFlat = lNumDaysFlat
                            lLocOfFlat = lLocationOfPeak
                            dDateOfFlat = gdGetNum(hTradeDate, lTrade)
                        End If
                        
                        lNumDaysFlat = 0&
                    Else
                        lNumDaysFlat = gdGetNum(hTradeDate, lTrade) - gdGetNum(hTradeDate, lPeakPos)
                        dDrawdown = dEquity - dPeak
                        If dDrawdown < dMaxDrawdown Then
                            dMaxDrawdown = dDrawdown
                            lLocationOfMax = lIndex
                            If m.strChartBy = "Trade" Then lLocationOfMax = lLocationOfMax - 1
                            dDateOfMax = gdGetNum(hTradeDate, lTrade)
                        End If
                    End If
                End If
                
                dPrevDate = gdGetNum(hTradeDate, lTrade)
            End If
        Next lIndex2
        
        ' Assign the XData and YData points to the chart all at once
        PEvset .hObject, PEP_faYDATA, YData(0), .Points * .Subsets
        PEvset .hObject, PEP_faXDATA, XData(0), .Points

        If lNumDaysFlat > lLongestFlat Then
            lLongestFlat = lNumDaysFlat
            lLocOfFlat = lTrade ' lIndex
            dDateOfFlat = gdGetNum(hTradeDate, lTrade) ' lIndex)
            If m.strChartBy = "Trade" Then lLocOfFlat = lLocOfFlat - 1
        End If
            
        '** Set SubsetLabels property array for 4 subsets **'
        .SubsetLabels(0) = "Unfiltered Equity"
        .SubsetColors(0) = vbRed
        
        If m.bShowMovingAverage Then
            .SubsetLabels(1) = "Equity Moving Average"
            .SubsetColors(1) = vbBlue
        End If
        If m.bShowFiltered Then
            .SubsetLabels(lFiltered) = "Filtered Equity"
            .SubsetColors(lFiltered) = QBColor(2)
        End If
        If m.strShowMaxIntra = "Yes" Then
            .SubsetLabels(lMI) = "Max Intra-Trade Drawdown"
            .SubsetColors(lMI) = vbBlack
        End If

        ' Draw a dotted zero line
        .GraphAnnotationType(0) = PEGAT_DOTLINE
        .GraphAnnotationX(0) = 0
        .GraphAnnotationY(0) = 0
        .GraphAnnotationColor(0) = vbBlack
        .GraphAnnotationType(1) = PEGAT_LINECONTINUE
        .GraphAnnotationX(1) = lItemsToChart
        .GraphAnnotationY(1) = 0
        .GraphAnnotationColor(1) = vbBlack
        
        ' Add the Drawdown Annotation
        .GraphAnnotationType(2) = PEGAT_SMALLUPTRIANGLESOLID
        .GraphAnnotationX(2) = lLocationOfMax + 1
        If m.bShowFiltered Then
            .GraphAnnotationY(2) = .YData(lFiltered, lLocationOfMax)
        Else
            .GraphAnnotationY(2) = .YData(0, lLocationOfMax)
        End If
        .GraphAnnotationColor(2) = vbBlack
        .LegendAnnotationType(0) = PEGAT_SMALLUPTRIANGLESOLID
        .LegendAnnotationText(0) = "Max Drawdown: " & Format(dMaxDrawdown, "$#,##0.00") & " on " & DateFormat(dDateOfMax)
        .LegendAnnotationColor(0) = vbBlack
        
        ' Add the Longest Flat Period to the Graph
        .GraphAnnotationType(3) = PEGAT_SMALLDOWNTRIANGLESOLID
        .GraphAnnotationX(3) = lLocOfFlat + 1
        If m.bShowFiltered Then
            .GraphAnnotationY(3) = .YData(lFiltered, lLocOfFlat)
        Else
            .GraphAnnotationY(3) = .YData(0, lLocOfFlat)
        End If
        .GraphAnnotationColor(3) = vbBlack
        .LegendAnnotationType(1) = PEGAT_SMALLDOWNTRIANGLESOLID
        .LegendAnnotationText(1) = "Longest Flat: " & Format(lLongestFlat, "#,##0") & " days ending " & DateFormat(dDateOfFlat)
        .LegendAnnotationColor(1) = vbBlack
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSummary.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

#End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowFromPoint
'' Description: Get the row number in the grid for the point double clicked on
'' Inputs:      Location of the Mouse Double-Click
'' Returns:     Row in the grid to highlight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowFromPoint(ByVal lIndex As Double, ByVal Y As Double) As Long
On Error GoTo ErrSection:

    RowFromPoint = -1
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptSummary.RowFromPoint", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Function

' TLB: testing for tooltips -- in case we want to use them in the future
Public Function ToolTip(ByVal nRow&, ByVal nCol&) As String

    On Error Resume Next
    Dim strTip$
    
    'strTip = "Row " & Str(nRow) & ", Col " & Str(nCol)
    ToolTip = strTip

End Function
