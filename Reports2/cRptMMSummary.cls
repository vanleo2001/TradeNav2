VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMSummary"
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
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary
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
    RaiseError "cRptMMSummary.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    'Report options
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    
    'Calculate Statistics
    With m.Stats
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        
        .MinProfit = 0
        .StartDate = m.RptForm.FromDate
        .EndDate = m.RptForm.ToDate
        .Calculate True
    End With
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    ShowReport
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMSummary.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .Rows = 21 '30
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
    RaiseError "cRptMMSummary.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .BackColorAlternate = vbDefault
        .WordWrap = False
        
        'OVERALL section
        m.lSubHeadRow = 0
        OutLookTitle "Overall", m.lSubHeadRow, 0, 1, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        
        'Total net profit/Profit Factor
        .TextMatrix(m.lSubHeadRow + 1, 0) = "Total Net Profit:"
        .TextMatrix(m.lSubHeadRow + 1, 1) = Format(m.Stats.WinTotal + m.Stats.LossTotal, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 1, 3) = "Profit Factor ($Wins/$Losses):"
        If m.Stats.LossTotal < 0 Then
            .TextMatrix(m.lSubHeadRow + 1, 4) = Format(Abs(m.Stats.WinTotal / m.Stats.LossTotal), "#,##0.00")
        Else
            .TextMatrix(m.lSubHeadRow + 1, 4) = "0"
        End If
        .Cell(flexcpFontBold, m.lSubHeadRow + 1, 4) = True
        ColorNegValue m.vsGrid, m.Stats.WinTotal + m.Stats.LossTotal, m.lSubHeadRow + 1, 1
        
        'Total Trades/Winning Pct
        .TextMatrix(m.lSubHeadRow + 2, 0) = "Total Trades:"
        .TextMatrix(m.lSubHeadRow + 2, 1) = Format(m.Stats.TotTrades, "#,##0")
        .TextMatrix(m.lSubHeadRow + 2, 3) = "Winning Percentage:"
        .TextMatrix(m.lSubHeadRow + 2, 4) = Format(m.Stats.WinPct, "##0.0%")
        .Cell(flexcpFontBold, m.lSubHeadRow + 2, 4) = True
        
        'Average Trade/Payout Ratio (avgwin/avgloss)
        .TextMatrix(m.lSubHeadRow + 3, 0) = "Average Trade:"
        .TextMatrix(m.lSubHeadRow + 3, 1) = Format(m.Stats.AvgTrade, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 3, 3) = "Payout Ratio (Avg Win/Loss):"
        .TextMatrix(m.lSubHeadRow + 3, 4) = Format(m.Stats.PLRatio, "##0.00")
        .Cell(flexcpFontBold, m.lSubHeadRow + 3, 4) = True
        ColorNegValue m.vsGrid, m.Stats.AvgTrade, m.lSubHeadRow + 3, 1
        
        'Avg # Bars/Kelly percentage:
        .TextMatrix(m.lSubHeadRow + 4, 0) = "Avg # of Bars in Trade:"
        If m.Stats.AvgBarsInTrade = 0 Then
            .TextMatrix(m.lSubHeadRow + 4, 1) = "N/A"
        Else
            .TextMatrix(m.lSubHeadRow + 4, 1) = Format(m.Stats.AvgBarsInTrade, "0.00")
        End If
        '.TextMatrix(m.lSubHeadRow + 4, 3) = "Kelly ratio:"
        '.TextMatrix(m.lSubHeadRow + 4, 4) = Format(m.Stats.KellyRatio, "0.000")
        
        'Avg trades per year/Percent in the market
        .TextMatrix(m.lSubHeadRow + 5, 0) = "Avg # of Trades per Year:"
        .TextMatrix(m.lSubHeadRow + 5, 1) = Format(m.Stats.AvgTradesPerYear, "0.0")
        .TextMatrix(m.lSubHeadRow + 5, 3) = "Percent in the Market:"
        If m.Stats.PctInMarket = 0 Then
            .TextMatrix(m.lSubHeadRow + 5, 4) = "N/A"
        Else
            .TextMatrix(m.lSubHeadRow + 5, 4) = Format(m.Stats.PctInMarket, "0.0%")
        End If
        
        'Drawdown/Intraday-Drawdown
        .TextMatrix(m.lSubHeadRow + 6, 0) = "Max Closed-out Drawdown:"
        .TextMatrix(m.lSubHeadRow + 6, 1) = Format(m.Stats.MaxDrawdownClosed, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 6, 3) = "Max Intra-trade Drawdown:"
        If m.Stats.MaxDrawdownIntra = 0 Then
            .TextMatrix(m.lSubHeadRow + 6, 4) = "N/A"
        Else
            .TextMatrix(m.lSubHeadRow + 6, 4) = Format(m.Stats.MaxDrawdownIntra, FormatDollar(m.strShowCents))
        End If
        ColorNegValue m.vsGrid, m.Stats.MaxDrawdownClosed, m.lSubHeadRow + 6, 1
        ColorNegValue m.vsGrid, m.Stats.MaxDrawdownIntra, m.lSubHeadRow + 6, 4
        
        'Account size required/Return on account
        .TextMatrix(m.lSubHeadRow + 7, 0) = "Account Size Required:"
        .TextMatrix(m.lSubHeadRow + 7, 1) = Format(m.Stats.AccountSizeReq, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 7, 3) = "Return Pct:"
        .TextMatrix(m.lSubHeadRow + 7, 4) = Format(m.Stats.ReturnPctMM, "#,##0.0%")
        .Cell(flexcpFontBold, m.lSubHeadRow + 7, 4) = True
        
        
        'WINNING/LOSING SECTION
        '======================
        'Format Winning/Losing Statistics in four column table
        OutLookTitle "Winning Trades", m.lSubHeadRow + 9, 0, 1, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        OutLookTitle "Losing Trades", m.lSubHeadRow + 9, 3, 4, m.vsGrid, _
            vbWhite, vbBlue, "Arial", 9, True, 1.3
        
        'Total Winners/Total Losers
        .TextMatrix(m.lSubHeadRow + 10, 0) = "Total Winners:"
        .TextMatrix(m.lSubHeadRow + 10, 1) = Format(m.Stats.Wins, "##0")
        .TextMatrix(m.lSubHeadRow + 10, 3) = "Total Losers:"
        .TextMatrix(m.lSubHeadRow + 10, 4) = Format(m.Stats.Losses, "##0")

        'Total Gross profit/loss
        .TextMatrix(m.lSubHeadRow + 11, 0) = "Gross Profit:"
        .TextMatrix(m.lSubHeadRow + 11, 1) = Format(m.Stats.WinTotal, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 11, 3) = "Gross Loss:"
        .TextMatrix(m.lSubHeadRow + 11, 4) = Format(m.Stats.LossTotal, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinTotal, m.lSubHeadRow + 11, 1
        ColorNegValue m.vsGrid, m.Stats.LossTotal, m.lSubHeadRow + 11, 4
        
        'Average Win/Loss
        .TextMatrix(m.lSubHeadRow + 12, 0) = "Average Win:"
        .TextMatrix(m.lSubHeadRow + 12, 1) = Format(m.Stats.WinAvg, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 12, 3) = "Average Loss:"
        .TextMatrix(m.lSubHeadRow + 12, 4) = Format(m.Stats.LossAvg, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.LossAvg, m.lSubHeadRow + 12, 4
        
        'Largest Win/Loss
        .TextMatrix(m.lSubHeadRow + 13, 0) = "Largest Win:"
        .TextMatrix(m.lSubHeadRow + 13, 1) = Format(m.Stats.WinLargest, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 13, 3) = "Largest Loss:"
        .TextMatrix(m.lSubHeadRow + 13, 4) = Format(m.Stats.LossLargest, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.LossLargest, m.lSubHeadRow + 13, 4
        
        'Largest Drawdown in win/Peak in Loss
        .TextMatrix(m.lSubHeadRow + 14, 0) = "Largest Drawdown in Win:"
        .TextMatrix(m.lSubHeadRow + 14, 1) = Format(m.Stats.WinMaxDrawdown, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 14, 3) = "Largest Peak in Loss:"
        .TextMatrix(m.lSubHeadRow + 14, 4) = Format(m.Stats.LossMaxPeak, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinMaxDrawdown, m.lSubHeadRow + 14, 1
        ColorNegValue m.vsGrid, m.Stats.LossMaxPeak, m.lSubHeadRow + 14, 4
        
        'Avg Drawdown in win/Peak in Loss
        .TextMatrix(m.lSubHeadRow + 15, 0) = "Avg Drawdown in Win:"
        .TextMatrix(m.lSubHeadRow + 15, 1) = Format(m.Stats.WinAvgDrawdown, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 15, 3) = "Avg Peak in Loss:"
        .TextMatrix(m.lSubHeadRow + 15, 4) = Format(m.Stats.LossAvgPeak, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinAvgDrawdown, m.lSubHeadRow + 15, 1
        ColorNegValue m.vsGrid, m.Stats.LossAvgPeak, m.lSubHeadRow + 15, 4
        
        'Avg Run up in win/loss
        .TextMatrix(m.lSubHeadRow + 16, 0) = "Avg Run Up in Win:"
        .TextMatrix(m.lSubHeadRow + 16, 1) = Format(m.Stats.WinRunUpAvg, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 16, 3) = "Avg Run Up in Loss:"
        .TextMatrix(m.lSubHeadRow + 16, 4) = Format(m.Stats.LossRunUpAvg, FormatDollar(m.strShowCents))
        ColorNegValue m.vsGrid, m.Stats.WinRunUpAvg, m.lSubHeadRow + 16, 1
        ColorNegValue m.vsGrid, m.Stats.LossRunUpAvg, m.lSubHeadRow + 16, 4
        
        'Avg Run down in win/loss
        .TextMatrix(m.lSubHeadRow + 17, 0) = "Avg Run Down in Win:"
        .TextMatrix(m.lSubHeadRow + 17, 1) = Format(m.Stats.WinRunDownAvg, FormatDollar(m.strShowCents))
        .TextMatrix(m.lSubHeadRow + 17, 3) = "Avg Run Down in Loss:"
        .TextMatrix(m.lSubHeadRow + 17, 4) = Format(m.Stats.LossRunDownAvg, "$#,##0.00")
        ColorNegValue m.vsGrid, m.Stats.WinRunDownAvg, m.lSubHeadRow + 17, 1
        ColorNegValue m.vsGrid, m.Stats.LossRunDownAvg, m.lSubHeadRow + 17, 4
        
        'Most Consec Wins/Losses
        .TextMatrix(m.lSubHeadRow + 18, 0) = "Most Consec Wins:"
        .TextMatrix(m.lSubHeadRow + 18, 1) = Format(m.Stats.WinConsec, "##0")
        .TextMatrix(m.lSubHeadRow + 18, 3) = "Most Consec Losses:"
        .TextMatrix(m.lSubHeadRow + 18, 4) = Format(m.Stats.LossConsec, "##0")
        
        'Average Number of Consec Wins/Losses
        .TextMatrix(m.lSubHeadRow + 19, 0) = "Avg # of Consec Wins:"
        If m.Stats.WinRuns > 0 Then
            .TextMatrix(m.lSubHeadRow + 19, 1) = Format((m.Stats.Wins / m.Stats.WinRuns), "##0.00")
        Else
            .TextMatrix(m.lSubHeadRow + 19, 1) = "0"
        End If
        .TextMatrix(m.lSubHeadRow + 19, 3) = "Avg # of Consec Losses:"
        If m.Stats.LossRuns > 0 Then
            .TextMatrix(m.lSubHeadRow + 19, 4) = Format((m.Stats.Losses / m.Stats.LossRuns), "##0.00")
        Else
            .TextMatrix(m.lSubHeadRow + 19, 4) = "0"
        End If
        
        'Average Number of Consec Wins/Losses
        .TextMatrix(m.lSubHeadRow + 20, 0) = "Avg # of Bars in Wins:"
        .TextMatrix(m.lSubHeadRow + 20, 1) = Format(m.Stats.WinAvgBars, "#0.00")
        .TextMatrix(m.lSubHeadRow + 20, 3) = "Avg # of Bars in Losses:"
        .TextMatrix(m.lSubHeadRow + 20, 4) = Format(m.Stats.LossAvgBars, "#0.00")
        
        .Select m.lSubHeadRow + 10, 0, m.lSubHeadRow + 20, 4
        .CellFontBold = False
        .CellFontSize = 8
        
        .AutoSize 0, .Cols - 1
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMSummary.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing

End Sub
