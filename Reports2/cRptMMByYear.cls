VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMByYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    dMinProfit As Currency
    dBeginBalance As Currency
    lSubHeadRow As Long
    strStatToChart As String
    strReportName As String
    strShowCents As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary
    
    lRecord As Long
    aCustomText As cGdArray
    aTotTrades As cGdArray
    aWinPct As cGdArray
    aAvgTrade As cGdArray
    aWinAvg As cGdArray
    aPLRatio As cGdArray
    aPF As cGdArray
    aEndBal As cGdArray
    aReturnPct As cGdArray
    aCumReturn As cGdArray
    aNetProfit As cGdArray
    aDrawdown As cGdArray
    aLossConsec As cGdArray
    aLossLargest As cGdArray
    aSharpe As cGdArray
    aSterling As cGdArray
    
    hTradeDate As Long
    hAccountBalance As Long
    hSkipRpt As Long
    hSorted As Long
End Type
Private m As mPrivate
     
Private Enum eGDCols
    eGDCol_Year = 0
    eGDCol_Return = 1
    eGDCol_CumReturn = 2
    eGDCol_EndBal = 3
    eGDCol_Sharpe = 4
    eGDCol_Sterling = 5
    eGDCol_Drawdown = 6
    eGDCol_Trades = 7
    eGDCol_WinPct = 8
    eGDCol_AvgTrade = 9
    eGDCol_WinAvg = 10
    eGDCol_PL = 11
    eGDCol_PF = 12
    eGDCol_NetProfit = 13
    eGDCol_CL = 14
    eGDCol_MaxLoss = 15
End Enum
Private Const kGridCols = 16

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
'' InPuts:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Resize()
On Error Resume Next

    m.RptForm.ExtendCustomColumn m.vsGrid

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
    RaiseError "cRptMMByYear.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    m.RptForm.CustomColumn = GDCol(eGDCol_Year)
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hTradeDate = .FieldHandle(entd_TradeDate)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hSkipRpt = .FieldHandle(entd_SkipRpt)
        m.hSorted = .SortHandle
    End With
    
    'Report options
    m.dMinProfit = m.RptForm.OptionValue("ProfitGT")
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strStatToChart = m.RptForm.OptionValue("StatToChart")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    
    'Autosense mode (money mgt vs single unit)
    If gdGetNum(m.hAccountBalance, 1) <= 0 Then
        m.RptForm.SetToSingleUnit True
    End If
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    CalculateStats
    ShowReport
    BuildChart
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMByYear.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    'Calculate Items to chart (number of rows in grid)
    lItemsToChart = m.vsGrid.Rows - 1
    ShowChart lItemsToChart, m.Chart, m.RptForm.lblNoTrades, m.RptForm.lblNoMMTrades
    If lItemsToChart = 0 Then Exit Sub
    
    '** Set Subsets, Points, and fill with YData **'
    With m.Chart
        .PEactions = 20     'Empty chart and reinitialize
        .Subsets = 1
        .Points = lItemsToChart
        
        .MainTitle = m.strReportName
        .SubTitle = "by " & m.strStatToChart
        .YAxisLabel = m.strStatToChart
        .AxisFormatY = "|,|"
        .XAxisLabel = "Years"
        
        SetChartSettings m.Chart, GPM_BAR
        
        Select Case m.strStatToChart
            Case "Cummulative Return Pct"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aCumReturn.Num(lTrade + 1)
                Next lTrade
            Case "Ending Equity"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aEndBal.Num(lTrade + 1)
                Next lTrade
            Case "Win Pct"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aWinPct.Num(lTrade + 1)
                Next lTrade
            Case "PL Ratio"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aPLRatio.Num(lTrade + 1)
                Next lTrade
            Case "Profit Factor"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aPF.Num(lTrade + 1)
                Next lTrade
            Case "Net Profit"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aNetProfit.Num(lTrade + 1)
                Next lTrade
            Case "Return Pct"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aReturnPct.Num(lTrade + 1)
                Next lTrade
            Case "Win Avg"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aWinAvg.Num(lTrade + 1)
                Next lTrade
            Case "Drawdown"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aDrawdown.Num(lTrade + 1)
                Next lTrade
            Case "Avg Trade"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aAvgTrade.Num(lTrade + 1)
                Next lTrade
        End Select
        
        '** Set PointLabels property array for 12 points **'
        For lTrade = 1 To m.aCustomText.Size - 1
            .PointLabels(lTrade - 1) = m.aCustomText.Item(lTrade)
        Next lTrade
    
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMByYear.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .FixedCols = 0
        .Cols = kGridCols
        .FrozenCols = 1
        m.lSubHeadRow = 0
        
        .ColAlignment(GDCol(eGDCol_Year)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Year)) = "Year"
        
        .ColAlignment(GDCol(eGDCol_Trades)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Trades)) = "Trades"
        
        .ColAlignment(GDCol(eGDCol_WinPct)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WinPct)) = "Win Pct"
        
        .ColAlignment(GDCol(eGDCol_AvgTrade)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AvgTrade)) = "AvgTrade"
        
        .ColAlignment(GDCol(eGDCol_WinAvg)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WinAvg)) = "Win Avg"
        
        .ColAlignment(GDCol(eGDCol_PL)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_PL)) = "PayoutR" '"P/L"
        
        .ColAlignment(GDCol(eGDCol_PF)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_PF)) = "ProfitF" '"P/F"
        
        .ColAlignment(GDCol(eGDCol_EndBal)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EndBal)) = "End Bal"
        
        .ColAlignment(GDCol(eGDCol_Return)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Return)) = "Return%"
        
        .ColAlignment(GDCol(eGDCol_CumReturn)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_CumReturn)) = "Tot Rtn%"
        
        .ColAlignment(GDCol(eGDCol_Sharpe)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Sharpe)) = "Sharpe"
        
        .ColAlignment(GDCol(eGDCol_Sterling)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Sterling)) = "Sterling"
        
        .ColAlignment(GDCol(eGDCol_NetProfit)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfit)) = "Net Profit"
        
        .ColAlignment(GDCol(eGDCol_Drawdown)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Drawdown)) = "Drawdown"
        
        .ColAlignment(GDCol(eGDCol_CL)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_CL)) = "C/L"
        
        .ColAlignment(GDCol(eGDCol_MaxLoss)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_MaxLoss)) = "Max Loss"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMByYear.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStats
'' Description: Calculate the statistics
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStats()
On Error GoTo ErrSection:
    
    Dim X           As Long
    Dim BeginDate   As Date
    Dim EndDate     As Date
    Dim FrDate      As Date
    Dim FrDate2     As Date
    Dim ToDate      As Date
    Dim ToDate2     As Date
    Dim First       As Boolean
    Dim RecalcEndOfMthBalances As Boolean
    
    RecalcEndOfMthBalances = True
        
    ' Set properties for loading trades to analyze
    m.Stats.MinProfit = m.dMinProfit
    First = True
    
    ' Run statistics for each year in the system
    m.dBeginBalance = gdGetNum(m.hAccountBalance, 0)
    'BeginDate = gdGetNum(m.hTradeDate, 1)
    'EndDate = gdGetNum(m.hTradeDate, m.RptForm.Trades.NumRecords - 1)
    BeginDate = m.RptForm.FromDate
    EndDate = m.RptForm.ToDate
    FrDate = DateValue("1/1" & Str(Year(BeginDate)))
    ToDate = DateValue("12/31" & Str(Year(BeginDate)))
    Do While FrDate < EndDate
    
        ' Override for exact beginning/ending dates from prompt
        FrDate2 = FrDate
        ToDate2 = ToDate
        If FrDate < BeginDate Then
            FrDate2 = BeginDate
        End If
        If ToDate > EndDate Then
            ToDate2 = EndDate
        End If
        If InStr(1, ToDate2, ".") <= 0 Then
            ToDate2 = ToDate + 0.9999
        End If
        
        'Process all trades for current year
        For X = 1 To m.RptForm.Trades.NumRecords - 1
            'If .SignalType(X) = gEntrySignal Then
                If CDate(gdGetNum(m.hTradeDate, X)) >= FrDate2 And _
                   CDate(gdGetNum(m.hTradeDate, X)) <= ToDate2 Then
                    gdSetNum m.hSkipRpt, X, 0
                Else
                    gdSetNum m.hSkipRpt, X, 1
                End If
            'End If
        Next X
        If Not First Then
            'm.Stats.BeginBalance = m.Stats.EndingBalance
            gdSetNum m.hAccountBalance, 0, m.Stats.EndingBalance
        End If
        First = False
        m.Stats.SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        m.Stats.SetSortHandle m.RptForm.Trades.SortHandle
        m.Stats.EndDate = ToDate2
        m.Stats.StartDate = FrDate2
        m.Stats.Calculate RecalcEndOfMthBalances
        SaveResults Year(FrDate)
        RecalcEndOfMthBalances = False
        
        FrDate = DateAdd("yyyy", 1, FrDate)     'add one year to from date
        ToDate = DateValue("12/31" & Str(Year(FrDate)))
    Loop
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMByYear.CalculateStats", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveResults
'' Description: Save the consecutive statistics
'' Inputs:      Custom Text
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveResults(ByVal strCustomText As String)
On Error GoTo ErrSection

    With m.Stats
        m.lRecord = m.lRecord + 1
        m.aCustomText.Add strCustomText, m.lRecord
        m.aTotTrades.Add .TotTrades, m.lRecord
        m.aWinPct.Add .WinPct, m.lRecord
        m.aWinAvg.Add .WinAvg, m.lRecord
        m.aEndBal.Add .EndingBalance, m.lRecord
        m.aReturnPct.Add .ReturnPctMM, m.lRecord
        If m.dBeginBalance <> 0 Then
            m.aCumReturn.Add (.EndingBalance - m.dBeginBalance) / m.dBeginBalance, m.lRecord
        Else
            m.aCumReturn.Add 0, m.lRecord
        End If
        m.aPLRatio.Add .PLRatio, m.lRecord
        m.aPF.Add .ProfitFactor, m.lRecord
        m.aAvgTrade.Add .AvgTrade, m.lRecord
        m.aNetProfit.Add .TotalNetProfit, m.lRecord
        m.aDrawdown.Add .PeakToValleyDrawdown, m.lRecord
        m.aLossConsec.Add .LossConsec, m.lRecord
        m.aLossLargest.Add .LossLargest, m.lRecord
        m.aSharpe.Add .SharpeRatio, m.lRecord
        m.aSterling.Add .SterlingRatioCust, m.lRecord
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMByYear.SaveResults", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .Rows = .Rows + m.aCustomText.Size - 1
        For lIndex = 1 To m.aCustomText.Size - 1
            lRow = m.lSubHeadRow + lIndex
            .TextMatrix(lRow, GDCol(eGDCol_Year)) = m.aCustomText.Item(lIndex)
            .TextMatrix(lRow, GDCol(eGDCol_Trades)) = Format(m.aTotTrades.Num(lIndex), "#,##0")
            .TextMatrix(lRow, GDCol(eGDCol_WinPct)) = Format(m.aWinPct.Num(lIndex), "##0.0%")
            .TextMatrix(lRow, GDCol(eGDCol_WinAvg)) = Format(m.aWinAvg.Num(lIndex), FormatDollar(m.strShowCents))
            .TextMatrix(lRow, GDCol(eGDCol_PL)) = Format(m.aPLRatio.Num(lIndex), "##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_PF)) = Format(m.aPF.Num(lIndex), "##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_EndBal)) = Format(m.aEndBal.Num(lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.aEndBal.Num(lIndex), .Rows - 1, GDCol(eGDCol_EndBal)
            .TextMatrix(lRow, GDCol(eGDCol_Return)) = Format(m.aReturnPct.Num(lIndex), "#,##0.0%")
            ColorNegValue m.vsGrid, m.aReturnPct.Num(lIndex), .Rows - 1, GDCol(eGDCol_Return)
            .TextMatrix(lRow, GDCol(eGDCol_CumReturn)) = Format(m.aCumReturn.Num(lIndex), "#,##0.0%")
            ColorNegValue m.vsGrid, m.aCumReturn.Num(lIndex), .Rows - 1, GDCol(eGDCol_CumReturn)
            .TextMatrix(lRow, GDCol(eGDCol_AvgTrade)) = Format(m.aAvgTrade.Num(lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.aAvgTrade.Num(lIndex), .Rows - 1, GDCol(eGDCol_AvgTrade)
            .TextMatrix(lRow, GDCol(eGDCol_NetProfit)) = Format(m.aNetProfit.Num(lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.aNetProfit.Num(lIndex), .Rows - 1, GDCol(eGDCol_NetProfit)
            .TextMatrix(lRow, GDCol(eGDCol_Drawdown)) = Format(m.aDrawdown.Num(lIndex) * -1, "#,##0.0%")
            .TextMatrix(lRow, GDCol(eGDCol_CL)) = Format(m.aLossConsec.Num(lIndex), "#,##0")
            .TextMatrix(lRow, GDCol(eGDCol_MaxLoss)) = Format(m.aLossLargest.Num(lIndex) * -1, FormatDollar(m.strShowCents))
            .TextMatrix(lRow, GDCol(eGDCol_Sharpe)) = Format(m.aSharpe.Num(lIndex), "#,##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_Sterling)) = Format(m.aSterling.Num(lIndex), "#,##0.00")
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMByYear.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary
    Set m.aCustomText = New cGdArray
    Set m.aTotTrades = New cGdArray
    Set m.aWinPct = New cGdArray
    Set m.aWinAvg = New cGdArray
    Set m.aEndBal = New cGdArray
    Set m.aReturnPct = New cGdArray
    Set m.aCumReturn = New cGdArray
    Set m.aPLRatio = New cGdArray
    Set m.aPF = New cGdArray
    Set m.aAvgTrade = New cGdArray
    Set m.aLossConsec = New cGdArray
    Set m.aLossLargest = New cGdArray
    Set m.aSharpe = New cGdArray
    Set m.aSterling = New cGdArray
    Set m.aDrawdown = New cGdArray
    Set m.aNetProfit = New cGdArray
    m.aCustomText.Create eGDARRAY_Strings
    m.aTotTrades.Create eGDARRAY_Longs
    m.aWinPct.Create eGDARRAY_Doubles
    m.aAvgTrade.Create eGDARRAY_Doubles
    m.aWinAvg.Create eGDARRAY_Doubles
    m.aPLRatio.Create eGDARRAY_Doubles
    m.aPF.Create eGDARRAY_Doubles
    m.aEndBal.Create eGDARRAY_Doubles
    m.aReturnPct.Create eGDARRAY_Doubles
    m.aCumReturn.Create eGDARRAY_Doubles
    m.aNetProfit.Create eGDARRAY_Doubles
    m.aDrawdown.Create eGDARRAY_Doubles
    m.aLossConsec.Create eGDARRAY_Doubles
    m.aLossLargest.Create eGDARRAY_Doubles
    m.aSharpe.Create eGDARRAY_Floats
    m.aSterling.Create eGDARRAY_Floats

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing

End Sub
