VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptLTStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    dMinProfit As Currency
    lSubHeadRow As Long
    strStatToChart As String
    strReportName As String
    strShowCents As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    
    Stats As cStatsSummary
    iForecastTrades As Long
    dTradeSize As Double
    
    dCurrentProfit As Currency
    lCurrentTrade As Long
    lNumConsec As Long
    lHighAmount As Long
    dLowAmt As Double
    lNDays As Long
    dNProfit As Double
    lNProfCnt As Long
    dNLoss As Double
    lNLossCnt As Long
    lNDaysTot As Long
    dNNProf As Double
    lNNProfCnt As Long
    dNNLoss As Double
    lNNLossCnt As Long
    lOffset As Long
    strTemp As String
    lWrkTrade As Long
    lTotTrades As Long
    
    hSignalType As Long
    hProfit As Long
    hSkip As Long
    hSkipEq As Long
    hSorted As Long
    
    Results As cGdTable
    lRecord As Long
End Type
Private m As mPrivate

Private Const mText = 0
Private Const mOcur = 1
Private Const mWin1 = 2
Private Const mPL1 = 3
Private Const mDraw1 = 4
Private Const mWin2 = 5
Private Const mPL2 = 6
Private Const mDraw2 = 7

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
    RaiseError "cRptLTStudy.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    m.RptForm.CustomColumn = mText
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hSignalType = .FieldHandle(entd_SignalType)
        m.hProfit = .FieldHandle(entd_Profit)
        m.hSkip = .FieldHandle(entd_Skip)
        m.hSkipEq = .FieldHandle(entd_SkipEqFilter)
        m.hSorted = .SortHandle
    End With
    
    m.RptForm.FilterTrades
    
    'Report options
    m.dMinProfit = m.RptForm.OptionValue("ProfitGT")
    m.strStatToChart = m.RptForm.OptionValue("StatToChart")
    m.iForecastTrades = m.RptForm.OptionValue("NextNTrades")
    m.dTradeSize = m.RptForm.OptionValue("TradeSize")
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")

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
    
    'Calculate Report contents...
    m.lTotTrades = m.RptForm.Trades.NumRecords - 1
    For m.lNumConsec = -6 To 6
        m.dLowAmt = m.lNumConsec * m.dTradeSize
        m.lHighAmount = (m.lNumConsec + 1) * m.dTradeSize
        CalcLT
        PostLT
    Next
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    ShowReport
    BuildChart
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptLTStudy.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
    lItemsToChart = m.vsGrid.Rows - 2
    ShowChart lItemsToChart, m.Chart, m.RptForm.lblNoTrades, m.RptForm.lblNoMMTrades
    If lItemsToChart = 0 Then Exit Sub
    
    '** Set Subsets, Points, and fill with YData **'
    With m.Chart
        .PEactions = 20     'Empty chart and reinitialize
        .Subsets = 2
        .Points = lItemsToChart
        
        .MainTitle = m.strReportName
        .SubTitle = "by " & m.strStatToChart
        .YAxisLabel = ""
        '.AxisFormatX = "|,|"
        .XAxisLabel = ""
        
        SetChartSettings m.Chart, GPM_HORIZONTALBAR
        
        Select Case m.strStatToChart
            Case "Win Pct"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.Results.Num(mWin1, lTrade + 1)
                Next lTrade
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(1, lTrade) = m.Results.Num(mWin2, lTrade + 1)
                Next lTrade
            Case "PL Ratio"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.Results.Num(mPL1, lTrade + 1)
                Next lTrade
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(1, lTrade) = m.Results.Num(mPL2, lTrade + 1)
                Next lTrade
            Case "Drawdown"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.Results.Num(mDraw1, lTrade + 1)
                Next lTrade
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(1, lTrade) = m.Results.Num(mDraw2, lTrade + 1)
                Next lTrade
        End Select
        
        '** Set PointLabels property array for 12 points **'
        For lTrade = 0 To lItemsToChart - 1
            .PointLabels(lTrade) = m.Results.Item(mText, lTrade + 1)
        Next lTrade
        .PointsToGraph = 15
        
        '.SubsetLabels(0) = m.strStatToChart & " Next Trade"
        '.SubsetLabels(1) = m.strStatToChart & " Next " & CStr(m.iForecastTrades) & " Trades"
        .SubsetLabels(0) = "Next Trade"
        .SubsetLabels(1) = "Next " & CStr(m.iForecastTrades) & " Trades"
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptLTStudy.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PostLT
'' Description: Save the statistics
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PostLT()
On Error GoTo ErrSection:
    
    'Save results...
    m.lRecord = m.lRecord + 1
    m.Results.Item(mText, m.lRecord) = "Trades of " & Format(m.dLowAmt, "$#,##0") & " to " & _
            Format(m.lHighAmount, "$#,##0") & _
            " occur every " & Format(Divide((m.lNDaysTot), (m.lNLossCnt) + (m.lNProfCnt)), "#,##0.0") & " trades"
    m.Results.Num(mOcur, m.lRecord) = (m.lNProfCnt + m.lNLossCnt)
    m.Results.Num(mWin1, m.lRecord) = Divide((m.lNProfCnt), (m.lNLossCnt) + (m.lNProfCnt))
    m.Results.Num(mPL1, m.lRecord) = Divide(m.dNProfit, (m.lNProfCnt))
    m.Results.Num(mDraw1, m.lRecord) = PLRatio(m.dNProfit, m.lNProfCnt, m.dNLoss, m.lNLossCnt)
    m.Results.Num(mWin2, m.lRecord) = Divide((m.lNNProfCnt), (m.lNNLossCnt) + (m.lNNProfCnt))
    m.Results.Num(mPL2, m.lRecord) = Divide(m.dNNProf, (m.lNNProfCnt))
    m.Results.Num(mDraw2, m.lRecord) = PLRatio(m.dNNProf, m.lNNProfCnt, m.dNNLoss, m.lNNLossCnt)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptLTStudy.PostLT", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcLT
'' Description: Calculate the statistics
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcLT()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim adProfit As cGdArray
    
    'Build array of adProfit amounts
    Set adProfit = New cGdArray
    adProfit.Create eGDARRAY_Doubles
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If gdGetNum(m.hSignalType, lIndex) = gExitSignal Then
            If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
                adProfit.Add gdGetNum(m.hProfit, lIndex)
            End If
        End If
    Next lIndex2
    m.lTotTrades = adProfit.Size - 1
    
    m.lNDays = 0: m.lNDaysTot = 0
    m.dNProfit = 0: m.lNProfCnt = 0: m.dNLoss = 0: m.lNLossCnt = 0
    m.dNNProf = 0: m.lNNProfCnt = 0: m.dNNLoss = 0: m.lNNLossCnt = 0
    
    For m.lCurrentTrade = 0 To m.lTotTrades
        
        m.dCurrentProfit = adProfit.Num(m.lCurrentTrade)
        If m.dCurrentProfit < m.dLowAmt Or m.dCurrentProfit > m.lHighAmount Then
            m.lNDays = m.lNDays + 1
        Else
            m.lNDaysTot = m.lNDaysTot + m.lNDays
            m.lNDays = 0
          
            ' Record "Next trade" win or loss
            If m.lCurrentTrade + 1 <= m.lTotTrades Then
                m.lWrkTrade = m.lCurrentTrade + 1
                m.dCurrentProfit = adProfit.Num(m.lCurrentTrade + 1)
                If m.dCurrentProfit > m.dTradeSize Then
                    m.dNProfit = m.dNProfit + m.dCurrentProfit
                    m.lNProfCnt = m.lNProfCnt + 1
                Else
                    m.dNLoss = m.dNLoss + Abs(m.dCurrentProfit)
                    m.lNLossCnt = m.lNLossCnt + 1
                End If
            End If
          
            ' Record next n days wins and losses
            For lIndex = 1 To m.iForecastTrades
                If m.lCurrentTrade + lIndex <= m.lTotTrades Then
                    m.lWrkTrade = m.lCurrentTrade + lIndex
                    m.dCurrentProfit = adProfit.Num(m.lCurrentTrade + lIndex)
                    If m.dCurrentProfit > m.dTradeSize Then
                        m.dNNProf = m.dNNProf + m.dCurrentProfit
                        m.lNNProfCnt = m.lNNProfCnt + 1
                    Else
                        m.dNNLoss = m.dNNLoss + Abs(m.dCurrentProfit)
                        m.lNNLossCnt = m.lNNLossCnt + 1
                    End If
                End If
            Next
            
        End If
    Next

ErrExit:
    Set adProfit = Nothing
    Exit Sub

ErrSection:
    RaiseError "cRptLTStudy.CalcLT", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Divide
'' Description: Divide two numbers making sure that the divisor is non-zero
'' Inputs:      Two numbers to divide
'' Returns:     Result (0 if divisor is 0)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Divide(n1 As Double, n2 As Double) As Double
On Error GoTo ErrSection:

    If n2 <> 0 Then
        Divide = Abs(n1 / n2)
    Else
        Divide = 0
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptLTStudy.Divide", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PLRatio
'' Description: Figure out the Profit/Loss Ratio
'' Inputs:      Profit, Profit Count, Loss, Loss Count
'' Returns:     Result (0 if divisor is 0)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PLRatio(Prof As Double, ProfCnt As Long, Loss As Double, _
        LossCnt As Long) As Double
On Error GoTo ErrSection:
    
    If Divide(Loss, (LossCnt)) <> 0 Then
        PLRatio = Divide(Prof, (ProfCnt)) / Divide(Loss, (LossCnt))
    Else
        PLRatio = 0
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptLTStudy.PLRatio", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Function

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
        
        .WordWrap = False 'True
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .AutoSizeMode = flexAutoSizeColWidth
        .ScrollBars = flexScrollBarBoth
        .BackColorAlternate = g.nAltGridRowColor '&HC0FFFF
        .ExtendLastCol = False 'True
        .Rows = 2
        .FixedRows = 2
        .FixedCols = 0
        .Cols = 8
        .FrozenCols = 1
        m.lSubHeadRow = 1
        
        '1st column can be customized
        .ColWidth(0) = 4000
        .ColSort(0) = flexSortStringAscending
        .TextMatrix(m.lSubHeadRow, 0) = "Profit Ranges"
        .TextMatrix(m.lSubHeadRow - 1, 0) = "Profit Ranges"
        .ColAlignment(0) = flexAlignLeftTop
        
        'Standard Columns that don't changed between report types...
        .ColWidth(1) = 500
        .ColAlignment(1) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, 1) = "Count"
        .TextMatrix(m.lSubHeadRow - 1, 1) = "Count"
        
        .ColWidth(2) = 700
        .ColDataType(2) = flexDTDouble
        .TextMatrix(m.lSubHeadRow, 2) = "Win%"
        .ColAlignment(2) = flexAlignRightTop
        
        .ColWidth(3) = 1000
        .ColDataType(3) = flexDTCurrency
        .TextMatrix(m.lSubHeadRow, 3) = "Avg Win"
        .ColAlignment(3) = flexAlignRightTop
        
        .ColWidth(4) = 700
        .ColDataType(4) = flexDTDouble
        .TextMatrix(m.lSubHeadRow, 4) = "PL Ratio"
        .ColAlignment(4) = flexAlignRightTop
        
        .ColWidth(5) = 700
        .ColDataType(5) = flexDTDouble
        .TextMatrix(m.lSubHeadRow, 5) = "Win%"
        .ColAlignment(5) = flexAlignRightTop
        
        .ColWidth(6) = 1000
        .ColDataType(6) = flexDTCurrency
        .TextMatrix(m.lSubHeadRow, 6) = "Avg Win"
        .ColAlignment(6) = flexAlignRightTop
        
        .ColWidth(7) = 700
        .ColDataType(7) = flexDTDouble
        .TextMatrix(m.lSubHeadRow, 7) = "PL Ratio"
        .ColAlignment(7) = flexAlignRightTop
        
        .TextMatrix(m.lSubHeadRow - 1, 2) = "The Next Trade"
        .TextMatrix(m.lSubHeadRow - 1, 3) = "The Next Trade"
        .TextMatrix(m.lSubHeadRow - 1, 4) = "The Next Trade"
        .TextMatrix(m.lSubHeadRow - 1, 5) = "Next " & CStr(m.iForecastTrades) & " Trades"
        .TextMatrix(m.lSubHeadRow - 1, 6) = "Next " & CStr(m.iForecastTrades) & " Trades"
        .TextMatrix(m.lSubHeadRow - 1, 7) = "Next " & CStr(m.iForecastTrades) & " Trades"
        .MergeCells = flexMergeFixedOnly
        .MergeRow(m.lSubHeadRow - 1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .Cell(flexcpFontBold, m.lSubHeadRow - 1, 2, m.lSubHeadRow - 1, 7) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignLeftTop
        .Cell(flexcpAlignment, 0, 0, 0, 1) = flexAlignLeftCenter
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptLTStudy.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReport
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowReport()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = .Rows + m.Results.NumRecords - 1
        For lIndex = 1 To m.Results.NumRecords - 1
            lRow = m.lSubHeadRow + lIndex
            .TextMatrix(lRow, 0) = m.Results.Item(mText, lIndex)
            .TextMatrix(lRow, 1) = Format(m.Results.Num(mOcur, lIndex), "##0")
            .TextMatrix(lRow, 2) = Format(m.Results.Num(mWin1, lIndex), "##0.0%")
            .TextMatrix(lRow, 3) = Format(m.Results.Num(mPL1, lIndex), FormatDollar(m.strShowCents))
            .TextMatrix(lRow, 4) = Format(m.Results.Num(mDraw1, lIndex), "##0.00")
            .TextMatrix(lRow, 5) = Format(m.Results.Num(mWin2, lIndex), "##0.0%")
            .TextMatrix(lRow, 6) = Format(m.Results.Num(mPL2, lIndex), FormatDollar(m.strShowCents))
            .TextMatrix(lRow, 7) = Format(m.Results.Num(mDraw2, lIndex), "##0.00")
        Next lIndex
        .RowHeight(0) = .RowHeight(2)
        .RowHeight(1) = .RowHeight(2)
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptLTStudy.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary
    Set m.Results = New cGdTable
    
    With m.Results
        .CreateField eGDARRAY_Strings, mText, "Text"
        .CreateField eGDARRAY_Longs, mOcur, "Ocur"
        .CreateField eGDARRAY_Floats, mWin1, "Win1"
        .CreateField eGDARRAY_Doubles, mPL1, "PL1"
        .CreateField eGDARRAY_Doubles, mDraw1, "Draw1"
        .CreateField eGDARRAY_Floats, mWin2, "Win2"
        .CreateField eGDARRAY_Doubles, mPL2, "PL2"
        .CreateField eGDARRAY_Doubles, mDraw2, "Draw2"
    End With

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing
    Set m.Results = Nothing

End Sub
