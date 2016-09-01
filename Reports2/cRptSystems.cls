VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptSystems"
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
    
    lRecord As Long
    aSysNbr As cGdArray
    aSysName As cGdArray
    aTotTrades As cGdArray
    aWinPct As cGdArray
    aWinAvg As cGdArray
    aReturnPct As cGdArray
    aPLRatio As cGdArray
    aPF As cGdArray
    aAvgTrade As cGdArray
    aLossConsec As cGdArray
    aLossLargest As cGdArray
    aDrawdown As cGdArray
    aNetProfit As cGdArray
    
    hSysNbr As Long
    hSkipRpt As Long
    hSorted As Long
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_SysNbr = 0
    eGDCol_SysName = 1
    eGDCol_TotTrades = 2
    eGDCol_WinPct = 3
    eGDCol_WinAvg = 4
    eGDCol_ReturnPct = 5
    eGDCol_PLRatio = 6
    eGDCol_PF = 7
    eGDCol_NetProfit = 8
    eGDCol_AvgTrade = 9
    eGDCol_LossConsec = 10
    eGDCol_LossLargest = 11
    eGDCol_Drawdown = 12
End Enum
Private Const kGridCols = 13

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
    RaiseError "cRptSystems.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    m.RptForm.CustomColumn = GDCol(eGDCol_SysName)
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hSysNbr = .FieldHandle(entd_SysNbr)
        m.hSkipRpt = .FieldHandle(entd_SkipRpt)
        m.hSorted = .SortHandle
    End With
    
    m.RptForm.SetToSingleUnit False
    m.RptForm.FilterTrades
    
    'Report options
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strStatToChart = m.RptForm.OptionValue("StatToChart")
    m.dMinProfit = m.RptForm.OptionValue("ProfitGT")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    CalculateStats
    ShowReport
    BuildChart
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSystems.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .XAxisLabel = "Strategies"
        
        SetChartSettings m.Chart, GPM_BAR
        
        Select Case m.strStatToChart
            Case "Trades"
                For lTrade = 0 To lItemsToChart - 1
                    m.Chart.YData(0, lTrade) = m.aTotTrades.Num(lTrade + 1)
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
        For lTrade = 1 To m.aSysNbr.Size - 1
            '.PointLabels(lTrade - 1) = m.aSysNbr.Item(lTrade)
            .PointLabels(lTrade - 1) = lTrade
        Next lTrade
    
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSystems.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .BackColorAlternate = g.nAltGridRowColor ' &HC0FFFF
        .AutoSizeMode = flexAutoSizeColWidth
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .Cols = kGridCols
        .FrozenCols = 1
        m.lSubHeadRow = 0
        
        .ColWidth(GDCol(eGDCol_SysNbr)) = 500
        .ColAlignment(GDCol(eGDCol_SysNbr)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, 0) = "Sys#"
        
        .ColWidth(GDCol(eGDCol_SysName)) = 1200
        .ColAlignment(GDCol(eGDCol_SysName)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SysName)) = "Strategy Name"
        
        .ColWidth(GDCol(eGDCol_TotTrades)) = 800
        .ColAlignment(GDCol(eGDCol_TotTrades)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotTrades)) = "Trades"
        
        .ColWidth(GDCol(eGDCol_WinPct)) = 800
        .ColAlignment(GDCol(eGDCol_WinPct)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WinPct)) = "Win Pct"
        
        .ColWidth(GDCol(eGDCol_WinAvg)) = 800
        .ColAlignment(GDCol(eGDCol_WinAvg)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WinAvg)) = "Win Avg"
        
        .ColWidth(GDCol(eGDCol_PLRatio)) = 600
        .ColAlignment(GDCol(eGDCol_PLRatio)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_PLRatio)) = "PayoutR" '"P/L"
        
        .ColWidth(GDCol(eGDCol_PF)) = 600
        .ColAlignment(GDCol(eGDCol_PF)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_PF)) = "ProfitF" '"P/F"
        
        .ColWidth(GDCol(eGDCol_ReturnPct)) = 800
        .ColAlignment(GDCol(eGDCol_ReturnPct)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ReturnPct)) = "Return%"
        
        .ColWidth(GDCol(eGDCol_AvgTrade)) = 800
        .ColAlignment(GDCol(eGDCol_AvgTrade)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AvgTrade)) = "AvgTrade"
        
        .ColWidth(GDCol(eGDCol_NetProfit)) = 1000
        .ColAlignment(GDCol(eGDCol_NetProfit)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfit)) = "Net Profit"
        
        .ColWidth(GDCol(eGDCol_Drawdown)) = 900
        .ColAlignment(GDCol(eGDCol_Drawdown)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Drawdown)) = "Drawdown"
        
        .ColWidth(GDCol(eGDCol_LossConsec)) = 650
        .ColAlignment(GDCol(eGDCol_LossConsec)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LossConsec)) = "C/L"
        
        .ColWidth(GDCol(eGDCol_LossLargest)) = 900
        .ColAlignment(GDCol(eGDCol_LossLargest)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LossLargest)) = "Lrg Loss"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSystems.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim bFound As Boolean
    
    If m.RptForm.Systems.NumRecords - 1 = 0 Then Exit Sub
    With m.Stats
        .MinProfit = m.dMinProfit
        For lIndex = 1 To m.RptForm.Systems.NumRecords - 1
        
            'Filter Trades by entry signal
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                If gdGetNum(m.hSysNbr, lIndex2) = m.RptForm.Systems.Num(lIndex, ensy_SystemNumber) Then
                    gdSetNum m.hSkipRpt, lIndex2, 0
                Else
                    gdSetNum m.hSkipRpt, lIndex2, 1
                End If
            Next lIndex2
            .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            .SetSortHandle m.RptForm.Trades.SortHandle
            
            .Calculate
            SaveResults lIndex, m.RptForm.Systems.Item(lIndex, ensy_SystemName)
        Next lIndex
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSystems.CalculateStats", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveResults
'' Description: Save the consecutive statistics
'' Inputs:      System Number, System Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveResults(ByVal lSysNbr As Long, ByVal strSysName As String)
On Error GoTo ErrSection

    With m.Stats
        m.lRecord = m.lRecord + 1
        m.aSysNbr.Add lSysNbr, m.lRecord
        m.aSysName.Add strSysName, m.lRecord
        m.aTotTrades.Add .TotTrades, m.lRecord
        m.aWinPct.Add .WinPct, m.lRecord
        m.aWinAvg.Add .WinAvg, m.lRecord
        m.aReturnPct.Add .ReturnPct, m.lRecord
        m.aPLRatio.Add .PLRatio, m.lRecord
        m.aPF.Add .ProfitFactor, m.lRecord
        m.aAvgTrade.Add .AvgTrade, m.lRecord
        m.aNetProfit.Add .TotalNetProfit, m.lRecord
        m.aDrawdown.Add .MaxDrawdownIntra, m.lRecord
        m.aLossConsec.Add .LossConsec, m.lRecord
        m.aLossLargest.Add .LossLargest, m.lRecord
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptSystems.SaveResults", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .Rows = .Rows + m.aSysNbr.Size - 1
        For lIndex = 1 To m.aSysNbr.Size - 1
            lRow = m.lSubHeadRow + lIndex
            .TextMatrix(lRow, GDCol(eGDCol_SysNbr)) = Str(m.aSysNbr.Num(lIndex))
            .TextMatrix(lRow, GDCol(eGDCol_SysName)) = m.aSysName.Item(lIndex)
            .TextMatrix(lRow, GDCol(eGDCol_TotTrades)) = Format(m.aTotTrades.Num(lIndex), "#,##0")
            .TextMatrix(lRow, GDCol(eGDCol_WinPct)) = Format(m.aWinPct.Num(lIndex), "##0.0%")
            .TextMatrix(lRow, GDCol(eGDCol_WinAvg)) = Format(m.aWinAvg.Num(lIndex), FormatDollar(m.strShowCents))
            .TextMatrix(lRow, GDCol(eGDCol_PLRatio)) = Format(m.aPLRatio.Num(lIndex), "##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_PF)) = Format(m.aPF.Num(lIndex), "##0.00")
            .TextMatrix(lRow, GDCol(eGDCol_ReturnPct)) = Format(m.aReturnPct.Num(lIndex), "##0.0%")
            If m.aReturnPct.Num(lIndex) < 0 Then
                m.vsGrid.Cell(flexcpForeColor, lRow, GDCol(eGDCol_ReturnPct)) = vbRed
            Else
                m.vsGrid.Cell(flexcpForeColor, lRow, GDCol(eGDCol_ReturnPct)) = vbDefault
            End If
            .TextMatrix(lRow, GDCol(eGDCol_AvgTrade)) = Format(m.aAvgTrade.Num(lIndex), FormatDollar(m.strShowCents))
            If m.aAvgTrade.Num(lIndex) < 0 Then
                m.vsGrid.Cell(flexcpForeColor, lRow, GDCol(eGDCol_AvgTrade)) = vbRed
            Else
                m.vsGrid.Cell(flexcpForeColor, lRow, GDCol(eGDCol_AvgTrade)) = vbDefault
            End If
            .TextMatrix(lRow, GDCol(eGDCol_NetProfit)) = Format(m.aNetProfit.Num(lIndex), FormatDollar(m.strShowCents))
            If m.aNetProfit.Num(lIndex) < 0 Then
                m.vsGrid.Cell(flexcpForeColor, lRow, GDCol(eGDCol_NetProfit)) = vbRed
            Else
                m.vsGrid.Cell(flexcpForeColor, lRow, GDCol(eGDCol_NetProfit)) = vbDefault
            End If
            .TextMatrix(lRow, GDCol(eGDCol_Drawdown)) = Format(m.aDrawdown.Num(lIndex) * -1, FormatDollar(m.strShowCents))
            .TextMatrix(lRow, GDCol(eGDCol_LossConsec)) = Format(m.aLossConsec.Num(lIndex), "#,##0")
            .TextMatrix(lRow, GDCol(eGDCol_LossLargest)) = Format(m.aLossLargest.Num(lIndex) * -1, FormatDollar(m.strShowCents))
        Next lIndex
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMDDAnalysis.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary
    Set m.aSysNbr = New cGdArray
    Set m.aSysName = New cGdArray
    Set m.aTotTrades = New cGdArray
    Set m.aWinPct = New cGdArray
    Set m.aWinAvg = New cGdArray
    Set m.aReturnPct = New cGdArray
    Set m.aPLRatio = New cGdArray
    Set m.aPF = New cGdArray
    Set m.aAvgTrade = New cGdArray
    Set m.aLossConsec = New cGdArray
    Set m.aLossLargest = New cGdArray
    Set m.aDrawdown = New cGdArray
    Set m.aNetProfit = New cGdArray
    m.aSysNbr.Create eGDARRAY_Longs
    m.aSysName.Create eGDARRAY_Strings
    m.aTotTrades.Create eGDARRAY_Longs
    m.aWinPct.Create eGDARRAY_Doubles
    m.aWinAvg.Create eGDARRAY_Doubles
    m.aReturnPct.Create eGDARRAY_Doubles
    m.aPLRatio.Create eGDARRAY_Doubles
    m.aPF.Create eGDARRAY_Doubles
    m.aAvgTrade.Create eGDARRAY_Doubles
    m.aLossConsec.Create eGDARRAY_Longs
    m.aLossLargest.Create eGDARRAY_Doubles
    m.aDrawdown.Create eGDARRAY_Doubles
    m.aNetProfit.Create eGDARRAY_Doubles

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing

End Sub
