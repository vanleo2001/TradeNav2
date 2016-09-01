VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptAllocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    dMinProfit As Double
    lSubHeadRow As Long
    strReportName As String
    strShowCents As String
    RptForm As Form
    Chart As Pepcoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary
    Results As cGdTable
    lRecord As Long
    hSysNbr As Long
    hSkipRpt As Long
    hSorted As Long
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_SysNbr = 0
    eGDCol_SysName = 1
    eGDCol_Market = 2
    eGDCol_TotTrades = 3
    eGDCol_TotTradesPct = 4
    eGDCol_NetProfit = 5
    eGDCol_NetProfitPct = 6
End Enum
Private Const kGridCols = 7

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
    RaiseError "cRptAllocations.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Set m.Chart = RptForm!PieChart
    Set m.vsGrid = RptForm!vsReport
    
    m.RptForm.CustomColumn = GDCol(eGDCol_SysName)
    
    'Get Field handles
    With m.RptForm.Trades
        m.hSysNbr = .FieldHandle(entd_SysNbr)
        m.hSkipRpt = .FieldHandle(entd_SkipRpt)
        m.hSorted = .SortHandle
    End With
    
    m.RptForm.SetToSingleUnit False
    m.RptForm.FilterTrades
    
    'Report options
    m.dMinProfit = m.RptForm.OptionValue("ProfitGT")
    m.strReportName = m.RptForm.OptionValue("ReportName")
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
    RaiseError "cRptAllocations.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
    If lItemsToChart = 0 Then Exit Sub
    
    With m.Chart
        .PEactions = 20     'Empty chart and reinitialize
        .Subsets = 2
        .Points = lItemsToChart
    
        For lTrade = 0 To lItemsToChart - 1
            m.Chart.XData(0, lTrade) = m.Results.Num(GDCol(eGDCol_TotTrades), lTrade + 1)
        Next lTrade
        For lTrade = 0 To lItemsToChart - 1
            If m.Results.Num(GDCol(eGDCol_NetProfit), lTrade + 1) < 0 Then
                m.Chart.XData(1, lTrade) = 0
            Else
                m.Chart.XData(1, lTrade) = m.Results.Num(GDCol(eGDCol_NetProfit), lTrade + 1)
            End If
        Next lTrade
        
        '** Set Various Properties **'
        'Pego1.DeskColor = RGB(192, 192, 192)
        '.DeskColor = RGB(255, 255, 255)
        .GraphBackColor = RGB(255, 255, 255)
        .GraphForeColor = 0
    
        .MainTitle = m.strReportName
        .SubTitle = ""
        .SubsetLabels(0) = "Total Trades"
        .SubsetLabels(1) = "Net Profit"
        
        '** Set PointLabels property array for 12 points **'
        Dim Col As Byte
        Col = 7
        For lTrade = 1 To m.Results.NumRecords - 1
            .PointLabels(lTrade - 1) = Left(m.Results.Item(GDCol(eGDCol_SysName), lTrade), 15)
            Col = Col + 1
            If Col = 16 Then Col = 8
            .SubsetColors(lTrade - 1) = QBColor(Col)
        Next lTrade
    
        '** Set various other properties **'
        .FocalRect = False
        .DataPrecision = 1
        .GroupingPercent = 3
        .DeskColor = RGB(192, 192, 192)
        .DataShadows = PEDS_3D
        .AutoExplode = PEAE_ALLSUBSETS
        .ShadowColor = .DeskColor
        .AllowCustomization = False
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptAllocations.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

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
        
        .ColWidth(GDCol(eGDCol_SysNbr)) = 700
        .ColAlignment(GDCol(eGDCol_SysNbr)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, 0) = "Sys#"
        
        .ColWidth(GDCol(eGDCol_SysName)) = 2000
        .ColAlignment(GDCol(eGDCol_SysName)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SysName)) = "Strategy Name"
        
        .ColWidth(GDCol(eGDCol_Market)) = 700
        .ColAlignment(GDCol(eGDCol_Market)) = flexAlignLeftTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Market)) = "Symbol"
        
        .ColWidth(GDCol(eGDCol_TotTrades)) = 800
        .ColAlignment(GDCol(eGDCol_TotTrades)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotTrades)) = "Trades"
        
        .ColWidth(GDCol(eGDCol_TotTradesPct)) = 1000
        .ColAlignment(GDCol(eGDCol_TotTradesPct)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotTradesPct)) = "Trades Pct"
        
        .ColWidth(GDCol(eGDCol_NetProfit)) = 1200
        .ColAlignment(GDCol(eGDCol_NetProfit)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfit)) = "Net Profit"
        
        .ColWidth(GDCol(eGDCol_NetProfitPct)) = 1200
        .ColAlignment(GDCol(eGDCol_NetProfitPct)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfitPct)) = "Net Prof Pct"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptAllocations.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStats
'' Description: Calculate the statistics that we will need for this report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStats()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim dTotal As Double                ' Total trades
    
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
            
            m.Results.Num(GDCol(eGDCol_SysNbr), lIndex) = lIndex
            m.Results.Item(GDCol(eGDCol_SysName), lIndex) = m.RptForm.Systems.Item(lIndex, ensy_SystemName)
            m.Results.Item(GDCol(eGDCol_Market), lIndex) = m.RptForm.Systems.Item(lIndex, ensy_Symbol)
            m.Results.Num(GDCol(eGDCol_TotTrades), lIndex) = .TotTrades
            m.Results.Num(GDCol(eGDCol_NetProfit), lIndex) = .TotalNetProfit
        Next lIndex
    End With
    
    'Calculate percentages for Total trades and net profit
    For lIndex = 1 To m.RptForm.Systems.NumRecords - 1
        dTotal = dTotal + m.Results.Num(GDCol(eGDCol_TotTrades), lIndex)
    Next lIndex
    For lIndex = 1 To m.RptForm.Systems.NumRecords - 1
        If dTotal > 0 Then
            m.Results.Num(GDCol(eGDCol_TotTradesPct), lIndex) = m.Results.Num(GDCol(eGDCol_TotTrades), lIndex) / dTotal
        Else
            m.Results.Num(GDCol(eGDCol_TotTradesPct), lIndex) = 0
        End If
    Next lIndex
    
    dTotal = 0
    For lIndex = 1 To m.RptForm.Systems.NumRecords - 1
        dTotal = dTotal + m.Results.Num(GDCol(eGDCol_NetProfit), lIndex)
    Next lIndex
    For lIndex = 1 To m.RptForm.Systems.NumRecords - 1
        If dTotal > 0 Then
            m.Results.Num(GDCol(eGDCol_NetProfitPct), lIndex) = m.Results.Num(GDCol(eGDCol_NetProfit), lIndex) / dTotal
        Else
            m.Results.Num(GDCol(eGDCol_NetProfitPct), lIndex) = 0
        End If
    Next lIndex
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptAllocations.CalculateStats", eGDRaiseError_Raise, m.RptForm.AppPath

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
        For lIndex = 1 To m.Results.NumRecords - 1
            lRow = m.lSubHeadRow + lIndex
            .TextMatrix(lRow, GDCol(eGDCol_SysNbr)) = Str(m.Results.Num(GDCol(eGDCol_SysNbr), lIndex))
            .TextMatrix(lRow, GDCol(eGDCol_SysName)) = m.Results.Item(GDCol(eGDCol_SysName), lIndex)
            .TextMatrix(lRow, GDCol(eGDCol_Market)) = m.Results.Item(GDCol(eGDCol_Market), lIndex)
            .TextMatrix(lRow, GDCol(eGDCol_TotTrades)) = Format(m.Results.Num(GDCol(eGDCol_TotTrades), lIndex), "#,##0")
            .TextMatrix(lRow, GDCol(eGDCol_TotTradesPct)) = Format(m.Results.Num(GDCol(eGDCol_TotTradesPct), lIndex), "##0.0%")
            .TextMatrix(lRow, GDCol(eGDCol_NetProfit)) = Format(m.Results.Num(GDCol(eGDCol_NetProfit), lIndex), FormatDollar(m.strShowCents))
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_NetProfit), lIndex), lRow, GDCol(eGDCol_NetProfit)
            .TextMatrix(lRow, GDCol(eGDCol_NetProfitPct)) = Format(m.Results.Num(GDCol(eGDCol_NetProfitPct), lIndex), "##0.0%")
            ColorNegValue m.vsGrid, m.Results.Num(GDCol(eGDCol_NetProfitPct), lIndex), lRow, GDCol(eGDCol_NetProfitPct)
        Next lIndex
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptAllocations.ShowReport", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

Private Sub Class_Initialize()
    
    Set m.Stats = New cStatsSummary
    Set m.Results = New cGdTable
    
    With m.Results
        .CreateField eGDARRAY_Longs, GDCol(eGDCol_SysNbr), "SysNbr"
        .CreateField eGDARRAY_Strings, GDCol(eGDCol_SysName), "SysName"
        .CreateField eGDARRAY_Strings, GDCol(eGDCol_Market), "Market"
        .CreateField eGDARRAY_Longs, GDCol(eGDCol_TotTrades), "TotTrades"
        .CreateField eGDARRAY_Floats, GDCol(eGDCol_TotTradesPct), "TotTradesPct"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_NetProfit), "NetProfit"
        .CreateField eGDARRAY_Doubles, GDCol(eGDCol_NetProfitPct), "NetProfitPct"
    End With

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing
    Set m.Results = Nothing

End Sub
