VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMDDAnalysis"
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
    lPeriods As Long
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Stats As cStatsSummary
    Results As cGdTable
    
    hTradeDate As Long
    hAccountBalance As Long
    hSkipRpt As Long
    hSorted As Long
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Depth = 0
    eGDCol_Length = 1
    eGDCol_Recovery = 2
    eGDCol_StartMth = 3
    eGDCol_EndMth = 4
    eGDCol_Trades = 5
    eGDCol_WinPct = 6
    eGDCol_AvgTrade = 7
    eGDCol_WinAvg = 8
    eGDCol_PL = 9
    eGDCol_PF = 10
    eGDCol_NetProfit = 11
    eGDCol_CL = 12
    eGDCol_MaxLoss = 13
End Enum
Private Const kGridCols = 14

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
    RaiseError "cRptMMDDAnalysis.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    m.RptForm.CustomColumn = GDCol(eGDCol_Depth)
    
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
    m.lPeriods = m.RptForm.OptionValue("Periods")
    
    'Autosense mode (money mgt vs single unit)
    If gdGetNum(m.hAccountBalance, 1) <= 0 Then
        m.RptForm.SetToSingleUnit True
    End If
    
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    CalculateStats
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMDDAnalysis.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
        
        .ColAlignment(GDCol(eGDCol_Depth)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Depth)) = "Depth%"
        
        .ColAlignment(GDCol(eGDCol_Length)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Length)) = "Months"
        
        .ColAlignment(GDCol(eGDCol_Recovery)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Recovery)) = "Recovery"
        
        .ColAlignment(GDCol(eGDCol_StartMth)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_StartMth)) = "Start"
        
        .ColAlignment(GDCol(eGDCol_EndMth)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EndMth)) = "End"
        
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
        
        .ColAlignment(GDCol(eGDCol_NetProfit)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfit)) = "Net Profit"
        
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
    RaiseError "cRptMMDDAnalysis.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    Dim X               As Long
    Dim Y               As Long
    Dim z               As Long
    Dim BeginDate       As Double
    Dim EndDate         As Double
    Dim EquityPeak      As Double
    Dim Valley          As Double
    Dim StartMth        As Double
    Dim ValleyMth       As Double
    Dim ValleyBegBal    As cGdArray
    Dim R               As Long
    Dim SortInd         As cGdArray
    Dim ptr             As Long
    Dim RetVal          As Long
    Dim lRedraw         As Long
    Dim lIndex          As Long
    
    'set to single unit
    If gdGetNum(m.hAccountBalance, 1) <= 0 Then
        m.RptForm.SetToSingleUnit True
    End If
    
    'Initialize first drawdown period
    Set ValleyBegBal = New cGdArray
    ValleyBegBal.Create eGDARRAY_Doubles
    StartMth = gdGetNum(m.hTradeDate, 1)
    Valley = gdGetNum(m.hAccountBalance, 0)
    EquityPeak = gdGetNum(m.hAccountBalance, 0)
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone

        For lIndex = 1 To m.RptForm.Trades.NumRecords - 1
            X = gdGetNum(m.hSorted, lIndex)
        
            'Check for new equity peak when a loss occurs.
            If gdGetNum(m.hAccountBalance, X) > EquityPeak Then
            
                'Record peak to valley percentage change and distance
                'in mths from start to end
                If Valley < EquityPeak Then
                    R = R + 1
                    m.Results.Num(GDCol(eGDCol_StartMth), R) = StartMth
                    m.Results.Num(GDCol(eGDCol_Depth), R) = (EquityPeak - Valley) / EquityPeak
                    m.Results.Num(GDCol(eGDCol_EndMth), R) = gdGetNum(m.hTradeDate, X - 1)
                    m.Results.Num(GDCol(eGDCol_Length), R) = DateDiff("D", StartMth, ValleyMth) / 30
                    m.Results.Num(GDCol(eGDCol_Recovery), R) = DateDiff("D", ValleyMth, gdGetNum(m.hTradeDate, X - 1)) / 30
                    ValleyBegBal.Add Valley, R
                End If
                
                'Initialize a new drawdown period
                StartMth = gdGetNum(m.hTradeDate, X)
                Valley = gdGetNum(m.hAccountBalance, X)
                EquityPeak = gdGetNum(m.hAccountBalance, X)
            Else
                If gdGetNum(m.hAccountBalance, X) < Valley Then
                    Valley = gdGetNum(m.hAccountBalance, X)
                    ValleyMth = gdGetNum(m.hTradeDate, X)
                End If
            End If
            
        Next lIndex
        
        'Sort table by depth
        Set SortInd = New cGdArray
        Set SortInd = m.Results.CreateIndex(GDCol(eGDCol_Depth))
        m.Results.SortIndex SortInd, GDCol(eGDCol_Depth), eGdSort_Descending
        
        For Y = 0 To SortInd.Size - 1
            ptr = SortInd.Num(Y)
            
            'Only process top n requests
            If z < m.lPeriods Then
                z = z + 1
                BeginDate = m.Results.Num(GDCol(eGDCol_StartMth), ptr)
                EndDate = m.Results.Num(GDCol(eGDCol_EndMth), ptr)
                
                'Process all trades for current year
                For X = 1 To m.RptForm.Trades.NumRecords - 1
                    If gdGetNum(m.hTradeDate, X) >= BeginDate And _
                       gdGetNum(m.hTradeDate, X) <= EndDate Then
                        gdSetNum m.hSkipRpt, X, 0
                    Else
                        gdSetNum m.hSkipRpt, X, 1
                    End If
                Next X
                
                gdSetNum m.hAccountBalance, 0, ValleyBegBal.Num(Y)
                
                With m.Stats
                    .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
                    .SetSortHandle m.RptForm.Trades.SortHandle
                    .Calculate
                End With
            
                'Add result to row in grid
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Depth)) = Format(m.Results.Num(GDCol(eGDCol_Depth), ptr), "#,##0.0%")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Length)) = Format(m.Results.Num(GDCol(eGDCol_Length), ptr), "#,##0.0")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Recovery)) = Format(m.Results.Num(GDCol(eGDCol_Recovery), ptr), "#,##0.0")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_StartMth)) = Format(m.Results.Num(GDCol(eGDCol_StartMth), ptr), "mmm-yyyy")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_EndMth)) = Format(m.Results.Num(GDCol(eGDCol_EndMth), ptr), "mmm-yyyy")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Trades)) = Format(m.Stats.TotTrades, "#,##0")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_WinPct)) = Format(m.Stats.WinPct, "##0.0%")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_WinAvg)) = Format(m.Stats.WinAvg, FormatDollar(m.strShowCents))
                .TextMatrix(.Rows - 1, GDCol(eGDCol_PL)) = Format(m.Stats.PLRatio, "##0.00")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_PF)) = Format(m.Stats.ProfitFactor, "##0.00")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_AvgTrade)) = Format(m.Stats.AvgTrade, FormatDollar(m.strShowCents))
                ColorNegValue m.vsGrid, m.Stats.AvgTrade, .Rows - 1, GDCol(eGDCol_AvgTrade)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_NetProfit)) = Format(m.Stats.TotalNetProfit, FormatDollar(m.strShowCents))
                ColorNegValue m.vsGrid, m.Stats.TotalNetProfit, .Rows - 1, GDCol(eGDCol_NetProfit)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_CL)) = Format(m.Stats.LossConsec, "#,##0")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_MaxLoss)) = Format(m.Stats.LossLargest, FormatDollar(m.strShowCents))
                ColorNegValue m.vsGrid, m.Stats.LossLargest, .Rows - 1, GDCol(eGDCol_MaxLoss)
            End If
        Next Y
    
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With
    
ErrExit:
    Set SortInd = Nothing
    Exit Sub

ErrSection:
    RaiseError "cRptMMDDAnalysis.CalculateStats", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Dim RetVal As Variant
    
    Set m.Stats = New cStatsSummary
    Set m.Results = New cGdTable
    
    With m.Results
        RetVal = .CreateField(eGDARRAY_Doubles, 0, "Depth")
        RetVal = .CreateField(eGDARRAY_Floats, 1, "Length")
        RetVal = .CreateField(eGDARRAY_Floats, 2, "Recovery")
        RetVal = .CreateField(eGDARRAY_Doubles, 3, "StartMth")
        RetVal = .CreateField(eGDARRAY_Doubles, 4, "EndMth")
        RetVal = .CreateField(eGDARRAY_Longs, 5, "TotTrades")
        RetVal = .CreateField(eGDARRAY_Doubles, 6, "WinPct")
        RetVal = .CreateField(eGDARRAY_Doubles, 7, "AvgTrade")
        RetVal = .CreateField(eGDARRAY_Doubles, 8, "WinAvg")
        RetVal = .CreateField(eGDARRAY_Doubles, 9, "PLRatio")
        RetVal = .CreateField(eGDARRAY_Doubles, 10, "PF")
        RetVal = .CreateField(eGDARRAY_Doubles, 11, "NetProfit")
        RetVal = .CreateField(eGDARRAY_Shorts, 12, "LossConsec")
        RetVal = .CreateField(eGDARRAY_Doubles, 13, "LossLargest")
    End With

End Sub

Private Sub Class_Terminate()
    
    Set m.Stats = Nothing
    Set m.Results = Nothing

End Sub
