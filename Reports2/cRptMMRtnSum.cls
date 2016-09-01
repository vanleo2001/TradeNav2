VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMRtnSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type mPrivate
    lSubHeadRow As Long
    strStatToChart As String
    strReportName As String
    strShowCents As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Balances As cGdTable
    Table As cGdTable
    
    hTradeDate As Long
    hAccountBalance As Long
    hSorted As Long
End Type
Private m As mPrivate

Private Enum eGDBalField
    eGDBalField_TradeDate = 0
    eGDBalField_AccBal = 1
    eGDBalField_TradesOccured = 2
    eGDBalField_ReturnPct = 3
End Enum

Private Enum eGDCols
    eGDCol_Months = 0
    eGDCol_Best = 1
    eGDCol_BestStart = 2
    eGDCol_BestEnd = 3
    eGDCol_Worst = 4
    eGDCol_WorstStart = 5
    eGDCol_WorstEnd = 6
    eGDCol_Avg = 7
    'eGDCol_Std = 8
    eGDCol_Latest = 8
End Enum
Private Const kGridCols = 9

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
    RaiseError "cRptMMRtnSum.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    m.RptForm.CustomColumn = GDCol(eGDCol_Months)
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hTradeDate = .FieldHandle(entd_TradeDate)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hSorted = .SortHandle
    End With
    
    'Report options
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
    m.vsGrid.Redraw = flexRDBuffered
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMRtnSum.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
        .ExtendLastCol = False 'True
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
        
        .ColAlignment(GDCol(eGDCol_Months)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Months)) = "Months"
        
        .ColAlignment(GDCol(eGDCol_Best)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Best)) = "Best"
        
        .ColAlignment(GDCol(eGDCol_BestStart)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_BestStart)) = "Start"
        
        .ColAlignment(GDCol(eGDCol_BestEnd)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_BestEnd)) = "End"
        
        .ColAlignment(GDCol(eGDCol_Worst)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Worst)) = "Worst"
        
        .ColAlignment(GDCol(eGDCol_WorstStart)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WorstStart)) = "Start"
        
        .ColAlignment(GDCol(eGDCol_WorstEnd)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_WorstEnd)) = "End"
        
        .ColAlignment(GDCol(eGDCol_Avg)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Avg)) = "Avg"
        
        '.ColAlignment(GDCol(eGDCol_Std)) = flexAlignRightTop
        '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Std)) = "StDev"
        
        .ColAlignment(GDCol(eGDCol_Latest)) = flexAlignRightTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Latest)) = "Recent"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMRtnSum.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Dim R               As Long
    Dim RetVal          As Long
    Dim curMonth        As Byte
    Dim NextMonth       As Byte
    Dim Periods()       As Byte
    Dim Best            As Double
    Dim Worst           As Double
    Dim Avg             As Double
    Dim Latest          As Double
    Dim BestStart       As Double
    Dim BestEnd         As Double
    Dim WorstStart      As Double
    Dim WorstEnd        As Double
    Dim curYear         As Long
    Dim Nextyear        As Long
    Dim Continue        As Boolean
    Dim lRedraw         As Long
    Dim Stats           As cStatsSummary
    
    'Build the end of month balances table
    Set Stats = New cStatsSummary
    With Stats
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        .CalcEndOfMthBalances
        Set m.Balances = New cGdTable
        m.Balances.CopyFromHandle .MthReturnValuesHandle
    End With
    
    ReDim Periods(16)
    Periods(1) = 1
    Periods(2) = 2
    Periods(3) = 3
    Periods(4) = 4
    Periods(5) = 5
    Periods(6) = 6
    Periods(7) = 7
    Periods(8) = 8
    Periods(9) = 9
    Periods(10) = 10
    Periods(11) = 11
    Periods(12) = 12
    Periods(13) = 18
    Periods(14) = 24
    Periods(15) = 36
    Periods(16) = 48
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For X = 1 To UBound(Periods)
            CalcReturn Periods(X), Best, BestStart, BestEnd, _
                Worst, WorstStart, WorstEnd, Avg, Latest
            
            'Add result to row in grid
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Months)) = Str(Periods(X))
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Best)) = Format(Best, "#,##0.0%")
            ColorNegValue m.vsGrid, Best, .Rows - 1, GDCol(eGDCol_Best)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_BestStart)) = Format(BestStart, "mmm-yyyy")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_BestEnd)) = Format(BestEnd, "mmm-yyyy")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Worst)) = Format(Worst, "#,##0.0%")
            ColorNegValue m.vsGrid, Worst, .Rows - 1, GDCol(eGDCol_Worst)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_WorstStart)) = Format(WorstStart, "mmm-yyyy")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_WorstEnd)) = Format(WorstEnd, "mmm-yyyy")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Avg)) = Format(Avg, "#,##0.0%")
            ColorNegValue m.vsGrid, Avg, .Rows - 1, GDCol(eGDCol_Avg)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Latest)) = Format(Latest, "#,##0.0%")
            ColorNegValue m.vsGrid, Latest, .Rows - 1, GDCol(eGDCol_Latest)
        Next X
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMRtnSum.CalculateStats", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcReturn
'' Description: Calculate the statistics
'' Inputs:      Stats to calculate
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcReturn(pMths As Byte, _
    Best As Double, BestStart As Double, BestEnd As Double, _
    Worst As Double, WorstStart As Double, WorstEnd As Double, _
    Avg As Double, Latest As Double)
On Error GoTo ErrSection:
    
    Dim X           As Long
    Dim ReturnPct   As Double
    Dim TotalRet    As Double
    Dim RetCnt      As Long
    
    Best = -999999
    Worst = 999999
    For X = 1 To m.Balances.NumRecords - 1
        If X + pMths > m.Balances.NumRecords - 1 Then Exit For
        
        'Calculate return pct between months
        If m.Balances.Num(eGDBalField_TradeDate, X - 1) <> 0 Then
            ReturnPct = (m.Balances.Num(eGDBalField_AccBal, X + pMths - 1) - m.Balances.Num(eGDBalField_AccBal, X - 1)) / _
                m.Balances.Num(eGDBalField_AccBal, X - 1)
        Else
            ReturnPct = 0
        End If
        
        If ReturnPct > Best Then
            Best = ReturnPct
            BestStart = m.Balances.Num(eGDBalField_TradeDate, X - 1)
            BestEnd = m.Balances.Num(eGDBalField_TradeDate, X + pMths - 1)
        End If
        If ReturnPct < Worst Then
            Worst = ReturnPct
            WorstStart = m.Balances.Num(eGDBalField_TradeDate, X - 1)
            WorstEnd = m.Balances.Num(eGDBalField_TradeDate, X + pMths - 1)
        End If
        
        RetCnt = RetCnt + 1
        TotalRet = TotalRet + ReturnPct
    Next X
    
    If RetCnt <> 0 Then
        Avg = TotalRet / RetCnt
    End If
    Latest = ReturnPct
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMRtnSum.CalcReturn", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub
