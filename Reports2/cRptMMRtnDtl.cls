VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMRtnDtl"
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
    strMonthIncs As String
    lMonthInc As Long
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    Results As cGdTable
    
    hTradeDate As Long
    hAccountBalance As Long
    hSkipRpt As Long
    hSignalType As Long
    hSorted As Long
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Year = 0
    eGDCol_Jan = 1
    eGDCol_Feb = 2
    eGDCol_Mar = 3
    eGDCol_Apl = 4
    eGDCol_May = 5
    eGDCol_Jun = 6
    eGDCol_Jul = 7
    eGDCol_Aug = 8
    eGDCol_Sep = 9
    eGDCol_Oct = 10
    eGDCol_Nov = 11
    eGDCol_Dec = 12
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
    RaiseError "cRptMMRtnDtl.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath

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
        m.hSignalType = .FieldHandle(entd_SignalType)
        m.hSorted = .SortHandle
    End With
    
    'Report options
    m.strReportName = m.RptForm.OptionValue("ReportName")
    m.strStatToChart = m.RptForm.OptionValue("StatToChart")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    m.strMonthIncs = m.RptForm.OptionValue("MonthInc")
    Select Case m.strMonthIncs
        Case "Monthly": m.lMonthInc = 1
        Case "Every 2 Months": m.lMonthInc = 2
        Case "Quarterly": m.lMonthInc = 3
        Case "Tri-annual": m.lMonthInc = 4
        Case "Semi-annual": m.lMonthInc = 6
    End Select
    
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
    RaiseError "cRptMMRtnDtl.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
        
        .ColAlignment(GDCol(eGDCol_Year)) = flexAlignCenterTop
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Year)) = "Year"
        
        For lIndex = 1 To 12 / m.lMonthInc
            .ColAlignment(lIndex) = flexAlignRightTop
        Next lIndex
        
        Select Case m.lMonthInc
            Case 1
                .TextMatrix(m.lSubHeadRow, 1) = "Jan"
                .TextMatrix(m.lSubHeadRow, 2) = "Feb"
                .TextMatrix(m.lSubHeadRow, 3) = "Mar"
                .TextMatrix(m.lSubHeadRow, 4) = "Apr"
                .TextMatrix(m.lSubHeadRow, 5) = "May"
                .TextMatrix(m.lSubHeadRow, 6) = "Jun"
                .TextMatrix(m.lSubHeadRow, 7) = "Jul"
                .TextMatrix(m.lSubHeadRow, 8) = "Aug"
                .TextMatrix(m.lSubHeadRow, 9) = "Sep"
                .TextMatrix(m.lSubHeadRow, 10) = "Oct"
                .TextMatrix(m.lSubHeadRow, 11) = "Nov"
                .TextMatrix(m.lSubHeadRow, 12) = "Dec"
            Case 2
                .TextMatrix(m.lSubHeadRow, 1) = "Jan/Feb"
                .TextMatrix(m.lSubHeadRow, 2) = "Mar/Apr"
                .TextMatrix(m.lSubHeadRow, 3) = "May/Jun"
                .TextMatrix(m.lSubHeadRow, 4) = "Jul/Aug"
                .TextMatrix(m.lSubHeadRow, 5) = "Sep/Oct"
                .TextMatrix(m.lSubHeadRow, 6) = "Nov/Dec"
                .Cols = 8
            Case 3
                .TextMatrix(m.lSubHeadRow, 1) = "Quarter1"
                .TextMatrix(m.lSubHeadRow, 2) = "Quarter2"
                .TextMatrix(m.lSubHeadRow, 3) = "Quarter3"
                .TextMatrix(m.lSubHeadRow, 4) = "Quarter4"
                .Cols = 6
            Case 4
                .TextMatrix(m.lSubHeadRow, 1) = "Jan-Apr"
                .TextMatrix(m.lSubHeadRow, 2) = "May-Aug"
                .TextMatrix(m.lSubHeadRow, 3) = "Sep-Dec"
                .Cols = 5
            Case 6
                .TextMatrix(m.lSubHeadRow, 1) = "1st Half"
                .TextMatrix(m.lSubHeadRow, 2) = "2nd Half"
                .Cols = 4
        End Select
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMRtnDtl.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    Dim TradeActivityFound  As Byte
    Dim X               As Long
    Dim EOM             As cGdTable
    Dim lRedraw         As Long
    Dim Col             As Long
    Dim svYear          As Long
    Dim ReturnPct       As Double
    Dim Stats           As cStatsSummary
    
    Const C_TradeDate = 0
    Const C_AccBal = 1
    Const C_TradesOccured = 2
    Const C_ReturnPct = 3
    
    'Build the end of month balances table
    Set Stats = New cStatsSummary
    With Stats
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.RptForm.Trades.SortHandle
        .CalcEndOfMthBalances
        Set EOM = New cGdTable
        EOM.CopyFromHandle .MthReturnValuesHandle
    End With
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        svYear = Year(EOM.Num(C_TradeDate, 1))
        Col = m.lMonthInc
        
        For X = 1 To EOM.NumRecords - 1
            If X + m.lMonthInc > EOM.NumRecords - 1 Then Exit For
            
            'If this is a new year then insert a row in grid
            If Year(EOM.Num(C_TradeDate, X)) <> svYear Or X = 1 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Year)) = Format(Year(EOM.Num(C_TradeDate, X)), "##0")
                Col = 1
                svYear = Year(EOM.Num(C_TradeDate, X))
            End If
            
            'Calculate return pct between months
            If EOM.Num(C_TradesOccured, X) = 0 Then
                .TextMatrix(.Rows - 1, Col) = ""
            Else
                If EOM.Num(C_AccBal, X) <> 0 Then
                    ReturnPct = (EOM.Num(C_AccBal, X + m.lMonthInc - 1) - EOM.Num(C_AccBal, X - 1)) / EOM.Num(C_AccBal, X - 1)
                Else
                    ReturnPct = 0
                End If
                .TextMatrix(.Rows - 1, Col) = Format(ReturnPct, "#,##0.00%")
                ColorNegValue m.vsGrid, ReturnPct, .Rows - 1, Col
            End If
    
            'Add result to row in grid
            Col = Col + 1
            
            X = X + m.lMonthInc - 1
        Next X
    
        ' Remove grid rows of Years prior to start date and after end date
        For X = .Rows - 1 To 1 Step -1
            If .Cell(flexcpValue, X, 0) < Year(m.RptForm.FromDate) Or _
               .Cell(flexcpValue, X, 0) > Year(m.RptForm.ToDate) Then
                .RemoveItem X
            End If
        Next X
        
        .AutoSize 0, .Cols - 1, False, 75
        m.RptForm.MinColWidth = .ColWidth(m.RptForm.CustomColumn)
        m.RptForm.ExtendCustomColumn m.vsGrid
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMRtnDtl.CalculateStats", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub
