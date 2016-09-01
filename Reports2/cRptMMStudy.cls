VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Implements IVSFlexDataSource

Private Type mPrivate
    aIndex As cGdArray
    alGraphToTrades As cGdArray

    'Report Options
    dScaleMult As Double
    dBeginBalance As Double
    lUnitSize As Long
    lStartUnits As Long
    lMaxUnits As Long
    strPriceFormat As String
    strShowTime As String
    strReportName As String
    strShowCents As String
    strSubClass As String
    strChartBy As String
    strView As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    fgMonthly As VSFlexGrid
    fgYearly As VSFlexGrid
    
    strDefaultFields As String
    strDefaultFieldsTrade As String
    strDefaultFieldsMonthly As String
    strDefaultFieldsYearly As String
    strFields As String
    strFieldsTrade As String
    strFieldsMonthly As String
    strFieldsYearly As String
    lAppID As Long
    bAutoSize As Boolean
    bAutoSizeTrade As Boolean
    bAutoSizeMonthly As Boolean
    bAutoSizeYearly As Boolean
    
    hSymbolHandle As Long

    'Local to report
    Stats As cStatsSummary
    lSubHeadRow As Long
    
    strRisk As String
    dRiskUp As Double
    dRiskDown As Double
    dDelta As Double
    dMinProfit As Double
    dRiskPerUnit As Double
    cAccBal As Currency
    dRiskPct As Double
    dLargestLosing As Double
    dMargin As Double
    
    ' TLB 3/25/2011 & 7/18/2011
    dRiskBasedOn As Double ' 0=StopLoss, -1=MaxLoss, -2=AvgLoss, >0 means user-specified $ amount
    aRiskBasedOn As cGdArray ' to keep track of RiskBasedOn (e.g. Max or Avg Loss) for each basket item
    aMargins As cGdArray ' to store the margin for each basket item
    dDelayedStartDate As Double ' used to delay the start of money management trading (ignore trades before this)
    
    aOpen As cGdArray
    aUnits As cGdArray
    aTotProfit As cGdArray
    aDrawDownPct As cGdArray
    aIncLevel As cGdArray
    aDecLevel As cGdArray
    aAccBalance As cGdArray
    aOptimalf As cGdArray
    
    lUnitsNextTrade As Long

    'Trade field handles
    hTradeNum As Long
    hTradeDate As Long
    hPosition As Long
    hSignalType As Long
    hEntryExitPtr As Long
    hProfit As Long
    hSystemNumber As Long
    hSymbolIndex As Long
    hSkip As Long
    hSkipEq As Long
    hShow As Long
    hSkipReport As Long
    hSignalsTotal As Long
    hOpenTradesTotal As Long
    hAccountBalance As Long
    hSorted As Long
    hUnits As Long
    hHeaderIndex As Long
End Type
Private m As mPrivate

Private Enum eGD_Cols
    eGDCol_Number = 0
    eGDCol_Date
    eGDCol_EntryType
    eGDCol_Position
    eGDCol_EntryDate
    eGDCol_ExitDate
    eGDCol_SystemNumber
    eGDCol_Pointer
    eGDCol_Symbol
    eGDCol_OpenPosition
    eGDCol_RiskPercent
    eGDCol_LargestLosing
    eGDCol_Profit
    eGDCol_Units
    eGDCol_TotalProfit
    eGDCol_Balance
    eGDCol_DecLevel
    eGDCol_IncLevel
    eGDCol_Skip
    eGDCol_MaxDip
    eGDCol_OptimalF
    eGDCol_SkipEq
    eGDCol_NumCols
End Enum

Private Enum eGDMonthlyCols
    eGDMonthlyCol_Month = 0
    eGDMonthlyCol_TotTrades
    eGDMonthlyCol_WinPct
    eGDMonthlyCol_WinAvg
    eGDMonthlyCol_LossAvg
    eGDMonthlyCol_AvgRunUp
    eGDMonthlyCol_AvgRunDown
    eGDMonthlyCol_PLRatio
    eGDMonthlyCol_PF
    eGDMonthlyCol_AvgTrade
    eGDMonthlyCol_LossConsec
    eGDMonthlyCol_LossLargest
    eGDMonthlyCol_Profit
    eGDMonthlyCol_NetProfit
    eGDMonthlyCol_AvgNumBars
    eGDMonthlyCol_ROI
    eGDMonthlyCol_NumCols
End Enum

Private Enum eGDYearlyCols
    eGDYearlyCol_Year = 0
    eGDYearlyCol_TotTrades
    eGDYearlyCol_WinPct
    eGDYearlyCol_WinAvg
    eGDYearlyCol_LossAvg
    eGDYearlyCol_AvgRunUp
    eGDYearlyCol_AvgRunDown
    eGDYearlyCol_PLRatio
    eGDYearlyCol_PF
    eGDYearlyCol_AvgTrade
    eGDYearlyCol_LossConsec
    eGDYearlyCol_LossLargest
    eGDYearlyCol_Profit
    eGDYearlyCol_NetProfit
    eGDYearlyCol_AvgNumBars
    eGDYearlyCol_ROI
    eGDYearlyCol_NumCols
End Enum

Private Function GDCol(ByVal lColumn As eGD_Cols) As Long
    GDCol = lColumn
End Function
Private Function MonthCol(ByVal Col As eGDMonthlyCols) As Long
    MonthCol = Col
End Function
Private Function YearCol(ByVal Col As eGDYearlyCols) As Long
    YearCol = Col
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    
    IVSFlexDataSource_GetFieldCount = GDCol(eGDCol_NumCols) + 1

End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    
    IVSFlexDataSource_GetRecordCount = m.aIndex.Size + 1 'm.RptForm.Trades.NumRecords - 1

End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
End Function

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal R As Long) As String
On Error GoTo ErrSection:
    
    Dim lExit As Long
    
    'Set 1st row to beginning balance row
    If R = 0 Then
        If Field = GDCol(eGDCol_Balance) Then
            IVSFlexDataSource_GetData = Format(m.dBeginBalance, FormatDollar(m.strShowCents))
        Else
            IVSFlexDataSource_GetData = ""
        End If
        Exit Function
    End If
    
    R = m.aIndex(R - 1)
    lExit = gdGetNum(m.hEntryExitPtr, R)
    
    Select Case Field
        Case GDCol(eGDCol_Number)
            IVSFlexDataSource_GetData = R
        Case GDCol(eGDCol_Date)
            If gdGetNum(m.hSignalsTotal, R) > 1 Then
                IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R))) & " " & Format(gdGetNum(m.hTradeDate, R), "hh:mm AM/PM")
            Else
                If m.strShowTime = "Yes" Then
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R))) & " " & Format(gdGetNum(m.hTradeDate, R), "hh:mm AM/PM")
                Else
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R)))
                End If
            End If
        Case GDCol(eGDCol_Position)
            If gdGetNum(m.hPosition, R) = 1 Then
                IVSFlexDataSource_GetData = "Long"
            Else
                IVSFlexDataSource_GetData = "Short"
            End If
        Case GDCol(eGDCol_EntryDate)
            If gdGetNum(m.hSignalsTotal, R) > 1 Then
                IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R))) & " " & Format(gdGetNum(m.hTradeDate, R), "hh:mm AM/PM")
            Else
                If m.strShowTime = "Yes" Then
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R))) & " " & Format(gdGetNum(m.hTradeDate, R), "hh:mm AM/PM")
                Else
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R)))
                End If
            End If
        Case GDCol(eGDCol_ExitDate)
            If gdGetNum(m.hSignalsTotal, lExit) > 1 Then
                IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, lExit))) & " " & Format(gdGetNum(m.hTradeDate, lExit), "hh:mm AM/PM")
            Else
                If m.strShowTime = "Yes" Then
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, lExit))) & " " & Format(gdGetNum(m.hTradeDate, lExit), "hh:mm AM/PM")
                Else
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, lExit)))
                End If
            End If
        Case GDCol(eGDCol_OpenPosition)
            IVSFlexDataSource_GetData = Str(m.aOpen.Num(R))
        Case GDCol(eGDCol_EntryType)
            If gdGetNum(m.hSignalType, R) = gEntrySignal Then
                IVSFlexDataSource_GetData = "Entry"
            Else
                IVSFlexDataSource_GetData = "Exit"
            End If
        Case GDCol(eGDCol_Pointer)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hEntryExitPtr, R))
        Case GDCol(eGDCol_SystemNumber)
            IVSFlexDataSource_GetData = Str(m.RptForm.SystemArrayInd(gdGetNum(m.hSystemNumber, R)))
        Case GDCol(eGDCol_Symbol)
            IVSFlexDataSource_GetData = gdGetStr(m.hSymbolHandle, gdGetNum(m.hSymbolIndex, R))
        Case GDCol(eGDCol_RiskPercent)
            IVSFlexDataSource_GetData = Format(m.dRiskPct, "0.00%")
        Case GDCol(eGDCol_LargestLosing)
            If UCase(m.strSubClass) = "STANDARD" Then
                R = gdGetNum(m.hHeaderIndex, R)
                IVSFlexDataSource_GetData = Format(m.aRiskBasedOn.Num(R), FormatDollar(m.strShowCents))
            Else
                IVSFlexDataSource_GetData = Format(m.dLargestLosing, FormatDollar(m.strShowCents))
            End If
        Case GDCol(eGDCol_Profit)
            If m.strView = "Trade" Then
                IVSFlexDataSource_GetData = Format(gdGetNum(m.hProfit, lExit), FormatDollar(m.strShowCents))
            Else
                If gdGetNum(m.hSignalType, R) = gEntrySignal Then
                    IVSFlexDataSource_GetData = Format(0, FormatDollar(m.strShowCents))
                Else
                    IVSFlexDataSource_GetData = Format(gdGetNum(m.hProfit, R), FormatDollar(m.strShowCents))
                End If
            End If
        Case GDCol(eGDCol_Units)
            IVSFlexDataSource_GetData = Format(m.aUnits.Num(R), "#,##0")
        Case GDCol(eGDCol_TotalProfit)
            If m.strView = "Trade" Then
                IVSFlexDataSource_GetData = Format(m.aTotProfit.Num(lExit), FormatDollar(m.strShowCents))
            Else
                IVSFlexDataSource_GetData = Format(m.aTotProfit.Num(R), FormatDollar(m.strShowCents))
            End If
        Case GDCol(eGDCol_Balance)
            If m.strView = "Trade" Then
                IVSFlexDataSource_GetData = Format(m.aAccBalance.Num(lExit), FormatDollar(m.strShowCents))
            Else
                IVSFlexDataSource_GetData = Format(m.aAccBalance.Num(R), FormatDollar(m.strShowCents))
            End If
        Case GDCol(eGDCol_MaxDip)
            If m.strView = "Trade" Then
                IVSFlexDataSource_GetData = Format(m.aDrawDownPct.Num(lExit), "0.00%")
            Else
                IVSFlexDataSource_GetData = Format(m.aDrawDownPct.Num(R), "0.00%")
            End If
        Case GDCol(eGDCol_DecLevel)
            IVSFlexDataSource_GetData = Format(m.aDecLevel.Num(R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_IncLevel)
            IVSFlexDataSource_GetData = Format(m.aIncLevel.Num(R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_Skip)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hSkip, R))
        Case GDCol(eGDCol_SkipEq)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hSkipEq, R))
        Case GDCol(eGDCol_OptimalF)
            IVSFlexDataSource_GetData = Format(m.aOptimalf.Num(R), "0.00%")
    End Select
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cRptMMStudy.GetData", eGDRaiseError_Raise, m.RptForm.AppPath

End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
End Sub
Public Sub BeforeScrollTip(ByVal Row As Long)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AfterScroll
'' Description: This routine formats the current page of grid rows
'' Inputs:      Old Top Row and Left Col, New Top Row and Left Col, Form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, _
    ByVal NewTopRow As Long, ByVal NewLeftCol As Long, ByVal pForm As Object)
On Error GoTo ErrSection:

    ''If Not pForm Is Nothing Then ResizeColWidths
    SetBackColor
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.AfterScroll", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResizeColWidths
'' Description: Resize the column widths using the text on the current page
'' Inputs:      Form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResizeColWidths()
On Error GoTo ErrSection:
    
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strText As String               ' Text from the current cell
    Dim lWidth As Long                  ' New width of the cell
    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        If .Rows > 1 Then '.FixedRows Then
            If .TopRow <> 0 Then
                For lCol = .LeftCol To .RightCol
                    strText = .Cell(flexcpTextDisplay, 0, lCol)
                    lWidth = m.RptForm.TextWidth(strText) + 180
                    If lWidth > .ColWidth(lCol) Then .ColWidth(lCol) = lWidth
                Next lCol
            End If
            
            For lRow = .TopRow To .BottomRow
                For lCol = .LeftCol To .RightCol
                    strText = .Cell(flexcpTextDisplay, lRow, lCol)
                    lWidth = m.RptForm.TextWidth(strText) + 180
                    If lWidth > .ColWidth(lCol) Then .ColWidth(lCol) = lWidth
                Next lCol
            Next lRow
        End If
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.ResizeColWidths", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetColor
'' Description: Set the color of the text according to the sign of the value
'' Inputs:      Row and Column of cell to edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetColor(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.vsGrid.Cell(flexcpValue, Row, Col) < 0 Then
        m.vsGrid.Cell(flexcpForeColor, Row, Col) = vbRed
    Else
        m.vsGrid.Cell(flexcpForeColor, Row, Col) = vbDefault
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.SetColor", eGDRaiseError_Raise, m.RptForm.AppPath
    
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
    
    Dim s$, lIndex As Long                 ' Index into a for loop
    
    Set m.RptForm = RptForm
    Set m.Chart = RptForm!BarChart
    Set m.vsGrid = RptForm!vsImplements
    Set m.fgMonthly = RptForm!fgMonthly
    Set m.fgYearly = RptForm!fgYearly
    
    RptForm.ShowViewOptions = True
    
    ' Get some information from the form...
    m.strDefaultFields = m.RptForm.OptionValue("DefaultFields")
    m.strDefaultFieldsTrade = m.RptForm.OptionValue("DefaultFieldsTrade")
    m.strDefaultFieldsMonthly = m.RptForm.OptionValue("DefaultFieldsMonthly")
    m.strDefaultFieldsYearly = m.RptForm.OptionValue("DefaultFieldsYearly")
    m.lAppID = m.RptForm.ReportID
    
    m.hSymbolHandle = m.RptForm.Trades.SymbolHandle
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hTradeNum = .FieldHandle(entd_TradeNbr)
        m.hPosition = .FieldHandle(entd_Position)
        m.hSignalType = .FieldHandle(entd_SignalType)
        m.hTradeDate = .FieldHandle(entd_TradeDate)
        m.hProfit = .FieldHandle(entd_Profit)
        m.hSymbolIndex = .FieldHandle(entd_SymbolIndex)
        m.hSystemNumber = .FieldHandle(entd_SysNbr)
        m.hSkip = .FieldHandle(entd_Skip)
        m.hSkipEq = .FieldHandle(entd_SkipEqFilter)
        m.hSkipReport = .FieldHandle(entd_SkipRpt)
        m.hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        m.hOpenTradesTotal = .FieldHandle(entd_OpenTradesTotal)
        m.hSignalsTotal = .FieldHandle(entd_SignalsTot)
        m.hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hShow = .FieldHandle(entd_Show)
        m.hSorted = .SortHandle
        m.hUnits = .FieldHandle(entd_Units)
        m.hHeaderIndex = .FieldHandle(entd_HeaderIndex)
    End With
    
    m.strSubClass = m.RptForm.OptionValue("SubClass")
    m.strView = m.RptForm.OptionValue("ViewBy")
    
    ' TLB 7/18/2011: store the margin for each basket item
    m.aMargins.Size = gdNumRecords(m.RptForm.Trades.HdrTableHandle)
    For lIndex = 0 To m.aMargins.Size - 1
        m.aMargins.Num(lIndex) = gdGetTableNum(m.RptForm.Trades.HdrTableHandle, enth_Margin, lIndex)
    Next
    
    If m.strView = "Trade" And m.RptForm.Pyramiding Then
        With m.vsGrid
            .Redraw = flexRDNone
            ClearGrid m.vsGrid
            .Rows = 1
            .Cols = 1
            .FixedCols = 0
            .FixedRows = 0
            .ExtendLastCol = True
            .TextMatrix(0, 0) = "This report cannot currently be run when pyramiding.  " & _
                    "Please change the 'View By' option to 'Transaction'."
            .Redraw = flexRDBuffered
        End With
    Else
        'If portfolio trades exist with money mgt and some trades are marked
        'as skipped then ignore this since this is a custom money mgt report
        'trades could have been skipped in the money mgt rules.  That's way
        'this is necessary
        For lIndex = 1 To m.RptForm.Trades.NumRecords - 1
            gdSetNum m.hSkip, lIndex, 0
            gdSetNum m.hSkipReport, lIndex, 0
            gdSetNum m.hShow, lIndex, 1
        Next lIndex
        
        m.RptForm.FilterTrades
        
        If m.strView = "Trade" Then
            For lIndex = 1 To m.RptForm.Trades.NumRecords - 1
                If gdGetNum(m.hSignalType, lIndex) = gExitSignal Then
                    gdSetNum m.hShow, lIndex, 0
                End If
            Next lIndex
        End If
        
        ' Set the index on the table based on the filter
        m.aIndex.SetArrayHandle gdCreateTableIndex(m.RptForm.Trades.DtlTableHandle, entd_Show), True
        If m.strView = "Trade" Then
            gdSortTableIndex m.RptForm.Trades.DtlTableHandle, m.aIndex.ArrayHandle, entd_SortKey2, eGdSort_Default
        Else
            gdSortTableIndex m.RptForm.Trades.DtlTableHandle, m.aIndex.ArrayHandle, entd_SortKey, eGdSort_Default
        End If
        
        m.vsGrid.Redraw = flexRDNone
        InitGrid
        
        'Report options
        m.dBeginBalance = m.RptForm.OptionValue("BeginBalance")
        m.lMaxUnits = m.RptForm.OptionValue("MaxUnits")
        m.strPriceFormat = m.RptForm.OptionValue("PriceFormat")
        m.strShowTime = m.RptForm.OptionValue("ShowTime")
        m.strReportName = m.RptForm.OptionValue("ReportName")
        m.strShowCents = m.RptForm.OptionValue("ShowCents")
        m.strChartBy = m.RptForm.OptionValue("ChartBy")
        
        InitMonthlyGrid
        InitYearlyGrid
        
        Select Case UCase(m.strSubClass)
            Case "FIXEDFRACTIONAL"
                m.lUnitSize = m.RptForm.OptionValue("UnitSize")
                m.lStartUnits = m.RptForm.OptionValue("StartUnits")
                m.dRiskUp = m.RptForm.OptionValue("RiskUp")
                m.dRiskDown = m.RptForm.OptionValue("RiskDown")
                RunFixedFractional
                BuildChart
                If m.vsGrid.Rows > m.vsGrid.FixedRows Then
                    m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
                End If
                'ResizeColWidths
                
                If m.RptForm.optMonthly Then CalculateMonthly
                If m.RptForm.optYearly Then CalculateYearly
                        
                SetUpColumns
                SetBackColor
            
            Case "FIXEDRATIO"
                m.lUnitSize = m.RptForm.OptionValue("UnitSize")
                m.lStartUnits = m.RptForm.OptionValue("StartUnits")
                m.dDelta = ValOfText(m.RptForm.OptionValue("Delta"))
                RunFixedRatio
                BuildChart
                If m.vsGrid.Rows > m.vsGrid.FixedRows Then
                    m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
                End If
                'ResizeColWidths
                
                If m.RptForm.optMonthly Then CalculateMonthly
                If m.RptForm.optYearly Then CalculateYearly
                
                SetUpColumns
                SetBackColor
            
            Case "OPTIMALF"
                m.dMinProfit = m.RptForm.OptionValue("MinProfit")
                m.dRiskPerUnit = m.RptForm.OptionValue("RiskPerUnit")
                If RunOptimalF Then
                    BuildChart
                    If m.vsGrid.Rows > m.vsGrid.FixedRows Then
                        m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
                    End If
                    'ResizeColWidths
                    
                    If m.RptForm.optMonthly Then CalculateMonthly
                    If m.RptForm.optYearly Then CalculateYearly
                    
                    SetUpColumns
                    SetBackColor
                End If
            
            Case "WILLIAMSRISK"
                m.dRiskPct = m.RptForm.OptionValue("RiskPct")
                If RunWilliamsRisk Then
                    BuildChart
                    If m.vsGrid.Rows > m.vsGrid.FixedRows Then
                        m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
                    End If
                    'ResizeColWidths
                    
                    If m.RptForm.optMonthly Then CalculateMonthly
                    If m.RptForm.optYearly Then CalculateYearly
                    
                    SetUpColumns
                    SetBackColor
                End If
                
            Case "STANDARD"
                m.strRisk = m.RptForm.OptionValue("RiskPct")
                s = UCase(m.RptForm.OptionValue("RiskBasedOn"))
                If Left(s, 7) = "AVERAGE" Then
                    ' risk based on Average Loss
                    m.dRiskBasedOn = -2
                    m.vsGrid.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LargestLosing)) = "Average Loss"
                ElseIf ValOfText(s) > 0 Then
                    ' risk based on user-specifed Dollar Amount
                    m.dRiskBasedOn = Abs(ValOfText(s))
                    m.vsGrid.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LargestLosing)) = "Risk Based On"
                ElseIf Left(s, 7) = "LARGEST" Or Left(s, 3) = "MAX" Then
                    ' risk based on Largest Loss
                    m.dRiskBasedOn = -1
                    m.vsGrid.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LargestLosing)) = "Largest Loss"
                Else
                    ' risk based on Stop Loss (or Max Loss if system has no Stop Loss)
                    m.dRiskBasedOn = 0
                    m.vsGrid.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LargestLosing)) = "Stop Loss"
                End If
                m.dMargin = m.RptForm.OptionValue("Margin")
                RunStandard
                BuildChart
                If m.vsGrid.Rows > m.vsGrid.FixedRows Then
                    m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
                End If
                'ResizeColWidths
                
                If m.RptForm.optMonthly Then CalculateMonthly
                If m.RptForm.optYearly Then CalculateYearly
                
                SetUpColumns
                SetBackColor
        
        End Select
        
        m.RptForm.txtReportName(1).Caption = m.RptForm.txtReportName(1).Caption & " - Units for Next Trade: " & Format(m.lUnitsNextTrade, "#,##0")
        If m.lUnitsNextTrade = 0& Then
            m.RptForm.elMsg.Caption = "Not enough money left to trade (need to adjust the starting balance, risk, margin, etc.)"
            m.RptForm.elMsg.Visible = True
        Else
            m.RptForm.elMsg.Caption = ""
            m.RptForm.elMsg.Visible = False
        End If
        
        m.RptForm.vsElastic1.Align = 5
        m.RptForm.vsElastic1.Refresh
        
        m.vsGrid.Redraw = flexRDBuffered
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.Run", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildChart
'' Description: Intialize and load the chart
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildChart()
On Error GoTo ErrSection:
    
    Dim lItemsToChart As Long
    Dim lIndex As Long
    Dim lDate As Long
    Dim lIndex2 As Long
    Dim lIndex3 As Long
    Dim dEquity As Double
    Dim lAnnot As Long
    Dim lPrev As Long
    Dim dPrevDate As Double
    Dim dStartDate As Double
    Dim dEndDate As Double
    
    Dim dMaxDrawdownPct As Double
    Dim lLocOfMaxDD As Long
    Dim dDateOfMaxDD As Double
    Dim lLocOfAnnot As Long
    Dim dMaxDD As Double
    Dim lUnitsOfFirstTrade As Long
    Dim dMax As Double
    Dim dMin As Double
       
    'Calculate Items to chart (number of rows in grid)
    For lIndex3 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex3)
        
        If gdGetNum(m.hSignalType, lIndex) = 1 Then
            If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
                If lUnitsOfFirstTrade = 0 Then
                    If m.aUnits.Num(lIndex) > 0 Then
                        lUnitsOfFirstTrade = m.aUnits.Num(lIndex)
                    End If
                End If
                If dStartDate = 0 Then
                    dStartDate = gdGetNum(m.hTradeDate, lIndex)
                End If
                dEndDate = gdGetNum(m.hTradeDate, lIndex)
                lIndex2 = lIndex2 + 1
            End If
        End If
        
        If gdGetNum(m.aDrawDownPct.ArrayHandle, lIndex) > dMaxDrawdownPct Then
            dMaxDrawdownPct = gdGetNum(m.aDrawDownPct.ArrayHandle, lIndex)
            lLocOfMaxDD = lIndex
            dDateOfMaxDD = gdGetNum(m.hTradeDate, lIndex)
            dMaxDD = gdGetNum(m.aAccBalance.ArrayHandle, lIndex)
        End If
    Next lIndex3
    
    If m.strChartBy = "Date" Then
        lItemsToChart = Int(dEndDate) - Int(dStartDate) + 1
    Else
        lItemsToChart = lIndex2
    End If
    
    ShowChart lItemsToChart, m.Chart, m.RptForm.lblNoTrades, m.RptForm.lblNoMMTrades, lUnitsOfFirstTrade
    If lItemsToChart = 0 Then Exit Sub
    
    Set m.alGraphToTrades = New cGdArray
    m.alGraphToTrades.Create eGDARRAY_Longs, lItemsToChart
    
    '** Set Subsets, Points, and fill with YData **'
    With m.Chart
        .PEactions = 20     'Empty chart and reinitialize
        .Subsets = 2
        .Points = lItemsToChart
    
        .MainTitle = "Equity Curve"
        .SubTitle = ""
        .YAxisLabel = "" ' = "dEquity"
        .XAxisLabel = "" ' = "Trades"
        
        SetChartSettings m.Chart, GPM_LINE
        
        If m.strChartBy = "Date" Then
            .DateTimeMode = PEDTM_VB
            .StartTime = dStartDate
        Else
            .DateTimeMode = PEDTM_NONE
            .AltFrequencies(0) = 10
            .AltFrequencies(1) = 50
            .AltFrequencies(2) = 100
            .AltFrequencies(3) = 500
            .AltFrequencies(4) = 1000
            .FirstPtLabelOffset = 1
        End If
    
        lIndex2 = 0
        dEquity = m.dBeginBalance
        For lIndex3 = 1 To m.RptForm.Trades.NumRecords - 1
            lIndex = gdGetNum(m.hSorted, lIndex3)
            
            If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
                If gdGetNum(m.hSignalType, lIndex) = 1 Then
                    If m.strChartBy = "Date" Then
                        If dPrevDate = 0 Then dPrevDate = dStartDate
                        For lDate = Int(dPrevDate) To Int(gdGetNum(m.hTradeDate, lIndex)) - 1
                            lIndex2 = lIndex2 + 1
                            .YData(0, lIndex2 - 1) = dEquity
                            .XData(0, lIndex2) = lDate
                            If lDate = CLng(dDateOfMaxDD) Then lLocOfAnnot = lIndex2 - 1
                            m.alGraphToTrades(lIndex2) = lPrev
                            
                            If Month(lDate) = 1 And Day(lDate) = 1 Then
                                .VertLineAnnotationType(lAnnot) = PELAT_GRIDLINE
                                .VertLineAnnotation(lAnnot) = lIndex2
                                .VertLineAnnotationText(lAnnot) = "|B" & Format(lDate, "yyyy")
                                .VertLineAnnotationColor(lAnnot) = 1
                                lAnnot = lAnnot + 1
                            End If
                        Next lDate
                        lPrev = lIndex
                        dPrevDate = gdGetNum(m.hTradeDate, lIndex)
                        dEquity = dEquity + gdGetNum(m.hProfit, lIndex)
                    Else
                        lIndex2 = lIndex2 + 1
                        dEquity = dEquity + gdGetNum(m.hProfit, lIndex)
                        .YData(0, lIndex2 - 1) = dEquity
                        .XData(0, lIndex2 - 1) = lIndex2 - 1
                        If lIndex = lLocOfMaxDD Then lLocOfAnnot = lIndex2 - 1
                        .PointLabels(lIndex2 - 1) = lIndex2 - 1
                        m.alGraphToTrades(lIndex2 - 1) = gdGetNum(m.hTradeNum, lIndex)
                    End If
                End If
            End If
        Next lIndex3
        
        dPrevDate = 0
        lIndex2 = 0
        For lIndex3 = 1 To m.RptForm.Trades.NumRecords - 1
            lIndex = gdGetNum(m.hSorted, lIndex3)
            
            If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
                If gdGetNum(m.hSignalType, lIndex) = 1 Then
                    If m.strChartBy = "Date" Then
                        If dPrevDate = 0 Then dPrevDate = dStartDate
                        For lDate = Int(dPrevDate) To Int(gdGetNum(m.hTradeDate, lIndex)) - 1
                            lIndex2 = lIndex2 + 1
                            .YData(1, lIndex2 - 1) = m.aAccBalance.Num(lIndex)
                        Next lDate
                        dPrevDate = gdGetNum(m.hTradeDate, lIndex)
                    Else
                        lIndex2 = lIndex2 + 1
                        .YData(1, lIndex2 - 1) = m.aAccBalance.Num(lIndex)
                    End If
                End If
            End If
        Next lIndex3
        
        ' TLB 5/11/2014: use logarithmic scale (unless goes negative)
        m.dScaleMult = 1
        dMax = 0
        dMin = 0
        For lIndex = 0 To .Points - 1
            If .YData(0, lIndex) > dMax Then
                dMax = .YData(0, lIndex)
            End If
            If .YData(1, lIndex) > dMax Then
                dMax = .YData(1, lIndex)
            End If
            If .YData(0, lIndex) < dMin Then
                dMin = .YData(0, lIndex)
            End If
            If .YData(1, lIndex) < dMin Then
                dMin = .YData(1, lIndex)
            End If
        Next
        .WorkingAxis = 0
        If dMin < 0 Then
            .YAxisScaleControl = PEAC_NORMAL
        Else
            ' for Logarithmic scale
            .YAxisScaleControl = PEAC_LOG
            .LogScaleExpLabels = False
            '.LogScaleExpLabels = True
            '.FontSize = PEFS_SMALL
            If dMax > 10000000 Then
                m.dScaleMult = 1000000
                .AxisFormatY = "$|,|M"
            ElseIf dMax > 10000 Then
                m.dScaleMult = 1000
                .AxisFormatY = "$|,|K"
            Else
                .AxisFormatY = "$|,|"
            End If
            If m.dScaleMult > 1 Then
                For lIndex = 0 To .Points - 1
                    .YData(0, lIndex) = .YData(0, lIndex) / m.dScaleMult
                    .YData(1, lIndex) = .YData(1, lIndex) / m.dScaleMult
                Next
            End If
        End If
    
        '** Set SubsetLabels property array for 4 subsets **'
        .SubsetLabels(0) = "No Money Management"
        .SubsetLabels(1) = "With Money Management"
        
        ' Add the Drawdown Annotation
        .GraphAnnotationType(0) = PEGAT_SMALLUPTRIANGLESOLID
        .GraphAnnotationX(0) = lLocOfAnnot '- 1
        .GraphAnnotationY(0) = dMaxDD '.YData(1, lLocOfAnnot)
        .GraphAnnotationColor(0) = vbBlack
        .LegendAnnotationType(0) = PEGAT_SMALLUPTRIANGLESOLID
        .LegendAnnotationText(0) = "Max Equity Dip: " & Format(dMaxDrawdownPct, "0.00%") & " on " & DateFormat(dDateOfMaxDD)
        .LegendAnnotationColor(0) = vbBlack
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunFixedFractional
'' Description: Test the system given the largest loss
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunFixedFractional()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lUnits As Long
    Dim dProfit As Double
    Dim dTotalProfit As Double
    Dim dAccountBalance As Double
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim dMaxDrawPct As Double
    Dim dIncLevel As Double
    Dim dDecLevel As Double
    Dim dIncAmt As Double
    Dim dDecAmt As Double
    Dim lOpenPos As Long
    
    gdSetNum m.hAccountBalance, 0, m.dBeginBalance
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
        
        .Calculate
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    ' Determine largest losers over history
    m.dLargestLosing = 0.001
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipReport, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
            If gdGetNum(m.hProfit, lIndex) < m.dLargestLosing Then
                m.dLargestLosing = gdGetNum(m.hProfit, lIndex)
            End If
        End If
    Next lIndex2
        
    dAccountBalance = m.dBeginBalance
    dMaxAccount = m.dBeginBalance
    dMaxDrawPct = 0
    
    lUnits = m.lStartUnits
    
    ' Increase/Decrease levels
    If m.dRiskUp = 0# Then
        dIncAmt = 0#
    Else
        dIncAmt = Abs(m.dLargestLosing) / m.dRiskUp
    End If
    If m.dRiskDown = 0# Then
        dDecAmt = 0#
    Else
        dDecAmt = Abs(m.dLargestLosing) / m.dRiskDown
    End If
    dIncLevel = dAccountBalance + (dIncAmt * lUnits)
    dDecLevel = dAccountBalance - (dDecAmt * lUnits)
    If dIncLevel < 0 Then dIncLevel = 0
    If dDecLevel < 0 Then dDecLevel = 0
    
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        ' If trade marked as skipped then carry forward previous trade stats...
        If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
            m.aUnits(lIndex) = 0
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            m.aOpen(lIndex) = lOpenPos
            
        ' Entry Signal: Calculate Units and new inc/dec levels
        ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
        
            ' Check inc/dec level add one unit if either reached
            If dAccountBalance > dIncLevel Then
                lUnits = lUnits + m.lUnitSize
                dIncLevel = dAccountBalance + (dIncAmt * lUnits)
                dDecLevel = dAccountBalance - (dDecAmt * lUnits)
                If dIncLevel < 0 Then dIncLevel = 0
                If dDecLevel < 0 Then dDecLevel = 0
            ElseIf dAccountBalance < dDecLevel Then
                lUnits = lUnits - m.lUnitSize
                If lUnits < 0 Then lUnits = 0
                dIncLevel = dAccountBalance + (dIncAmt * lUnits)
                dDecLevel = dAccountBalance - (dDecAmt * lUnits)
                If dIncLevel < 0 Then dIncLevel = 0
                If dDecLevel < 0 Then dDecLevel = 0
            End If
            If lUnits > m.lMaxUnits Then
                lUnits = m.lMaxUnits
            End If
            
            lOpenPos = lOpenPos + 1
            m.aOpen(lIndex) = lOpenPos
            m.aUnits(lIndex) = lUnits
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            
        ' Exit Signal: Retrieve entry lUnits and calculate total dProfit
        Else
            lOpenPos = lOpenPos - 1
            m.aOpen(lIndex) = lOpenPos
            lUnits = m.aUnits.Item(gdGetNum(m.hEntryExitPtr, lIndex))
            m.aUnits(lIndex) = lUnits
            dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
            m.aTotProfit(lIndex) = dTotalProfit
    
            ' Calculate new account balance...
            If dAccountBalance > 0 Then
                dAccountBalance = dAccountBalance + dTotalProfit
            End If
            
            ' Calculate closed drawdown per trade...
            If dAccountBalance > dMaxAccount Then
                dMaxAccount = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            Else
                dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                If dDrawPct > dMaxDrawPct Then dMaxDrawPct = dDrawPct
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            End If
            
            m.aAccBalance(lIndex) = dAccountBalance
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
        End If
        
        gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
    Next lIndex2
        
    ' Calculate number of units one more time for next trade...
    If dAccountBalance > dIncLevel Then
        lUnits = lUnits + m.lUnitSize
        dIncLevel = dAccountBalance + (dIncAmt * lUnits)
        dDecLevel = dAccountBalance - (dDecAmt * lUnits)
        If dIncLevel < 0 Then dIncLevel = 0
        If dDecLevel < 0 Then dDecLevel = 0
    ElseIf dAccountBalance < dDecLevel Then
        lUnits = lUnits - m.lUnitSize
        If lUnits < 0 Then lUnits = 0
        dIncLevel = dAccountBalance + (dIncAmt * lUnits)
        dDecLevel = dAccountBalance - (dDecAmt * lUnits)
        If dIncLevel < 0 Then dIncLevel = 0
        If dDecLevel < 0 Then dDecLevel = 0
    End If
    If lUnits > m.lMaxUnits Then
        lUnits = m.lMaxUnits
    End If
    m.lUnitsNextTrade = lUnits

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.RunFixedFractional", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunFixedRatio
'' Description: Test the system given the largest loss
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunFixedRatio()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lUnits As Long
    Dim dProfit As Double
    Dim dTotalProfit As Double
    Dim dAccountBalance As Double
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim dMaxDrawPct As Double
    Dim dIncLevel As Double
    Dim dDecLevel As Double
    Dim lOpenPos As Long
    Dim dDelta As Double
    Dim dBaseBalance As Double
    
    gdSetNum m.hAccountBalance, 0, m.dBeginBalance
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
        
        .Calculate
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    dAccountBalance = m.dBeginBalance
    dMaxAccount = m.dBeginBalance
    dMaxDrawPct = 0
    
    If m.dDelta = 0# Then m.dDelta = Max(Abs(m.Stats.MaxDrawdownClosed) / 2#, m.dMargin)
    dDelta = m.dDelta
       
#If 0 Then
    ' OLD METHOD ...
    
    'Increase/Decrease levels
    lUnits = m.lStartUnits
    dIncLevel = Max(dAccountBalance + (dDelta * lUnits), 0)
    dDecLevel = Max(dAccountBalance - (dDelta * lUnits), 0)

    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        'If trade marked as skipped then carry forward previous trade stats
        If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
            m.aUnits(lIndex) = 0
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            m.aOpen(lIndex) = lOpenPos
            
        ' Entry Signal: Calculate lUnits and new inc/dec levels
        ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
        
            ' Check inc/dec level add one unit if either reached
            If dAccountBalance > dIncLevel Then
                lUnits = lUnits + m.lUnitSize
                dIncLevel = Max(dAccountBalance + (dDelta * lUnits), 0)
                dDecLevel = Max(dAccountBalance - (dDelta * lUnits), 0)
            ElseIf dAccountBalance <= 0 Then
                lUnits = 0&
                dIncLevel = 0#
                dDecLevel = 0#
            ElseIf dAccountBalance < dDecLevel Then
                lUnits = CLng(Max(lUnits - m.lUnitSize, 0))
                dIncLevel = Max(dAccountBalance + (dDelta * lUnits), 0)
                dDecLevel = Max(dAccountBalance - (dDelta * lUnits), 0)
            End If
            If lUnits > m.lMaxUnits Then
                lUnits = m.lMaxUnits
            End If
            
            lOpenPos = lOpenPos + 1
            m.aOpen(lIndex) = lOpenPos
            m.aUnits(lIndex) = lUnits
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            
        ' Exit Signal: Retrieve entry lUnits and calculate total dProfit
        Else
            lOpenPos = lOpenPos - 1
            m.aOpen(lIndex) = lOpenPos
            lUnits = m.aUnits.Item(gdGetNum(m.hEntryExitPtr, lIndex))
            m.aUnits(lIndex) = lUnits
            dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
            m.aTotProfit(lIndex) = dTotalProfit
    
            ' Calculate new account balance...
            If dAccountBalance > 0 Then
                dAccountBalance = dAccountBalance + dTotalProfit
            End If
            
            ' Calculate closed drawdown per trade...
            If dAccountBalance > dMaxAccount Then
                dMaxAccount = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            Else
                dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                If dDrawPct > dMaxDrawPct Then dMaxDrawPct = dDrawPct
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            End If
            
            m.aAccBalance(lIndex) = dAccountBalance
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
        End If
        
        gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
    Next lIndex2
        
    ' Calculate number of units one more time for next trade...
    If dAccountBalance > dIncLevel Then
        lUnits = lUnits + m.lUnitSize
        dIncLevel = Max(dAccountBalance + (dDelta * lUnits), 0)
        dDecLevel = Max(dAccountBalance - (dDelta * lUnits), 0)
    ElseIf dAccountBalance <= 0 Then
        lUnits = 0&
        dIncLevel = 0#
        dDecLevel = 0#
    ElseIf dAccountBalance < dDecLevel Then
        lUnits = CLng(Max(lUnits - m.lUnitSize, 0))
        dIncLevel = Max(dAccountBalance + (dDelta * lUnits), 0)
        dDecLevel = Max(dAccountBalance - (dDelta * lUnits), 0)
    End If
    If lUnits > m.lMaxUnits Then
        lUnits = m.lMaxUnits
    End If
    m.lUnitsNextTrade = lUnits

#Else
    ' TLB 3/11/2015: new method with some major fixes to correct Fixed Ratio ...
    
    ' Set BaseBalance (for 1 contract) and Increase/Decrease levels
    If m.lUnitSize < 1 Then m.lUnitSize = 1
    If m.lStartUnits < 1 Then m.lStartUnits = 1
    lUnits = m.lStartUnits
    dBaseBalance = m.dBeginBalance - dDelta * (lUnits * lUnits - lUnits) / 2
    dDecLevel = dBaseBalance + dDelta * (lUnits * lUnits - lUnits) / 2
    dIncLevel = dBaseBalance + dDelta * ((lUnits + m.lUnitSize) * (lUnits + m.lUnitSize) - (lUnits + m.lUnitSize)) / 2

    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        'If trade marked as skipped then carry forward previous trade stats
        If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
            m.aUnits(lIndex) = 0
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            m.aOpen(lIndex) = lOpenPos
            
        ' Entry Signal: Calculate lUnits and new inc/dec levels
        ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            If dAccountBalance <= 0 Then ' Abs(m.Stats.MaxDrawdownClosed) Then
                ' can't trade anymore
                lUnits = 0&
                dIncLevel = 0#
                dDecLevel = 0#
            Else
                ' Check inc/dec level add one unit if either reached
                Do While dAccountBalance < dDecLevel And lUnits > m.lUnitSize
                    lUnits = lUnits - m.lUnitSize
                    dDecLevel = dBaseBalance + dDelta * (lUnits * lUnits - lUnits) / 2
                    dIncLevel = dBaseBalance + dDelta * ((lUnits + m.lUnitSize) * (lUnits + m.lUnitSize) - (lUnits + m.lUnitSize)) / 2
                Loop
                Do While dAccountBalance >= dIncLevel And lUnits < m.lMaxUnits
                    lUnits = lUnits + m.lUnitSize
                    If lUnits > m.lMaxUnits Then
                        lUnits = m.lMaxUnits
                    End If
                    dDecLevel = dBaseBalance + dDelta * (lUnits * lUnits - lUnits) / 2
                    dIncLevel = dBaseBalance + dDelta * ((lUnits + m.lUnitSize) * (lUnits + m.lUnitSize) - (lUnits + m.lUnitSize)) / 2
                Loop
            End If
            
            lOpenPos = lOpenPos + 1
            m.aOpen(lIndex) = lOpenPos
            m.aUnits(lIndex) = lUnits
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            
        ' Exit Signal: Retrieve entry lUnits and calculate total dProfit
        Else
            lOpenPos = lOpenPos - 1
            m.aOpen(lIndex) = lOpenPos
            lUnits = m.aUnits.Item(gdGetNum(m.hEntryExitPtr, lIndex))
            m.aUnits(lIndex) = lUnits
            dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
            m.aTotProfit(lIndex) = dTotalProfit
    
            ' Calculate new account balance...
            If dAccountBalance > 0 Then
                dAccountBalance = dAccountBalance + dTotalProfit
            End If
            
            ' Calculate closed drawdown per trade...
            If dAccountBalance > dMaxAccount Then
                dMaxAccount = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            Else
                dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                If dDrawPct > dMaxDrawPct Then dMaxDrawPct = dDrawPct
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            End If
            
            m.aAccBalance(lIndex) = dAccountBalance
            m.aIncLevel(lIndex) = dIncLevel
            m.aDecLevel(lIndex) = dDecLevel
        End If
        
        gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
    Next lIndex2
        
    ' Calculate number of units one more time for next trade...
    If dAccountBalance <= 0 Then ' Abs(m.Stats.MaxDrawdownClosed) Then
        ' can't trade anymore
        lUnits = 0&
        dIncLevel = 0#
        dDecLevel = 0#
    Else
        ' Check inc/dec level add one unit if either reached
        Do While dAccountBalance < dDecLevel And lUnits > m.lUnitSize
            lUnits = lUnits - m.lUnitSize
            dDecLevel = dBaseBalance + dDelta * (lUnits * lUnits - lUnits) / 2
            dIncLevel = dBaseBalance + dDelta * ((lUnits + m.lUnitSize) * (lUnits + m.lUnitSize) - (lUnits + m.lUnitSize)) / 2
        Loop
        Do While dAccountBalance >= dIncLevel And lUnits < m.lMaxUnits
            lUnits = lUnits + m.lUnitSize
            If lUnits > m.lMaxUnits Then
                lUnits = m.lMaxUnits
            End If
            dDecLevel = dBaseBalance + dDelta * (lUnits * lUnits - lUnits) / 2
            dIncLevel = dBaseBalance + dDelta * ((lUnits + m.lUnitSize) * (lUnits + m.lUnitSize) - (lUnits + m.lUnitSize)) / 2
        Loop
    End If
    m.lUnitsNextTrade = lUnits

#End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.RunFixedRatio", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunOptimalF
'' Description: Test the system given the largest loss
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunOptimalF() As Boolean
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lUnits As Long
    Dim dProfit As Double
    Dim dTotalProfit As Double
    Dim dAccountBalance As Double
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim dMaxDrawPct As Double
    Dim lOpenPos As Long
    Dim dOptimalF As Double
    
    'Calculate OptimalF
    gdSetNum m.hAccountBalance, 0, m.dBeginBalance
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
        
        .Calculate
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    If m.Stats.OptimalF < 0 Then
        With m.vsGrid
            ClearGrid m.vsGrid
            .Rows = 2
            .Cols = 1
            .FixedCols = 0
            .FixedRows = 0
            .ExtendLastCol = True
            .BackColorAlternate = .BackColor
            .TextMatrix(0, 0) = "This strategy does not have an optimal risk percentage.  " & _
                    "OptimalF was calculated to be less than zero."
            .TextMatrix(1, 0) = "Either the strategy had no winning trades, or the " & _
                    "Risk per Unit value needs to be increased."
            .Cell(flexcpAlignment, 0, 0, 1, 0) = flexAlignCenterCenter
            DoEvents
        End With
        
        RunOptimalF = False
    Else
        dAccountBalance = m.dBeginBalance
        dMaxAccount = m.dBeginBalance
        dMaxDrawPct = 0
        dOptimalF = m.Stats.OptimalF
        
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lIndex = gdGetNum(m.hSorted, lIndex2)
            
            'If trade marked as skipped then carry forward previous trade stats
            If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
                m.aUnits(lIndex) = 0
                m.aTotProfit(lIndex) = 0
                m.aAccBalance(lIndex) = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
                m.aOpen(lIndex) = lOpenPos
                m.aOptimalf(lIndex) = dOptimalF
                
            'Entry Signal: Calculate lUnits and new inc/dec levels
            ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            
                If m.dRiskPerUnit <> 0 Then
                    lUnits = Int((dAccountBalance * dOptimalF) / m.dRiskPerUnit)
                Else
                    lUnits = 0
                End If
                If lUnits > m.lMaxUnits Then
                    lUnits = m.lMaxUnits
                End If
                
                lOpenPos = lOpenPos + 1
                m.aOpen(lIndex) = lOpenPos
                m.aUnits(lIndex) = lUnits
                m.aTotProfit(lIndex) = 0
                m.aAccBalance(lIndex) = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
                m.aOptimalf(lIndex) = dOptimalF
                
            'Exit Signal: Retrieve entry lUnits and calculate total dProfit
            Else
                lOpenPos = lOpenPos - 1
                m.aOpen(lIndex) = lOpenPos
                lUnits = m.aUnits.Item(gdGetNum(m.hEntryExitPtr, lIndex))
                m.aUnits(lIndex) = lUnits
                dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
                m.aTotProfit(lIndex) = dTotalProfit
        
                'Calculate new account balance
                If dAccountBalance > 0 Then
                    dAccountBalance = dAccountBalance + dTotalProfit
                End If
                
                'Calc closed drawdown per trade
                If dAccountBalance > dMaxAccount Then
                    dMaxAccount = dAccountBalance
                    If dMaxDrawPct = 0 Then
                        m.aDrawDownPct(lIndex) = 0
                    Else
                        m.aDrawDownPct(lIndex) = dMaxDrawPct
                    End If
                Else
                    If dAccountBalance > 0 Then
                        If dMaxAccount <> 0 Then
                            dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                        Else
                            dDrawPct = 0
                        End If
                        If dDrawPct > dMaxDrawPct Then
                            m.aDrawDownPct(lIndex) = dDrawPct
                            dMaxDrawPct = dDrawPct
                        Else
                            m.aDrawDownPct(lIndex) = dMaxDrawPct
                        End If
                    Else
                        m.aDrawDownPct(lIndex) = 0
                    End If
                End If
                
                m.aAccBalance(lIndex) = dAccountBalance
                m.aOptimalf(lIndex) = dOptimalF
            End If
            
            gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
        Next lIndex2
        
        RunOptimalF = True
    End If
        
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cRptMMStudy.RunOptimalF", eGDRaiseError_Raise, m.RptForm.AppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunWilliamsRisk
'' Description: Test the system given the largest loss
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunWilliamsRisk() As Boolean
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lUnits As Long
    Dim dProfit As Double
    Dim dTotalProfit As Double
    Dim dAccountBalance As Double
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim dMaxDrawPct As Double
    Dim lOpenPos As Long
    
    gdSetNum m.hAccountBalance, 0, m.dBeginBalance
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
        
        .Calculate
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    'Determine largest losers over history
    m.dLargestLosing = 0.001
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipReport, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
            If gdGetNum(m.hProfit, lIndex) < m.dLargestLosing Then
                m.dLargestLosing = gdGetNum(m.hProfit, lIndex)
            End If
        End If
    Next lIndex2
    m.dLargestLosing = Abs(m.dLargestLosing)
        
    If m.dLargestLosing = 0# Then
        With m.vsGrid
            ClearGrid m.vsGrid
            .Rows = 1
            .Cols = 1
            .FixedCols = 0
            .FixedRows = 0
            .ExtendLastCol = True
            .BackColorAlternate = .BackColor
            .TextMatrix(0, 0) = "Williams Risk/Ratio cannot be calculated unless the " & _
                    "largest loss is greater than zero."
            .Cell(flexcpAlignment, 0, 0, 0, 0) = flexAlignCenterCenter
            DoEvents
        End With
        
        RunWilliamsRisk = False
    Else
        dAccountBalance = m.dBeginBalance
        dMaxAccount = m.dBeginBalance
        dMaxDrawPct = 0
        
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lIndex = gdGetNum(m.hSorted, lIndex2)
            
            ' If trade marked as skipped then carry forward previous trade stats
            If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
                m.aUnits(lIndex) = 0
                m.aTotProfit(lIndex) = 0
                m.aAccBalance(lIndex) = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
                m.aOpen(lIndex) = lOpenPos
                
            ' Entry Signal: Calculate lUnits and new inc/dec levels
            ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            
                If m.dLargestLosing = 0.001 Then
                    lUnits = m.lMaxUnits
                Else
                    lUnits = Int((dAccountBalance * m.dRiskPct) / m.dLargestLosing)
                    If lUnits > m.lMaxUnits Then lUnits = m.lMaxUnits
                End If
                
                lOpenPos = lOpenPos + 1
                m.aOpen(lIndex) = lOpenPos
                m.aUnits(lIndex) = lUnits
                m.aTotProfit(lIndex) = 0
                m.aAccBalance(lIndex) = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
                
            ' Exit Signal: Retrieve entry lUnits and calculate total dProfit
            Else
                lOpenPos = lOpenPos - 1
                m.aOpen(lIndex) = lOpenPos
                lUnits = m.aUnits.Item(gdGetNum(m.hEntryExitPtr, lIndex))
                m.aUnits(lIndex) = lUnits
                dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
                m.aTotProfit(lIndex) = dTotalProfit
        
                ' Calculate new account balance
                If dAccountBalance > 0 Then
                    dAccountBalance = dAccountBalance + dTotalProfit
                End If
                
                ' Calc closed drawdown per trade
                If dAccountBalance > dMaxAccount Then
                    dMaxAccount = dAccountBalance
                    m.aDrawDownPct(lIndex) = dMaxDrawPct
                Else
                    dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                    If dDrawPct > dMaxDrawPct Then dMaxDrawPct = dDrawPct
                    m.aDrawDownPct(lIndex) = dMaxDrawPct
                End If
                
                m.aAccBalance(lIndex) = dAccountBalance
            End If
            
            gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
        Next lIndex2
        
        RunWilliamsRisk = True
    End If
        
    ' Calculate number of units one more time for next trade...
    If m.dLargestLosing = 0.001 Then
        lUnits = m.lMaxUnits
    Else
        lUnits = Int((dAccountBalance * m.dRiskPct) / m.dLargestLosing)
        If lUnits > m.lMaxUnits Then lUnits = m.lMaxUnits
    End If
    m.lUnitsNextTrade = lUnits

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cRptMMStudy.RunWilliamsRisk", eGDRaiseError_Raise, m.RptForm.AppPath

End Function

#If 1 Then

' TLB 3/31/2011: modified to be based on either largest or average loss, and also to do
' each strategy basket item separately (i.e. #units based on max or avg loss for just that item)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunStandard
'' Description: Test the system given the largest loss
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunStandard()
On Error GoTo ErrSection:
    
    Dim lHdrIndex As Long
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lUnits As Long
    Dim dProfit As Double
    Dim dTotalProfit As Double
    Dim dAccountBalance As Double
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim dMaxDrawPct As Double
    Dim lOpenPos As Long
    Dim dDivisor As Double
    Dim dDate As Double
    Dim bSkip As Boolean
    Dim d#, dMargin#, dRisk#
    
    Dim aAvgLosses As New cGdArray
    Dim aLossCounts As New cGdArray
       
    m.aRiskBasedOn.Size = 0
    aAvgLosses.Create eGDARRAY_Doubles, 0, 0
    aLossCounts.Create eGDARRAY_Longs, 0, 0
    
    gdSetNum m.hAccountBalance, 0, m.dBeginBalance
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
        
        .Calculate
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    'Determine largest losers over history
    m.dLargestLosing = 0 ' = 0.001
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipReport, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
            dProfit = gdGetNum(m.hProfit, lIndex)
            If dProfit < m.dLargestLosing Then
                m.dLargestLosing = dProfit
            End If
            If dProfit <= 0 Then
                ' examine this loss compared to other losses for this basket item
                lHdrIndex = gdGetNum(m.hHeaderIndex, lIndex)
                If lHdrIndex >= m.aRiskBasedOn.Size Then
                    m.aRiskBasedOn.Size = lHdrIndex + 1
                    aAvgLosses.Size = lHdrIndex + 1
                    aLossCounts.Size = lHdrIndex + 1
                End If
                aAvgLosses.Num(lHdrIndex) = aAvgLosses.Num(lHdrIndex) + dProfit
                aLossCounts.Num(lHdrIndex) = aLossCounts.Num(lHdrIndex) + 1
                ' for now, store Largest Loss for each system for the $Risk
                If dProfit < m.aRiskBasedOn.Num(lHdrIndex) Then
                    m.aRiskBasedOn.Num(lHdrIndex) = dProfit
                End If
            End If
        End If
    Next
    
    ' assign $Risk for each system
    For lHdrIndex = 1 To m.aRiskBasedOn.Size - 1
        ' calc the Avg Loss
        dDivisor = aLossCounts.Num(lHdrIndex)
        If dDivisor > 1 Then
            aAvgLosses.Num(lHdrIndex) = aAvgLosses.Num(lHdrIndex) / dDivisor
        End If
        ' calc $Risk
        If m.dRiskBasedOn > 0 Then
            ' if risk based on user-specified Dollar Amount
            dRisk = m.dRiskBasedOn
        ElseIf m.dRiskBasedOn = -2 Then
            ' if risk based on Avg Loss
            If m.aRiskBasedOn.Size = 2 Then
                ' if only 1 basket item, just get the AvgLoss from the stats
                ' (so will be more exact -- i.e. will not include an open equity loss)
                dRisk = m.Stats.LossAvg
            Else
                dRisk = aAvgLosses.Num(lHdrIndex)
            End If
        ElseIf m.dRiskBasedOn = -1 Then
            ' if risk based on Largest Loss
            dRisk = m.aRiskBasedOn.Num(lHdrIndex)
        Else
            ' Default: if using StopLoss
            dRisk = 0
            d = gdGetTableNum(m.RptForm.Trades.HdrTableHandle, enth_LongStopLoss, lHdrIndex)
            If d > 0 And -d < dRisk Then
                dRisk = -d
            End If
            d = gdGetTableNum(m.RptForm.Trades.HdrTableHandle, enth_ShortStopLoss, lHdrIndex)
            If d > 0 And -d < dRisk Then
                dRisk = -d
            End If
            If dRisk = 0 Then
                ' but if no StopLoss amount identified for this system, then just use Largest Loss
                dRisk = m.aRiskBasedOn.Num(lHdrIndex)
            ElseIf dRisk > aAvgLosses.Num(lHdrIndex) Then
                ' and it has to be at least as big as the AvgLoss
                dRisk = aAvgLosses.Num(lHdrIndex)
            End If
        End If
        m.aRiskBasedOn.Num(lHdrIndex) = dRisk
    Next
    
#If 0 Then
    If m.dRiskBasedOn = 0 Then
        ' if risk based on Average Loss
        If m.aRiskBasedOn.Size = 2 Then
            ' if only 1 basket item, just get the AvgLoss from the stats
            ' (so will be more exact -- i.e. will not include an open equity loss)
            m.aRiskBasedOn.Num(1) = m.Stats.LossAvg
        Else
            ' divide cumulative losses by # of losses for each basket item
            For lHdrIndex = 1 To m.aRiskBasedOn.Size - 1
                dDivisor = aLossCounts.Num(lHdrIndex)
                If dDivisor > 1 Then
                    m.aRiskBasedOn.Num(lHdrIndex) = m.aRiskBasedOn.Num(lHdrIndex) / dDivisor
                End If
            Next
        End If
    ElseIf m.dRiskBasedOn > 0 Then
        ' if risk based on user-specified Dollar Amount
        For lHdrIndex = 1 To m.aRiskBasedOn.Size - 1
            m.aRiskBasedOn.Num(lHdrIndex) = m.dRiskBasedOn
        Next
    End If
#End If
    
    Select Case UCase(m.strRisk)
        Case "KELLY RATIO"
            m.dRiskPct = m.Stats.KellyRatio
        Case "OPTIMAL F"
            m.dRiskPct = m.Stats.OptimalF
        Case Else
            m.dRiskPct = ValOfText(m.strRisk)
    End Select
    
    If m.dRiskPct < 0 Then m.dRiskPct = 0
    dAccountBalance = m.dBeginBalance
    dMaxAccount = m.dBeginBalance
    dMaxDrawPct = 0
    
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        lHdrIndex = gdGetNum(m.hHeaderIndex, lIndex)
        
        ' see if skipping this trade
        bSkip = False
        If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
            bSkip = True
        ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            ' skip trades with entry prior to user-specified start date
            dDate = gdGetNum(m.hTradeDate, lIndex)
            If dDate < m.dDelayedStartDate Then
                bSkip = True
            End If
        End If
        
        ' If trade marked as skipped then carry forward previous trade stats
        If bSkip Then
            m.aUnits(lIndex) = 0
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            m.aOpen(lIndex) = lOpenPos
            
        ' Entry Signal: Calculate Units and new inc/dec levels
        ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
        
            dDivisor = Abs(m.aRiskBasedOn.Num(lHdrIndex))
        
            If dDivisor <> 0 And dAccountBalance >= 0 Then
                lUnits = Int((dAccountBalance * m.dRiskPct) / dDivisor)
            ElseIf dDivisor = 0 Then
                lUnits = m.lMaxUnits
            Else
                lUnits = 0
            End If
            If lUnits > m.lMaxUnits Then
                lUnits = m.lMaxUnits
            End If
            ' TLB 7/18/2011: but can't trade more contracts than margin allows for
            dMargin = m.aMargins.Num(lHdrIndex)
            If dMargin > 0 And dAccountBalance > 0 Then
                If lUnits > Int(dAccountBalance / dMargin) Then
                    lUnits = Int(dAccountBalance / dMargin)
                End If
            End If
            
            lOpenPos = lOpenPos + 1
            m.aOpen(lIndex) = lOpenPos
            m.aUnits(lIndex) = lUnits
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            
        ' Exit Signal: Retrieve entry Units and calculate Total Profit
        Else
            lOpenPos = lOpenPos - 1
            m.aOpen(lIndex) = lOpenPos
            lUnits = m.aUnits(gdGetNum(m.hEntryExitPtr, lIndex))
            m.aUnits(lIndex) = lUnits
            dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
            m.aTotProfit(lIndex) = dTotalProfit
    
            ' Calculate new account balance
            If dAccountBalance > 0 Then
                dAccountBalance = dAccountBalance + dTotalProfit
            End If
            
            ' Calculate closed drawdown per trade
            If dAccountBalance > dMaxAccount Then
                dMaxAccount = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            Else
                dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                If dDrawPct > dMaxDrawPct Then
                    m.aDrawDownPct(lIndex) = dDrawPct
                    dMaxDrawPct = dDrawPct
                Else
                    m.aDrawDownPct(lIndex) = dMaxDrawPct
                End If
            End If
            
            m.aAccBalance(lIndex) = dAccountBalance
        End If
        
        gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
    Next lIndex2
        
    ' Calculate the number of units one more time for the next trade...
    If dDivisor <> 0 And dAccountBalance >= 0 Then
        lUnits = Int((dAccountBalance * m.dRiskPct) / dDivisor)
    Else
        lUnits = 0
    End If
    If lUnits > m.lMaxUnits Then
        lUnits = m.lMaxUnits
    End If
    m.lUnitsNextTrade = lUnits
    
ErrExit:
    Set aLossCounts = Nothing
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.RunStandard", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

#Else

' TLB 3/31/2011: this was a version that was going to be based on a percentile of losses,
' but the need to do each strategy basket item (code above) was more important for now.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunStandard
'' Description: Test the system given the largest loss
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunStandard_TEST()
On Error GoTo ErrSection:
    
    Dim lHdrIndex As Long
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim lUnits As Long
    Dim dProfit As Double
    Dim dTotalProfit As Double
    Dim dAccountBalance As Double
    Dim dMaxAccount As Double
    Dim dDrawPct As Double
    Dim dMaxDrawPct As Double
    Dim lOpenPos As Long
    Dim dDivisor As Double
    
    Dim dPercentile As Double
    Dim bUseAvgLoss As Boolean
    Dim aLosses As New cGdArray
    aLosses.Create eGDARRAY_Doubles, 0, 0
    
If IsIDE Then
    dPercentile = 95
    bUseAvgLoss = True
End If
    
    gdSetNum m.hAccountBalance, 0, m.dBeginBalance
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
        
        .Calculate
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    'Determine largest losers over history
    m.dLargestLosing = 0 ' = 0.001
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        lHdrIndex = gdGetNum(m.hHeaderIndex, lIndex2)
        If lHdrIndex <> 1 Then
            lHdrIndex = lHdrIndex
        End If
        
        If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipReport, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
            dProfit = gdGetNum(m.hProfit, lIndex)
            If dProfit < m.dLargestLosing Then
                m.dLargestLosing = dProfit
            End If
            
            If dProfit < 0 And dPercentile > 0 Then
                aLosses.Add dProfit
            End If
        End If
    Next lIndex2
    
    If dPercentile > 0 Then
        ' TLB: get the Nth percentile loss
        aLosses.Sort eGdSort_Descending
        lIndex2 = Int(aLosses.Size * dPercentile / 100#) ' for this context, we do NOT want to round
        If lIndex2 > aLosses.Size - 1 Then
            lIndex2 = aLosses.Size - 1
        End If
        m.dLargestLosing = aLosses(lIndex2)
        dDivisor = Abs(m.dLargestLosing)
    Else
        If m.dMargin > Abs(m.dLargestLosing) Then
            dDivisor = m.dMargin
        Else
            dDivisor = Abs(m.dLargestLosing)
        End If
    End If
        
    
    Select Case UCase(m.strRisk)
        Case "KELLY RATIO"
            m.dRiskPct = m.Stats.KellyRatio
         
        Case "OPTIMAL F"
            m.dRiskPct = m.Stats.OptimalF
        
        Case Else
            m.dRiskPct = ValOfText(m.strRisk)
    
    End Select
    
    If m.dRiskPct < 0 Then m.dRiskPct = 0
    dAccountBalance = m.dBeginBalance
    dMaxAccount = m.dBeginBalance
    dMaxDrawPct = 0
    
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        ' If trade marked as skipped then carry forward previous trade stats
        If (gdGetNum(m.hSkip, lIndex) = 1) Or (gdGetNum(m.hSkipEq, lIndex) = 1) Then
            m.aUnits(lIndex) = 0
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            m.aOpen(lIndex) = lOpenPos
            
        ' Entry Signal: Calculate Units and new inc/dec levels
        ElseIf gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
        
            If dDivisor <> 0 And dAccountBalance >= 0 Then
                lUnits = Int((dAccountBalance * m.dRiskPct) / dDivisor)
            ElseIf dDivisor = 0 Then
                lUnits = m.lMaxUnits
            Else
                lUnits = 0
            End If
            If lUnits > m.lMaxUnits Then
                lUnits = m.lMaxUnits
            End If
            
            lOpenPos = lOpenPos + 1
            m.aOpen(lIndex) = lOpenPos
            m.aUnits(lIndex) = lUnits
            m.aTotProfit(lIndex) = 0
            m.aAccBalance(lIndex) = dAccountBalance
            m.aDrawDownPct(lIndex) = dMaxDrawPct
            
        ' Exit Signal: Retrieve entry Units and calculate Total Profit
        Else
            lOpenPos = lOpenPos - 1
            m.aOpen(lIndex) = lOpenPos
            lUnits = m.aUnits(gdGetNum(m.hEntryExitPtr, lIndex))
            m.aUnits(lIndex) = lUnits
            dTotalProfit = lUnits * gdGetNum(m.hProfit, lIndex)
            m.aTotProfit(lIndex) = dTotalProfit
    
            ' Calculate new account balance
            If dAccountBalance > 0 Then
                dAccountBalance = dAccountBalance + dTotalProfit
            End If
            
            ' Calculate closed drawdown per trade
            If dAccountBalance > dMaxAccount Then
                dMaxAccount = dAccountBalance
                m.aDrawDownPct(lIndex) = dMaxDrawPct
            Else
                dDrawPct = (dMaxAccount - dAccountBalance) / dMaxAccount
                If dDrawPct > dMaxDrawPct Then
                    m.aDrawDownPct(lIndex) = dDrawPct
                    dMaxDrawPct = dDrawPct
                Else
                    m.aDrawDownPct(lIndex) = dMaxDrawPct
                End If
            End If
            
            m.aAccBalance(lIndex) = dAccountBalance
        End If
        
        gdSetNum m.hUnits, lIndex, m.aUnits(lIndex)
    Next lIndex2
        
    ' Calculate the number of units one more time for the next trade...
    If dDivisor <> 0 And dAccountBalance >= 0 Then
        lUnits = Int((dAccountBalance * m.dRiskPct) / dDivisor)
    Else
        lUnits = 0
    End If
    If lUnits > m.lMaxUnits Then
        lUnits = m.lMaxUnits
    End If
    m.lUnitsNextTrade = lUnits
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.RunStandard", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitGrid()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim bSysNav As Boolean              ' Are we called from System Navigator?
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        ClearGrid m.vsGrid
        
        .Editable = True
        .BackColorAlternate = g.nAltGridRowColor '&HC0FFFF
        .ExplorerBar = flexExSortShowAndMove
        .GridLines = flexGridNone
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .HighLight = flexHighlightAlways
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols) + 1
        .FrozenCols = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AllowUserResizing = flexResizeColumns
        .ScrollTips = True
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        
        m.lSubHeadRow = 0

        For lIndex = 0 To .Cols - 1
            .ColHidden(lIndex) = False
        Next lIndex

        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Number)) = "Num"
        If UCase(m.strView) = "TRADE" Then
            .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Date)) = "Date "
        Else
            .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Date)) = "Date"
        End If
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntryType)) = "Type"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Position)) = "Pos"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntryDate)) = "Entry Date"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ExitDate)) = "Exit Date"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Pointer)) = "Ptr"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SystemNumber)) = "Strategy"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OpenPosition)) = "Open"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_RiskPercent)) = "Risk %"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_LargestLosing)) = "Largest Loss"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Profit)) = "Profit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Units)) = "Units"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotalProfit)) = "Total Profit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Balance)) = "Balance"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_MaxDip)) = "Equity Dip"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_DecLevel)) = "Decrease"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_IncLevel)) = "Increase"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OptimalF)) = "Optimal f"
        
        bSysNav = (UCase(m.RptForm.AppName) = "SYSTEM NAVIGATOR")
        
        Select Case UCase(m.strSubClass)
            Case "FIXEDFRACTIONAL"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NumCols)) = "Equity Dip"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OptimalF)) = "Equity Dip"
                .ColHidden(GDCol(eGDCol_DecLevel)) = False
                .ColHidden(GDCol(eGDCol_IncLevel)) = False
                .ColHidden(GDCol(eGDCol_OptimalF)) = True
                .ColHidden(GDCol(eGDCol_RiskPercent)) = True
                .ColHidden(GDCol(eGDCol_LargestLosing)) = True
            
            Case "FIXEDRATIO"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NumCols)) = "Equity Dip"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OptimalF)) = "Equity Dip"
                .ColHidden(GDCol(eGDCol_DecLevel)) = False
                .ColHidden(GDCol(eGDCol_IncLevel)) = False
                .ColHidden(GDCol(eGDCol_OptimalF)) = True
                .ColHidden(GDCol(eGDCol_RiskPercent)) = True
                .ColHidden(GDCol(eGDCol_LargestLosing)) = True
            
            Case "OPTIMALF"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NumCols)) = "Optimal f"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OptimalF)) = "Optimal f"
                .ColHidden(GDCol(eGDCol_DecLevel)) = True
                .ColHidden(GDCol(eGDCol_IncLevel)) = True
                .ColHidden(GDCol(eGDCol_OptimalF)) = False
                .ColHidden(GDCol(eGDCol_RiskPercent)) = True
                .ColHidden(GDCol(eGDCol_LargestLosing)) = True
            
            Case "WILLIAMSRISK"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NumCols)) = "Equity Dip"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OptimalF)) = "Equity Dip"
                .ColHidden(GDCol(eGDCol_DecLevel)) = True
                .ColHidden(GDCol(eGDCol_IncLevel)) = True
                .ColHidden(GDCol(eGDCol_OptimalF)) = True
                .ColHidden(GDCol(eGDCol_RiskPercent)) = False
                .ColHidden(GDCol(eGDCol_LargestLosing)) = False
                
            Case "STANDARD"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NumCols)) = "Equity Dip"
                '.TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OptimalF)) = "Equity Dip"
                .ColHidden(GDCol(eGDCol_DecLevel)) = True
                .ColHidden(GDCol(eGDCol_IncLevel)) = True
                .ColHidden(GDCol(eGDCol_OptimalF)) = True
                .ColHidden(GDCol(eGDCol_RiskPercent)) = False
                .ColHidden(GDCol(eGDCol_LargestLosing)) = False
        
        End Select
        
        If m.strView = "Trade" Then
            .ColHidden(GDCol(eGDCol_Number)) = True
            .ColHidden(GDCol(eGDCol_Date)) = True
            .ColHidden(GDCol(eGDCol_EntryType)) = True
        Else
            .ColHidden(GDCol(eGDCol_EntryDate)) = True
            .ColHidden(GDCol(eGDCol_ExitDate)) = True
        End If
        
        .ColHidden(GDCol(eGDCol_OpenPosition)) = bSysNav
        .ColHidden(GDCol(eGDCol_Pointer)) = bSysNav And Not m.RptForm.Pyramiding
        .ColHidden(GDCol(eGDCol_SystemNumber)) = bSysNav
        .ColHidden(GDCol(eGDCol_Symbol)) = bSysNav And (m.RptForm.Trades.HeaderNumRecords = 2)
        .ColHidden(GDCol(eGDCol_Skip)) = True
        .ColHidden(GDCol(eGDCol_SkipEq)) = True
        
        For lIndex = 0 To .Cols - 1
            .ColData(lIndex) = lIndex
        Next lIndex
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        
        .FlexDataSource = Me
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Private Sub Class_Initialize()
    
    Set m.aIndex = New cGdArray
    Set m.alGraphToTrades = New cGdArray
    Set m.Stats = New cStatsSummary
    Set m.aUnits = New cGdArray
    Set m.aOpen = New cGdArray
    Set m.aTotProfit = New cGdArray
    Set m.aDrawDownPct = New cGdArray
    Set m.aAccBalance = New cGdArray
    Set m.aIncLevel = New cGdArray
    Set m.aDecLevel = New cGdArray
    Set m.aOptimalf = New cGdArray
    Set m.aRiskBasedOn = New cGdArray
    Set m.aMargins = New cGdArray
    
    m.alGraphToTrades.Create eGDARRAY_Longs
    m.aUnits.Create eGDARRAY_Longs
    m.aOpen.Create eGDARRAY_Longs
    m.aTotProfit.Create eGDARRAY_Doubles
    m.aDrawDownPct.Create eGDARRAY_Doubles
    m.aAccBalance.Create eGDARRAY_Doubles
    m.aIncLevel.Create eGDARRAY_Doubles
    m.aDecLevel.Create eGDARRAY_Doubles
    m.aOptimalf.Create eGDARRAY_Doubles
    
    m.aRiskBasedOn.Create eGDARRAY_Doubles, 0, 0
    m.aMargins.Create eGDARRAY_Doubles, 0, 0

End Sub

Private Sub Class_Terminate()
    
    Set m.aIndex = Nothing
    Set m.alGraphToTrades = Nothing
    Set m.Stats = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetBackColor
'' Description: Set the colors of the grid appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetBackColor()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim lIndex2 As Long

    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        If .Rows > 1 Then '.FixedRows Then
            For lIndex2 = .TopRow To .BottomRow
                SetColor lIndex2, ColNum(eGDCol_Profit)
                SetColor lIndex2, ColNum(eGDCol_TotalProfit)
                SetColor lIndex2, ColNum(eGDCol_MaxDip)
                
                If .Cell(flexcpValue, lIndex2, GDCol(eGDCol_Skip)) = 1 Then
                    .Cell(flexcpForeColor, lIndex2, 0, lIndex2, .Cols - 1) = vbRed
                ElseIf .Cell(flexcpValue, lIndex2, GDCol(eGDCol_SkipEq)) = 1 Then
                    .Cell(flexcpForeColor, lIndex2, 0, lIndex2, .Cols - 1) = RGB(125, 125, 125)
                End If
                
                If lIndex2 = 1 Then
                    .Cell(flexcpBackColor, lIndex2, 0, lIndex2, .Cols - 1) = .BackColor
                Else
                    If .Cell(flexcpBackColor, lIndex2 - 1, 0, lIndex2 - 1, 0) = .BackColor Then
                        .Cell(flexcpBackColor, lIndex2, 0, lIndex2, .Cols - 1) = .BackColorAlternate
                    Else
                        .Cell(flexcpBackColor, lIndex2, 0, lIndex2, .Cols - 1) = .BackColor
                    End If
                End If
            Next lIndex2
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.SetBackColor", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowFromPoint
'' Description: Get the row number in the grid for the point double clicked on
'' Inputs:      Location of the Mouse Double-Click
'' Returns:     Row in the grid to highlight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowFromPoint(ByVal lIndex2 As Double, ByVal Y As Double) As Long
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in grid to show
    Dim lTradeNbr As Long               ' Trade number of the trade
    Dim lPos As Long                    ' Location in table of the trade number
    
    lTradeNbr = gdGetNum(m.hTradeNum, m.alGraphToTrades(lIndex2))
    If m.strView = "Trade" Then lTradeNbr = gdGetNum(m.hEntryExitPtr, lTradeNbr)
    
    If gdBinarySearchAsIndex(m.aIndex.ArrayHandle, m.hTradeNum, lTradeNbr, lPos, eGdSort_Default, 0, 99999999) Then
        RowFromPoint = lPos + 2
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptMMStudy.RowFromPoint", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Resize
'' Description: When the grid gets resized, recolor the rows
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Resize()
On Error Resume Next

    With m.vsGrid
        AfterScroll .TopRow, .LeftCol, .TopRow, .LeftCol, m.RptForm
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayValues
'' Description: Display the Transaction Number, Date, and dEquity at the current
''              mouse location
'' Inputs:      Current Mouse Location, Label to display values in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayValues(ByVal X As Double, ByVal Y As Double, lbl As Object)
On Error GoTo ErrSection:

    Dim lTradeNum As Long               ' Transaction Number
    Dim dDate As Double                 ' Date of the transaction
    Dim dEquity As Double               ' dEquity without Money Management
    Dim dMMEquity As Double             ' dEquity with Money Management

    lTradeNum = m.alGraphToTrades(X - 1)
    dDate = gdGetNum(m.hTradeDate, m.alGraphToTrades(X - 1))
    dEquity = m.Chart.YData(0, X - 1)
    dMMEquity = m.Chart.YData(1, X - 1)

    With lbl
        .Caption = "Transaction: " & lTradeNum & _
                ", Date: " & DateFormat(dDate) & _
                ", Equity (w/o MM): " & Format(dEquity * m.dScaleMult, "$#,##0.00") & _
                ", Equity (with MM): " & Format(dMMEquity * m.dScaleMult, "$#,##0.00") & vbCrLf & _
                "(Double-Click on the Graph to go to the Transaction)"
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortOnCol
'' Description: Sort a column in a particular direction
'' Inputs:      Column to Sort, Direction to Sort in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SortOnCol(ByVal lCol As Long, Optional ByVal lDirection As Long = 0&)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.SortOnCol", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

Public Sub DblClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Row > 1 Then
        'm.RptForm.ShowDateInChart gdGetNum(m.hTradeDate, gdGetNum(m.aIndex.ArrayHandle, Row - 2)), 0#, gdGetNum(m.hHeaderIndex, gdGetNum(m.aIndex.ArrayHandle, Row - 2))
        
        ' TLB 7/18/2011: double-clicking a date will delay the start of MM trading
        m.dDelayedStartDate = gdGetNum(m.hTradeDate, gdGetNum(m.aIndex.ArrayHandle, Row - 2))
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMStudy.DblClick", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

Private Function Max(ByVal dOne As Double, ByVal dTwo As Double) As Double
    If dOne > dTwo Then
        Max = dOne
    Else
        Max = dTwo
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeFields
'' Description: Allow the user to change the order of the fields
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeFields()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of field information
    Dim astrUsed As New cGdArray        ' Array of fields to send to add/remove form
    Dim astrAvailable As New cGdArray   ' Array of available fields
    Dim astrDefault As New cGdArray     ' Array of default fields
    Dim astrNew As New cGdArray         ' Final array of column information
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strTemp As String               ' Temporary string
    Dim strIniFile As String            ' File name and path of the ini file

    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    
    ' Compile list of default fields...
    Select Case True
        Case m.RptForm.optTrades
            If UCase(m.strView) = "TRADE" Then
                astrFields.SplitFields m.strDefaultFieldsTrade, ","
            Else
                astrFields.SplitFields m.strDefaultFields, ","
            End If
        Case m.RptForm.optMonthly
            astrFields.SplitFields m.strDefaultFieldsMonthly, ","
        Case m.RptForm.optYearly
            astrFields.SplitFields m.strDefaultFieldsYearly, ","
    End Select
    For lIndex = 0 To astrFields.Size - 1
        If Parse(astrFields(lIndex), ";", 2) = "0" Then
            astrDefault.Add Parse(astrFields(lIndex), ";", 1)
        End If
    Next lIndex
    
    ' Compile lists of used and available fields...
    Select Case True
        Case m.RptForm.optTrades
            If UCase(m.strView) = "TRADE" Then
                astrFields.SplitFields m.strFieldsTrade, ","
            Else
                astrFields.SplitFields m.strFields, ","
            End If
        Case m.RptForm.optMonthly
            astrFields.SplitFields m.strFieldsMonthly, ","
        Case m.RptForm.optYearly
            astrFields.SplitFields m.strFieldsYearly, ","
    End Select
    For lIndex = 0 To astrFields.Size - 1
        Select Case Parse(astrFields(lIndex), ";", 2)
            Case "1"
                astrNew.Add astrFields(lIndex)
            Case "0"
                astrUsed.Add Parse(astrFields(lIndex), ";", 1)
            Case "-1"
                astrAvailable.Add Parse(astrFields(lIndex), ";", 1)
        End Select
    Next lIndex
    
    ' Show the add/remove form...
    If frmAddRemove.ShowMe(astrAvailable, astrUsed, eOrderMode_Ordered, astrDefault) Then
        ' If the user hit OK, recompile the fields list and setup the columns...
        For lIndex = 0 To astrUsed.Size - 1
            For lIndex2 = 0 To astrFields.Size - 1
                If Parse(astrFields(lIndex2), ";", 1) = astrUsed(lIndex) Then
                    If Len(Parse(astrFields(lIndex2), ";", 3)) > 0 Then
                        astrNew.Add astrUsed(lIndex) & ";0;" & Parse(astrFields(lIndex2), ";", 3)
                    Else
                        astrNew.Add astrUsed(lIndex) & ";0"
                    End If
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        
        For lIndex = 0 To astrAvailable.Size - 1
            For lIndex2 = 0 To astrFields.Size - 1
                If Parse(astrFields(lIndex2), ";", 1) = astrAvailable(lIndex) Then
                    If Len(Parse(astrFields(lIndex2), ";", 3)) > 0 Then
                        astrNew.Add astrAvailable(lIndex) & ";-1;" & Parse(astrFields(lIndex2), ";", 3)
                    Else
                        astrNew.Add astrAvailable(lIndex) & ";-1"
                    End If
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        
        Select Case True
            Case m.RptForm.optTrades
                If UCase(m.strView) = "TRADE" Then
                    m.strFieldsTrade = astrNew.JoinFields(",")
                    SetIniFileProperty "FieldsTrade", m.strFieldsTrade, Str(m.lAppID), strIniFile
                Else
                    m.strFields = astrNew.JoinFields(",")
                    SetIniFileProperty "Fields", m.strFields, Str(m.lAppID), strIniFile
                End If
            Case m.RptForm.optMonthly
                m.strFieldsMonthly = astrNew.JoinFields(",")
                SetIniFileProperty "FieldsMonthly", m.strFieldsMonthly, Str(m.lAppID), strIniFile
            Case m.RptForm.optYearly
                m.strFieldsYearly = astrNew.JoinFields(",")
                SetIniFileProperty "FieldsYearly", m.strFieldsYearly, Str(m.lAppID), strIniFile
        End Select
        
        SetUpColumns False
    End If
    
ErrExit:
    Set astrFields = Nothing
    Set astrUsed = Nothing
    Set astrAvailable = Nothing
    Set astrDefault = Nothing
    Set astrNew = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    Set astrUsed = Nothing
    Set astrAvailable = Nothing
    Set astrDefault = Nothing
    Set astrNew = Nothing
    RaiseError "cRptMMStudy.ChangeFields", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpColumns
'' Description: Set up the column order/width/visibility according to spec
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpColumns(Optional ByVal bDoAllThree As Boolean = True)
On Error GoTo ErrSection:

    If m.RptForm.optTrades Or bDoAllThree Then SetUpTradeColumns
    If m.RptForm.optMonthly Or bDoAllThree Then SetUpMonthlyColumns
    If m.RptForm.optYearly Or bDoAllThree Then SetUpYearlyColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.SetUpColumns", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveCols
'' Description: Save the column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveCols()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of field information
    Dim lIndex As Long                  ' Index into a for loop
    Dim bContinue As Boolean            ' Should we continue?
    
    astrFields.Create eGDARRAY_Strings
    
    Select Case True
        Case m.RptForm.optTrades
            For lIndex = 0 To m.vsGrid.Cols - 1
                If Len(m.vsGrid.TextMatrix(0, lIndex)) > 0 Then
                    bContinue = True
                    If UCase(m.strView) = "TRADE" Then
                        If InStr(m.strFieldsTrade, m.vsGrid.TextMatrix(0, lIndex) & ";") = 0 Then
                            bContinue = False
                        End If
                    Else
                        If InStr(m.strFields, m.vsGrid.TextMatrix(0, lIndex) & ";") = 0 Then
                            bContinue = False
                        End If
                    End If
                    
                    If bContinue Then
                        If m.vsGrid.ColHidden(lIndex) = True Then
                            astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";-1;" & Str(m.vsGrid.ColWidth(lIndex))
                        Else
                            astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";0;" & Str(m.vsGrid.ColWidth(lIndex))
                        End If
                    End If
                End If
            Next lIndex
        
            If UCase(m.strView) = "TRADE" Then
                m.strFieldsTrade = astrFields.JoinFields(",")
                SetIniFileProperty "FieldsTrade", m.strFieldsTrade, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
                m.bAutoSizeTrade = False
            Else
                m.strFields = astrFields.JoinFields(",")
                SetIniFileProperty "Fields", m.strFields, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
                m.bAutoSize = False
            End If
            
        Case m.RptForm.optMonthly
            For lIndex = 0 To m.fgMonthly.Cols - 1
                If Len(m.fgMonthly.TextMatrix(0, lIndex)) > 0 Then
                    bContinue = True
                    If InStr(m.strFieldsMonthly, m.fgMonthly.TextMatrix(0, lIndex) & ";") = 0 Then
                        bContinue = False
                    End If
                    
                    If bContinue Then
                        If m.fgMonthly.ColHidden(lIndex) = True Then
                            astrFields.Add m.fgMonthly.TextMatrix(0, lIndex) & ";-1;" & Str(m.fgMonthly.ColWidth(lIndex))
                        Else
                            astrFields.Add m.fgMonthly.TextMatrix(0, lIndex) & ";0;" & Str(m.fgMonthly.ColWidth(lIndex))
                        End If
                    End If
                End If
            Next lIndex
        
            m.strFieldsMonthly = astrFields.JoinFields(",")
            SetIniFileProperty "FieldsMonthly", m.strFieldsMonthly, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
            m.bAutoSizeMonthly = False
            
        Case m.RptForm.optYearly
            For lIndex = 0 To m.fgYearly.Cols - 1
                If Len(m.fgYearly.TextMatrix(0, lIndex)) > 0 Then
                    bContinue = True
                    If InStr(m.strFieldsYearly, m.fgYearly.TextMatrix(0, lIndex) & ";") = 0 Then
                        bContinue = False
                    End If
                    
                    If bContinue Then
                        If m.fgYearly.ColHidden(lIndex) = True Then
                            astrFields.Add m.fgYearly.TextMatrix(0, lIndex) & ";-1;" & Str(m.fgYearly.ColWidth(lIndex))
                        Else
                            astrFields.Add m.fgYearly.TextMatrix(0, lIndex) & ";0;" & Str(m.fgYearly.ColWidth(lIndex))
                        End If
                    End If
                End If
            Next lIndex
        
            m.strFieldsYearly = astrFields.JoinFields(",")
            SetIniFileProperty "FieldsYearly", m.strFieldsYearly, Str(m.lAppID), AddSlash(g.strAppPath) & "Reports.INI"
            m.bAutoSizeYearly = False
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.SaveCols", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Function ColNum(ByVal Col As eGD_Cols) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    For lIndex = 0 To m.vsGrid.Cols - 1
        If m.vsGrid.ColData(lIndex) = Col Then
            ColNum = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptMMStudy.ColNum", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateMonthly
'' Description: Calculate the monthly statistics for the money management run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalculateMonthly()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lMonth As Long                  ' Index into a for loop
    Dim lExit As Long                   ' Index for the corresponding exit signal
    Dim Stats As New cStatsSummary      ' Stats summary class to utilize
    Dim dBeginDate As Double            ' Earliest trade date
    Dim dEndDate As Double              ' Latest trade date
    Dim dFromDate As Double             ' Date to start current calculation
    Dim dToDate As Double               ' Date to stop current calculation
    Dim dActualFrom As Double           ' Actual starting date
    Dim dActualTo As Double             ' Actual ending date
    Dim dDate As Double                 ' Date of the trade
    Dim bFirst As Boolean               ' Is this the first trade?
    Dim dNetProfit As Double            ' Net Profit
    Dim dPctIncrease As Double          ' Percent increase
    Dim lMousePointer As Long           ' Current state of the mouse pointer
    
    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    ' Figure out the beginning and ending dates...
    For lIndex2 = 1 To gdGetSize(m.hTradeDate) - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
                If dBeginDate = 0 Then dBeginDate = gdGetNum(m.hTradeDate, lIndex)
                dEndDate = gdGetNum(m.hTradeDate, lIndex)
            End If
        End If
    Next lIndex2
    
    Stats.MinProfit = m.dMinProfit
    bFirst = True
    dNetProfit = m.dBeginBalance
    
    ' Figure out the from and to dates for the first calculation...
    dFromDate = DateValue("1/1/" & Str(Year(dBeginDate)))
    dToDate = DateValue("12/31/" & Str(Year(dBeginDate)))
    
    Do While dFromDate < dEndDate
        ' Calculate the actual to and from values to calculate...
        dActualFrom = dFromDate
        dActualTo = dToDate
        If dFromDate < dBeginDate Then dActualFrom = dBeginDate
        If dToDate > dEndDate Then dActualTo = dEndDate
        If InStr(1, dActualTo, ".") <= 0 Then dActualTo = dToDate + 0.9999
        
        For lMonth = 1 To 12
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                lIndex = gdGetNum(m.hSorted, lIndex2)
                lExit = gdGetNum(m.hEntryExitPtr, lIndex)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    dDate = CDate(gdGetNum(m.hTradeDate, lExit))
                    If dDate >= dActualFrom And dDate <= dActualTo And Month(dDate) = lMonth Then
                        gdSetNum m.hSkipReport, lIndex, 0
                    Else
                        gdSetNum m.hSkipReport, lIndex, 1
                    End If
                End If
            Next lIndex2
        
            If Not bFirst Then
                gdSetNum m.hAccountBalance, 0, m.Stats.EndingBalance
            End If
            bFirst = False
            
            Stats.SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            Stats.SetSortHandle m.RptForm.Trades.SortHandle
            Stats.Calculate False, False, False, True
            
            If dNetProfit <> 0 Then
                dPctIncrease = Stats.TotalNetProfit / dNetProfit
            End If
            dNetProfit = dNetProfit + Stats.TotalNetProfit
            AddLineToMonthly Format(DateOf((Year(dFromDate) * 10000) + (lMonth * 100) + 1), "mmm yyyy"), Stats, dNetProfit, dPctIncrease
        Next lMonth
        
        dFromDate = DateAdd("yyyy", 1, dFromDate)     'add one year to from date
        dToDate = DateValue("12/31/" & Str(Year(dFromDate)))
    Loop
    
    m.fgMonthly.AutoSize 0, m.fgMonthly.Cols - 1, False, 75
    Screen.MousePointer = lMousePointer
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.CalculateMonthly", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateYearly
'' Description: Calculate the yearly statistics for the money management run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalculateYearly()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lMonth As Long                  ' Index into a for loop
    Dim lExit As Long                   ' Index for the corresponding exit signal
    Dim Stats As New cStatsSummary      ' Stats summary class to utilize
    Dim dBeginDate As Double            ' Earliest trade date
    Dim dEndDate As Double              ' Latest trade date
    Dim dFromDate As Double             ' Date to start current calculation
    Dim dToDate As Double               ' Date to stop current calculation
    Dim dActualFrom As Double           ' Actual starting date
    Dim dActualTo As Double             ' Actual ending date
    Dim dDate As Double                 ' Date of the trade
    Dim bFirst As Boolean               ' Is this the first trade?
    Dim dNetProfit As Double            ' Net Profit
    Dim dPctIncrease As Double          ' Percent increase
    Dim lMousePointer As Long           ' Current state of the mouse pointer
    
    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    ' Figure out the beginning and ending dates...
    For lIndex2 = 1 To gdGetSize(m.hTradeDate) - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
            If (gdGetNum(m.hSkip, lIndex) = 0) And (gdGetNum(m.hSkipEq, lIndex) = 0) Then
                If dBeginDate = 0 Then dBeginDate = gdGetNum(m.hTradeDate, lIndex)
                dEndDate = gdGetNum(m.hTradeDate, lIndex)
            End If
        End If
    Next lIndex2
    
    Stats.MinProfit = m.dMinProfit
    bFirst = True
    dNetProfit = m.dBeginBalance
    
    ' Figure out the from and to dates for the first calculation...
    dFromDate = DateValue("1/1/" & Str(Year(dBeginDate)))
    dToDate = DateValue("12/31/" & Str(Year(dBeginDate)))
    
    Do While dFromDate < dEndDate
        ' Calculate the actual to and from values to calculate...
        dActualFrom = dFromDate
        dActualTo = dToDate
        If dFromDate < dBeginDate Then dActualFrom = dBeginDate
        If dToDate > dEndDate Then dActualTo = dEndDate
        If InStr(1, dActualTo, ".") <= 0 Then dActualTo = dToDate + 0.9999
        
        'For lMonth = 1 To 12
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                lIndex = gdGetNum(m.hSorted, lIndex2)
                lExit = gdGetNum(m.hEntryExitPtr, lIndex)
                
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    dDate = CDate(gdGetNum(m.hTradeDate, lExit))
                    If dDate >= dActualFrom And dDate <= dActualTo Then ' And Month(dDate) = lMonth Then
                        gdSetNum m.hSkipReport, lIndex, 0
                    Else
                        gdSetNum m.hSkipReport, lIndex, 1
                    End If
                End If
            Next lIndex2
        
            If Not bFirst Then
                gdSetNum m.hAccountBalance, 0, m.Stats.EndingBalance
            End If
            bFirst = False
            
            Stats.SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
            Stats.SetSortHandle m.RptForm.Trades.SortHandle
            Stats.Calculate False, False, False, True
            
            If dNetProfit <> 0 Then
                dPctIncrease = Stats.TotalNetProfit / dNetProfit
            End If
            dNetProfit = dNetProfit + Stats.TotalNetProfit
            AddLineToYearly Str(Year(dFromDate)), Stats, dNetProfit, dPctIncrease
        'Next lMonth
        
        dFromDate = DateAdd("yyyy", 1, dFromDate)     'add one year to from date
        dToDate = DateValue("12/31/" & Str(Year(dFromDate)))
    Loop
    
    m.fgYearly.AutoSize 0, m.fgYearly.Cols - 1, False, 75
    Screen.MousePointer = lMousePointer
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.CalculateYearly", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitMonthlyGrid
'' Description: Initialize the monthly grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitMonthlyGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    'Format grid columns and headings...
    With m.fgMonthly
        lRedraw = .Redraw
        .Redraw = flexRDNone
        ClearGrid m.fgMonthly
        
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = g.nAltGridRowColor '&HC0FFFF
        .AutoSizeMode = flexAutoSizeColWidth
        .Rows = 2
        .FixedRows = 1
        .Cols = MonthCol(eGDMonthlyCol_NumCols)
        .FixedCols = 0
        .FrozenCols = 1
        
        .TextMatrix(0, MonthCol(eGDMonthlyCol_Month)) = "Month"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_TotTrades)) = "Trades"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_WinPct)) = "Win Pct"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_WinAvg)) = "Win Avg"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_LossAvg)) = "Loss Avg"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_PLRatio)) = "PayoutR" '"P/L"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_PF)) = "ProfitF" '"P/F"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_AvgTrade)) = "AvgTrade"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_AvgRunUp)) = "Run Up"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_AvgRunDown)) = "Run Dn"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_Profit)) = "Profit"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_NetProfit)) = "Net Profit"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_LossConsec)) = "C/L"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_LossLargest)) = "Max Loss"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_AvgNumBars)) = "Avg Bars"
        .TextMatrix(0, MonthCol(eGDMonthlyCol_ROI)) = "%Increase"
        
        .TextMatrix(1, MonthCol(eGDMonthlyCol_NetProfit)) = Format(m.dBeginBalance, FormatDollar(m.strShowCents))
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.InitMonthlyGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitYearlyGrid
'' Description: Initialize the yearly grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitYearlyGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    'Format grid columns and headings...
    With m.fgYearly
        lRedraw = .Redraw
        .Redraw = flexRDNone
        ClearGrid m.fgYearly
        
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = False
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = g.nAltGridRowColor ' &HC0FFFF
        .AutoSizeMode = flexAutoSizeColWidth
        .Rows = 2
        .FixedRows = 1
        .Cols = YearCol(eGDYearlyCol_NumCols)
        .FixedCols = 0
        .FrozenCols = 1
        
        .TextMatrix(0, YearCol(eGDYearlyCol_Year)) = "Year"
        .TextMatrix(0, YearCol(eGDYearlyCol_TotTrades)) = "Trades"
        .TextMatrix(0, YearCol(eGDYearlyCol_WinPct)) = "Win Pct"
        .TextMatrix(0, YearCol(eGDYearlyCol_WinAvg)) = "Win Avg"
        .TextMatrix(0, YearCol(eGDYearlyCol_LossAvg)) = "Loss Avg"
        .TextMatrix(0, YearCol(eGDYearlyCol_PLRatio)) = "PayoutR" '"P/L"
        .TextMatrix(0, YearCol(eGDYearlyCol_PF)) = "ProfitF" '"P/F"
        .TextMatrix(0, YearCol(eGDYearlyCol_AvgTrade)) = "AvgTrade"
        .TextMatrix(0, YearCol(eGDYearlyCol_AvgRunUp)) = "Run Up"
        .TextMatrix(0, YearCol(eGDYearlyCol_AvgRunDown)) = "Run Dn"
        .TextMatrix(0, YearCol(eGDYearlyCol_Profit)) = "Profit"
        .TextMatrix(0, YearCol(eGDYearlyCol_NetProfit)) = "Net Profit"
        .TextMatrix(0, YearCol(eGDYearlyCol_LossConsec)) = "C/L"
        .TextMatrix(0, YearCol(eGDYearlyCol_LossLargest)) = "Max Loss"
        .TextMatrix(0, YearCol(eGDYearlyCol_AvgNumBars)) = "Avg Bars"
        .TextMatrix(0, YearCol(eGDYearlyCol_ROI)) = "%Increase"
                
        .TextMatrix(1, MonthCol(eGDMonthlyCol_NetProfit)) = Format(m.dBeginBalance, FormatDollar(m.strShowCents))
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.InitYearlyGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLineToMonthly
'' Description: Add a line to the monthly grid
'' Inputs:      Month, Stats, Net Profit, %Increase
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLineToMonthly(ByVal strMonth As String, ByVal Stats As cStatsSummary, ByVal dNetProfit As Double, ByVal dPctIncrease As Double)
On Error GoTo ErrSection:

    With m.fgMonthly
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_Month)) = strMonth
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_TotTrades)) = Format(Stats.TotTrades, "#,##0")
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_WinPct)) = Format(Stats.WinPct, "##0.00%")
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_WinAvg)) = Format(Stats.WinAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_LossAvg)) = Format(Stats.LossAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_PLRatio)) = Format(Stats.PLRatio, "##0.00")
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_PF)) = Format(Stats.ProfitFactor, "##0.00")
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_AvgTrade)) = Format(Stats.AvgTrade, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_AvgRunUp)) = Format(Stats.WinRunUpAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_AvgRunDown)) = Format(Stats.LossRunDownAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_Profit)) = Format(Stats.TotalNetProfit, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_NetProfit)) = Format(dNetProfit, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_LossConsec)) = Format(Stats.LossConsec, "#,##0")
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_LossLargest)) = Format(Stats.LossLargest, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_AvgNumBars)) = Format(Stats.AvgBarsInTrade, "#,##0.00")
        .TextMatrix(.Rows - 1, MonthCol(eGDMonthlyCol_ROI)) = Format(dPctIncrease, "#,##0.00%")
        
        ColorNegValue m.fgMonthly, Stats.WinAvg, .Rows - 1, MonthCol(eGDMonthlyCol_WinAvg)
        ColorNegValue m.fgMonthly, Stats.LossAvg, .Rows - 1, MonthCol(eGDMonthlyCol_LossAvg)
        ColorNegValue m.fgMonthly, Stats.AvgTrade, .Rows - 1, MonthCol(eGDMonthlyCol_AvgTrade)
        ColorNegValue m.fgMonthly, Stats.WinRunUpAvg, .Rows - 1, MonthCol(eGDMonthlyCol_AvgRunUp)
        ColorNegValue m.fgMonthly, Stats.LossRunDownAvg, .Rows - 1, MonthCol(eGDMonthlyCol_AvgRunDown)
        ColorNegValue m.fgMonthly, Stats.TotalNetProfit, .Rows - 1, MonthCol(eGDMonthlyCol_Profit)
        ColorNegValue m.fgMonthly, dNetProfit, .Rows - 1, MonthCol(eGDMonthlyCol_NetProfit)
        ColorNegValue m.fgMonthly, Stats.LossLargest, .Rows - 1, MonthCol(eGDMonthlyCol_LossLargest)
        ColorNegValue m.fgMonthly, dPctIncrease, .Rows - 1, MonthCol(eGDMonthlyCol_ROI)
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.AddLineToMonthly", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLineToYearly
'' Description: Add a line to the yearly grid
'' Inputs:      Month, Stats, Net Profit, %Increase
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLineToYearly(ByVal strYear As String, ByVal Stats As cStatsSummary, ByVal dNetProfit As Double, ByVal dPctIncrease As Double)
On Error GoTo ErrSection:

    With m.fgYearly
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_Year)) = strYear
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_TotTrades)) = Format(Stats.TotTrades, "#,##0")
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_WinPct)) = Format(Stats.WinPct, "##0.00%")
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_WinAvg)) = Format(Stats.WinAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_LossAvg)) = Format(Stats.LossAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_PLRatio)) = Format(Stats.PLRatio, "##0.00")
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_PF)) = Format(Stats.ProfitFactor, "##0.00")
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_AvgTrade)) = Format(Stats.AvgTrade, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_AvgRunUp)) = Format(Stats.WinRunUpAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_AvgRunDown)) = Format(Stats.LossRunDownAvg, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_Profit)) = Format(Stats.TotalNetProfit, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_NetProfit)) = Format(dNetProfit, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_LossConsec)) = Format(Stats.LossConsec, "#,##0")
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_LossLargest)) = Format(Stats.LossLargest, FormatDollar(m.strShowCents))
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_AvgNumBars)) = Format(Stats.AvgBarsInTrade, "#,##0.00")
        .TextMatrix(.Rows - 1, YearCol(eGDYearlyCol_ROI)) = Format(dPctIncrease, "#,##0.00%")
        
        ColorNegValue m.fgYearly, Stats.WinAvg, .Rows - 1, YearCol(eGDYearlyCol_WinAvg)
        ColorNegValue m.fgYearly, Stats.LossAvg, .Rows - 1, YearCol(eGDYearlyCol_LossAvg)
        ColorNegValue m.fgYearly, Stats.AvgTrade, .Rows - 1, YearCol(eGDYearlyCol_AvgTrade)
        ColorNegValue m.fgYearly, Stats.WinRunUpAvg, .Rows - 1, YearCol(eGDYearlyCol_AvgRunUp)
        ColorNegValue m.fgYearly, Stats.LossRunDownAvg, .Rows - 1, YearCol(eGDYearlyCol_AvgRunDown)
        ColorNegValue m.fgYearly, Stats.TotalNetProfit, .Rows - 1, YearCol(eGDYearlyCol_Profit)
        ColorNegValue m.fgYearly, dNetProfit, .Rows - 1, YearCol(eGDYearlyCol_NetProfit)
        ColorNegValue m.fgYearly, Stats.LossLargest, .Rows - 1, YearCol(eGDYearlyCol_LossLargest)
        ColorNegValue m.fgYearly, dPctIncrease, .Rows - 1, YearCol(eGDYearlyCol_ROI)
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMStudy.AddLineToYearly", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpTradeColumns
'' Description: Set up the columns for the by trade (or transaction) report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpTradeColumns()
On Error GoTo ErrSection:

    Dim strIniFile As String            ' File path and name of the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Array of field information
    Dim strColName As String            ' Column Name
    Dim strHidden As String             ' Is Column Hidden?
    Dim lColWidth As Long               ' Width of the column
    Dim lTotWidth As Long               ' Total width from the fields string
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColPos As Long                 ' Column position
    
    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    m.strFields = GetIniFileProperty("Fields", m.strDefaultFields, Str(m.lAppID), strIniFile)
    m.strFieldsTrade = GetIniFileProperty("FieldsTrade", m.strDefaultFieldsTrade, Str(m.lAppID), strIniFile)
    lTotWidth = 0&
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If UCase(m.strView) = "TRADE" Then
            astrFields.SplitFields m.strFieldsTrade, ","
        Else
            astrFields.SplitFields m.strFields, ","
        End If
        
        lColPos = 0&
        For lIndex = 0 To astrFields.Size - 1
            strColName = Parse(astrFields(lIndex), ";", 1)
            strHidden = Parse(astrFields(lIndex), ";", 2)
            lColWidth = CLng(ValOfText(Parse(astrFields(lIndex), ";", 3)))
            
            If strHidden <> "1" Then lColPos = lColPos + 1
            
            If strColName = "Largest Loss" Then
                If m.dRiskBasedOn = -2 Then
                    strColName = "Average Loss"
                ElseIf m.dRiskBasedOn = 0 Then
                    strColName = "Stop Loss"
                ElseIf m.dRiskBasedOn > 0 Then
                    strColName = "Risk Based On"
                End If
            End If
            
            For lCol = 0 To .Cols - 2
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    If strHidden <> "1" Then
                        .ColPosition(lCol) = lColPos
                        lTotWidth = lTotWidth + lColWidth
                        If strHidden = "-1" Then
                            .ColHidden(lColPos) = True
                        Else
                            .ColHidden(lColPos) = False
                        End If
                        .ColWidth(lColPos) = lColWidth
                    Else
                        .ColWidth(lCol) = lColWidth
                    End If
                    Exit For
                End If
            Next lCol
        Next lIndex
        
        If lTotWidth = 0& Then
            ResizeColWidths
            
            If UCase(m.strView) = "TRADE" Then
                m.bAutoSizeTrade = True
            Else
                m.bAutoSize = True
            End If
        Else
            If UCase(m.strView) = "TRADE" Then
                m.bAutoSizeTrade = False
            Else
                m.bAutoSize = False
            End If
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Set astrFields = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "cRptMMStudy.SetUpTradeColumns", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpMonthlyColumns
'' Description: Set up the columns for the monthly report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpMonthlyColumns()
On Error GoTo ErrSection:

    Dim strIniFile As String            ' File path and name of the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Array of field information
    Dim strColName As String            ' Column Name
    Dim strHidden As String             ' Is Column Hidden?
    Dim lColWidth As Long               ' Width of the column
    Dim lTotWidth As Long               ' Total width from the fields string
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColPos As Long                 ' Column position
    
    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    m.strFieldsMonthly = GetIniFileProperty("FieldsMonthly", m.strDefaultFieldsMonthly, Str(m.lAppID), strIniFile)
    lTotWidth = 0&
    
    With m.fgMonthly
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        astrFields.SplitFields m.strFieldsMonthly, ","
        
        lColPos = 0&
        For lIndex = 0 To astrFields.Size - 1
            strColName = Parse(astrFields(lIndex), ";", 1)
            strHidden = Parse(astrFields(lIndex), ";", 2)
            lColWidth = CLng(ValOfText(Parse(astrFields(lIndex), ";", 3)))
            
            If strHidden <> "1" Then lColPos = lColPos + 1
            
            For lCol = 0 To .Cols - 2
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    If strHidden <> "1" Then
                        .ColPosition(lCol) = lColPos
                        lTotWidth = lTotWidth + lColWidth
                        If strHidden = "-1" Then
                            .ColHidden(lColPos) = True
                        Else
                            .ColHidden(lColPos) = False
                        End If
                        .ColWidth(lColPos) = lColWidth
                    Else
                        .ColWidth(lCol) = lColWidth
                    End If
                    Exit For
                End If
            Next lCol
        Next lIndex
        
        If lTotWidth = 0& Then
            m.bAutoSizeMonthly = True
            .AutoSize 0, .Cols - 1, False, 75
        Else
            m.bAutoSizeMonthly = False
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Set astrFields = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "cRptMMStudy.SetUpMonthlyColumns", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpYearlyColumns
'' Description: Set up the columns for the yearly report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpYearlyColumns()
On Error GoTo ErrSection:

    Dim strIniFile As String            ' File path and name of the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Array of field information
    Dim strColName As String            ' Column Name
    Dim strHidden As String             ' Is Column Hidden?
    Dim lColWidth As Long               ' Width of the column
    Dim lTotWidth As Long               ' Total width from the fields string
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColPos As Long                 ' Column position
    
    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    m.strFieldsYearly = GetIniFileProperty("FieldsYearly", m.strDefaultFieldsYearly, Str(m.lAppID), strIniFile)
    lTotWidth = 0&
    
    With m.fgYearly
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        astrFields.SplitFields m.strFieldsYearly, ","
        
        lColPos = 0&
        For lIndex = 0 To astrFields.Size - 1
            strColName = Parse(astrFields(lIndex), ";", 1)
            strHidden = Parse(astrFields(lIndex), ";", 2)
            lColWidth = CLng(ValOfText(Parse(astrFields(lIndex), ";", 3)))
            
            If strHidden <> "1" Then lColPos = lColPos + 1
            
            For lCol = 0 To .Cols - 2
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    If strHidden <> "1" Then
                        .ColPosition(lCol) = lColPos
                        lTotWidth = lTotWidth + lColWidth
                        If strHidden = "-1" Then
                            .ColHidden(lColPos) = True
                        Else
                            .ColHidden(lColPos) = False
                        End If
                        .ColWidth(lColPos) = lColWidth
                    Else
                        .ColWidth(lCol) = lColWidth
                    End If
                    Exit For
                End If
            Next lCol
        Next lIndex
        
        If lTotWidth = 0& Then
            m.bAutoSizeYearly = True
            .AutoSize 0, .Cols - 1, False, 75
        Else
            m.bAutoSizeYearly = False
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Set astrFields = Nothing
    Exit Sub
    
ErrSection:
    Set astrFields = Nothing
    RaiseError "cRptMMStudy.SetUpYearlyColumns", eGDRaiseError_Raise
    
End Sub

Public Property Get DelayedStartDate() As Double
    DelayedStartDate = m.dDelayedStartDate
End Property

Public Property Let DelayedStartDate(ByVal dDelayedStartDate As Double)
    m.dDelayedStartDate = dDelayedStartDate
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportToCSV
'' Description: Export the grid to a comma separated format
'' Inputs:      Filename to Export to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExportToCSV(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array to dump to the file
    Dim astrLine As New cGdArray        ' Array for column headers
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim s$, strFmt$
    
    astrFile.Create eGDARRAY_Strings
    astrLine.Create eGDARRAY_Strings
    
    With m.vsGrid
        For lRow = 0 To .Rows - 1
            astrLine.Size = 0
            For lCol = 0 To .Cols - 1
                If .ColHidden(lCol) = False Then
                    s = .TextMatrix(lRow, lCol)
#If 0 Then ' CustomText formatting is in cRptSummaryBy, but not supported here
                    If lRow > 0 And astrLine.Size = 0 Then
                        strFmt = .ColFormat(GDCol(eGDCol_CustomText))
                        If Len(strFmt) > 0 Then
                            s = Format(Val(s), strFmt)
                        End If
                    End If
#End If
                    astrLine.Add s
                End If
            Next lCol
            astrFile.Add astrLine.JoinFields(vbTab)
        Next
    End With
    
    astrFile.ToFile strFileName

ErrExit:
    Set astrFile = Nothing
    Exit Sub
    
ErrSection:
    Set astrFile = Nothing
    RaiseError "cRptMMStudy.ExportToCSV", eGDRaiseError_Raise, g.strAppPath
End Sub

