VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptMMTrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Implements IVSFlexDataSource

Private Type mPrivate
    dMinProfit As Currency
    dBeginBalance As Currency
    dReportName As String
    strShowCents As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    RtfMsg As RichTextBox
    elMsg As vsElastic
    elGrid As vsElastic
    
    hRuleNames As Long
    hSymbolHandle As Long
    
    'Work variables required for report to run
    strPriceFormat As String
    lSubHeadRow As Long
    strShowTime As String
    
    'Trade field handles
    hTradeNbr As Long
    hPosition As Long
    hSignalType As Long
    hTradeDate As Long
    hPrice As Long
    hRuleID As Long
    hSignalIndex As Long
    hUnits As Long
    hProfit As Long
    hTotalProfit As Long
    hAccountBalance As Long
    hSymbolIndex As Long
    hSysNbr As Long
    hAllocation As Long
    hMaxProfit As Long
    hMaxLoss As Long
    hBarsInTrade As Long
    hSkip As Long
    hShow As Long
    hSkipRpt As Long
    hTDOW As Long
    hTDOM As Long
    hTDOY As Long
    hRank As Long
    hEquityAvail As Long
    hConflict As Long
    hOpenTrade As Long
    hOpenTradesTotal As Long
    hSignalsTot As Long
    hLink As Long
    hMsg As Long
    hEntryExitPtr As Long
    hSorted As Long
    hHeaderIndex As Long

    aIndex As cGdArray
    strChartBy As String
    lBottomRow As Long
    alGraphToTrades As cGdArray
    
    BeginDate As Double
    EndDate As Double
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Nbr = 0
    eGDCol_Date = 1
    eGDCol_EntryType = 2
    eGDCol_Pos = 3
    eGDCol_SystemNumber = 4
    eGDCol_Ptr = 5
    eGDCol_Msg = 6
    eGDCol_Symbol = 7
    eGDCol_Conflict = 8
    eGDCol_OpenPos = 9
    eGDCol_Rank = 10
    eGDCol_Profit = 11
    eGDCol_Units = 12
    eGDCol_TotProfit = 13
    eGDCol_AccBal = 14
    eGDCol_Allocation = 15
    eGDCol_EquityAvail = 16
    eGDCol_MaxProfit = 17
    eGDCol_MaxLoss = 18
    eGDCol_Signal = 19
    eGDCol_Price = 20
    eGDCol_TradeID = 21
    eGDCol_Skip = 22
    eGDCol_Link = 23
End Enum
Private Const kGridCols = 24

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    
    IVSFlexDataSource_GetFieldCount = kGridCols

End Function
Private Function IVSFlexDataSource_GetRecordCount() As Long
    
    IVSFlexDataSource_GetRecordCount = m.RptForm.Trades.NumRecords - 1

End Function
Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    
    IVSFlexDataSource_GetFieldName = m.vsGrid.TextMatrix(0, Field)

End Function

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal R As Long) As String
On Error GoTo ErrSection:

    R = R + 1
    
    Select Case Field
        Case GDCol(eGDCol_Link): IVSFlexDataSource_GetData = Str(gdGetNum(m.hLink, R))
        Case GDCol(eGDCol_Nbr): IVSFlexDataSource_GetData = Str(R)
        Case GDCol(eGDCol_Date)
            If gdGetNum(m.hSignalsTot, R) > 1 Then
                IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R))) & " " & Format(gdGetNum(m.hTradeDate, R), "hh:mm AM/PM")
            Else
                If m.strShowTime = "Yes" Then
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R))) & " " & Format(gdGetNum(m.hTradeDate, R), "hh:mm AM/PM")
                Else
                    IVSFlexDataSource_GetData = DateFormat(CDate(gdGetNum(m.hTradeDate, R)))
                End If
            End If
            
        Case GDCol(eGDCol_Rank)
            If gdGetNum(m.hRank, R) < 0 Then
                IVSFlexDataSource_GetData = "0"
            Else
                IVSFlexDataSource_GetData = Format(gdGetNum(m.hRank, R), "##0")
            End If
        Case GDCol(eGDCol_Conflict)
            IVSFlexDataSource_GetData = Format(gdGetNum(m.hSignalsTot, R), "##0")
        Case GDCol(eGDCol_OpenPos)
            IVSFlexDataSource_GetData = Format(gdGetNum(m.hOpenTradesTotal, R), "##0")
        Case GDCol(eGDCol_EntryType)
            If gdGetNum(m.hSignalType, R) = gEntrySignal Then
                IVSFlexDataSource_GetData = "E"
            Else
                IVSFlexDataSource_GetData = "X"
            End If
        Case GDCol(eGDCol_Pos)
            If gdGetNum(m.hPosition, R) = 1 Then
                IVSFlexDataSource_GetData = "L"
            Else
                IVSFlexDataSource_GetData = "S"
            End If
        Case GDCol(eGDCol_Msg):
            If gdGetStr(m.hMsg, R) <> "" Then
                IVSFlexDataSource_GetData = "Y"
            End If
        Case GDCol(eGDCol_Ptr):
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hEntryExitPtr, R))
        Case GDCol(eGDCol_Symbol):
            IVSFlexDataSource_GetData = gdGetStr(m.hSymbolHandle, gdGetNum(m.hSymbolIndex, R))
        Case GDCol(eGDCol_Price)
            Dim i As Long
            i = m.RptForm.SystemArrayInd(gdGetNum(m.hSysNbr, R))
            If m.strPriceFormat = "Ticks" Then
                IVSFlexDataSource_GetData = _
                    gdFormatPrice(gdGetNum(m.hPrice, R), _
                    m.RptForm.Systems.Num(i, ensy_TickMove), m.RptForm.Systems.Num(i, ensy_MinMoveInTicks), 0)
            Else
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, R), _
                    m.RptForm.Systems.Num(i, ensy_TickMove), m.RptForm.Systems.Num(i, ensy_MinMoveInTicks), -1)
            End If
        Case GDCol(eGDCol_Signal):
            IVSFlexDataSource_GetData = gdGetStr(m.hRuleNames, gdGetNum(m.hSignalIndex, R))
        Case GDCol(eGDCol_Profit)
            If gdGetNum(m.hSignalType, R) = gEntrySignal Then
                 IVSFlexDataSource_GetData = Format(0, FormatDollar(m.strShowCents))
            Else
                 IVSFlexDataSource_GetData = Format(gdGetNum(m.hProfit, R), FormatDollar(m.strShowCents))
            End If
        Case GDCol(eGDCol_Units): IVSFlexDataSource_GetData = Format(gdGetNum(m.hUnits, R), "#.#")
        Case GDCol(eGDCol_TotProfit): IVSFlexDataSource_GetData = Format(gdGetNum(m.hTotalProfit, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_AccBal): IVSFlexDataSource_GetData = Format(gdGetNum(m.hAccountBalance, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_Allocation): IVSFlexDataSource_GetData = Format(gdGetNum(m.hAllocation, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_EquityAvail): IVSFlexDataSource_GetData = Format(gdGetNum(m.hEquityAvail, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_MaxProfit): IVSFlexDataSource_GetData = Format(gdGetNum(m.hMaxProfit, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_MaxLoss): IVSFlexDataSource_GetData = Format(gdGetNum(m.hMaxLoss, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_SystemNumber): IVSFlexDataSource_GetData = Format(m.RptForm.SystemArrayInd(gdGetNum(m.hSysNbr, R)), "##0")
        Case GDCol(eGDCol_TradeID): IVSFlexDataSource_GetData = Str(gdGetNum(m.hTradeNbr, R))
        Case GDCol(eGDCol_Skip): IVSFlexDataSource_GetData = Str(gdGetNum(m.hSkip, R))
    End Select
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cRptMMTrades.GetData", eGDRaiseError_Raise, m.RptForm.AppPath

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
    ByVal NewTopRow As Long, ByVal NewLeftCol As Long, pForm As Object)
On Error GoTo ErrSection:

    ''If Not pForm Is Nothing Then ResizeColWidths
    SetBackColor

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMTrades.AfterScroll", eGDRaiseError_Raise, m.RptForm.AppPath

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
    RaiseError "cRptMMTrades.ResizeColWidths", eGDRaiseError_Raise, m.RptForm.AppPath

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
    RaiseError "cRptMMTrades.SetColor", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowColChange
'' Description: Show messages for current trade selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RowColChange()
On Error GoTo ErrSection:

    If m.vsGrid.Row <= 0 Then Exit Sub
    If gdGetStr(m.hMsg, m.vsGrid.Row) = "" Then Exit Sub
    m.RtfMsg.Text = gdGetStr(m.hMsg, m.vsGrid.Row)
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMTrades.RowColChange", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim bFound As Boolean               ' Whether the entity was found
    
    Set m.RptForm = RptForm
    Set m.Chart = RptForm!BarChart
    Set m.vsGrid = RptForm!vsImplements
    Set m.RtfMsg = RptForm!rtfMessages
    Set m.elMsg = RptForm!elMessages
    Set m.elGrid = RptForm!elImplements
    
    m.hRuleNames = m.RptForm.Trades.RuleNameHandle
    m.hSymbolHandle = m.RptForm.Trades.SymbolHandle
    
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hTradeNbr = .FieldHandle(entd_TradeNbr)
        m.hPosition = .FieldHandle(entd_Position)
        m.hSignalType = .FieldHandle(entd_SignalType)
        m.hTradeDate = .FieldHandle(entd_TradeDate)
        m.hPrice = .FieldHandle(entd_Price)
        m.hRuleID = .FieldHandle(entd_RuleID)
        m.hSignalIndex = .FieldHandle(entd_SignalIndex)
        m.hUnits = .FieldHandle(entd_Units)
        m.hProfit = .FieldHandle(entd_Profit)
        m.hTotalProfit = .FieldHandle(entd_TotalProfit)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hSymbolIndex = .FieldHandle(entd_SymbolIndex)
        m.hSysNbr = .FieldHandle(entd_SysNbr)
        m.hAllocation = .FieldHandle(entd_Allocation)
        m.hMaxProfit = .FieldHandle(entd_MaxProfit)
        m.hMaxLoss = .FieldHandle(entd_MaxLoss)
        m.hBarsInTrade = .FieldHandle(entd_BarsInTrade)
        m.hSkip = .FieldHandle(entd_Skip)
        m.hSkipRpt = .FieldHandle(entd_SkipRpt)
        m.hTDOW = .FieldHandle(entd_TradeDayOfWeek)
        m.hTDOM = .FieldHandle(entd_TradeDayOfMonth)
        m.hTDOY = .FieldHandle(entd_TradeDayOfYear)
        m.hRank = .FieldHandle(entd_Rank)
        m.hEquityAvail = .FieldHandle(entd_EquityAvail)
        m.hConflict = .FieldHandle(entd_Conflict)
        m.hLink = .FieldHandle(entd_Link)
        m.hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        m.hOpenTradesTotal = .FieldHandle(entd_OpenTradesTotal)
        m.hSignalsTot = .FieldHandle(entd_SignalsTot)
        m.hMsg = .FieldHandle(entd_Msg)
        m.hShow = .FieldHandle(entd_Show)
        m.hSorted = .SortHandle
        m.hHeaderIndex = .FieldHandle(entd_HeaderIndex)
    End With
    
    'Report options
    m.dBeginBalance = m.RptForm.OptionValue("BeginBalance")
    m.strPriceFormat = m.RptForm.OptionValue("PriceFormat")
    m.strShowTime = m.RptForm.OptionValue("ShowTime")
    m.dMinProfit = m.RptForm.OptionValue("MinProfit")
    m.dReportName = m.RptForm.OptionValue("ReportName")
    m.strShowCents = m.RptForm.OptionValue("ShowCents")
    m.strChartBy = m.RptForm.OptionValue("ChartBy")
    m.BeginDate = m.RptForm.FromDate
    m.EndDate = m.RptForm.ToDate
        
    m.vsGrid.Redraw = flexRDNone
    InitGrid
    
    'Autosense mode (money mgt vs single unit)
    If gdGetNum(m.hAccountBalance, 1) <= 0 Then
        m.RptForm.SetToSingleUnit True
    End If
    
    BuildChart
    
    'Search for at last one message on a trade.  If found, show the Msg
    'column vsImplements and show vsMessages grid.
    If Not m.elMsg Is Nothing Then
        For lIndex = 1 To m.RptForm.Trades.NumRecords - 1
            If gdGetStr(m.hMsg, lIndex) <> "" Then
                bFound = True
                Exit For
            End If
        Next lIndex
        
        If bFound Then
            m.elMsg.Visible = True
            m.vsGrid.ColHidden(GDCol(eGDCol_Msg)) = False
        Else
            m.elMsg.Visible = False
            m.vsGrid.ColHidden(GDCol(eGDCol_Msg)) = True
        End If
        
        'This refreshes the implements grid.
        m.elGrid.Align = 5       '5-asFill
        m.elGrid.Refresh
    End If
    
    m.vsGrid.FrozenCols = 0
    ResizeColWidths
    m.vsGrid.FrozenCols = 3
    SetBackColor
    m.vsGrid.Redraw = flexRDBuffered
    
    'This searches for the first Non-skipped records BEFORE the
    'start date and advances to this.  When the row is displayed, it
    'is located at the end of the grid visually.  To bring it "near"
    'the top of the grid, add 20 rows to FirstRow
    Dim FirstRow As Long
    Dim RetVal As Long
    FirstRow = 0
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lIndex = gdGetNum(m.hSorted, lIndex2)
        
        If gdGetNum(m.hTradeDate, lIndex) > m.BeginDate Then
            FirstRow = lIndex
            Exit For
        End If
    Next lIndex2
    m.vsGrid.Row = FirstRow + 23
    RetVal = m.vsGrid.CellTop   '<bring row into view
    m.vsGrid.Row = FirstRow

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMTrades.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
    
    Dim lItemsToChart As Long           ' Number of items to chart
    Dim lTrade As Long                  ' Index into a for loop
    Dim dEquity As Double               ' Running Equity
    Dim dMaxIntra As Double             ' Max Intra-Trade Drawdown
    Dim dEntryDate As Double            ' Entry Date for the System
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
    
    Dim dStartDate As Double            ' Starting date to chart
    Dim dEndDate As Double              ' Ending date to chart
    Dim dPrevDate As Double             ' Previous date processed
    Dim lDate As Long                   ' Index into a for loop
    Dim lAnnot As Long                  ' Annotation index
    Dim lPrev As Long                   ' Previous index processed
    Dim YData() As Single               ' Array to assign to the Y-Axis
    Dim XData() As Single               ' Array to assign to the X-Axis
    
    ' Calculate Items to chart (number of rows in grid)
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords
        lTrade = gdGetNum(m.hSorted, lIndex2)
    
        If gdGetNum(m.hSignalType, lTrade) = 1 Then
            If gdGetNum(m.hSkip, lTrade) = 0 Then
                If dStartDate = 0 Then dStartDate = gdGetNum(m.hTradeDate, gdGetNum(m.hEntryExitPtr, lTrade))
                dEndDate = gdGetNum(m.hTradeDate, lTrade)
                lIndex = lIndex + 1
            End If
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
        .Points = lItemsToChart
        
        ReDim YData(lItemsToChart * .Subsets) As Single
        ReDim XData(lItemsToChart) As Single

        .MainTitle = "Equity Curve"
        .SubTitle = ""
        .YAxisLabel = "" ' = "Equity"
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

        lIndex = 0
        dEquity = m.dBeginBalance
        dPeak = -999999#
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords
            lTrade = gdGetNum(m.hSorted, lIndex2)
            
            If gdGetNum(m.hSkip, lTrade) = 0 Then
                If gdGetNum(m.hSignalType, lTrade) = 1 Then
                    If m.strChartBy = "Date" Then
                        If dPrevDate = 0 Then dPrevDate = dStartDate
                        dEntryDate = gdGetNum(m.hTradeDate, gdGetNum(m.hEntryExitPtr, lTrade))
                        dMaxIntra = dEquity + gdGetNum(m.hMaxLoss, lTrade)
                        For lDate = Int(dPrevDate) To Int(gdGetNum(m.hTradeDate, lTrade)) - 1
                            lIndex = lIndex + 1
                            YData(lIndex - 1) = dEquity
                            XData(lIndex - 1) = lDate
                            m.alGraphToTrades(lIndex) = lPrev
                            
                            If Month(lDate) = 1 And Day(lDate) = 1 Then
                                .VertLineAnnotationType(lAnnot) = PELAT_GRIDLINE
                                .VertLineAnnotation(lAnnot) = lIndex
                                .VertLineAnnotationText(lAnnot) = "|B" & Format(lDate, "yyyy")
                                .VertLineAnnotationColor(lAnnot) = 1
                                lAnnot = lAnnot + 1
                            End If
                        Next lDate
                        lPrev = lTrade
                        dPrevDate = gdGetNum(m.hTradeDate, lTrade)
                        dEquity = dEquity + gdGetNum(m.hTotalProfit, lTrade)
                    Else
                        lIndex = lIndex + 1
                        dEquity = dEquity + gdGetNum(m.hTotalProfit, lTrade)
                        dMaxIntra = dEquity + gdGetNum(m.hMaxLoss, lTrade)
                        YData(lIndex - 1) = dEquity
                        XData(lIndex - 1) = lIndex - 1
                        .PointLabels(lIndex - 1) = lIndex - 1
                        m.alGraphToTrades(lIndex - 1) = gdGetNum(m.hTradeNbr, lTrade)
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
                            dDateOfFlat = gdGetNum(m.hTradeDate, lTrade)
                        End If
                        
                        lNumDaysFlat = 0&
                    Else
                        lNumDaysFlat = gdGetNum(m.hTradeDate, lTrade) - gdGetNum(m.hTradeDate, lPeakPos)
                        dDrawdown = dEquity - dPeak
                        If dDrawdown < dMaxDrawdown Then
                            dMaxDrawdown = dDrawdown
                            lLocationOfMax = lIndex
                            If m.strChartBy = "Trade" Then lLocationOfMax = lLocationOfMax - 1
                            dDateOfMax = gdGetNum(m.hTradeDate, lTrade)
                        End If
                    End If
                End If
            End If
        Next lIndex2
        If m.strChartBy = "Date" Then
            YData(lIndex) = dEquity
            XData(lIndex) = dPrevDate
            m.alGraphToTrades(lIndex + 1) = lPrev
        End If
        
        ' Assign the XData and YData points to the chart all at once
        PEvset .hObject, PEP_faYDATA, YData(0), .Points * .Subsets
        PEvset .hObject, PEP_faXDATA, XData(0), .Points

        If lNumDaysFlat > lLongestFlat Then
            lLongestFlat = lNumDaysFlat
            lLocOfFlat = lIndex
            If m.strChartBy = "Trade" Then lLocOfFlat = lLocOfFlat - 1
        End If
            
        '** Set SubsetLabels property array for 4 subsets **'
        .SubsetLabels(0) = "Running Total of Equity"
        
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
        .GraphAnnotationY(2) = .YData(0, lLocationOfMax)  'dPeak + dMaxDrawdown
        .GraphAnnotationColor(2) = vbBlack
        .LegendAnnotationType(0) = PEGAT_SMALLUPTRIANGLESOLID
        .LegendAnnotationText(0) = "Max Drawdown: " & Format(dMaxDrawdown, "$#,##0.00") & " on " & DateFormat(dDateOfMax)
        .LegendAnnotationColor(0) = vbBlack
        
        ' Add the Longest Flat Period to the Graph
        .GraphAnnotationType(3) = PEGAT_SMALLDOWNTRIANGLESOLID
        .GraphAnnotationX(3) = lLocOfFlat + 1
        .GraphAnnotationY(3) = .YData(0, lLocOfFlat)
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
    RaiseError "cRptTrades2.BuildChart", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

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
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        ClearGrid m.vsGrid
        
        .BackColorAlternate = g.nAltGridRowColor '&HC0FFFF
        .ExplorerBar = flexExSortShowAndMove
        .GridLines = flexGridNone
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .HighLight = flexHighlightAlways
        .Rows = 1
        .FixedRows = 1
        .Cols = kGridCols
        .FixedCols = 0
        .FrozenCols = 3
        .AutoSizeMode = flexAutoSizeColWidth
        .AllowUserResizing = flexResizeColumns
        .ScrollTips = True
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True

        For lIndex = 0 To .Cols - 1
            .ColHidden(lIndex) = False
        Next lIndex
        
        m.lSubHeadRow = 0
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Nbr)) = "Num"
        .ColWidth(GDCol(eGDCol_Nbr)) = 600
        .ColAlignment(GDCol(eGDCol_Nbr)) = flexAlignLeftTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Date)) = "Date"
        .ColDataType(GDCol(eGDCol_Date)) = flexDTDate
        .ColWidth(GDCol(eGDCol_Date)) = 900
        .ColAlignment(GDCol(eGDCol_Date)) = flexAlignLeftTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntryType)) = "Type"
        .ColWidth(GDCol(eGDCol_EntryType)) = 350
        .ColAlignment(GDCol(eGDCol_EntryType)) = flexAlignCenterTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Msg)) = "M"
        .ColWidth(GDCol(eGDCol_Msg)) = 200
        .ColAlignment(GDCol(eGDCol_Msg)) = flexAlignCenterTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Pos)) = "Pos"
        .ColWidth(GDCol(eGDCol_Pos)) = 350
        .ColAlignment(GDCol(eGDCol_Pos)) = flexAlignCenterTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Ptr)) = "Ptr"
        .ColDataType(GDCol(eGDCol_Ptr)) = flexDTLong
        .ColWidth(GDCol(eGDCol_Ptr)) = 475
        .ColAlignment(GDCol(eGDCol_Ptr)) = flexAlignCenterTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SystemNumber)) = "Strategy"
        .ColWidth(GDCol(eGDCol_SystemNumber)) = 400
        .ColAlignment(GDCol(eGDCol_SystemNumber)) = flexAlignCenterTop
        .ColHidden(GDCol(eGDCol_SystemNumber)) = (UCase(m.RptForm.AppName) = "SYSTEM NAVIGATOR")
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Symbol)) = "Sym"
        .ColWidth(GDCol(eGDCol_Symbol)) = 500
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Conflict)) = "C"
        .ColAlignment(GDCol(eGDCol_Conflict)) = flexAlignCenterTop
        .ColWidth(GDCol(eGDCol_Conflict)) = 300
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_OpenPos)) = "O"
        .ColAlignment(GDCol(eGDCol_OpenPos)) = flexAlignCenterTop
        .ColWidth(GDCol(eGDCol_OpenPos)) = 300
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Price)) = "Price"
        .ColDataType(GDCol(eGDCol_Price)) = flexDTDouble
        .ColAlignment(GDCol(eGDCol_Price)) = flexAlignRightTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Profit)) = "Profit"
        .ColDataType(GDCol(eGDCol_Profit)) = flexDTCurrency
        .ColAlignment(GDCol(eGDCol_Profit)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_Units)) = 900
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Units)) = "Units"
        .ColDataType(GDCol(eGDCol_Units)) = flexDTDouble
        .ColAlignment(GDCol(eGDCol_Units)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_Units)) = 700
   
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotProfit)) = "Tot Profit"
        .ColDataType(GDCol(eGDCol_TotProfit)) = flexDTCurrency
        .ColAlignment(GDCol(eGDCol_TotProfit)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_TotProfit)) = 1000

        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_AccBal)) = "Equity"
        .ColDataType(GDCol(eGDCol_AccBal)) = flexDTCurrency
        .ColAlignment(GDCol(eGDCol_AccBal)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_AccBal)) = 1250

        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Allocation)) = "Allocation"
        .ColDataType(GDCol(eGDCol_Allocation)) = flexDTCurrency
        .ColAlignment(GDCol(eGDCol_Allocation)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_Allocation)) = 1000

        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EquityAvail)) = "Equity Avail"
        .ColDataType(GDCol(eGDCol_EquityAvail)) = flexDTCurrency
        .ColAlignment(GDCol(eGDCol_EquityAvail)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_EquityAvail)) = 1250

        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_MaxProfit)) = "Max Profit"
        .ColDataType(GDCol(eGDCol_MaxProfit)) = flexDTCurrency
        .ColAlignment(GDCol(eGDCol_MaxProfit)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_MaxProfit)) = 1000

        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_MaxLoss)) = "Max Loss"
        .ColAlignment(GDCol(eGDCol_MaxLoss)) = flexAlignRightTop
        .ColWidth(GDCol(eGDCol_MaxLoss)) = 1000
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Signal)) = "Signal"
        .ColAlignment(GDCol(eGDCol_Signal)) = flexAlignLeftTop
        .ColWidth(GDCol(eGDCol_Signal)) = 2000
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Rank)) = "R"
        .ColWidth(GDCol(eGDCol_Rank)) = 300
        .ColAlignment(GDCol(eGDCol_Rank)) = flexAlignCenterTop
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TradeID)) = "Trade ID"
        .ColAlignment(GDCol(eGDCol_TradeID)) = flexAlignLeftTop
        .ColWidth(GDCol(eGDCol_TradeID)) = 700
        
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Link)) = "Link"
    
        .ColHidden(GDCol(eGDCol_Skip)) = True
        .ColHidden(GDCol(eGDCol_TradeID)) = True
        .ColHidden(GDCol(eGDCol_Link)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .FlexDataSource = Me
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMTrades.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetBackColor
'' Description: Set the background color of the rows appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetBackColor()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop

    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        If .Rows > .FixedRows Then
            For lIndex = .TopRow To .BottomRow
                SetColor lIndex, GDCol(eGDCol_Profit)
                SetColor lIndex, GDCol(eGDCol_TotProfit)
                SetColor lIndex, GDCol(eGDCol_AccBal)
                SetColor lIndex, GDCol(eGDCol_MaxProfit)
                SetColor lIndex, GDCol(eGDCol_MaxLoss)
                
                If .Cell(flexcpValue, lIndex, GDCol(eGDCol_Skip)) = 1 Then
                    .Cell(flexcpForeColor, lIndex, 0, lIndex, .Cols - 1) = vbRed
                End If
                
                If lIndex = 1 Then
                    .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = .BackColor
                Else
                    If .Cell(flexcpBackColor, lIndex - 1, 0, lIndex - 1, 0) = .BackColor Then
                        .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = .BackColorAlternate
                    Else
                        .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = .BackColor
                    End If
                End If
            Next lIndex
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMTrades.SetBackColor", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowFromPoint
'' Description: Get the row number in the grid for the point double clicked on
'' Inputs:      Location of the Mouse Double-Click
'' Returns:     Row in the grid to highlight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowFromPoint(ByVal X As Double, ByVal Y As Double) As Long
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in grid to show
    Dim lTradeNbr As Long               ' Trade number of the trade
    Dim lPos As Long                    ' Location in table of the trade number
    
    lTradeNbr = gdGetNum(m.hTradeNbr, m.alGraphToTrades(X))
    RowFromPoint = lTradeNbr
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptMMTrades.RowFromPoint", eGDRaiseError_Raise, m.RptForm.AppPath
    
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
'' Description: Display the Transaction Number, Date, and Equity at the current
''              mouse location
'' Inputs:      Current Mouse Location, Label to display values in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayValues(ByVal X As Double, ByVal Y As Double, lbl As Object)
On Error GoTo ErrSection:

    Dim lTradeNum As Long               ' Transaction Number
    Dim dProfit As Double               ' Profit of the current transaction
    Dim dDate As Double                 ' Date of the current transaction
    Dim dEquity As Double               ' Equity of the current transaction

    lTradeNum = gdGetNum(m.hTradeNbr, m.alGraphToTrades(X))
    dProfit = gdGetNum(m.hTotalProfit, m.alGraphToTrades(X))
    dDate = gdGetNum(m.hTradeDate, m.alGraphToTrades(X))
    dEquity = m.Chart.YData(0, X)

    With lbl
        .Caption = "Transaction: " & lTradeNum & _
                ", Date: " & DateFormat(dDate) & _
                ", Equity: " & Format(dEquity, "$#,##0.00") & vbCrLf & _
                "(Double-Click on the Graph to go to the Transaction)"
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptMMTrades.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath
    
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
    RaiseError "cRptMMTrades.SortOnCol", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Public Sub DblClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Row > 1 Then
        m.RptForm.ShowDateInChart gdGetNum(m.hTradeDate, gdGetNum(m.aIndex.ArrayHandle, Row - 2)), gdGetNum(m.hPrice, gdGetNum(m.aIndex.ArrayHandle, Row - 2)), gdGetNum(m.hHeaderIndex, gdGetNum(m.aIndex.ArrayHandle, Row - 2))
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptMMTrades.DblClick", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

