VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptTrades2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Implements IVSFlexDataSource

Private Type mPrivate
    aIndex As cGdArray
    lBottomRow As Long
    alGraphToTrades As cGdArray
    Stats As cStatsSummary

    'Work variables required for report to run
    strPriceFormat As String
    lSubHeadRow As Long
    strShowTime As String
    strShowCents As String
    strView As String
    RptForm As Form
    Chart As Pegoa
    vsGrid As VSFlexGrid
    RtfMsg As RichTextBox
    
    strDefaultFields As String
    strDefaultFieldsTrade As String
    strFields As String
    strFieldsTrade As String
    lAppID As Long
    bAutoSize As Boolean
    bAutoSizeTrade As Boolean
    
    ' Handles into Trades lookup arrays..
    hRuleIndex As Long
    hSymbolHandle As Long

    ' Trade field handles
    hTradeNumber As Long
    hPosition As Long
    hSignalType As Long
    hTradeDate As Long
    hPrice As Long
    hRuleID As Long
    hUnits As Long
    hNumShares As Long
    hProfit As Long
    hTotalProfit As Long
    hAccountBalance As Long
    hSymbolIndex As Long
    hSystemNumber As Long
    hAllocation As Long
    hMaxProfit As Long
    hMaxLoss As Long
    hBarsInTrade As Long
    hSkip As Long
    hShow As Long
    hSkipReport As Long
    hTDOW As Long
    hTDOM As Long
    hTDOY As Long
    hRank As Long
    hEquityAvailable As Long
    hConflict As Long
    hOpenTrade As Long
    hOpenTradesTotal As Long
    hSignalsTotal As Long
    hLink As Long
    hEntryExitPtr As Long
    hMsg As Long
    hSignalIndex As Long
    hEquity As Long
    hEquityMA As Long
    hSkipEqFilter As Long
    hFilteredEquity As Long
    hSortIndex As Long
    hHeaderIndex As Long
    
    alNum As cGdArray
    adProfit As cGdArray
    adNetProfit As cGdArray
    adEntryDate As cGdArray
    adEntryPrice As cGdArray
    alEntrySignal As cGdArray
    
    iSortedCol As Long
    bSortedDescending As Boolean
    bRecalcIndex As Boolean
End Type
Private m As mPrivate

Private Enum eGD_Cols
    eGDCol_Number = 0
    eGDCol_Skip
    eGDCol_TradeID
    eGDCol_Link
    eGDCol_Date
    eGDCol_EntryType
    eGDCol_Position
    eGDCol_Price
    eGDCol_SystemNumber
    eGDCol_Pointer
    eGDCol_Symbol
    eGDCol_EntryDate
    eGDCol_EntryPrice
    eGDCol_ExitDate
    eGDCol_ExitPrice
    eGDCol_NumShares
    eGDCol_Units
    eGDCol_Profit
    eGDCol_TotalProfit
    eGDCol_NetProfit
    eGDCol_MaxProfit
    eGDCol_MaxLoss
    eGDCol_EntrySignal
    eGDCol_ExitSignal
    eGDCol_SkipEq
    eGDCol_NumCols
End Enum

Private Function GDCol(ByVal lColumn As eGD_Cols)
    GDCol = lColumn
End Function

Private Sub Class_Initialize()

    Set m.Stats = New cStatsSummary
    Set m.aIndex = New cGdArray
    m.aIndex.Create eGDARRAY_Longs
    Set m.alGraphToTrades = New cGdArray
    m.alGraphToTrades.Create eGDARRAY_Longs
    Set m.alNum = New cGdArray
    m.alNum.Create eGDARRAY_Longs
    Set m.adProfit = New cGdArray
    m.adProfit.Create eGDARRAY_Doubles
    Set m.adNetProfit = New cGdArray
    m.adNetProfit.Create eGDARRAY_Doubles
    
    Set m.adEntryDate = New cGdArray
    m.adEntryDate.Create eGDARRAY_Doubles
    Set m.adEntryPrice = New cGdArray
    m.adEntryPrice.Create eGDARRAY_Doubles
    Set m.alEntrySignal = New cGdArray
    m.alEntrySignal.Create eGDARRAY_Longs
        
End Sub

Private Sub Class_Terminate()

    Set m.Stats = Nothing
    Set m.aIndex = Nothing
    Set m.alGraphToTrades = Nothing
    Set m.alNum = Nothing
    Set m.adProfit = Nothing
    Set m.adNetProfit = Nothing
    
    Set m.adEntryDate = Nothing
    Set m.adEntryPrice = Nothing
    Set m.alEntrySignal = Nothing

End Sub

Private Function IVSFlexDataSource_GetFieldCount() As Long

    IVSFlexDataSource_GetFieldCount = GDCol(eGDCol_NumCols)
    
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long

    IVSFlexDataSource_GetRecordCount = m.aIndex.Size
    
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    
    IVSFlexDataSource_GetFieldName = m.vsGrid.TextMatrix(0, Field)

End Function

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal R As Long) As String
On Error GoTo ErrSection:
    
    Dim lExit As Long
    Dim i As Long
       
    If m.strView = "Trade" Then
        ' TLB 8/14/2012: for Trades, the Exit is now getting passed in, so find the Entry
        lExit = m.aIndex(R)
        R = gdGetNum(m.hEntryExitPtr, lExit)
    Else
        R = m.aIndex(R)
        'R = m.aIndex(gdGetNum(m.hSortIndex, R))
        lExit = gdGetNum(m.hEntryExitPtr, R)
    End If
    
    Select Case Field
        Case GDCol(eGDCol_Number)
            IVSFlexDataSource_GetData = Str(R)
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
        Case GDCol(eGDCol_EntryType)
            If gdGetNum(m.hSignalType, R) = gEntrySignal Then
                IVSFlexDataSource_GetData = "E"
            Else
                IVSFlexDataSource_GetData = "X"
            End If
        Case GDCol(eGDCol_Position)
            If gdGetNum(m.hPosition, R) = 1 Then
                IVSFlexDataSource_GetData = "L"
            Else
                IVSFlexDataSource_GetData = "S"
            End If
        Case GDCol(eGDCol_Price)
            i = m.RptForm.SystemArrayInd(gdGetNum(m.hSystemNumber, R))
            If m.strPriceFormat = "Ticks" Then
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, R), _
                        m.RptForm.Systems.Num(i, ensy_TickMove), , 0)
            Else
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, R), _
                        m.RptForm.Systems.Num(i, ensy_TickMove), , -1)
            End If
        Case GDCol(eGDCol_SystemNumber)
            IVSFlexDataSource_GetData = Str(m.RptForm.SystemArrayInd(gdGetNum(m.hSystemNumber, R)))
        Case GDCol(eGDCol_Pointer)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hEntryExitPtr, R))
        Case GDCol(eGDCol_Symbol)
            IVSFlexDataSource_GetData = gdGetStr(m.hSymbolHandle, gdGetNum(m.hSymbolIndex, R))
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
        Case GDCol(eGDCol_EntryPrice)
            i = m.RptForm.SystemArrayInd(gdGetNum(m.hSystemNumber, R))
            If m.strPriceFormat = "Ticks" Then
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, R), _
                        m.RptForm.Systems.Num(i, ensy_TickMove), , 0)
            Else
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, R), _
                        m.RptForm.Systems.Num(i, ensy_TickMove), , -1)
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
        Case GDCol(eGDCol_ExitPrice)
            i = m.RptForm.SystemArrayInd(gdGetNum(m.hSystemNumber, lExit))
            If m.strPriceFormat = "Ticks" Then
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, lExit), _
                        m.RptForm.Systems.Num(i, ensy_TickMove), , 0)
            Else
                IVSFlexDataSource_GetData = gdFormatPrice(gdGetNum(m.hPrice, lExit), _
                        m.RptForm.Systems.Num(i, ensy_TickMove), , -1)
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
        Case GDCol(eGDCol_NumShares)
            i = gdGetNum(m.hNumShares, R)
            If i > 0 Then
                IVSFlexDataSource_GetData = Str(i)
            Else
                IVSFlexDataSource_GetData = "" ' display blank if no shares (e.g. a Future in a basket with a Stock)
            End If
        Case GDCol(eGDCol_Units)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hUnits, R))
        Case GDCol(eGDCol_TotalProfit)
            IVSFlexDataSource_GetData = Format(gdGetNum(m.hTotalProfit, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_NetProfit)
            If m.strView = "Trade" Then
                IVSFlexDataSource_GetData = Format(gdGetNum(m.hAccountBalance, lExit), FormatDollar(m.strShowCents))
            Else
                IVSFlexDataSource_GetData = Format(gdGetNum(m.hAccountBalance, R), FormatDollar(m.strShowCents))
            End If
        Case GDCol(eGDCol_MaxProfit)
            IVSFlexDataSource_GetData = Format(gdGetNum(m.hMaxProfit, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_MaxLoss)
            IVSFlexDataSource_GetData = Format(gdGetNum(m.hMaxLoss, R), FormatDollar(m.strShowCents))
        Case GDCol(eGDCol_EntrySignal)
            IVSFlexDataSource_GetData = gdGetStr(m.hRuleIndex, gdGetNum(m.hSignalIndex, R))
        Case GDCol(eGDCol_ExitSignal)
            IVSFlexDataSource_GetData = gdGetStr(m.hRuleIndex, gdGetNum(m.hSignalIndex, lExit))
        Case GDCol(eGDCol_Skip)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hSkip, R))
        Case GDCol(eGDCol_TradeID)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hTradeNumber, R))
        Case GDCol(eGDCol_Link)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hLink, R))
        Case GDCol(eGDCol_SkipEq)
            IVSFlexDataSource_GetData = Str(gdGetNum(m.hSkipEqFilter, R))
    End Select
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cRptTrades2.GetData", eGDRaiseError_Raise, m.RptForm.AppPath

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
    RaiseError "cRptTrades2.AfterScroll", eGDRaiseError_Raise, m.RptForm.AppPath

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
    RaiseError "cRptTrades2.ResizeColWidths", eGDRaiseError_Raise, m.RptForm.AppPath

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
    RaiseError "cRptTrades2.SetColor", eGDRaiseError_Raise, m.RptForm.AppPath
    
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
    RaiseError "cRptTrades2.RowColChange", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Dim lExit As Long
    Dim lEntry As Long
    
    Set m.RptForm = RptForm
    Set m.Chart = RptForm!BarChart
    Set m.vsGrid = RptForm!vsImplements
    Set m.RtfMsg = RptForm!rtfMessages
    
    RptForm.ShowViewOptions = False
    
    ' Get some information from the form...
    m.strDefaultFields = m.RptForm.OptionValue("DefaultFields")
    m.strDefaultFieldsTrade = m.RptForm.OptionValue("DefaultFieldsTrade")
    m.lAppID = m.RptForm.ReportID
    
    ' Get the header to the Rule Name array from the Trades...
    m.hRuleIndex = m.RptForm.Trades.RuleNameHandle
    m.hSymbolHandle = m.RptForm.Trades.SymbolHandle
        
    'Save global Field addresses from Trades class
    With m.RptForm.Trades
        m.hTradeNumber = .FieldHandle(entd_TradeNbr)
        m.hPosition = .FieldHandle(entd_Position)
        m.hSignalType = .FieldHandle(entd_SignalType)
        m.hTradeDate = .FieldHandle(entd_TradeDate)
        m.hPrice = .FieldHandle(entd_Price)
        m.hRuleID = .FieldHandle(entd_RuleID)
        m.hSignalIndex = .FieldHandle(entd_SignalIndex)
        m.hNumShares = .FieldHandle(entd_NumShares)
        m.hUnits = .FieldHandle(entd_Units)
        m.hProfit = .FieldHandle(entd_Profit)
        m.hTotalProfit = .FieldHandle(entd_TotalProfit)
        m.hAccountBalance = .FieldHandle(entd_AccountBalance)
        m.hSymbolIndex = .FieldHandle(entd_SymbolIndex)
        m.hSystemNumber = .FieldHandle(entd_SysNbr)
        m.hAllocation = .FieldHandle(entd_Allocation)
        m.hMaxProfit = .FieldHandle(entd_MaxProfit)
        m.hMaxLoss = .FieldHandle(entd_MaxLoss)
        m.hBarsInTrade = .FieldHandle(entd_BarsInTrade)
        m.hSkip = .FieldHandle(entd_Skip)
        m.hSkipReport = .FieldHandle(entd_SkipRpt)
        m.hTDOW = .FieldHandle(entd_TradeDayOfWeek)
        m.hTDOM = .FieldHandle(entd_TradeDayOfMonth)
        m.hTDOY = .FieldHandle(entd_TradeDayOfYear)
        m.hRank = .FieldHandle(entd_Rank)
        m.hEquityAvailable = .FieldHandle(entd_EquityAvail)
        m.hConflict = .FieldHandle(entd_Conflict)
        m.hLink = .FieldHandle(entd_Link)
        m.hEntryExitPtr = .FieldHandle(entd_EntryExitPtr)
        m.hOpenTradesTotal = .FieldHandle(entd_OpenTradesTotal)
        m.hSignalsTotal = .FieldHandle(entd_SignalsTot)
        m.hMsg = .FieldHandle(entd_Msg)
        m.hShow = .FieldHandle(entd_Show)
        m.hEquity = .FieldHandle(entd_Equity)
        m.hEquityMA = .FieldHandle(entd_EquityMA)
        m.hSkipEqFilter = .FieldHandle(entd_SkipEqFilter)
        m.hFilteredEquity = .FieldHandle(entd_FilteredEquity)
        m.hSortIndex = .SortHandle
        m.hHeaderIndex = .FieldHandle(entd_HeaderIndex)
    End With
            
    ' Calculate Statistics...
    With m.Stats
        .SetTradeHandles m.RptForm.Trades.DtlTableHandle, m.RptForm.Trades.HdrTableHandle
        .SetSortHandle m.hSortIndex
        
        .StartDate = m.RptForm.OptionValue("FromDate")
        .EndDate = m.RptForm.OptionValue("ToDate")
        
        .MovAvgPeriod = m.RptForm.EquityFilter.MovingAveragePeriod
        .MovAvgType = m.RptForm.EquityFilter.MovingAverageType
        .EquityFilterOn = m.RptForm.EquityFilter.EquityFilterOn
        .EquityFilterMode = m.RptForm.EquityFilter.EquityFilterMode
                
        .Calculate True
    
        m.RptForm.TakeNextTrade .TakeNextTrade
    End With
    
    m.strView = m.RptForm.OptionValue("ViewBy")
    m.RptForm.txtReportName(1) = m.strView & "s"
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
        m.vsGrid.Redraw = flexRDNone
        m.RptForm.SetToSingleUnit False
        m.RptForm.FilterTrades
           
        ' Skip all of the exits
        If m.strView = "Trade" Then
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                lIndex = gdGetNum(m.hSortIndex, lIndex2)
                
#If 1 Then
                ' TLB 8/14/2012: use sortation based on the Exit dates
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    gdSetNum m.hShow, lIndex, 0 ' hide the entry rows
                Else
                    lExit = lIndex
                    lEntry = gdGetNum(m.hEntryExitPtr, lIndex)
                    gdSetNum m.adProfit.ArrayHandle, lIndex, gdGetNum(m.hProfit, lExit)
                    gdSetNum m.adNetProfit.ArrayHandle, lIndex, gdGetNum(m.hAccountBalance, lExit)
                    gdSetNum m.adEntryPrice.ArrayHandle, lIndex, gdGetNum(m.hPrice, lEntry)
                    gdSetNum m.alEntrySignal.ArrayHandle, lIndex, gdGetNum(m.hSignalIndex, lEntry)
                    gdSetNum m.adEntryDate.ArrayHandle, lIndex, gdGetNum(m.hTradeDate, lEntry)
                End If
#Else
                
                If gdGetNum(m.hSignalType, lIndex) = gExitSignal Then
                    gdSetNum m.hShow, lIndex, 0
                Else
                    lExit = gdGetNum(m.hEntryExitPtr, lIndex)
                    gdSetNum m.adExitDate.ArrayHandle, lIndex, gdGetNum(m.hTradeDate, lExit)
                    gdSetNum m.adExitPrice.ArrayHandle, lIndex, gdGetNum(m.hPrice, lExit)
                    gdSetNum m.adProfit.ArrayHandle, lIndex, gdGetNum(m.hProfit, lExit)
                    gdSetNum m.adNetProfit.ArrayHandle, lIndex, gdGetNum(m.hAccountBalance, lExit)
                    gdSetNum m.alExitSignal.ArrayHandle, lIndex, gdGetNum(m.hSignalIndex, lExit)
                End If
#End If
            Next lIndex2
        Else
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                lIndex = gdGetNum(m.hSortIndex, lIndex2)
                
                gdSetNum m.alNum.ArrayHandle, lIndex, lIndex
                If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                    gdSetNum m.adProfit.ArrayHandle, lIndex, 0#
                Else
                    gdSetNum m.adProfit.ArrayHandle, lIndex, gdGetNum(m.hProfit, lIndex)
                End If
            Next lIndex2
        End If
        
        ' Set the index on the table based on the filter
        m.aIndex.SetArrayHandle gdCreateTableIndex(m.RptForm.Trades.DtlTableHandle, entd_Show), True
        gdSortTableIndex m.RptForm.Trades.DtlTableHandle, m.aIndex.ArrayHandle, entd_SortKey, eGdSort_Default
        
        'Report options
        m.strPriceFormat = m.RptForm.OptionValue("PriceFormat")
        m.strShowTime = m.RptForm.OptionValue("ShowTime")
        m.strShowCents = m.RptForm.OptionValue("ShowCents")
        
        InitGrid
        
        With m.vsGrid
            .ColData(GDCol(eGDCol_Number)) = m.alNum.ArrayHandle
            .ColData(GDCol(eGDCol_Date)) = m.hTradeDate
            .ColData(GDCol(eGDCol_EntryType)) = m.hSignalType
            .ColData(GDCol(eGDCol_Position)) = m.hPosition
            .ColData(GDCol(eGDCol_Price)) = m.hPrice
            .ColData(GDCol(eGDCol_SystemNumber)) = m.hSystemNumber
            .ColData(GDCol(eGDCol_Pointer)) = m.hEntryExitPtr
            .ColData(GDCol(eGDCol_Symbol)) = m.hSymbolIndex
            .ColData(GDCol(eGDCol_EntryDate)) = m.adEntryDate.ArrayHandle ' m.hTradeDate
            .ColData(GDCol(eGDCol_EntryPrice)) = m.adEntryPrice.ArrayHandle 'm.hPrice
            .ColData(GDCol(eGDCol_ExitDate)) = m.hTradeDate 'm.adExitDate.ArrayHandle
            .ColData(GDCol(eGDCol_ExitPrice)) = m.hPrice 'm.adExitPrice.ArrayHandle
            .ColData(GDCol(eGDCol_Profit)) = m.adProfit.ArrayHandle
            If m.strView = "Trade" Then
                .ColData(GDCol(eGDCol_NetProfit)) = m.adNetProfit.ArrayHandle
            Else
                .ColData(GDCol(eGDCol_NetProfit)) = m.hAccountBalance
            End If
            .ColData(GDCol(eGDCol_NumShares)) = m.hNumShares
            .ColData(GDCol(eGDCol_Units)) = m.hUnits
            .ColData(GDCol(eGDCol_MaxProfit)) = m.hMaxProfit
            .ColData(GDCol(eGDCol_MaxLoss)) = m.hMaxLoss
            .ColData(GDCol(eGDCol_TotalProfit)) = m.hTotalProfit
            .ColData(GDCol(eGDCol_EntrySignal)) = m.alEntrySignal.ArrayHandle ' m.hSignalIndex ' m.hSignal
            .ColData(GDCol(eGDCol_ExitSignal)) = m.hSignalIndex ' m.alExitSignal.ArrayHandle
            .ColData(GDCol(eGDCol_Skip)) = m.hSkip
            .ColData(GDCol(eGDCol_TradeID)) = m.hTradeNumber
            .ColData(GDCol(eGDCol_Link)) = m.hLink
        End With
        
        BuildChart

#If 0 Then ' TLB 8/14/2012: this sorting no longer needed -- it's much better to do sorting upstream in cTrades.cls
        If m.strView = "Trade" Then
            'SortOnCol GDCol(eGDCol_EntryDate), 1
            ' TLB 2/23/2011: sort by symbol (secondary), then exit date (primary) -- so
            ' the sequence will match when running merged reports for a strategy basket
            SortOnCol GDCol(eGDCol_Symbol), 1
            SortOnCol GDCol(eGDCol_ExitDate), 1
        Else
            SortOnCol GDCol(eGDCol_Number), 1
        End If
#End If
        
        'ResizeColWidths
        SetUpColumns
        SetBackColor
        
        m.vsGrid.Redraw = flexRDBuffered
        If m.vsGrid.Rows > m.vsGrid.FixedRows Then
            m.vsGrid.ShowCell m.vsGrid.Rows - 1, 0
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptTrades2.Run", eGDRaiseError_Raise, m.RptForm.AppPath

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
    Dim bSysNav As Boolean              ' Are we called from System Navigator?
    Dim strSymbol$
    
    With m.vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .FlexDataSource = Nothing
        ClearGrid m.vsGrid
        
        '.Editable = True
        .BackColorAlternate = g.nAltGridRowColor '&HC0FFFF
        .ExplorerBar = flexExSortShowAndMove
        .GridLines = flexGridNone
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .HighLight = flexHighlightAlways
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        .FrozenCols = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AllowUserResizing = flexResizeColumns
        .ScrollTips = False
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True

        For lIndex = 0 To .Cols - 1
            .ColHidden(lIndex) = False
        Next lIndex
        
        m.lSubHeadRow = 0
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Number)) = "Num"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Skip)) = "Skip"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TradeID)) = "TradeID"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Link)) = "Link"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Date)) = "Date"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntryType)) = "Type"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Position)) = "Pos"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Price)) = "Price"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SystemNumber)) = "Strategy"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Pointer)) = "Ptr"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntryDate)) = "Entry"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntryPrice)) = "E Price"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ExitDate)) = "Exit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ExitPrice)) = "X Price"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NumShares)) = "Shares"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Units)) = "Units"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_Profit)) = "Profit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_TotalProfit)) = "Total Profit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_NetProfit)) = "Net Profit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_MaxProfit)) = "Max Profit"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_MaxLoss)) = "Max Loss"
        If m.strView = "Trade" Then
            .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntrySignal)) = "Entry Signal"
        Else
            .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_EntrySignal)) = "Signal"
        End If
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_ExitSignal)) = "Exit Signal"
        .TextMatrix(m.lSubHeadRow, GDCol(eGDCol_SkipEq)) = "Skip Eq"
        
        bSysNav = (m.RptForm.AppName = "System Navigator")
        If m.strView = "Trade" Then
            .ColHidden(GDCol(eGDCol_Number)) = True
            .ColHidden(GDCol(eGDCol_Date)) = True
            .ColHidden(GDCol(eGDCol_EntryType)) = True
            .ColHidden(GDCol(eGDCol_Price)) = True
            .ColHidden(GDCol(eGDCol_SystemNumber)) = True
            .ColHidden(GDCol(eGDCol_Pointer)) = True
            .ColHidden(GDCol(eGDCol_Symbol)) = (m.RptForm.Trades.HeaderNumRecords = 2)
            .ColHidden(GDCol(eGDCol_Units)) = True
        Else
            .ColHidden(GDCol(eGDCol_SystemNumber)) = bSysNav
            .ColHidden(GDCol(eGDCol_Pointer)) = bSysNav And Not m.RptForm.Pyramiding
            .ColHidden(GDCol(eGDCol_Symbol)) = bSysNav And (m.RptForm.Trades.HeaderNumRecords = 1)
            .ColHidden(GDCol(eGDCol_EntryDate)) = True
            .ColHidden(GDCol(eGDCol_EntryPrice)) = True
            .ColHidden(GDCol(eGDCol_ExitDate)) = True
            .ColHidden(GDCol(eGDCol_ExitPrice)) = True
            .ColHidden(GDCol(eGDCol_Units)) = bSysNav And Not m.RptForm.Pyramiding
            .ColHidden(GDCol(eGDCol_ExitSignal)) = True
        End If
        
        ' TLB 3/14/2011: display Shares column if trading a stock in TradeNav
        .ColHidden(GDCol(eGDCol_NumShares)) = True
        If bSysNav Then
            For lIndex = 0 To gdGetSize(m.hSymbolHandle) - 1
                strSymbol = Trim(gdGetStr(m.hSymbolHandle, lIndex))
                If Len(strSymbol) > 0 And Left(strSymbol, 1) <> "$" And InStr(strSymbol, "-") = 0 Then
                    .ColHidden(GDCol(eGDCol_NumShares)) = False
                    Exit For
                End If
            Next
        End If
        
        .ColHidden(GDCol(eGDCol_Skip)) = True
        .ColHidden(GDCol(eGDCol_TradeID)) = True
        .ColHidden(GDCol(eGDCol_Link)) = True
        .ColHidden(GDCol(eGDCol_SkipEq)) = True
                
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .FlexDataSource = Me
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptTrades2.InitGrid", eGDRaiseError_Raise, m.RptForm.AppPath

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
        If .Rows > 1 Then '.FixedRows Then
            For lIndex = .TopRow To .BottomRow
                SetColor lIndex, GDCol(eGDCol_Profit)
                SetColor lIndex, GDCol(eGDCol_NetProfit)
                
                If .Cell(flexcpValue, lIndex, GDCol(eGDCol_SkipEq)) = 1 Then
                    .Cell(flexcpForeColor, lIndex, 0, lIndex, .Cols - 1) = RGB(125, 125, 125)
                ElseIf .Cell(flexcpValue, lIndex, GDCol(eGDCol_Skip)) = 1 Then
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
            m.lBottomRow = .BottomRow
        End If
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.SetBackColor", eGDRaiseError_Raise, m.RptForm.AppPath
    
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
    
    lTradeNbr = gdGetNum(m.hTradeNumber, m.alGraphToTrades(Int(X)))
    If m.strView = "Trade" Then lTradeNbr = gdGetNum(m.hEntryExitPtr, lTradeNbr)
    
    If gdBinarySearchAsIndex(m.aIndex.ArrayHandle, m.hTradeNumber, lTradeNbr, lPos, eGdSort_Default, 0, 99999999) Then
        RowFromPoint = lPos + 1
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptTrades2.RowFromPoint", eGDRaiseError_Raise, m.RptForm.AppPath
    
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
    Dim dProfit As Double               ' Profit for this transaction
    Dim dDate As Double                 ' Date of this transaction
    Dim dEquity As Double               ' Equity at this point

    lTradeNum = gdGetNum(m.hTradeNumber, m.alGraphToTrades(X - 1))
    dProfit = gdGetNum(m.hTotalProfit, m.alGraphToTrades(X - 1))
    dDate = gdGetNum(m.hTradeDate, m.alGraphToTrades(X - 1))
    dEquity = m.Chart.YData(0, X - 1)

    With lbl
        '.Caption = "Transaction: " & lTradeNum & _
                ", Date: " & DateFormat(dDate) & _
                ", Equity: " & Format(dEquity, "$#,##0.00") & vbCrLf & _
                "(Double-Click on the Graph to go to the Transaction)"
        .Caption = "Transaction: " & lTradeNum & _
                ", Date: " & DateFormat(dDate) & _
                ", Profit: " & Format(dEquity, "$#,##0.00") & vbCrLf & _
                "(Double-Click on the Graph to go to the Transaction)"
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.DisplayValues", eGDRaiseError_Raise, m.RptForm.AppPath
    
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

    Dim lCurRecord As Long              ' Get the current record to set selection back to
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lSortField As Long              ' Field being sorted
    Dim lIndex As Long                  ' Index into a for loop
    Dim eSortFlags As eGdSort_Flags     ' Flags to pass to the sort
    
    Static lPrevCol As Long             ' Previous column sorted
    Static bPrevDescending As Boolean   ' Previous sort order

    ' Set Back to the previous sort
    If lCol = -1 Then
        lCol = lPrevCol
        If bPrevDescending Then
            lDirection = -1
        Else
            lDirection = 1
        End If
        
    ' Save the current sort information as the previous
    Else
        lPrevCol = m.iSortedCol
        bPrevDescending = m.bSortedDescending
    End If

    ' Exit if invalid column
    If lCol < 0 Or lCol >= m.vsGrid.Cols Then Exit Sub

    ' Save some stuff to restore later
    lCurRecord = m.aIndex(m.vsGrid.Row - 1)
    lRedraw = m.vsGrid.Redraw
    m.vsGrid.Redraw = flexRDNone

    ' Determine if descending
    If lCol = m.iSortedCol And lDirection = 0 Then
        m.bSortedDescending = Not m.bSortedDescending
    ElseIf lDirection = -1 Then
        m.bSortedDescending = True
    Else
        m.bSortedDescending = False
    End If
    m.iSortedCol = lCol

    ' Create filtered index from filter field #
    If m.bRecalcIndex Then
        m.bRecalcIndex = False
        m.aIndex.SetArrayHandle gdCreateTableIndex(m.RptForm.Trades.DtlTableHandle, entd_Show), True
    End If
    
    ' Sort index
    lSortField = m.vsGrid.ColData(m.iSortedCol)
    eSortFlags = eGdSort_Stable
    If m.bSortedDescending Then eSortFlags = eSortFlags Or eGdSort_Descending
    gdSortAsIndex m.aIndex.ArrayHandle, lSortField, 0, eSortFlags, 0, m.aIndex.Size - 1

    ' Show sortation arrow in column heading
    With m.vsGrid
        .FillStyle = flexFillSingle
        For lIndex = 0 To .Cols - 1
            .Select 0, lIndex
            If lIndex = m.iSortedCol Then
                If m.bSortedDescending Then
                    .CellPicture = m.RptForm.Images.Item("kSortedDownRight").Picture
                Else
                    .CellPicture = m.RptForm.Images.Item("kSortedUpRight").Picture
                End If
                .CellPictureAlignment = flexPicAlignRightTop
                .PicturesOver = True
            Else
                .CellPicture = Nothing
            End If
        Next
    End With
    
    ' Always show the first row after sorting...
    If m.vsGrid.Rows > 1 Then m.vsGrid.Row = 1
    
    ' Restore things
    With m.vsGrid
        .Select .Row, 0, .Row, .Cols - 1
        'ResizeColWidths
        .ShowCell .Row, 0
        SetBackColor
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.SortOnCol", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Public Sub DblClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Row > 0 Then
        m.RptForm.ShowDateInChart gdGetNum(m.hTradeDate, gdGetNum(m.aIndex.ArrayHandle, Row - 1)), gdGetNum(m.hPrice, gdGetNum(m.aIndex.ArrayHandle, Row - 1)), gdGetNum(m.hHeaderIndex, gdGetNum(m.aIndex.ArrayHandle, Row - 1))
    End If
    'If m.strView = "Trade" Then
    '    m.RptForm.ShowDateInChart gdGetNum(m.hTradeDate, gdGetNum(m.aIndex.ArrayHandle, Row - 1))
    'Else
    '    m.RptForm.ShowDateInChart gdGetNum(m.hTradeDate, Row)
    'End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptTrades2.DblClick", eGDRaiseError_Raise, m.RptForm.AppPath
    
End Sub

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
    If UCase(m.strView) = "TRADE" Then
        astrFields.SplitFields m.strDefaultFieldsTrade, ","
    Else
        astrFields.SplitFields m.strDefaultFields, ","
    End If
    For lIndex = 0 To astrFields.Size - 1
        If Parse(astrFields(lIndex), ";", 2) = "0" Then
            astrDefault.Add Parse(astrFields(lIndex), ";", 1)
        End If
    Next lIndex
    
    ' Compile lists of used and available fields...
    If UCase(m.strView) = "TRADE" Then
        astrFields.SplitFields m.strFieldsTrade, ","
    Else
        astrFields.SplitFields m.strFields, ","
    End If
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
        
        If UCase(m.strView) = "TRADE" Then
            m.strFieldsTrade = astrNew.JoinFields(",")
            SetIniFileProperty "FieldsTrade", m.strFieldsTrade, Str(m.lAppID), strIniFile
        Else
            m.strFields = astrNew.JoinFields(",")
            SetIniFileProperty "Fields", m.strFields, Str(m.lAppID), strIniFile
        End If
        SetUpColumns
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
    RaiseError "cRptTrades2.ChangeFields", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpColumns
'' Description: Set up the column order/width/visibility according to spec
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpColumns()
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
            
            For lCol = 0 To .Cols - 1
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    If strHidden <> "1" Then
                        .ColPosition(lCol) = GDCol(eGDCol_Units) + lColPos
                        lTotWidth = lTotWidth + lColWidth
                        If strHidden = "-1" Then
                            .ColHidden(GDCol(eGDCol_Units) + lColPos) = True
                        Else
                            .ColHidden(GDCol(eGDCol_Units) + lColPos) = False
                        End If
                        .ColWidth(GDCol(eGDCol_Units) + lColPos) = lColWidth
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
    RaiseError "cRptTrades2.SetUpColumns", eGDRaiseError_Raise, g.strAppPath
    
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
    
    astrFields.Create eGDARRAY_Strings
    For lIndex = 0 To m.vsGrid.Cols - 1
        If lIndex <= GDCol(eGDCol_Units) Then
            If m.vsGrid.ColHidden(lIndex) = False Then
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";1;" & Str(m.vsGrid.ColWidth(lIndex))
            End If
        Else
            If m.vsGrid.ColHidden(lIndex) = True Then
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";-1;" & Str(m.vsGrid.ColWidth(lIndex))
            Else
                astrFields.Add m.vsGrid.TextMatrix(0, lIndex) & ";0;" & Str(m.vsGrid.ColWidth(lIndex))
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.SaveCols", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub BuildChart()
On Error GoTo ErrSection:

    Dim lTrade As Long                  ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lItemsToChart As Long           ' Number of items to chart
    Dim YData() As Single               ' Array to assign to the Y-Axis
    Dim XData() As Single               ' Array to assign to the X-Axis
    Dim dProfit As Double
    Dim dSumOfSquares As Double
    Dim dSum As Double
    Dim dStdDev As Double
    Dim dMaxProfit As Double
    Dim dMinProfit As Double
    Dim dValue As Double
    Dim lAnnotIndex As Long
    
    ' Calculate Items to chart (number of rows in grid)
    For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
        lTrade = gdGetNum(m.hSortIndex, lIndex2)
        If gdGetNum(m.hSignalType, lTrade) = 1 Then
            If (gdGetNum(m.hSkip, lTrade) = 0) And (gdGetNum(m.hSkipReport, lTrade) = 0) And (gdGetNum(m.hSkipEqFilter, lTrade) = 0) Then
                lIndex = lIndex + 1
            End If
        End If
    Next lIndex2
    lItemsToChart = lIndex
    
    ' If we decide to later on, limit the number of trades to 32765 so that the coloring
    ' works (the .PointColors array takes an Integer as the index)
    If 0 Then
        If lItemsToChart >= 32765 Then
            For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
                lTrade = gdGetNum(m.hSortIndex, lIndex2)
                If gdGetNum(m.hSignalType, lTrade) = 1 Then
                    gdSetNum m.hSkip, lTrade, 1
                    gdSetNum m.hSkip, gdGetNum(m.hEntryExitPtr, lTrade), 1
                    lItemsToChart = lItemsToChart - 1
                    If lItemsToChart < 32765 Then Exit For
                End If
            Next lIndex2
        End If
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
        
        ReDim YData(lItemsToChart) As Single
        ReDim XData(lItemsToChart) As Single

        .MainTitle = "Profit Per Trade"
        .SubTitle = ""
        .YAxisLabel = "" ' = "Equity"
        .XAxisLabel = "" ' = "Trades"
               
        SetChartSettings m.Chart, GPM_BAR
        .DataShadows = PEDS_NONE
        .SubsetColors(0) = RGB(128, 128, 128) 'vbBlack 'vbBlue
        .YAxisOnRight = True
        .ShowAnnotations = True
        .AnnotationsInFront = True
        
        .PointsToGraph = ((.Width / Screen.TwipsPerPixelX) * 0.75) / 3
        If .PointsToGraph > .Points Then
            .PointsToGraph = .Points
        Else
            .HorzScrollPos = .Points - .PointsToGraph + 1
        End If
        
        .GraphAnnotationType(-1) = 0
        .GraphAnnotationText(-1) = 0
        .HorzLineAnnotationText(-1) = 0
        .VertLineAnnotationText(-1) = 0
        .LegendAnnotationText(-1) = 0
        .YAxisAnnotationText(-1) = 0
        .XAxisAnnotationText(-1) = 0
        
        .DateTimeMode = PEDTM_NONE
        .AltFrequencies(0) = 10
        .AltFrequencies(1) = 50
        .AltFrequencies(2) = 100
        .AltFrequencies(3) = 500
        .AltFrequencies(4) = 1000
        .FirstPtLabelOffset = 1

        lIndex = 0
        dSum = 0#
        dSumOfSquares = 0#
        dMaxProfit = -999999#
        dMinProfit = 999999#
        For lIndex2 = 1 To m.RptForm.Trades.NumRecords - 1
            lTrade = gdGetNum(m.hSortIndex, lIndex2)
            
            If (gdGetNum(m.hSkip, lTrade) = 0) And (gdGetNum(m.hSkipReport, lTrade) = 0) And (gdGetNum(m.hSkipEqFilter, lTrade) = 0) Then
                If gdGetNum(m.hSignalType, lTrade) = 1 Then
                    lIndex = lIndex + 1
                    dProfit = gdGetNum(m.hProfit, lTrade)
                    
                    YData(lIndex - 1) = dProfit
                    XData(lIndex - 1) = lIndex - 1
                    .PointLabels(lIndex - 1) = lIndex - 1
                    m.alGraphToTrades(lIndex - 1) = gdGetNum(m.hTradeNumber, lTrade)
                    
                    ' Only color the trades if there is less than 32765 because the
                    ' .PointColors function takes an Integer as the index...
                    If lItemsToChart < 32765 Then
                        If gdGetNum(m.hPosition, lTrade) = 1 Then
                            .PointColors(0, lIndex - 1) = 16744576  'vbBlue
                        Else
                            .PointColors(0, lIndex - 1) = vbRed
                        End If
                    End If
                    
                    dSum = dSum + dProfit
                    dSumOfSquares = dSumOfSquares + (dProfit ^ 2)
                    
                    If dProfit > dMaxProfit Then dMaxProfit = dProfit
                    If dProfit < dMinProfit Then dMinProfit = dProfit
                End If
            End If
        Next lIndex2
        
        ' Calculate the standard deviation...
        dStdDev = Sqr(((lItemsToChart * dSumOfSquares) - (dSum ^ 2)) / (lItemsToChart ^ 2))
        
        ' Assign the XData and YData points to the chart all at once
        PEvset .hObject, PEP_faYDATA, YData(0), .Points
        PEvset .hObject, PEP_faXDATA, XData(0), .Points

        .GraphAnnotationTextSize = 80
        .LabelBold = True
        
        ' Draw a dotted zero line...
        .HorzLineAnnotationType(0) = PELT_MEDIUMSOLID
        .HorzLineAnnotation(0) = 0
        .HorzLineAnnotationColor(0) = vbBlack
        .HorzLineAnnotationText(0) = ""
        
        ' Draw a dotted line at the average trade value...
        .HorzLineAnnotationType(1) = PELT_DASH
        .HorzLineAnnotation(1) = m.Stats.AvgTrade
        .HorzLineAnnotationColor(1) = vbBlack
        .HorzLineAnnotationText(1) = "" ' "|lAvg Trade = " & Format(m.Stats.AvgTrade, "$#,##0.00")
        
        ' Draw a dotted line at the positive standard deviations...
        lAnnotIndex = 1
        dValue = m.Stats.AvgTrade
        For lIndex = 1 To 10
            dValue = dValue + dStdDev
            
            If dValue <= dMaxProfit Then
                lAnnotIndex = lAnnotIndex + 1
                .HorzLineAnnotationType(lAnnotIndex) = PELT_DOT
                .HorzLineAnnotation(lAnnotIndex) = dValue
                .HorzLineAnnotationColor(lAnnotIndex) = vbBlack
                .HorzLineAnnotationText(lAnnotIndex) = "|l+" & Str(lIndex) & " Std Dev" ' = " & Format(dValue, "$#,##0.00")
            Else
                Exit For
            End If
        Next lIndex
        
        ' Draw a dotted line at the negative standard deviations...
        dValue = m.Stats.AvgTrade
        For lIndex = -1 To -10 Step -1
            dValue = dValue - dStdDev
            
            If dValue >= dMinProfit Then
                lAnnotIndex = lAnnotIndex + 1
                .HorzLineAnnotationType(lAnnotIndex) = PELT_DOT
                .HorzLineAnnotation(lAnnotIndex) = dValue
                .HorzLineAnnotationColor(lAnnotIndex) = vbBlack
                .HorzLineAnnotationText(lAnnotIndex) = "|l" & Str(lIndex) & " Std Dev" ' = " & Format(dValue, "$#,##0.00")
            Else
                Exit For
            End If
        Next lIndex
        
        .SubsetsToLegend(0) = -1
        
        If lItemsToChart < 32765 Then
            .LegendAnnotationType(0) = PEGAT_SQUARESOLID
            .LegendAnnotationColor(0) = 16744576
            .LegendAnnotationText(0) = "Long Trades"
            
            .LegendAnnotationType(1) = PEGAT_SQUARESOLID
            .LegendAnnotationColor(1) = vbRed
            .LegendAnnotationText(1) = "Short Trades"
            
            .LegendAnnotationType(2) = PEGAT_DASHLINE
            .LegendAnnotationColor(2) = vbBlack
            .LegendAnnotationText(2) = "Avg Trade = " & Format(m.Stats.AvgTrade, "$#,##0.00")
        Else
            .LegendAnnotationType(0) = PEGAT_DASHLINE
            .LegendAnnotationColor(0) = vbBlack
            .LegendAnnotationText(0) = "Avg Trade = " & Format(m.Stats.AvgTrade, "$#,##0.00")
        End If
        
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.BuildChart", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub MoreBars()
On Error GoTo ErrSection:

    With m.Chart
        If .PointsToGraph + 50 <= .Points Then
            .PointsToGraph = .PointsToGraph + 50
            .PEactions = 0
        Else
            .PointsToGraph = .Points
            .PEactions = 0
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.MoreBars", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub LessBars()
On Error GoTo ErrSection:

    With m.Chart
        If .PointsToGraph - 50 >= 50 Then
            .PointsToGraph = .PointsToGraph - 50
            .PEactions = 0
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptTrades2.LessBars", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportRowCSV
'' Description: Export a row in the grid in comma separated format
'' Inputs:      Row to Export
'' Returns:     Comma Separated String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExportRowCSV(ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim astrLine As New cGdArray        ' Array of data from the line
    Dim lCol As Long                    ' Index into a for loop
    Dim lIndex As Long                  ' Index into the table
    Dim lExit As Long                   ' Exit pointer for an entry
    
    If m.strView = "Trade" Then
        ' TLB 8/14/2012: for Trades, the Exit is now getting passed in, so find the Entry
        lExit = m.aIndex(lRow - 1)
        lIndex = gdGetNum(m.hEntryExitPtr, lExit)
    Else
        lIndex = m.aIndex(lRow - 1)
        lExit = gdGetNum(m.hEntryExitPtr, lIndex)
    End If
    
    astrLine.Create eGDARRAY_Strings
    With m.vsGrid
        For lCol = 0 To .Cols - 1
            If .ColHidden(lCol) = False Then
                Select Case UCase(.TextMatrix(0, lCol))
                    Case "NUM"
                        astrLine.Add CStr(lIndex)
                    Case "SKIP"
                        astrLine.Add CStr(gdGetNum(m.hSkip, lIndex))
                    Case "TRADEID"
                        astrLine.Add CStr(gdGetNum(m.hTradeNumber, lIndex))
                    Case "LINK"
                        astrLine.Add CStr(gdGetNum(m.hLink, lIndex))
                    Case "DATE"
                        If gdGetNum(m.hSignalsTotal, lIndex) > 1 Or m.strShowTime = "Yes" Then
                            astrLine.Add DateFormat(gdGetNum(m.hTradeDate, lIndex), MM_DD_YYYY, HH_MM)
                        Else
                            astrLine.Add DateFormat(gdGetNum(m.hTradeDate, lIndex))
                        End If
                    Case "TYPE"
                        If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                            astrLine.Add "E"
                        Else
                            astrLine.Add "X"
                        End If
                    Case "POS"
                        If gdGetNum(m.hPosition, lIndex) = 1 Then
                            astrLine.Add "L"
                        Else
                            astrLine.Add "S"
                        End If
                    Case "PRICE"
                        astrLine.Add CStr(gdGetNum(m.hPrice, lIndex))
                    Case "STRATEGY"
                        astrLine.Add CStr(m.RptForm.SystemArrayInd(gdGetNum(m.hSystemNumber, lIndex)))
                    Case "PTR"
                        astrLine.Add CStr(gdGetNum(m.hEntryExitPtr, lIndex))
                    Case "SYMBOL"
                        astrLine.Add gdGetStr(m.hSymbolHandle, gdGetNum(m.hSymbolIndex, lIndex))
                    Case "ENTRY"
                        If gdGetNum(m.hSignalsTotal, lIndex) > 1 Or m.strShowTime = "Yes" Then
                            astrLine.Add DateFormat(gdGetNum(m.hTradeDate, lIndex), MM_DD_YYYY, HH_MM)
                        Else
                            astrLine.Add DateFormat(gdGetNum(m.hTradeDate, lIndex))
                        End If
                    Case "E PRICE"
                        astrLine.Add CStr(gdGetNum(m.hPrice, lIndex))
                    Case "EXIT"
                        If gdGetNum(m.hSignalsTotal, lExit) > 1 Or m.strShowTime = "Yes" Then
                            astrLine.Add DateFormat(gdGetNum(m.hTradeDate, lExit), MM_DD_YYYY, HH_MM)
                        Else
                            astrLine.Add DateFormat(gdGetNum(m.hTradeDate, lExit))
                        End If
                    Case "X PRICE"
                        astrLine.Add CStr(gdGetNum(m.hPrice, lExit))
                    Case "SHARES"
                        astrLine.Add CStr(gdGetNum(m.hNumShares, lIndex))
                    Case "UNITS"
                        astrLine.Add CStr(gdGetNum(m.hUnits, lIndex))
                    Case "PROFIT"
                        If m.strView = "Trade" Then
                            astrLine.Add CStr(gdGetNum(m.hProfit, lExit))
                        Else
                            If gdGetNum(m.hSignalType, lIndex) = gEntrySignal Then
                                astrLine.Add "0"
                            Else
                                astrLine.Add CStr(gdGetNum(m.hProfit, lIndex))
                            End If
                        End If
                    Case "TOTAL PROFIT"
                        astrLine.Add CStr(gdGetNum(m.hTotalProfit, lIndex))
                    Case "NET PROFIT"
                        If m.strView = "Trade" Then
                            astrLine.Add CStr(gdGetNum(m.hAccountBalance, lExit))
                        Else
                            astrLine.Add CStr(gdGetNum(m.hAccountBalance, lIndex))
                        End If
                    Case "MAX PROFIT"
                        astrLine.Add CStr(gdGetNum(m.hMaxProfit, lIndex))
                    Case "MAX LOSS"
                        astrLine.Add CStr(gdGetNum(m.hMaxLoss, lIndex))
                    Case "ENTRY SIGNAL"
                        astrLine.Add gdGetStr(m.hRuleIndex, gdGetNum(m.hSignalIndex, lIndex))
                    Case "SIGNAL"
                        astrLine.Add gdGetStr(m.hRuleIndex, gdGetNum(m.hSignalIndex, lIndex))
                    Case "EXIT SIGNAL"
                        astrLine.Add gdGetStr(m.hRuleIndex, gdGetNum(m.hSignalIndex, lExit))
                End Select
            End If
        Next lCol
    End With
    
    ExportRowCSV = astrLine.JoinFields(vbTab)

ErrExit:
    Set astrLine = Nothing
    Exit Function
    
ErrSection:
    Set astrLine = Nothing
    RaiseError "cRptTrades2.ExportRowCSV", eGDRaiseError_Raise, g.strAppPath
    
End Function

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
    
    astrFile.Create eGDARRAY_Strings
    astrLine.Create eGDARRAY_Strings
    
    With m.vsGrid
        For lCol = 0 To .Cols - 1
            If .ColHidden(lCol) = False Then astrLine.Add .TextMatrix(0, lCol)
        Next lCol
        astrFile.Add astrLine.JoinFields(vbTab)
        For lRow = .FixedRows To .Rows - 1
            astrFile.Add ExportRowCSV(lRow)
        Next lRow
    End With
    
    astrFile.ToFile strFileName

ErrExit:
    Set astrFile = Nothing
    Exit Sub
    
ErrSection:
    Set astrFile = Nothing
    RaiseError "cRptTrades2.ExportToCSV", eGDRaiseError_Raise, g.strAppPath
    
End Sub
