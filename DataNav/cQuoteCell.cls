VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuoteCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cQuoteCell.cls
'' Description: Controller class for a single quote cell in a quote cell board
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kBlinkTime = 500

'JM 02-17-2011: These two enumerated types are redundant, since numerically they are same.
'   For maintenace purposes, however, keeping them separate clarifies their purposes.
'   The color index enum has a direct one-to-one correspondence to array passed to grapheng.dll,
'   therefore the numerics must be kept in sync with the grapheng.dll. The update cell enum
'   is strictly for readability and has no counterpart in the grapheng.dll

Private Enum eGDUpdateCell
    eGDUpdateCell_Symbol = 0
    eGDUpdateCell_Open
    eGDUpdateCell_High
    eGDUpdateCell_Low
    eGDUpdateCell_Close
    eGDUpdateCell_Delta
    eGDUpdateCell_Bid
    eGDUpdateCell_Ask
    eGDUpdateCell_Time
End Enum

Private Enum eGDColorIndex
    eGDColorIndex_Symbol = 0
    eGDColorIndex_Open
    eGDColorIndex_High
    eGDColorIndex_Low
    eGDColorIndex_Close
    eGDColorIndex_Delta
    eGDColorIndex_Bid
    eGDColorIndex_Ask
    eGDColorIndex_Alert          '-1=none, 0=has inactive alert(s), >=1 indicate number of active alerts
    eGDColorIndex_Update
End Enum

Private Type mPrivate
    hQuoteObj As Long
    MyQbBoard As cQuoteCellBoard
    
    Bars As cGdBars
    lDelay As Long
    geCell As quote_cell
    tbCellTickCount As cGdTable
    Settles As cGdBars
    
    lUpColor As Long
    lDownColor As Long
    lUnchColor As Long
    lUpdateColor As Long
    lDeltaColor As Long
    
    lUseUpdateColor As Long     'if not zero then use delta color for OHLC (Jo Dinapoli--all red/green request)
    lColorSymbol As Long        'if zero then symbol is always black
    
    nActiveAlert As Long        'count of active alert with same symbol as this cell
    nInactiveAlert As Long      'count of inactive alert with same symbol as this cell
    
    lPrevRtCount As Long        'count of how many times ChangePrice was called with RT active - 5710
    
    dPrevBid As Double          'this is here instead of in quote_cell structure to keep grapheng.dll backwards compatible
    dPrevAsk As Double
End Type
Dim m As mPrivate

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructor & Destructor
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Bars = New cGdBars
    m.Bars.Size = 0

    m.geCell.CellID = -1&
    m.geCell.SymbolID = -1&
    
    m.geCell.gshSymbol = gdCreateArray(eGDARRAY_gdString)
    m.geCell.gshOpen = gdCreateArray(eGDARRAY_gdString)
    m.geCell.gshHigh = gdCreateArray(eGDARRAY_gdString)
    m.geCell.gshLow = gdCreateArray(eGDARRAY_gdString)
    
    m.geCell.gshBid = gdCreateArray(eGDARRAY_gdString)
'    m.geCell.gshAsk = gdCreateArray(eGDARRAY_gdString)
    
    m.geCell.gshLast = gdCreateArray(eGDARRAY_gdString)
    m.geCell.gshNetChange = gdCreateArray(eGDARRAY_gdString)
    m.geCell.gshTickTime = gdCreateArray(eGDARRAY_gdString)
    m.geCell.glhTextColors = gdCreateArray(eGDARRAY_Longs)
    
    Set m.tbCellTickCount = New cGdTable
    
    m.tbCellTickCount.CreateField eGDARRAY_Doubles, 0, "UpdateTime", 0
    m.tbCellTickCount.CreateField eGDARRAY_Doubles, 1, "BlinkTime", 0
    
    'number of records = items that has updated color changes and need ticktime saved to clear updated color
    'there is a record in table for the delta value, but it is not used as delta value is only up, down or unchanged color
    m.tbCellTickCount.NumRecords = eGDUpdateCell_Ask
        
    m.lDelay = -999999
    
    m.nActiveAlert = 0
    m.nInactiveAlert = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCell.Initialize", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim i&
    Dim Alert As cAlert

    If Not g.bUnloading Then
        For i = 1 To g.Alerts.Count
            Set Alert = g.Alerts(i)
            If Not Alert Is Nothing Then
                If Alert.Symbol = m.Bars.Prop(eBARS_Symbol) Then Alert.RemoveQbCell Me
            End If
        Next
    End If

    Set m.Settles = Nothing
    Set m.Bars = Nothing
    gdDestroyArray m.geCell.gshSymbol
    gdDestroyArray m.geCell.gshOpen
    gdDestroyArray m.geCell.gshHigh
    gdDestroyArray m.geCell.gshLow
    
    gdDestroyArray m.geCell.gshBid
'    gdDestroyArray m.geCell.gshAsk
    
    gdDestroyArray m.geCell.gshLast
    gdDestroyArray m.geCell.gshNetChange
    gdDestroyArray m.geCell.gshTickTime
    gdDestroyArray m.geCell.glhTextColors
    
    Set m.tbCellTickCount = Nothing
    
    geUpdateCell m.hQuoteObj, m.geCell
    
    Set m.MyQbBoard = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCell.Terminate", eGDRaiseError_Raise

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get hQuoteObj() As Long
    hQuoteObj = m.hQuoteObj
End Property
Public Property Let hQuoteObj(ByVal hObj As Long)
    m.hQuoteObj = hObj
End Property

Public Property Get ID() As Long
    ID = m.geCell.CellID
End Property
Public Property Let ID(ByVal lID As Long)
    m.geCell.CellID = lID
End Property

Public Property Get Row() As Long
    Row = m.geCell.Row
End Property
Public Property Let Row(ByVal lRow As Long)
    m.geCell.Row = lRow
End Property

Public Property Get Col() As Long
    Col = m.geCell.Col
End Property
Public Property Let Col(ByVal lCol As Long)
    m.geCell.Col = lCol
End Property

Public Property Get HasData() As Boolean
    HasData = (Len(Symbol) > 0)
End Property

Public Property Get IsHighlighted() As Long
    IsHighlighted = m.geCell.IsHighlighted
End Property

Public Property Get IsBold() As Long
    IsBold = m.geCell.IsBolded
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.geCell.SymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.geCell.SymbolID = lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = gdGetStr(m.geCell.gshSymbol)
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    gdSetStr m.geCell.gshSymbol, 0, strSymbol
End Property

Public Property Get Bars() As cGdBars
    Set Bars = m.Bars
End Property
Public Property Let Bars(NewBars As cGdBars)
    Set m.Bars = NewBars
End Property

Public Property Get UpColor() As Long
    UpColor = m.lUpColor
End Property
Public Property Let UpColor(ByVal lUpColor As Long)
    m.lUpColor = lUpColor
End Property

Public Property Get DownColor() As Long
    DownColor = m.lDownColor
End Property
Public Property Let DownColor(ByVal lDownColor As Long)
    m.lDownColor = lDownColor
End Property

Public Property Get UnchColor() As Long
    UnchColor = m.lUnchColor
End Property
Public Property Let UnchColor(ByVal lUnchColor As Long)
    m.lUnchColor = lUnchColor
End Property

Public Property Get UpdateColor() As Long
    UpdateColor = m.lUpdateColor
End Property
Public Property Let UpdateColor(ByVal lUpdateColor As Long)
    m.lUpdateColor = lUpdateColor
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Subs & Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByRef QbBoard As cQuoteCellBoard, ByVal hQuoteObj As Long, ByVal lRow As Long, ByVal lCol As Long, _
                ByVal lUpColor As Long, ByVal lDownColor As Long, ByVal lUnchColor As Long, _
                ByVal lUpdateColor As Long, ByVal lUseUpdateColor As Long, ByVal lColorSymbol As Long)
On Error GoTo ErrSection:

    Dim i&

    m.hQuoteObj = hQuoteObj
    Set m.MyQbBoard = QbBoard
    
    m.geCell.Row = lRow
    m.geCell.Col = lCol
    m.lUpColor = lUpColor
    m.lDownColor = lDownColor
    m.lUnchColor = lUnchColor
    m.lUpdateColor = lUpdateColor
    m.lUseUpdateColor = lUseUpdateColor
    m.lColorSymbol = lColorSymbol
    
    'initialize text colors
    For i = 0 To eGDColorIndex_Ask
        gdSetNum m.geCell.glhTextColors, i, 0
    Next
    gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, -1
    If 0 = lUseUpdateColor Then
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Update, -1
    Else
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Update, lUpdateColor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCell.Init", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateData
'' Description: Update the cell with the information in the bars
'' Inputs:      Bars
'' Returns:     Return code from geUpdateCell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateData(Bars As cGdBars, ByVal bFullHeight As Boolean, _
    ByVal eDisplayStyle As eGDQuoteStyle, ByVal nCompactQB As Long, ByVal bDisplay As Boolean) As Long
On Error GoTo ErrSection:

    Dim strMonth As String              ' Front contract for a continuous
    Dim strSymbol As String             ' Symbol from the Bars
    Dim strPrice As String              ' String representation of the price
    Dim dTime As Double
    Dim strTime As String
    Dim lIndex As Long
    Dim lDeltaColor As Long             'to use when flag to use same color for all text is set

    Set m.Bars = Bars
    strSymbol = m.Bars.Prop(eBARS_Symbol)
    m.geCell.SymbolID = m.Bars.Prop(eBARS_SymbolID)
    gdSetStr m.geCell.gshSymbol, 0, strSymbol
    
    If bDisplay = True Then
        If m.Bars.SecurityType = "F" Then
            If InStr(Symbol, "(") = 0 And m.Bars.Size > 0 Then
                strMonth = RollSymbolForDate(strSymbol, m.Bars(eBARS_DateTime, m.Bars.Size - 1))
                If strMonth <> strSymbol And Len(strMonth) > 0 Then
                    strMonth = MonthName(Val(Right(strMonth, 2)), True)
                    If Len(strMonth) > 0 Then
                        '09-03-2008: box style changed to have OHLC label only in left-most column
                        'per Glen/Pete: remove items inside parentheses from symbol name; okay to truncate Forex with @ in name
                        If nCompactQB = 1 Then
                            gdSetStr m.geCell.gshSymbol, 0, strSymbol   '& " (" & strMonth & ")"
                        Else
                            gdSetStr m.geCell.gshSymbol, 0, strSymbol & " (" & strMonth & ")"
                        End If
                    End If
                End If
            End If
'        ElseIf IsForex(strSymbol) And InStr(strSymbol, "@") <> 0 Then
'            '07-25-2012: don't remember why this is here; leave awhile then remove if all ok
'            gdSetStr m.geCell.gshSymbol, 0, Left(strSymbol, 10)
        End If
        
        If g.RealTime.Active Then
            If m.lPrevRtCount > 100 Then
                '07-22-2010: lPrevRtCount is just a counter to determine whether to use updated colors
                'updated colors is used when lPrevRtCount >3, resetting to 10 just so don't get too big
                m.lPrevRtCount = 10
            Else
                m.lPrevRtCount = m.lPrevRtCount + 1
            End If
            If m.lDelay < 0 Then m.lDelay = g.RealTime.SymbolDelay(strSymbol)
        Else
            m.lPrevRtCount = 0
            m.lDelay = -999999
        End If
        
        If m.Bars.Size > 0 Then
            If eDisplayStyle = eGDQuoteStyle_Forex Then
                ChangePrice eGDUpdateCell_Open, m.Bars(eBARS_Bid, m.Bars.Size - 1), eDisplayStyle
                ChangePrice eGDUpdateCell_Close, m.Bars(eBARS_Ask, m.Bars.Size - 1), eDisplayStyle
            Else
                ChangePrice eGDUpdateCell_Open, m.Bars(eBARS_Open, m.Bars.Size - 1), eDisplayStyle
                ChangePrice eGDUpdateCell_Close, m.Bars(eBARS_Close, m.Bars.Size - 1), eDisplayStyle
                
                If m.MyQbBoard.ShowExtraInfo = eGEQbExtraInfo_BidAsk Or m.MyQbBoard.ShowExtraInfo = eGEQbExtraInfo_All Then
                    ChangePrice eGDUpdateCell_Bid, m.Bars(eBARS_Bid, m.Bars.Size - 1), eDisplayStyle
                    ChangePrice eGDUpdateCell_Ask, m.Bars(eBARS_Ask, m.Bars.Size - 1), eDisplayStyle
                End If
            End If
            
            ChangePrice eGDUpdateCell_High, m.Bars(eBARS_High, m.Bars.Size - 1), eDisplayStyle
            ChangePrice eGDUpdateCell_Low, m.Bars(eBARS_Low, m.Bars.Size - 1), eDisplayStyle

            If bFullHeight Then
                m.geCell.dPrevOpen = m.Bars(eBARS_Open, m.Bars.Size - 1)
                m.geCell.dPrevHigh = m.Bars(eBARS_High, m.Bars.Size - 1)
                m.geCell.dPrevLow = m.Bars(eBARS_Low, m.Bars.Size - 1)
            Else
                m.geCell.dPrevOpen = m.Bars(eBARS_Open, m.Bars.Size - 2)
                m.geCell.dPrevHigh = m.Bars(eBARS_High, m.Bars.Size - 2)
                m.geCell.dPrevLow = m.Bars(eBARS_Low, m.Bars.Size - 2)
            End If
            'm.geCell.dPrevClose = m.Bars(eBARS_Close, m.Bars.Size - 2)
            m.geCell.dPrevClose = GetPrevCloseForQB(m.Bars)
            
            If m.geCell.dPrevClose <> kNullData Then
                ChangePrice eGDUpdateCell_Delta, m.geCell.dLast - m.geCell.dPrevClose, eDisplayStyle
            End If
            
            If m.lUseUpdateColor = 0 Then
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Open, m.lDeltaColor
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_High, m.lDeltaColor
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Low, m.lDeltaColor
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Close, m.lDeltaColor
            End If
        
            ' do this last since will change color of symbol to black when not trading
            ChangePrice eGDUpdateCell_Time, m.Bars.Prop(eBARS_LastTickTime), eDisplayStyle
        End If
        
        If eDisplayStyle = eGDQuoteStyle_Forex Then
            gdSetNum m.geCell.glhTextColors, eGDColorIndex_Symbol, vbWhite
            gdSetNum m.geCell.glhTextColors, eGDColorIndex_High, vbWhite
            gdSetNum m.geCell.glhTextColors, eGDColorIndex_Low, vbWhite
        End If
        
        UpdateData = geUpdateCell(m.hQuoteObj, m.geCell)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.UpdateData", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Update
'' Description: Update the cell
'' Inputs:      None
'' Returns:     Return code from geUpdateCell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Update() As Long
On Error GoTo ErrSection:

    Update = geUpdateCell(m.hQuoteObj, m.geCell)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.Update", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearCell
'' Description: Clear out a cell
'' Inputs:      None
'' Returns:     Return code from geUpdateCell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ClearCell() As Long
On Error GoTo ErrSection:

    m.dPrevBid = kNullData
    m.dPrevAsk = kNullData
    
    With m.geCell
        .SymbolID = -1&
        .CellID = -1&
        
        gdSetStr .gshSymbol, 0, ""
        gdSetStr .gshHigh, 0, ""
        gdSetStr .gshLast, 0, ""
        gdSetStr .gshLow, 0, ""
        gdSetStr .gshNetChange, 0, ""
        gdSetStr .gshOpen, 0, ""
        
        gdSetStr .gshBid, 0, ""
'        gdSetStr .gshAsk, 0, ""
        
        .dOpen = kNullData
        .dHigh = kNullData
        .dLow = kNullData
        .dLast = kNullData
        .dPrevOpen = kNullData
        .dPrevHigh = kNullData
        .dPrevLow = kNullData
        .dPrevClose = kNullData
        
'        .dBid = kNullData
'        .dAsk = kNullData
    End With
    
    Set m.Bars = New cGdBars
    m.Bars.Size = 0

    ClearCell = geUpdateCell(m.hQuoteObj, m.geCell)
'    Refresh

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.ClearCell", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Refresh the cell
'' Inputs:      Window Handle, Window Device Context
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Refresh(ByVal hWnd As Long, ByVal hDC As Long) As Long
On Error GoTo ErrSection:

    Refresh = geDrawQuoteCell(m.hQuoteObj, m.geCell, hWnd, hDC)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.Refresh", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Move
'' Description: Move the cell to a new location
'' Inputs:      Row and Column of new cell
'' Returns:     Return code from geUpdateCell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Move(ByVal NewRow As Long, ByVal NewCol As Long) As Long
On Error GoTo ErrSection:

    m.geCell.Row = NewRow
    m.geCell.Col = NewCol
    Move = geUpdateCell(m.hQuoteObj, m.geCell)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.Move", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Return a copy of the quote cell
'' Inputs:      None
'' Returns:     Copy of the quote cell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cQuoteCell
On Error GoTo ErrSection:

    Dim aCopy As New cQuoteCell
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.MakeCopy", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Subs & Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for a copy
'' Inputs:      Private data to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:
    
    m = mCopyFrom
    Set m.Bars = mCopyFrom.Bars.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCell.SetPrivateData", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Restore price colors
'' Inputs:      None
'' Returns:     True/False as to whether a text color was changed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ClearUpdatedColors() As Boolean
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim dTickDiff As Double
    
    If m.lUseUpdateColor = 0 Then Exit Function         '6474
    
    For lIndex = 1 To eGDColorIndex_Ask
        If frmQuotes.UseUpdatedColors Then
            If gdGetNum(m.geCell.glhTextColors, lIndex) = m.lUpdateColor Then
                dTickDiff = gdTickCount - m.tbCellTickCount(0, lIndex)
                If dTickDiff >= 0 And dTickDiff <= g.nUpdatedColorDuration Then
                Else
                    gdSetNum m.geCell.glhTextColors, lIndex, m.lUnchColor
                    ClearUpdatedColors = True
                End If
            End If
        ElseIf lIndex <> eGDColorIndex_Delta Then
            gdSetNum m.geCell.glhTextColors, lIndex, m.lUnchColor
            ClearUpdatedColors = True
        End If
    Next lIndex

    'clear or set high/low color and bold box if necessary
    ClearUpdatedColors = CheckHiLowPrice

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.ClearUpdatedColors", eGDRaiseError_Raise

End Function

Private Sub ChangePrice(ByVal CellPrice As eGDUpdateCell, ByVal dValue As Double, _
    ByVal eDisplayStyle As eGDQuoteStyle)
On Error GoTo ErrSection:

    Dim strPrice As String              ' String version of the price
    Dim strCurrent As String
    Dim dTickCount As Double            ' Previous tick count
    Dim dTickDiff As Double
    Dim lColor As Long
    Dim bUseUpdatedColors
    
    If m.lPrevRtCount > 3 Then
        bUseUpdatedColors = frmQuotes.UseUpdatedColors          '5710
    End If
    
    Select Case CellPrice
        Case eGDUpdateCell_Open
            If dValue <> m.geCell.dOpen Then
                If eDisplayStyle = eGDQuoteStyle_Forex Then
                    lColor = gdGetNum(m.geCell.glhTextColors, eGDColorIndex_Open)
                Else
                    gdSetNum m.geCell.glhTextColors, eGDColorIndex_Open, m.lUnchColor    'reset
                    lColor = m.lUpdateColor
                End If
                
                strPrice = m.Bars.PriceDisplay(dValue) & " "
                
                'ForexQB: open cell contains bid price (change to up/down color as appropriate)
                If eDisplayStyle = eGDQuoteStyle_Forex Then
                    If m.Bars.Prop(eBARS_MinMoveInTicks) < 1 Then
                        '07-25-2012: FP is flag for fractional pip (avaialble from IB at time of implementation)
                        strPrice = Trim(strPrice) & ";FP"
                    End If
                    If dValue > m.geCell.dOpen Then
                        lColor = m.lUpColor
                    ElseIf dValue < m.geCell.dOpen Then
                        lColor = m.lDownColor
                    'Else
                        'lColor = m.lUnchColor           'do this for now
                    End If
                End If
                
                gdSetStr m.geCell.gshOpen, 0, strPrice
                If m.geCell.dOpen <> dValue Then
                    '5710 - For some reason, this code gets executed right after RT is started or
                    '       when a new symbol is added even when dOpen >0 AND dOpen = dValue.
                    '       Once lPrevRtCount > 3 then the comparison behave as expected.
                    m.geCell.dOpen = dValue
                    
                    If bUseUpdatedColors Then
                        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Open, lColor      'm.lUpdateColor
                        m.tbCellTickCount(0, CellPrice) = gdTickCount
                    End If
                
                End If
            End If
        
        Case eGDUpdateCell_High
            If dValue <> m.geCell.dHigh Then
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_High, m.lUnchColor    'reset
                strPrice = m.Bars.PriceDisplay(dValue) & " "
                If CheckHiLowPrice Then
                    gdSetStr m.geCell.gshHigh, 0, strPrice
                    m.geCell.dHigh = dValue
                Else
                    If m.geCell.iBoldTextIdx = 2 Then m.geCell.iBoldTextIdx = 0
                    If m.geCell.dHigh <> dValue Then
                        m.geCell.dHigh = dValue
                        gdSetStr m.geCell.gshHigh, 0, strPrice
                        If bUseUpdatedColors Then
                            gdSetNum m.geCell.glhTextColors, eGDColorIndex_High, m.lUpdateColor
                            m.tbCellTickCount(0, CellPrice) = gdTickCount
                        End If
                    End If
                End If
            End If
        
        Case eGDUpdateCell_Low
            If dValue <> m.geCell.dLow Then
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Low, m.lUnchColor    'reset
                strPrice = m.Bars.PriceDisplay(dValue) & " "
                If CheckHiLowPrice Then
                    gdSetStr m.geCell.gshLow, 0, strPrice
                    m.geCell.dLow = dValue
                Else
                    If m.geCell.iBoldTextIdx = 3 Then m.geCell.iBoldTextIdx = 0
                    If m.geCell.dLow <> dValue Then
                        m.geCell.dLow = dValue
                        gdSetStr m.geCell.gshLow, 0, strPrice
                        If bUseUpdatedColors Then
                            gdSetNum m.geCell.glhTextColors, eGDColorIndex_Low, m.lUpdateColor
                            m.tbCellTickCount(0, CellPrice) = gdTickCount
                        End If
                    End If
                End If
            End If
        
        Case eGDUpdateCell_Close
            If dValue <> m.geCell.dLast Then
                If eDisplayStyle = eGDQuoteStyle_Forex Then
                    lColor = gdGetNum(m.geCell.glhTextColors, eGDColorIndex_Close)
                Else
                    gdSetNum m.geCell.glhTextColors, eGDColorIndex_Open, m.lUnchColor    'reset
                    lColor = m.lUpdateColor
                End If
                
                'ForexQB: close cell contains ask price (change to up/down color as appropriate)
                If eDisplayStyle = eGDQuoteStyle_Forex Then
                    strPrice = m.Bars.PriceDisplay(dValue) & " "
                    If m.Bars.Prop(eBARS_MinMoveInTicks) < 1 Then
                        '07-25-2012: FP is flag for fractional pip (avaialble from IB at time of implementation)
                        strPrice = Trim(strPrice) & ";FP"
                    End If
                    If dValue > m.geCell.dLast Then
                        lColor = m.lUpColor
                    ElseIf dValue < m.geCell.dLast Then
                        lColor = m.lDownColor
    '                Else
    '                    lColor = m.lUnchColor       'do this for now
                    End If
                Else
                    If m.Bars.Prop(eBARS_PriceHasSettled) <> 0 And m.Bars.SecurityType = "F" Then
                        strPrice = m.Bars.PriceDisplay(dValue) & "s"
                    Else
                        strPrice = m.Bars.PriceDisplay(dValue) & " "
                    End If
                End If
            
                gdSetStr m.geCell.gshLast, 0, strPrice
                If m.geCell.dLast <> dValue Then
                    m.geCell.dLast = dValue
                    If bUseUpdatedColors Then
                        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Close, lColor      'm.lUpdateColor
                        m.tbCellTickCount(0, CellPrice) = gdTickCount
                    End If
                End If
                CheckHiLowPrice
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JM 02-24-2011:
'   bid/ask prices are stored in same string
'   expected format is B=xxx.xxx A=xxx.xxx
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case eGDUpdateCell_Bid
'If m.Bars.Prop(eBARS_Symbol) = "HE-201302-S1" Then
'    dValue = dValue
'End If
If dValue = 0 Then
    dValue = dValue
End If
            If dValue > kNullData And dValue <> m.dPrevBid Then
            'If dValue > kNullData Then
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Bid, m.lUnchColor    'reset
            
                If m.dPrevBid <> dValue Then
                    m.dPrevBid = dValue
                    lColor = m.lUpdateColor
                Else
                    lColor = m.lUnchColor
                End If
                
                strPrice = Parse(gdGetStr(m.geCell.gshBid, 0), " ", 2)
                strPrice = "B=" & m.Bars.PriceDisplay(m.dPrevBid) & " " & strPrice & " "
                gdSetStr m.geCell.gshBid, 0, strPrice
                
                If bUseUpdatedColors Then
                    gdSetNum m.geCell.glhTextColors, eGDColorIndex_Bid, lColor
                    m.tbCellTickCount(0, CellPrice) = gdTickCount
                End If
            End If
            
        Case eGDUpdateCell_Ask
If dValue = 0 Then
    dValue = dValue
End If
            If dValue > kNullData And dValue <> m.dPrevAsk Then
            'If dValue > kNullData Then
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Ask, m.lUnchColor    'reset

                If m.dPrevAsk <> dValue Then
                    m.dPrevAsk = dValue
                    lColor = m.lUpdateColor
                Else
                    lColor = m.lUnchColor
                End If
                
                strPrice = Parse(gdGetStr(m.geCell.gshBid, 0), " ", 1)
                strPrice = strPrice & " A=" & m.Bars.PriceDisplay(m.dPrevAsk) & " "
                gdSetStr m.geCell.gshBid, 0, strPrice
                
                If bUseUpdatedColors Then
                    gdSetNum m.geCell.glhTextColors, eGDColorIndex_Ask, lColor
                    m.tbCellTickCount(0, CellPrice) = gdTickCount
                End If
            End If
        
        Case eGDUpdateCell_Delta
            strPrice = m.Bars.PriceDisplay(dValue) & " "
            If dValue > 0 And Left(strPrice, 1) <> "+" Then strPrice = "+" & strPrice
            gdSetStr m.geCell.gshNetChange, 0, strPrice
            If dValue > 0 Then
                lColor = m.lUpColor
            ElseIf dValue < 0 Then
                lColor = m.lDownColor
            Else
                lColor = m.lUnchColor
            End If
            gdSetNum m.geCell.glhTextColors, eGDColorIndex_Delta, lColor
            m.lDeltaColor = lColor
            
            If m.lColorSymbol Then
                ' if Realtime and have up or down trades for the day, then color symbol as an Up or Down tick
                If g.RealTime.Active And m.Bars(eBARS_High, m.Bars.Size - 1) <> m.Bars(eBARS_Low, m.Bars.Size - 1) Then
                    If m.Bars.Prop(eBARS_LastTickDown) = 0 Then
                        lColor = m.lUpColor
                    Else
                        lColor = m.lDownColor
                    End If
                Else
                    lColor = vbBlack
                End If
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Symbol, lColor
            Else
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Symbol, vbBlack
            End If
        
        Case eGDUpdateCell_Time
            strCurrent = gdGetStr(m.geCell.gshTickTime)
            If dValue = 0 Or m.Bars.Size = 0 Then
                dValue = m.Bars(eBARS_DateTime, m.Bars.Size - 1)
                If dValue > 0 Then
                    strPrice = DateFormat(dValue, MM_DD_YYYY)
                Else
                    strPrice = ""
                End If
            Else
                dValue = Int(m.Bars(eBARS_DateTime, m.Bars.Size - 1)) + dValue / 1440#
                If g.bShowInLocalTimeZone Then
                    dValue = ConvertTimeZone(dValue, m.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                End If
                dValue = gdFixDateTime(dValue)
                If g.RealTime.Active And m.lDelay > 0 Then
                    strPrice = "  (d=" & Str(m.lDelay) & ")"
                Else
                    strPrice = ""
                End If
                ' don't show seconds if seconds = 0 or if a delay (since not enough room)
                If (Second(dValue) <> 0 Or InStr(4, strCurrent, ":") > 0) And Len(strPrice) = 0 Then
                    strPrice = Format(dValue, "hh:mm:ss")
                Else
                    strPrice = Format(dValue, "hh:mm") & strPrice
                End If
            End If
            If strCurrent <> strPrice Then
                gdSetStr m.geCell.gshTickTime, 0, strPrice
                If bUseUpdatedColors Then
                    m.tbCellTickCount(0, CellPrice) = gdTickCount
                End If
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCell.ChangePrice", eGDRaiseError_Raise
    
End Sub

Private Function CheckHiLowPrice() As Boolean
On Error GoTo ErrSection:

    Dim iBoldText As Long
    
    'set high/low price to up/down color if equals current price (grapheng.dll will bold box)
    If Not m.Bars Is Nothing Then
        If m.Bars(eBARS_High, m.Bars.Size - 1) <> m.Bars(eBARS_Low, m.Bars.Size - 1) Then
            If m.Bars(eBARS_Close, m.Bars.Size - 1) = m.Bars(eBARS_High, m.Bars.Size - 1) Then
                iBoldText = 2
                If gdGetNum(m.geCell.glhTextColors, eGDColorIndex_High) = m.lUnchColor Then
                    gdSetNum m.geCell.glhTextColors, eGDColorIndex_High, m.lUpColor
                    CheckHiLowPrice = True
                End If
            ElseIf gdGetNum(m.geCell.glhTextColors, eGDColorIndex_High) = m.lUpColor Then
                If m.geCell.iBoldTextIdx = 2 Then iBoldText = 0
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_High, m.lUnchColor
                CheckHiLowPrice = True
            End If
            If m.Bars(eBARS_Close, m.Bars.Size - 1) = m.Bars(eBARS_Low, m.Bars.Size - 1) Then
                iBoldText = 3
                If gdGetNum(m.geCell.glhTextColors, eGDColorIndex_Low) = m.lUnchColor Then
                    gdSetNum m.geCell.glhTextColors, eGDColorIndex_Low, m.lDownColor
                    CheckHiLowPrice = True
                End If
            ElseIf gdGetNum(m.geCell.glhTextColors, eGDColorIndex_Low) = m.lDownColor Then
                If m.geCell.iBoldTextIdx = 3 Then iBoldText = 0
                gdSetNum m.geCell.glhTextColors, eGDColorIndex_Low, m.lUnchColor
                CheckHiLowPrice = True
            End If
        End If
    End If
    
    m.geCell.iBoldTextIdx = iBoldText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCell.CheckHiLowPrice", eGDRaiseError_Raise
    
End Function

'nAddRemove: -1=just update, 0=remove alert, 1=addalert
Public Sub QbCellAlertUpdate(oAlert As cAlert, ByVal nAddRemove&)
On Error GoTo ErrExit:

    If g.bUnloading Then Exit Sub

    If Not oAlert Is Nothing Then
        If oAlert.Symbol <> m.Bars.Prop(eBARS_Symbol) Then GoTo ErrExit
        If oAlert.Active Then
            Select Case nAddRemove
                Case -1:
                    m.nActiveAlert = m.nActiveAlert + 1
                    m.nInactiveAlert = m.nInactiveAlert - 1
                Case 0:
                    m.nActiveAlert = m.nActiveAlert - 1
                Case 1:
                    m.nActiveAlert = m.nActiveAlert + 1
                    oAlert.AddQbCell Me
            End Select
        Else
            Select Case nAddRemove
                Case -1:
                    m.nInactiveAlert = m.nInactiveAlert + 1
                    m.nActiveAlert = m.nActiveAlert - 1
                Case 0:
                    m.nInactiveAlert = m.nInactiveAlert - 1
                Case 1:
                    m.nInactiveAlert = m.nInactiveAlert + 1
                    oAlert.AddQbCell Me
            End Select
        End If
    End If

    If m.nActiveAlert >= 1 Then
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, 1
    ElseIf m.nInactiveAlert >= 1 Then
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, 0
    Else
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, -1
    End If
    
    If Not m.MyQbBoard Is Nothing Then m.MyQbBoard.DrawBoard

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cQuoteCell.QbCellAlertUpdate"

End Sub

Public Sub QbCellAlertInit()
On Error GoTo ErrSection:

    Dim i&, strBarsSymbol$
    Dim Alert As cAlert

    m.nActiveAlert = 0
    m.nInactiveAlert = 0
    
    If m.Bars Is Nothing Then GoTo ErrExit
    
    strBarsSymbol = m.Bars.Prop(eBARS_Symbol)
    If Len(strBarsSymbol) = 0 Then GoTo ErrExit
    
    For i = 1 To g.Alerts.Count
        Set Alert = g.Alerts(i)
        If Not Alert Is Nothing Then
            If Len(Alert.Symbol) > 0 Then
                If strBarsSymbol = Alert.Symbol Then
                    If Alert.Active Then
                        m.nActiveAlert = m.nActiveAlert + 1
                    Else
                        m.nInactiveAlert = m.nInactiveAlert + 1
                    End If
                    Alert.AddQbCell Me
                End If
            End If
        End If
    Next
    
    If m.nActiveAlert >= 1 Then
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, 1
    ElseIf m.nInactiveAlert >= 1 Then
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, 0
    Else
        gdSetNum m.geCell.glhTextColors, eGDColorIndex_Alert, -1
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cQuoteCell.QbCellAlertInit"

End Sub
