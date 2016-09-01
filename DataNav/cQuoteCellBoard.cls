VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuoteCellBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cQuoteCellBoard.cls
'' Description: Controller class for the thermometer style quote board
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDTblCols
    eGDTblCol_Symbol = 0
    eGDTblCol_Cells
    eGDTblCol_NumCols
End Enum

Private WithEvents gdsHorz As gdScrollBar
Attribute gdsHorz.VB_VarHelpID = -1
Private WithEvents gdsVert As gdScrollBar
Attribute gdsVert.VB_VarHelpID = -1

Private Type mPrivate
    pbQuotes As PictureBox
    
    tblIndex As cGdTable
    
    Cells As cGdTree
    
    QuoteWinStruct As quote_win
    hQuoteObj As Long
    
    lMouseRow As Long
    lMouseCol As Long
    lMouseID As Long
    lRow As Long
    lCol As Long
    lID As Long
    
    lVertVal As Long
    lHorzVal As Long
    
    bDragging As Boolean
    bHittestBell As Boolean
    bFullHeight As Boolean
    
    lUpdateColor As Long
    lUseUpdateColor As Long     'if not zero then use delta color for OHLC (Jo Dinapoli--all red/green request)
    lColorSymbol As Long        'if zero then symbol is always black
    
    lCompactBox As Long
    lCompactForex As Long
    
    bInitDone As Boolean
    bSkipScroll As Boolean
End Type
Dim m As mPrivate

Private Function TblCol(ByVal Col As eGDTblCols) As Long
    TblCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Constructor & Destructor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Set m.tblIndex = New cGdTable
    m.tblIndex.CreateField eGDARRAY_Strings, TblCol(eGDTblCol_Symbol), "Symbol"
    m.tblIndex.CreateField eGDARRAY_Strings, TblCol(eGDTblCol_Cells), "Cells"
    
    Set m.Cells = New cGdTree
    
    ' Create some gdArrays...
    m.QuoteWinStruct.glhWinDim = gdCreateArray(eGDARRAY_Longs, 4)
    m.QuoteWinStruct.glhTextColors = gdCreateArray(eGDARRAY_Longs, 4)
    m.QuoteWinStruct.gshTabAlerts = gdCreateArray(eGDARRAY_Strings)
    m.QuoteWinStruct.gshfName = gdCreateArray(eGDARRAY_gdString)
    
    SetDefaults
    
    m.lMouseCol = -1&
    m.lMouseRow = -1&
    m.lMouseID = -1&
    m.lCol = -1&
    m.lRow = -1&
    m.lID = -1&
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.Initialize", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:
    
    Set m.tblIndex = Nothing
    Set m.Cells = Nothing
    
    geCloseQuoteWin& m.hQuoteObj
    m.hQuoteObj = 0
    
    gdDestroyArray m.QuoteWinStruct.glhWinDim
    gdDestroyArray m.QuoteWinStruct.gshfName
    gdDestroyArray m.QuoteWinStruct.glhTextColors
    gdDestroyArray m.QuoteWinStruct.glhHorzScroll
    gdDestroyArray m.QuoteWinStruct.glhVertScroll
    gdDestroyArray m.QuoteWinStruct.gshTabAlerts
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.Terminate", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Properties
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get QuoteBox() As PictureBox
    QuoteBox = m.pbQuotes
End Property
Public Property Let QuoteBox(pbQuotes As PictureBox)
    Set m.pbQuotes = pbQuotes
    If m.hQuoteObj = 0 Then InitBoard True
End Property

Public Property Get HScrollBar() As gdScrollBar
    Set HScrollBar = gdsHorz
End Property
Public Property Let HScrollBar(HBar As gdScrollBar)
    Set gdsHorz = HBar
End Property

Public Property Get VScrollBar() As gdScrollBar
    Set VScrollBar = gdsVert
End Property
Public Property Let VScrollBar(VBar As gdScrollBar)
    Set gdsVert = VBar
End Property

Public Property Get QuoteBoardObj() As Long
    QuoteBoardObj = m.hQuoteObj
End Property

Public Property Get QuoteBoardStyle() As eGDQuoteStyle
    QuoteBoardStyle = m.QuoteWinStruct.qwinStyle
End Property

Public Property Let QuoteBoardStyle(ByVal qbStyle As eGDQuoteStyle)
On Error GoTo ErrSection:

    Dim qbPrevStyle As eGDQuoteStyle
    
    qbPrevStyle = m.QuoteWinStruct.qwinStyle
    m.QuoteWinStruct.qwinStyle = qbStyle
    
    ' forex box-style is different from all other box-styles (needs reinitialization)
    If (qbPrevStyle = eGDQuoteStyle_Forex And qbStyle <> eGDQuoteStyle_Forex) Or _
       (qbPrevStyle <> eGDQuoteStyle_Forex And qbStyle = eGDQuoteStyle_Forex) Then
        Reset
    Else
        DrawBoard
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cQuoteCellBoard.QuoteBoardStyle.Let", eGDRaiseError_Raise
    
End Property

Public Property Get MouseRow() As Long
    MouseRow = m.lMouseRow
End Property
Public Property Get MouseCol() As Long
    MouseCol = m.lMouseCol
End Property
Public Property Get Row() As Long
    Row = m.lRow
End Property
Public Property Let Row(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim nRow&, nCol&
    
    geGetBoldCell m.hQuoteObj, nRow, nCol
    
    If nRow < 0 Then
        m.lRow = lRow
    Else
        BoldCell m.lRow, m.lCol, 0
        m.lRow = lRow
        BoldCell m.lRow, m.lCol, 1
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cQuoteCellBoard.Row.Let", eGDRaiseError_Raise
    
End Property
Public Property Get Col() As Long
    Col = m.lCol
End Property
Public Property Let Col(ByVal lCol As Long)
On Error GoTo ErrSection:

    Dim nRow&, nCol&
    
    geGetBoldCell m.hQuoteObj, nRow, nCol
    
    If nCol < 0 Then
        m.lCol = lCol
    Else
        BoldCell m.lRow, m.lCol, 0
        m.lCol = lCol
        BoldCell m.lRow, m.lCol, 1
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cQuoteCellBoard.Col.Let", eGDRaiseError_Raise
    
End Property

Public Property Get Rows() As Long
    Rows = m.QuoteWinStruct.maxRow
End Property
Public Property Get Cols() As Long
    Cols = m.QuoteWinStruct.maxCol
End Property

Public Property Get Cell(ByVal lRow As Long, ByVal lCol As Long) As cQuoteCell
    Set Cell = m.Cells(Key(lRow, lCol))
End Property

Public Property Get Font() As StdFont
    Set Font = m.pbQuotes.Font
End Property

Public Property Get FirstCellInTree() As cQuoteCell
    Set FirstCellInTree = m.Cells(1)
End Property

Public Property Let Font(NewFont As StdFont)
On Error GoTo ErrSection:

    Dim SavedCells As New cGdTree       ' Saved off cells from the quote board
    Dim lIndex As Long                  ' Index into a for loop
    Dim OldCell As New cQuoteCell       ' Temporary cell object
    Dim NewCell As New cQuoteCell       ' Temporary cell object
    Dim lID As Long                     ' New Cell ID
    Dim lStyle As Long                  ' Font Style

    Set SavedCells = m.Cells.MakeCopy
    
    If Not m.pbQuotes Is Nothing Then
        Set m.pbQuotes.Font = NewFont
    End If
    
    lStyle = 0&
    
    ' Set up the font with the new information...
    gdSetStr m.QuoteWinStruct.gshfName, 0, NewFont.Name
    m.QuoteWinStruct.fSize = NewFont.Size
    m.QuoteWinStruct.fUnderline = NewFont.Underline
    If NewFont.Bold Then lStyle = lStyle Or 1
    If NewFont.Italic Then lStyle = lStyle Or 2
    m.QuoteWinStruct.fStyle = lStyle
    
    Reset
    
ErrExit:
    Set OldCell = Nothing
    Set NewCell = Nothing
    Set SavedCells = Nothing
    Exit Property
    
ErrSection:
    Set OldCell = Nothing
    Set NewCell = Nothing
    Set SavedCells = Nothing
    RaiseError "cQuoteCellBoard.Font.Let", eGDRaiseError_Raise
    
End Property

Public Property Get UpColor() As Long
    UpColor = m.QuoteWinStruct.netUpColor
End Property
Public Property Let UpColor(ByVal NewColor As Long)
    m.QuoteWinStruct.netUpColor = NewColor
End Property

Public Property Get DownColor() As Long
    DownColor = m.QuoteWinStruct.netDownColor
End Property
Public Property Let DownColor(ByVal NewColor As Long)
    m.QuoteWinStruct.netDownColor = NewColor
End Property

Public Property Get UnchColor() As Long
    UnchColor = m.QuoteWinStruct.netUnchColor
End Property
Public Property Let UnchColor(ByVal NewColor As Long)
    m.QuoteWinStruct.netUnchColor = NewColor
End Property

Public Property Get FullBarHeight() As Boolean
    FullBarHeight = m.bFullHeight
End Property
Public Property Let FullBarHeight(ByVal bFullBarHeight As Boolean)
On Error GoTo ErrSection:

    m.bFullHeight = bFullBarHeight
    UpdateAll
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cQuoteCellBoard.FullBarHeight.Let", eGDRaiseError_Raise
    
End Property

Public Property Get CellCount() As Long
    CellCount = m.Cells.Count
End Property

Public Property Get CellItem(ByVal vIndexOrKey As Variant) As cQuoteCell
    Set CellItem = m.Cells(vIndexOrKey)
End Property

Public Property Get LastDataRow() As Long
    LastDataRow = m.QuoteWinStruct.LastDataRow
End Property

Public Property Get LastDataCol() As Long
    LastDataCol = m.QuoteWinStruct.LastDataCol
End Property

Public Property Get UpdateColor() As Long
    UpdateColor = m.lUpdateColor
End Property
Public Property Let UpdateColor(ByVal lUpdateColor As Long)
    m.lUpdateColor = lUpdateColor
    m.QuoteWinStruct.thermoOpenColor = m.lUpdateColor
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Public Subs & Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'temporary test code (begin)
#If 0 Then
Public Sub PrintBoard()

    Screen.MousePointer = 11
    Printer.ScaleMode = 3  'Pixels
    Printer.Orientation = vbPRORLandscape   'vbPRORPortrait
    Printer.Print
    gePrintQuoteWin m.hQuoteObj, m.QuoteForm.hWnd, Printer.hDC
    Printer.EndDoc
    Screen.MousePointer = 0

End Sub
#End If
'temporary test code (end)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the controls for the quote cell board class
'' Inputs:      Picture Box, Horizontal Scroll Bar, Vertical Scroll Bar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(pbQuotes As PictureBox, Horz As gdScrollBar, Vert As gdScrollBar)
On Error GoTo ErrSection:

    Set m.pbQuotes = pbQuotes
    Set gdsHorz = Horz
    Set gdsVert = Vert
    If m.hQuoteObj = 0 Then InitBoard True
    Font = pbQuotes.Font

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.Init", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearAll
'' Description: Clear out the all of the cells in the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearAll()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    ' Walk through all of the cells in the collection and clear them out...
    For lIndex = m.Cells.Count To 1 Step -1
        If Not m.Cells(lIndex) Is Nothing Then
            m.Cells(lIndex).ClearCell
            m.Cells(lIndex).Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
            m.Cells.Remove lIndex
        End If
    Next
    
    ' Also clear out the cross reference table...
    m.tblIndex.NumRecords = 0
    
    ' Draw the Quote Board again...
    DrawBoard

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.ClearAll", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DrawBoard
'' Description: Draw the Quote Board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DrawBoard()
On Error GoTo ErrSection:

'    Dim bLocked As Boolean
    
    If Not m.bInitDone Then Exit Sub        'aardvark 6781

    If m.QuoteWinStruct.qwinStyle = eGDQuoteStyle_Forex Then
        If m.QuoteWinStruct.nCompactQB <> m.lCompactForex Then m.QuoteWinStruct.nCompactQB = m.lCompactForex
    ElseIf m.QuoteWinStruct.qwinStyle <> eGDQuoteStyle_Forex Then
        If m.QuoteWinStruct.nCompactQB <> m.lCompactBox Then m.QuoteWinStruct.nCompactQB = m.lCompactBox
    End If
    If g.nColorTheme = kDarkThemeColor Then
        m.QuoteWinStruct.bkColor = kDarkThemeColor
    Else
        m.QuoteWinStruct.bkColor = vbWhite
    End If
    
    'bLocked = LockWindowUpdate(m.pbQuotes.hWnd)
        SetupScroll
        geDrawQuoteWin m.hQuoteObj, m.pbQuotes.hWnd, m.pbQuotes.hDC
        'JM 02-03-2016: This is the which comes first, the chicken or the egg scenario.
        '   The visibility of scroll bars depend on last visible row & column.
        '   Last visible row & column cannot be known until graphics engine draws the board.
        '   Graphics engine cannot draw the board correctly until it knows what the upper left cell should be.
        '   1. First call to SetupScrollbar lets graphics engine know what the upper left cell should be.
        '   2. Second call to SetupScrollbar shows or hide scrollbars base on correct uppler left cell
        SetupScroll
        m.pbQuotes.Refresh
    'If bLocked Then LockWindowUpdate 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.DrawBoard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindCell
'' Description: Find a cell in the collection by row and column
'' Inputs:      Row, Column, ID
'' Returns:     Quote Cell object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindCell(ByVal lRow As Long, ByVal lCol As Long) As cQuoteCell
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim Cell As cQuoteCell              ' Quote Board Cell object
    Dim lID As Long                     ' ID of the cell
    
    ' Do some initializing...
    Set Cell = Nothing
    
    ' Walk through the cells in the collection until the row/column is found...
    For lIndex = 1 To m.Cells.Count
        Set Cell = m.Cells(lIndex)
        If Cell.Row = lRow And Cell.Col = lCol Then
            Cell.ID = lIndex
            Exit For
        End If
    Next
    
    ' Return the ID and the Quote Cell object...
    If Cell.ID = -1 Then Set Cell = Nothing
    Set FindCell = Cell

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.FindCell", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HighlightCell
'' Description: Select/Highlight a certain cell in the quote board
'' Inputs:      ID, Row, Column, Toggle
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HighlightCell(ByVal lRow As Long, ByVal lCol As Long, ByVal lToggle As Long)
On Error GoTo ErrSection:

    If lRow >= 0 And lCol >= 0 Then
        geSetHighlightCell m.hQuoteObj, m.pbQuotes.hWnd, m.pbQuotes.hDC, lRow, lCol, lToggle
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.HighlightCell", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BoldCell
'' Description: Select/Bold a certain cell in the quote board
'' Inputs:      ID, Row, Column, Toggle
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BoldCell(ByVal lRow As Long, ByVal lCol As Long, ByVal lToggle As Long)
On Error GoTo ErrSection:

    If lRow >= 0 And lCol >= 0 Then
        geSetBoldCell m.hQuoteObj, m.pbQuotes.hWnd, m.pbQuotes.hDC, lRow, lCol, lToggle
        m.pbQuotes.Refresh
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.BoldCell", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveCell
'' Description: Move a cell from one location to another
'' Inputs:      Destination Cell Object, Source Cell Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MoveCell(ByVal lOldRow As Long, ByVal lOldCol As Long, ByVal lNewRow As Long, ByVal lNewCol As Long)
On Error GoTo ErrSection:
    
    Dim NewCell As New cQuoteCell
    Dim OldCell As New cQuoteCell
    Dim lID As Long
    Dim Bars As New cGdBars
    Dim bNewRow As Boolean
    Dim bNewCol As Boolean
    
    ' Initialize variables...
    lID = -1&
    Set Bars = Nothing
    
    If lNewRow >= m.QuoteWinStruct.maxRow Then bNewRow = True
    If lNewCol >= m.QuoteWinStruct.maxCol Then bNewCol = True
    
    If Not Cell(lNewRow, lNewCol) Is Nothing Then
        Set NewCell = Cell(lNewRow, lNewCol)
        lID = NewCell.ID
        Set Bars = NewCell.Bars.MakeCopy
    Else
        Set NewCell = CreateCell(lNewRow, lNewCol)
    End If
    Set OldCell = Cell(lOldRow, lOldCol)
    
    If (Not NewCell Is Nothing) And (Not OldCell Is Nothing) Then
        With NewCell
            .ID = OldCell.ID
            .Bars = OldCell.Bars.MakeCopy
            
            OldCell.ClearCell
            OldCell.Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
            RemoveCell lOldRow, lOldCol
            Set OldCell = Nothing
            
            .UpdateData .Bars, m.bFullHeight, m.QuoteWinStruct.qwinStyle, m.QuoteWinStruct.nCompactQB, True
            .Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
        End With
        
        m.Cells.Add NewCell, Key(lNewRow, lNewCol)
        
        If lID >= 0 Then
            Set OldCell = CreateCell(lOldRow, lOldCol)
            With OldCell
                .ID = lID
                .Bars = Bars.MakeCopy
                .UpdateData .Bars, m.bFullHeight, m.QuoteWinStruct.qwinStyle, m.QuoteWinStruct.nCompactQB, True
                .Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
            End With
            m.Cells.Add OldCell, Key(lOldRow, lOldCol)
        End If
        RebuildXref
        'redo scroll
        If bNewRow Or bNewCol Then
            SetupScroll
        End If
    End If

ErrExit:
    Set Bars = Nothing
    Set OldCell = Nothing
    Set NewCell = Nothing
    Exit Sub

ErrSection:
    Set Bars = Nothing
    Set OldCell = Nothing
    Set NewCell = Nothing
    RaiseError "cQuoteCellBoard.MoveCell", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbol
'' Description: Insert a new symbol in the highlighted cell
'' Inputs:      Symbol ID/Symbol to add, Row and Column of cell
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSymbol(ByVal vSymbol As Variant, Optional ByVal lRow As Long = -1&, Optional ByVal lCol As Long = -1&)
On Error GoTo ErrSection:

'JM 06-15-2015:
'   this routine can be called with a single symbol or iteratively from BoardFromString
'   for tab-switching or loading from file, symbol validity is checked in BoardFromString
'   for individual symbol added by user, symbol validity is checked in quotes.frm code


    Dim rc As Long                      ' Return code
    Dim NewCell As cQuoteCell           ' Quote Board Cell object
    Dim strSymbol As String             ' Symbol to add
    Dim lSymbolID As Long               ' Symbol ID to add
    Dim lPos As Long                    ' Position to add to the index table
    Dim strLookup As String             ' String to look up in the table
    
    ' Get the requested cell from the collection...
    If lRow >= 0 And lCol >= 0 Then
        If Not Cell(lRow, lCol) Is Nothing Then
            RemoveCell lRow, lCol
        End If
        Set NewCell = CreateCell(lRow, lCol)
    Else
        Set NewCell = FindFirstFreeCell
    End If
    
    ' Update and refresh the cell that received the data...
    With NewCell
        .SymbolID = GetSymbolID(vSymbol)
        .Symbol = GetSymbol(vSymbol)
        .Update
        .Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
    End With
    AddCell NewCell
    
    ' Add the cross refrence of the cell index to the symbol...
    strLookup = SymbolOrSymbolID(NewCell.SymbolID, NewCell.Symbol)
    If m.tblIndex.FieldArray(TblCol(eGDTblCol_Symbol), False).BinarySearch(strLookup, lPos) = False Then
        m.tblIndex.AddRecord strLookup & vbTab & Str(NewCell.ID), lPos, vbTab
    Else
        m.tblIndex(TblCol(eGDTblCol_Cells), lPos) = m.tblIndex(TblCol(eGDTblCol_Cells), lPos) & "," & Str(NewCell.ID)
    End If
    
    If m.lRow < 0 Then Row = NewCell.Row
    If m.lCol < 0 Then Col = NewCell.Col

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.AddSymbol", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Refresh the cells in the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Walk through all of the cells in the collection and refresh them...
    For lIndex = 1 To m.Cells.Count
        m.Cells(lIndex).Refresh
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteBoardCell.Refresh", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResizeBoard
'' Description: Resize the dimensions of the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ResizeBoard(Optional ByVal bIncludeScrollBars As Boolean = True)
On Error GoTo ErrSection:

    gdSetNum m.QuoteWinStruct.glhWinDim, 0, 0
    gdSetNum m.QuoteWinStruct.glhWinDim, 1, 0
    gdSetNum m.QuoteWinStruct.glhWinDim, 2, m.pbQuotes.ScaleHeight / Screen.TwipsPerPixelY
    gdSetNum m.QuoteWinStruct.glhWinDim, 3, m.pbQuotes.ScaleWidth / Screen.TwipsPerPixelX
    
    If Not bIncludeScrollBars Then
        SetVertScrollDim 0, 0, 0, 0
        SetHorzScrollDim 0, 0, 0, 0
    End If
    
    DrawBoard
    UpdateAll

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteBoardCell.ResizeBoard", eGDRaiseError_Raise
    
End Sub

Public Sub GotFocus()
On Error GoTo ErrSection:

    BoldCell m.lRow, m.lCol, 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.GotFocus", eGDRaiseError_Raise
    
End Sub

Public Sub LostFocus()
On Error GoTo ErrSection:

    m.bDragging = False
    BoldCell m.lRow, m.lCol, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.LostFocus", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MouseDown
'' Description: If the user left clicks, highlight the cell and start drag
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim cInfo As cell_hittest_info
    Dim lReturn As Long
    Static bRightButton As Boolean
    
    
    If Button = vbLeftButton Then
        If m.lMouseCol > -1& And m.lMouseRow > -1& Then
            
            If Not m.bHittestBell Then BoldCell m.lRow, m.lCol, 0
            m.lRow = m.lMouseRow
            m.lCol = m.lMouseCol
            If Not m.bHittestBell Then BoldCell m.lRow, m.lCol, 1
            
            m.lID = m.lMouseID
            If Not bRightButton And Not m.bHittestBell Then m.bDragging = True
            bRightButton = False
        End If
    Else
        bRightButton = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteBoardCell.MouseDown", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MouseMove
'' Description: As the user moves the mouse, change the mouse row/col
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim i&
    Dim cInfo As cell_hittest_info
    Dim lReturn As Long
    Static strDesc As String
    
    Dim bBell As Boolean
    
    lReturn = geHitTestCell(m.hQuoteObj, m.pbQuotes.hDC, cInfo, eGEQuoteItem_iconBell, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
   
    If lReturn = 0 Then
        m.pbQuotes.ToolTipText = ""
        m.lMouseRow = -1&
        m.lMouseCol = -1&
        m.lMouseID = -1&
    Else
        m.bHittestBell = False
        m.lMouseRow = cInfo.Row
        m.lMouseCol = cInfo.Col
        m.lMouseID = cInfo.CellID
        Select Case cInfo.Item
            Case eGEQuoteItem_Symbol
                If Not m.bDragging Then
                    If cInfo.Item = eGEQuoteItem_Symbol Then
                        If m.pbQuotes.ToolTipText <> strDesc Or Len(m.pbQuotes.ToolTipText) = 0 Then
                            i = Cell(cInfo.Row, cInfo.Col).Bars.Prop(eBARS_SymbolID)
                            strDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbolID(i))
                            If Len(strDesc) = 0 Then
                                strDesc = Cell(cInfo.Row, cInfo.Col).Bars.Prop(eBARS_Desc)
                            End If
                            m.pbQuotes.ToolTipText = strDesc
                        End If
                    Else
                        m.pbQuotes.ToolTipText = ""
                    End If
                End If
            Case eGEQuoteItem_iconBell
                m.bHittestBell = True
            Case Else
                m.pbQuotes.ToolTipText = ""
        End Select
    End If

    If m.bDragging Then
        m.pbQuotes.MousePointer = 99
        m.pbQuotes.MouseIcon = Picture16(ToolbarIcon("kMove4Way"))
    ElseIf m.bHittestBell Then
        m.pbQuotes.MousePointer = 99
        m.pbQuotes.MouseIcon = Picture16(ToolbarIcon("kHand"))
    Else
        m.pbQuotes.MousePointer = 1
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MouseUp
'' Description: When the user presses the mouse, set up some properties
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     1 if cell moved, 0 otherwize
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) As Long
On Error GoTo ErrSection:

    Dim bLocked As Boolean
    Dim strSymbol$

    MouseUp = 0&

    If m.bDragging Then
        ' Turn dragging off...
        m.bDragging = False
        
        ' If the source has data and the user is in a new cell, move the cell...
        If Not Cell(m.lRow, m.lCol) Is Nothing Then
            If m.lMouseRow <> m.lRow Or m.lMouseCol <> m.lCol Then
                BoldCell m.lRow, m.lCol, 0
                bLocked = LockWindowUpdate(m.pbQuotes.hWnd)
                    MoveCell m.lRow, m.lCol, m.lMouseRow, m.lMouseCol
                    m.lRow = m.lMouseRow
                    m.lCol = m.lMouseCol
                    BoldCell m.lRow, m.lCol, 1
                    MouseUp = 1&
                If bLocked Then LockWindowUpdate 0
            End If
        End If
    ElseIf m.bHittestBell Then
        If Not Cell(m.lRow, m.lCol) Is Nothing Then strSymbol = Cell(m.lRow, m.lCol).Symbol
        If FormIsLoaded("frmAlertsSetup") Then
            frmAlertsSetup.LoadGrid strSymbol
        Else
            frmAlertsSetup.ShowMe strSymbol
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.MouseUp", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyDown
'' Description: Handle the navigation keys for the control
'' Inputs:      Code of the key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim lLastCol As Long                ' Last column to scroll to
    Dim lLastRow As Long                ' Last row to scroll to
    Dim lVisRows As Long                ' Number of visible rows
    Dim lVisCols As Long                ' Number of visible columns

    ' Calculate the last column...
    If m.QuoteWinStruct.LastDataCol > m.QuoteWinStruct.lastVisCol Then
        lLastCol = m.QuoteWinStruct.LastDataCol
    Else
        lLastCol = m.QuoteWinStruct.lastVisCol
    End If
'    If lLastCol > m.QuoteWinStruct.LastDataCol Then lLastCol = m.QuoteWinStruct.LastDataCol
    If lLastCol >= m.QuoteWinStruct.maxCol Then lLastCol = m.QuoteWinStruct.maxCol - 1
    
    ' Calculate the last row...
    If m.QuoteWinStruct.LastDataRow > m.QuoteWinStruct.lastVisRow Then
        lLastRow = m.QuoteWinStruct.LastDataRow
    Else
        lLastRow = m.QuoteWinStruct.lastVisRow
    End If
'    If lLastRow > m.QuoteWinStruct.LastDataRow Then lLastRow = m.QuoteWinStruct.LastDataRow
    If lLastRow >= m.QuoteWinStruct.maxRow Then lLastRow = m.QuoteWinStruct.maxRow - 1
    
    ' Calculate the number of visible rows...
    lVisCols = m.QuoteWinStruct.lastVisCol - m.QuoteWinStruct.firstVisCol + 1
    lVisRows = m.QuoteWinStruct.lastVisRow - m.QuoteWinStruct.firstVisRow + 1
    
    Select Case KeyCode
        Case vbKeyRight
            If m.lCol + 1 < m.QuoteWinStruct.maxCol Then
                Col = Col + 1
                If lVisCols > 1 Then m.bSkipScroll = True
                gdsHorz.Value = Col
            End If
        
        Case vbKeyLeft
            If m.lCol - 1 >= 0 Then
                Col = Col - 1
                If lVisCols > 1 Then m.bSkipScroll = True
                gdsHorz.Value = Col
            End If
            
        
        Case vbKeyUp
            If m.lRow - 1 >= 0 Then
                Row = Row - 1
                If lVisRows > 1 Then m.bSkipScroll = True
                gdsVert.Value = Row
            End If
        
        Case vbKeyDown
            If m.lRow + 1 < m.QuoteWinStruct.maxRow Then
                Row = Row + 1
                If lVisRows > 1 Then m.bSkipScroll = True
                gdsVert.Value = Row
            End If
            
        Case vbKeyEnd
            If Shift And vbCtrlMask Then
                'go to last row that has a cell with data then go to right-most cell on that row that has data
                BoldCell m.lRow, m.lCol, 0
                'm.lRow = lLastRow
                'm.lCol = lLastCol
                m.lRow = m.QuoteWinStruct.LastDataRow
                m.lCol = LastDataColForRow(m.lRow)
                BoldCell m.lRow, m.lCol, 1
                
                If m.lCol <= m.QuoteWinStruct.lastVisCol Then m.bSkipScroll = True
                gdsHorz.Value = m.lCol      ' - lVisCols
                If m.lRow <= m.QuoteWinStruct.lastVisRow Then
                    m.bSkipScroll = True
                Else
                    m.bSkipScroll = False
                End If
                gdsVert.Value = m.lRow      ' - lVisRows
                DrawBoard
            Else
                'go right-most cell on current row that has data
                BoldCell m.lRow, m.lCol, 0
                m.lCol = LastDataColForRow(m.lRow)
                BoldCell m.lRow, m.lCol, 1
                If m.lCol <= m.QuoteWinStruct.lastVisCol Then m.bSkipScroll = True
                gdsHorz.Value = m.lCol      '- lVisCols
            End If
        
        Case vbKeyHome
            If Shift And vbCtrlMask Then
                BoldCell m.lRow, m.lCol, 0
                m.lRow = 0
                m.lCol = 0
                BoldCell m.lRow, m.lCol, 1
                
                gdsVert.Value = 0
                gdsHorz.Value = 0
                DrawBoard
            Else
                Col = 0
                gdsHorz.Value = 0
            End If
            
        Case vbKeyPageUp
            'visible rows is calculated above by lastVisibleRow - firstVisibleRow
            'when first and last visible row are the same lastVisibleRow - firstVisibleRow = 0
            'which really means one row is visible
            If lVisRows = 0 Then lVisRows = 1
            If m.lRow - lVisRows - 1 >= 0 Then
                Row = m.lRow - lVisRows - 1
                gdsVert.Value = Row         'gdsVert.Value - lVisRows
            Else
                Row = 0
                gdsVert.Value = 0
            End If
        
        Case vbKeyPageDown
            If lVisRows = 0 Then lVisRows = 1
            If m.lRow + lVisRows + 1 <= m.QuoteWinStruct.maxRow Then
                Row = m.lRow + lVisRows + 1
                gdsVert.Value = Row      '.Value + lVisRows
            Else
                Row = lLastRow
                gdsVert.Value = lLastRow
            End If
    
    End Select
    
ErrExit:
    m.bSkipScroll = False
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.KeyDown", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveString
'' Description: Compile a string of information to save for next load
'' Inputs:      None
'' Returns:     String to save
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SaveString() As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrSave As New cGdArray        ' Array of elements to save
    Dim SaveCell As New cQuoteCell      ' Cell of the quote board
    Dim strTemp As String               ' Temporary string

    For lIndex = 1 To m.Cells.Count
        Set SaveCell = m.Cells(lIndex)
        If SaveCell.SymbolID = 0 Then
            strTemp = SaveCell.Symbol
        Else
            strTemp = Str(SaveCell.SymbolID)
        End If
        strTemp = strTemp & ";" & Str(SaveCell.Row) & ";" & Str(SaveCell.Col)
        astrSave.Add strTemp
    Next lIndex
    
    SaveString = astrSave.JoinFields(",")

ErrExit:
    Set astrSave = Nothing
    Set SaveCell = Nothing
    Exit Function
    
ErrSection:
    Set astrSave = Nothing
    Set SaveCell = Nothing
    RaiseError "cQuoteCellBoard.SaveString", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BoardFromString
'' Description: Set up the quote board from the string passed in
'' Inputs:      String to Load From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BoardFromString(ByVal strInfo As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrInfo As New cGdArray        ' Array of elements to load
    Dim strSymbol As String             ' Symbol to add
    Dim lRow As Long                    ' Row of the cell to add to
    Dim lCol As Long                    ' Column of the cell to add to
    Dim strTemp As String               ' Temporary string
    
    astrInfo.SplitFields strInfo, ","
    
    ClearAll
    Row = 0
    Col = 0
    
    For lIndex = 0 To astrInfo.Size - 1
        strSymbol = Parse(astrInfo(lIndex), ";", 1)
        strTemp = Parse(astrInfo(lIndex), ";", 2)
        If Len(strTemp) = 0 Then lRow = -1& Else lRow = CLng(ValOfText(strTemp))
        strTemp = Parse(astrInfo(lIndex), ";", 3)
        If Len(strTemp) = 0 Then lCol = -1& Else lCol = CLng(ValOfText(strTemp))
        
        If ValOfText(strSymbol) > 0 Then
            AddSymbol CLng(ValOfText(strSymbol)), lRow, lCol
        Else
            AddSymbol strSymbol, lRow, lCol
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.BoardFromString", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSymbol
'' Description: Update the cells in the collection with the symbol in the bars
'' Inputs:      Bars for the Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateSymbol(Bars As cGdBars, Optional ByVal bOnlyIfVisible As Boolean = True)
On Error GoTo ErrSection:

    Dim astrCells As New cGdArray       ' List of cells to update
    Dim lPos As Long                    ' Position in the table
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLookup As String             ' Lookup into the table
    Dim bVisible As Boolean             ' Is the cell visible?
    
    strLookup = SymbolOrSymbolID(Bars.Prop(eBARS_SymbolID), Bars.Prop(eBARS_Symbol))
    If m.tblIndex.FieldArray(TblCol(eGDTblCol_Symbol), False).BinarySearch(strLookup, lPos) Then
        astrCells.SplitFields m.tblIndex(TblCol(eGDTblCol_Cells), lPos), ","
        For lIndex = 0 To astrCells.Size - 1
            bVisible = CellIsVisible(m.Cells(CLng(astrCells(lIndex))))
            m.Cells(CLng(astrCells(lIndex))).UpdateData Bars, m.bFullHeight, m.QuoteWinStruct.qwinStyle, m.QuoteWinStruct.nCompactQB, bVisible
            If bVisible Or (Not bOnlyIfVisible) Then
                m.Cells(CLng(astrCells(lIndex))).Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
            End If
        Next lIndex
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.UpdateSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupScroll
'' Description: Set up scroll bars as necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupScroll()
On Error GoTo ErrSection:

    If m.pbQuotes.Visible = False Then Exit Sub
    
    ' Save off the current values...
    If gdsVert.Visible Then m.lVertVal = gdsVert.Value
    If gdsHorz.Visible Then m.lHorzVal = gdsHorz.Value

    ' Show/Hide the scroll bars as necessary...
    SetVertScroll
    SetHorzScroll
    
    ' If both scroll bars are visible, color the bottom left corner...
    If gdsHorz.Visible And gdsVert.Visible Then
        m.QuoteWinStruct.ScrollBarColor = RGB(226, 226, 226)     'light, light gray
    Else
        m.QuoteWinStruct.ScrollBarColor = -1
    End If
    
    If m.lVertVal > 0 Then
        gdsVert.Value = m.lVertVal
        gdsVert_Change
    End If
    
    If m.lHorzVal > 0 Then
        gdsHorz.Value = m.lHorzVal
        gdsHorz_Change
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.SetupScroll", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveCell
'' Description: Remove a cell from the box
'' Inputs:      Row and Column of the cell to remove
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveCell(Optional ByVal lRow As Long = -1&, Optional ByVal lCol As Long = -1&)
On Error GoTo ErrSection:

    Dim CellToRemove As New cQuoteCell

    ' Default to the current cell...
    If lRow = -1& Then lRow = m.lRow
    If lCol = -1& Then lCol = m.lCol

    ' Clear out the cell and refresh...
    Set CellToRemove = Cell(lRow, lCol)
    If Not CellToRemove Is Nothing Then
        CellToRemove.ClearCell
        CellToRemove.Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
        
        m.Cells.Remove Key(lRow, lCol)
        Set CellToRemove = Nothing
        
        RebuildXref
        DrawBoard
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.RemoveCell", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextFreeCell
'' Description: Set the row and column to the next available cell
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NextFreeCell(Optional ByVal bStartFromCurrentCell As Boolean = True)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim bExit As Boolean                ' Can we exit from the outer for loop?
    Dim lStartRow As Long               ' Row to start with
    Dim lStartCol As Long               ' Column to start with
    
    If bStartFromCurrentCell Then
        lStartRow = Row
    Else
        lStartRow = 0
    End If
    
    bExit = False
    For lRow = lStartRow To m.QuoteWinStruct.maxRow
        If lRow = lStartRow Then
            lStartCol = Col
        Else
            lStartCol = 0
        End If
        
        For lCol = lStartCol To m.QuoteWinStruct.lastVisCol
            If Cell(lRow, lCol) Is Nothing Then
                Row = lRow
                Col = lCol
                bExit = True
                
                Exit For
            End If
        Next lCol
        
        If bExit Then Exit For
    Next lRow

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cQuoteCellBoard.NextFreeCell", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Private Subs & Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitBoard
'' Description: Initialize the Quote Board
'' Inputs:      Initialize the Window?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitBoard(ByVal bInitWin As Boolean)
On Error GoTo ErrSection:
        
    ' Initialize the window if necessary...
    If bInitWin Then geInitWin
    
    ' Clear out the collection of cells...
    m.Cells.Clear
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.InitBoard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    geInitWin
'' Description: Initialize the window for the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub geInitWin()
On Error GoTo ErrSection:
        
    Dim i&
    
    If m.hQuoteObj <> 0 Then Exit Sub

    ' These scroll bars dimensions should be set to zeros when not visible
    m.QuoteWinStruct.glhHorzScroll = gdCreateArray(eGDARRAY_Longs, 4, 0)
    m.QuoteWinStruct.glhVertScroll = gdCreateArray(eGDARRAY_Longs, 4, 0)
    
    ' Set up the font information...
    gdSetStr m.QuoteWinStruct.gshfName, 0, m.pbQuotes.Font.Name
    m.QuoteWinStruct.fStyle = 0
    m.QuoteWinStruct.fUnderline = m.pbQuotes.Font.Underline
    m.QuoteWinStruct.fSize = m.pbQuotes.Font.Size
    
    ' Set window dimensions...
    ResizeBoard
    
    ' Initialize the quote window...
    m.hQuoteObj = geInitQuoteWin(m.QuoteWinStruct, m.pbQuotes.hDC)
    If m.hQuoteObj <> 0 Then geEnableFeature m.hQuoteObj, 1
    
    m.bInitDone = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.geInitWin", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindFirstFreeCell
'' Description: Find the first free cell in the Quote Board
'' Inputs:      None
'' Returns:     Quote Cell object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindFirstFreeCell() As cQuoteCell
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim FreeCell As cQuoteCell          ' Quote Board Cell object
    Dim lEndIdx As Long
    
    Set FindFirstFreeCell = Nothing
    Set FreeCell = Nothing
    
    If g.bStarting Then
        'JM 01-12-2015: This code will only get hit during startup if a quoteboard INF file does not exist
        '   and we are loading pre-set default symbols. At startup the initial dimension of the pbQuote hwnd
        '   is not yet fully realized so the last visible col calculated by grapheng.dll will be inaccurate.
        '   Use maxCol instead to ensure symbols will all show in one row.
        '
        '   This code will not get hit if a quoteboard INF file exist because the cell's position is saved
        '   to the INF file and will be restored to the saved (row,col); i.e. there is no need to find
        '   first free cell when loading a saved INF file.
        lEndIdx = m.QuoteWinStruct.maxCol
    Else
        lEndIdx = m.QuoteWinStruct.lastVisCol
    End If
    
    For lRow = 0 To m.QuoteWinStruct.maxRow
        For lCol = 0 To lEndIdx
            If Cell(lRow, lCol) Is Nothing Then
                Set FreeCell = CreateCell(lRow, lCol)
                Exit For
            End If
        Next lCol
        
        If Not FreeCell Is Nothing Then Exit For
    Next lRow
    
    ' Return the cell...
    Set FindFirstFreeCell = FreeCell

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.FindFirstFreeCell", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAll
'' Description: Update all of the cells (mainly to reset the full height flag)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UpdateAll(Optional ByVal bOnlyIfVisible As Boolean = True)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bVisible As Boolean             ' Is the cell visible?
    
    For lIndex = 1 To m.Cells.Count
        bVisible = CellIsVisible(m.Cells(lIndex))
        m.Cells(lIndex).UpdateData m.Cells(lIndex).Bars, m.bFullHeight, m.QuoteWinStruct.qwinStyle, m.QuoteWinStruct.nCompactQB, bVisible
        If bVisible Or (Not bOnlyIfVisible) Then
            m.Cells(lIndex).Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.UpdateAll", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetVertScroll
'' Description: Determine where to place vertical scroll bar and if its visible
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetVertScroll()
On Error GoTo ErrSection:

'    Dim bLocked As Boolean
    
    If m.pbQuotes.ScaleHeight < 500 Then Exit Sub       '5031
    
'bLocked = LockWindowUpdate(m.pbQuotes.hWnd)
    If m.QuoteWinStruct.firstVisRow > 0 Or m.QuoteWinStruct.lastVisRow < m.QuoteWinStruct.LastDataRow Then
        With gdsVert
            If m.QuoteWinStruct.firstVisCol > 0 Or m.QuoteWinStruct.lastVisCol < m.QuoteWinStruct.LastDataCol Then
                'horz scroll visible
                .Move m.pbQuotes.ScaleWidth - .Width, m.pbQuotes.ScaleTop, .Width, m.pbQuotes.ScaleHeight - gdsHorz.Height
            Else
                'horz scroll not visible
                .Move m.pbQuotes.ScaleWidth - .Width, m.pbQuotes.ScaleTop, .Width, m.pbQuotes.ScaleHeight
            End If
            .ZOrder
            .SmallChange = 1
            .LargeChange = 1
            .Max = m.QuoteWinStruct.maxRow - 1
            .Min = 0
            '.Value = 0
            If .Value < 0 Then .Value = 0
            If .Value > .Max Then .Value = .Max
            If .Max > 0 Then        'fix for aardvark 1200
                .Visible = True
                ' Graphics engine needs this to determine what is visible in the last col...
                SetVertScrollDim .Top, .Left, .Width, .Height
            Else
                .Visible = False
                ' lets graphics engine know scroll is not on
                SetVertScrollDim 0, 0, 0, 0
            End If
        End With
    Else
        With gdsVert
            .Visible = False
            '.Min = 0
            '.Max = m.QuoteWinStruct.lastDataRow - (m.QuoteWinStruct.lastVisRow - m.QuoteWinStruct.firstVisRow)
        End With
        SetVertScrollDim 0, 0, 0, 0 'lets graphics engine know vertical scroll not on
        m.QuoteWinStruct.firstVisRow = 0
    End If
    
'If bLocked Then LockWindowUpdate 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.SetVertScroll", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetHorzScroll
'' Description: Determine where to place horizontal scroll bar and if its visible
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetHorzScroll()
On Error GoTo ErrSection:
    
'    Dim bLocked As Boolean
    
    If m.pbQuotes.Width < 500 Then Exit Sub         '5031

'bLocked = LockWindowUpdate(m.pbQuotes.hWnd)
    If m.QuoteWinStruct.firstVisCol > 0 Or m.QuoteWinStruct.lastVisCol < m.QuoteWinStruct.LastDataCol Then
        With gdsHorz
            If m.QuoteWinStruct.firstVisRow > 0 Or m.QuoteWinStruct.LastDataRow > m.QuoteWinStruct.lastVisRow Then
                'vert scroll visible
                .Move m.pbQuotes.ScaleLeft, m.pbQuotes.ScaleHeight - .Height, m.pbQuotes.ScaleWidth - gdsVert.Width
            Else
                'vert scroll not visible
                .Move m.pbQuotes.ScaleLeft, m.pbQuotes.ScaleHeight - .Height, m.pbQuotes.ScaleWidth
            End If
            .ZOrder
            .SmallChange = 1
            .LargeChange = 1
            .Max = m.QuoteWinStruct.maxCol - 1
            .Min = 0
            If .Value < 0 Then .Value = 0
            If .Value > .Max Then .Value = .Max
            If .Max > 0 Then        'fix for aardvark 1200
                .Visible = True
                ' Graphics engine needs this to determine what is visible in the last col...
                SetHorzScrollDim .Top, .Left, .Width, .Height
            Else
                .Visible = False
                ' lets graphics engine know scroll is not on
                SetHorzScrollDim 0, 0, 0, 0
            End If
        End With
    Else
        With gdsHorz
            .Visible = False
            '.Min = 0
            '.Max = m.QuoteWinStruct.lastDataCol - (m.QuoteWinStruct.lastVisCol - m.QuoteWinStruct.firstVisCol)
        End With
        SetHorzScrollDim 0, 0, 0, 0 'lets graphics engine know scroll is not on
        m.QuoteWinStruct.firstVisCol = 0
    End If
    
'If bLocked Then LockWindowUpdate 0
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.SetHorzScroll", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdsHorz_Change
'' Description: Handle a change in the horizontal scroll bar
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdsHorz_Change()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    Static nPrev As Long
    
    Dim nMaxFirstCol&       'the largest value that the first visible col can be

    If Not bInProgress Then
        bInProgress = True
        
        nMaxFirstCol = m.QuoteWinStruct.maxCol - (m.QuoteWinStruct.lastVisCol - m.QuoteWinStruct.firstVisCol)
        nMaxFirstCol = nMaxFirstCol - 1     'col is zero-based in graphics engine
        
        With gdsHorz
            If .Value > .Max Then
                .Value = .Max
            ElseIf .Value < .Min Then
                .Value = .Min
            ElseIf .Value <> .Max And (.Value > m.QuoteWinStruct.lastVisCol Or .Value < m.QuoteWinStruct.firstVisCol) Then
                If .Value <= nMaxFirstCol Then
                    m.QuoteWinStruct.firstVisCol = .Value
                Else
                    m.QuoteWinStruct.firstVisCol = nMaxFirstCol
                End If
                DrawBoard
                UpdateAll
            ElseIf m.bSkipScroll Then
                If .Value > m.QuoteWinStruct.lastVisCol Then
                    m.QuoteWinStruct.firstVisCol = nMaxFirstCol
                    DrawBoard
                    UpdateAll
                End If
            Else
                If Abs(.Value - nPrev) = 1 Then
                    If .Value > nPrev Then
                        If m.QuoteWinStruct.firstVisCol + 1 > .Value Then
                            .Value = m.QuoteWinStruct.firstVisCol + 1
                        End If
                    ElseIf m.QuoteWinStruct.firstVisCol > 0 Then
                        .Value = m.QuoteWinStruct.firstVisCol - 1
                    Else
                        .Value = 0
                    End If
                End If
                If .Value <= nMaxFirstCol Then
                    If .Value > m.QuoteWinStruct.lastVisCol Then
                        m.QuoteWinStruct.firstVisCol = .Value
                    End If
                Else
                    m.QuoteWinStruct.firstVisCol = nMaxFirstCol
                    .Value = .Max
                End If
                DrawBoard
                UpdateAll
            End If
            nPrev = .Value
        End With
    
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.gdsHorz.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdsHorz_GotFocus
'' Description: Move the focus back to the picture box when we get the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdsHorz_GotFocus()
On Error GoTo ErrSection:

    MoveFocus m.pbQuotes
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.gdsHorz.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdsVert_Change
'' Description: Handle a change in the vertical scroll bar
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdsVert_Change()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    Static nPrev As Long
    
    Dim nMaxFirstRow As Long        'the largest value that the first visible row can be

    If Not bInProgress Then
        bInProgress = True
        
        nMaxFirstRow = m.QuoteWinStruct.maxRow - (m.QuoteWinStruct.lastVisRow - m.QuoteWinStruct.firstVisRow)
        nMaxFirstRow = nMaxFirstRow - 1     'row is zero-based in graphics engine
        
        With gdsVert
            If .Value > .Max Then
                .Value = .Max
            ElseIf .Value < .Min Then
                .Value = .Min
            ElseIf .Value <> .Max And (.Value > m.QuoteWinStruct.lastVisRow Or .Value < m.QuoteWinStruct.firstVisRow) Then
                If .Value <= nMaxFirstRow Then
                    m.QuoteWinStruct.firstVisRow = .Value
                Else
                    m.QuoteWinStruct.firstVisRow = nMaxFirstRow
                End If
                DrawBoard
                UpdateAll
            ElseIf m.bSkipScroll Then
                If .Value > nMaxFirstRow Then
                    m.QuoteWinStruct.firstVisRow = nMaxFirstRow
                    DrawBoard
                    UpdateAll
                End If
            Else
                If Abs(.Value - nPrev) = 1 Then
                    If .Value > nPrev Then
                        m.QuoteWinStruct.firstVisRow = m.QuoteWinStruct.firstVisRow + 1
                    ElseIf m.QuoteWinStruct.firstVisRow > 0 Then
                        m.QuoteWinStruct.firstVisRow = m.QuoteWinStruct.firstVisRow - 1
                    Else
                        .Value = 0
                    End If
                End If
                If m.QuoteWinStruct.firstVisRow >= nMaxFirstRow Or .Value = .Max Then
                    m.QuoteWinStruct.firstVisRow = nMaxFirstRow
                End If
                DrawBoard
                UpdateAll
            End If
            nPrev = .Value
        End With
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.gdsVert.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdsVert_GotFocus
'' Description: Move the focus back to the picture box when we get the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdsVert_GotFocus()
On Error GoTo ErrSection:

    MoveFocus m.pbQuotes
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.gdsVert.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetHorzScrollDim
'' Description: Set the dimensions for the horizontal scroll bar
'' Inputs:      Top, Left, Width, and Height of the scroll bar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetHorzScrollDim(ByVal lTop&, ByVal lLeft&, ByVal lWidth&, ByVal lHeight&)
On Error GoTo ErrSection:

    With m.QuoteWinStruct
        gdSetNum .glhHorzScroll, 0, lTop / Screen.TwipsPerPixelY
        gdSetNum .glhHorzScroll, 1, lLeft / Screen.TwipsPerPixelX
        gdSetNum .glhHorzScroll, 2, lHeight / Screen.TwipsPerPixelY
        gdSetNum .glhHorzScroll, 3, lWidth / Screen.TwipsPerPixelX
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.SetHorzScrollDim", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetVertScrollDim
'' Description: Set the dimensions for the vertical scroll bar
'' Inputs:      Top, Left, Width, and Height of the scroll bar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetVertScrollDim(ByVal lTop&, ByVal lLeft&, ByVal lWidth&, ByVal lHeight&)
On Error GoTo ErrSection:

    With m.QuoteWinStruct
        gdSetNum .glhVertScroll, 0, lTop / Screen.TwipsPerPixelY
        gdSetNum .glhVertScroll, 1, lLeft / Screen.TwipsPerPixelX
        gdSetNum .glhVertScroll, 2, lHeight / Screen.TwipsPerPixelY
        gdSetNum .glhVertScroll, 3, lWidth / Screen.TwipsPerPixelX
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.SetVertScrollDim", eGDRaiseError_Raise
    
End Sub

Private Function Key(ByVal lRow As Long, ByVal lCol As Long) As String
    Key = Str(lRow) & ";" & Str(lCol)
End Function

Private Function CreateCell(ByVal lRow As Long, ByVal lCol As Long) As cQuoteCell
On Error GoTo ErrSection:

'    Dim lSameColorFlag As Long
'
'    If m.QuoteWinStruct.qwinStyle = eGDQuoteStyle_Forex Then
'        lSameColorFlag = 0
'    Else
'        lSameColorFlag = m.lSameColorText
'    End If

    Set CreateCell = New cQuoteCell
    
    CreateCell.Init Me, m.hQuoteObj, lRow, lCol, _
            m.QuoteWinStruct.netUpColor, m.QuoteWinStruct.netDownColor, _
            m.QuoteWinStruct.netUnchColor, m.lUpdateColor, m.lUseUpdateColor, m.lColorSymbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.CreateCell", eGDRaiseError_Raise
    
End Function

Private Function AddCell(CellToAdd As cQuoteCell) As Long
On Error GoTo ErrSection:
    
    CellToAdd.ID = m.Cells.Count + 1
    m.Cells.Add CellToAdd, Key(CellToAdd.Row, CellToAdd.Col)
    AddCell = CellToAdd.ID

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.AddCell", eGDRaiseError_Raise
    
End Function

Private Sub Reset()
On Error GoTo ErrSection:

    Dim strCurrent As String            ' Current state of the cell structure

    ' Save off current cell structure...
    strCurrent = SaveString

    ' Close and reinitialize the quote board...
    If m.hQuoteObj <> 0 Then
        geCloseQuoteWin m.hQuoteObj
        m.hQuoteObj = 0
        ResizeBoard
    End If
    
    If Not m.pbQuotes Is Nothing Then
        m.hQuoteObj = geInitQuoteWin(m.QuoteWinStruct, m.pbQuotes.hDC)
        If m.hQuoteObj <> 0 Then geEnableFeature m.hQuoteObj, 1
    End If
    
    ' If successfully reinitialized, load the cell structure back in...
    If m.hQuoteObj <> 0 Then
        BoardFromString strCurrent
        DrawBoard
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.Reset", eGDRaiseError_Raise
    
End Sub

Private Sub RebuildXref()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the table to update
    Dim strLookup As String             ' Lookup into the table
    
    ' Clear out the cells field of the table...
    m.tblIndex.NumRecords = 0&
    
    ' Walk through the cells collection, reset the index, and refill the cells field...
    For lIndex = 1 To m.Cells.Count
        If m.Cells(lIndex).ID <> lIndex Then
            m.Cells(lIndex).ID = lIndex
            m.Cells(lIndex).Update
        End If
        
        strLookup = SymbolOrSymbolID(m.Cells(lIndex).SymbolID, m.Cells(lIndex).Symbol)
        If m.tblIndex.FieldArray(TblCol(eGDTblCol_Symbol), False).BinarySearch(strLookup, lPos) = False Then
            m.tblIndex.AddRecord strLookup & vbTab & Str(m.Cells(lIndex).ID), lPos, vbTab
        Else
            m.tblIndex(TblCol(eGDTblCol_Cells), lPos) = m.tblIndex(TblCol(eGDTblCol_Cells), lPos) & "," & Str(m.Cells(lIndex).ID)
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.RebuildXRef", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated color of a price
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.Cells.Count
        If m.Cells(lIndex).ClearUpdatedColors Then
            m.Cells(lIndex).Refresh m.pbQuotes.hWnd, m.pbQuotes.hDC
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.ClearUpdatedColors", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolOrSymbolID
'' Description: Figure out the key into the index table
'' Inputs:      Symbol ID, Symbol
'' Returns:     Symbol ID if non-zero, Symbol otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolOrSymbolID(ByVal lSymbolID As Long, ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    If lSymbolID <> 0& Then
        SymbolOrSymbolID = Str(lSymbolID)
    Else
        SymbolOrSymbolID = Parse(strSymbol, "(", 1)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.SymbolOrSymbolID", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CellIsVisible
'' Description: Is the cell passed in visible right now?
'' Inputs:      Cell to Check
'' Returns:     True if Visible, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CellIsVisible(ByVal Cell As cQuoteCell, Optional ByVal bIncludePartiallyVisible As Boolean = True) As Boolean
On Error GoTo ErrSection:

    If Cell.Row >= m.QuoteWinStruct.firstVisRow And Cell.Row <= m.QuoteWinStruct.lastVisRow + Abs(bIncludePartiallyVisible) Then
        If Cell.Col >= m.QuoteWinStruct.firstVisCol And Cell.Col <= m.QuoteWinStruct.lastVisCol + Abs(bIncludePartiallyVisible) Then
            CellIsVisible = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.CellIsVisible", eGDRaiseError_Raise
    
End Function

Public Property Get MaxRows() As Long
    MaxRows = m.QuoteWinStruct.maxRow
End Property

Public Property Get MaxCols() As Long
    MaxCols = m.QuoteWinStruct.maxCol
End Property

Public Property Get ForexTextColor() As Long
    ForexTextColor = gdGetNum(m.QuoteWinStruct.glhTextColors, 4)
End Property

Public Property Let ForexTextColor(ByVal nColor&)
    gdSetNum m.QuoteWinStruct.glhTextColors, 4, nColor
End Property

Public Property Get ForexBkColor() As Long
    ForexBkColor = gdGetNum(m.QuoteWinStruct.glhTextColors, 0)
End Property

Public Property Let ForexBkColor(ByVal nColor&)
    gdSetNum m.QuoteWinStruct.glhTextColors, 0, nColor
End Property

Public Property Get ForexUpDownBk() As Long
    ForexUpDownBk = m.QuoteWinStruct.nForexUpDownBk
End Property

Public Property Let ForexUpDownBk(ByVal lFlag&)
    m.QuoteWinStruct.nForexUpDownBk = lFlag
End Property

Public Function SortSymbols(ByVal strInfo$, ByVal eOrder As eGdSort_Flags) As String
On Error GoTo ErrSection:

    'info string format: symId;row;col,symId;row;col etc.

    Dim i&, j&, nID&
    Dim strSymbol$
    
    Dim tbSymInfo As New cGdTable
    Dim aInput As New cGdArray
    Dim aIndex As cGdArray

    tbSymInfo.CreateField eGDARRAY_Strings, 0, "SymName"
    tbSymInfo.CreateField eGDARRAY_Longs, 1, "SymID", -1
    tbSymInfo.CreateField eGDARRAY_Longs, 2, "CellRow", -1
    tbSymInfo.CreateField eGDARRAY_Longs, 3, "CellCol", -1
    
    'populate table with symbol name & ID from input string
    aInput.SplitFields strInfo, ","
    For i = 0 To aInput.Size - 1
        nID = ValOfText(Parse(aInput(i), ";", 1))
        strSymbol = GetSymbol(nID)
        If Len(strSymbol) > 0 Then
            tbSymInfo.AddRecord ""
            j = tbSymInfo.NumRecords - 1
            tbSymInfo(0, j) = strSymbol
            tbSymInfo(1, j) = nID
        End If
    Next
    
    'sort table by symbol name based on passed in sort flag
    If tbSymInfo.NumRecords > 0 Then
        Set aIndex = tbSymInfo.CreateSortedIndex(0, eOrder)
    Else
        Set aInput = Nothing
        Set tbSymInfo = Nothing
        Exit Function                   'precautionary
    End If
    
    'precautionary (theoretically should never happen)
    If aIndex Is Nothing Then
        Set aInput = Nothing
        Set tbSymInfo = Nothing
        Exit Function
    End If
    
    If aIndex.Size <> tbSymInfo.NumRecords Then
        Set aInput = Nothing
        Set aIndex = Nothing
        Set tbSymInfo = Nothing
        Exit Function
    End If
    
    Dim nVisibleCells&, nCols&, nRows&
    'determine if visible rows & cols are sufficient to show all symbols (adjust as necessary)
    nCols = m.QuoteWinStruct.lastVisCol + 1
    nRows = m.QuoteWinStruct.lastVisRow + 1
    nVisibleCells = nCols * nRows
    j = tbSymInfo.NumRecords
    While nVisibleCells < j
        If nRows + 1 < m.QuoteWinStruct.maxRow Then
            nRows = nRows + 1
            nVisibleCells = nCols * nRows
        ElseIf nCols + 1 <= m.QuoteWinStruct.maxCol Then
            nCols = nCols + 1
            nVisibleCells = nCols * nRows
        Else
            nRows = nRows + 1
            nVisibleCells = nCols * nRows
        End If
    Wend
    
    Dim iRow&, iCol&
    Dim bAddToString As Boolean
    'populate row/col values and rebuild input string with new cell location info
    aInput.Clear
    j = tbSymInfo.NumRecords - 1
    iRow = 0
    iCol = 0
    For i = 0 To j
        bAddToString = True
        nID = tbSymInfo(1, aIndex(i))
        tbSymInfo(2, aIndex(i)) = iRow
        tbSymInfo(3, aIndex(i)) = iCol
        If iCol + 1 < nCols Then
            iCol = iCol + 1
        ElseIf iRow + 1 <= nRows Then
            iRow = iRow + 1
            iCol = 0
        Else
            'theoretically should never get here, but if does get here then do nothing
            bAddToString = False
        End If
        
        If bAddToString Then
            aInput.Add Str(nID) & ";" & Str(tbSymInfo(2, aIndex(i))) & ";" & Str(tbSymInfo(3, aIndex(i)))
        End If
    Next
        
    strSymbol = aInput.JoinFields(",")
    
    Set aInput = Nothing
    Set aIndex = Nothing
    Set tbSymInfo = Nothing
        
    SortSymbols = Trim(strSymbol)
    
    Exit Function
    
ErrSection:
    Set aInput = Nothing
    Set aIndex = Nothing
    Set tbSymInfo = Nothing
    
    RaiseError "cQuoteCellBoard.SortSymbols"
        
End Function

Public Sub CopyProperties(qbSource As cQuoteCellBoard)
On Error GoTo ErrSection:

    m.QuoteWinStruct.qwinStyle = qbSource.QuoteBoardStyle
    m.QuoteWinStruct.netDownColor = qbSource.DownColor
    m.QuoteWinStruct.netUpColor = qbSource.UpColor
    m.QuoteWinStruct.netUnchColor = qbSource.UnchColor
    m.QuoteWinStruct.qwinStyle = qbSource.QuoteBoardStyle
    m.lUseUpdateColor = qbSource.UseUpdateColor
    m.lColorSymbol = qbSource.ColorSymbol
    Me.Font = qbSource.Font
    
    m.bFullHeight = qbSource.FullBarHeight
    m.lUpdateColor = qbSource.UpdateColor
    
    m.lCompactBox = qbSource.CompactQB(eGDQuoteStyle_OHLC)
    m.lCompactForex = qbSource.CompactQB(eGDQuoteStyle_Forex)
    
    If m.QuoteWinStruct.qwinStyle = eGDQuoteStyle_Forex Then        '4625
        m.QuoteWinStruct.nCompactQB = m.lCompactForex
    Else
        m.QuoteWinStruct.nCompactQB = m.lCompactBox
    End If
    
    m.QuoteWinStruct.nShowExtraInfo = qbSource.ShowExtraInfo
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.CopyProperties", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindCellBySymbol
'' Description: Find a cell in the collection by row and column
'' Inputs:      Row, Column, ID
'' Returns:     Quote Cell object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindCellBySymbol(Optional ByVal strSymbol As String = "", _
    Optional ByVal nSymbolID As Long = -1) As cQuoteCell
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim Cell As cQuoteCell              ' Quote Board Cell object
    Dim lID As Long                     ' ID of the cell
    
    ' Do some initializing...
    Set Cell = Nothing
    
    If nSymbolID <= 0 And Len(strSymbol) = 0 Then Exit Function
    
    ' Walk through the cells in the collection until cell with matching symbol is found...
    For lIndex = 1 To m.Cells.Count
        Set Cell = m.Cells(lIndex)
        If Not Cell Is Nothing Then
            If Cell.SymbolID = nSymbolID Or Cell.Symbol = strSymbol Then
                Cell.ID = lIndex
                Exit For
            End If
        End If
    Next
    
    ' Return the ID and the Quote Cell object...
    If Cell.ID = -1 Then Set Cell = Nothing
    Set FindCellBySymbol = Cell

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.FindCellBySymbol", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BoardToString
'' Description: store board info to string format
'' Inputs:      None
'' Returns:     string that should work for saving to
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BoardToString() As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim Cell As cQuoteCell              ' Quote Board Cell object
    Dim strSymPeriod


    For lIndex = 1 To m.Cells.Count
        Set Cell = m.Cells(lIndex)
        If Not Cell Is Nothing Then
            strSymPeriod = strSymPeriod & Cell.SymbolID & ";0,"
        End If
    Next

    BoardToString = strSymPeriod
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.BoardToString", eGDRaiseError_Raise

End Function

Public Property Get UseUpdateColor() As Long
    UseUpdateColor = m.lUseUpdateColor
End Property

Public Property Let UseUpdateColor(ByVal lFlag As Long)
    m.lUseUpdateColor = lFlag
End Property

Public Property Get ColorSymbol() As Long
    ColorSymbol = m.lColorSymbol
End Property

Public Property Let ColorSymbol(ByVal lFlag As Long)
    m.lColorSymbol = lFlag
End Property

Public Property Get ShowExtraInfo() As eGEQbExtraInfo
    ShowExtraInfo = m.QuoteWinStruct.nShowExtraInfo
End Property

Public Property Let ShowExtraInfo(ByVal lFlag As eGEQbExtraInfo)
    m.QuoteWinStruct.nShowExtraInfo = lFlag
End Property

Public Property Get CompactQB(ByVal eStyle As eGDQuoteStyle) As Long

    If eStyle = eGDQuoteStyle_Forex Then
        CompactQB = m.lCompactForex
    Else
        CompactQB = m.lCompactBox
    End If

End Property

Public Property Let CompactQB(ByVal eStyle As eGDQuoteStyle, ByVal lFlag&)

    If eStyle = eGDQuoteStyle_Forex Then
        m.lCompactForex = lFlag
    Else
        m.lCompactBox = lFlag
    End If

End Property

Private Sub SetDefaults()
On Error GoTo ErrSection:

    ' Set default values for window...
    m.QuoteWinStruct.FgColor = RGB(160, 160, 160)
    m.QuoteWinStruct.bkColor = vbWhite
    m.QuoteWinStruct.ScrollBarColor = -1
    m.QuoteWinStruct.boldColor = 0
    m.QuoteWinStruct.qwinStyle = eGDQuoteStyle_Thermometer

    ' default colors for forex style
    If gdGetNum(m.QuoteWinStruct.glhTextColors, 0) < 0 Then
        gdSetNum m.QuoteWinStruct.glhTextColors, 0, RGB(128, 128, 128)  'dark gray (color around symbol name)
    End If
    gdSetNum m.QuoteWinStruct.glhTextColors, 1, RGB(219, 219, 219)      'light gray (color around bid/ask, high/low labels)
    gdSetNum m.QuoteWinStruct.glhTextColors, 2, vbWhite                 'background color around bid/ask values
    gdSetNum m.QuoteWinStruct.glhTextColors, 3, vbBlack                 'label colors
    If gdGetNum(m.QuoteWinStruct.glhTextColors, 4) < 0 Then
        gdSetNum m.QuoteWinStruct.glhTextColors, 4, vbWhite             'text color for symbol & time stamp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteCellBoard.SetDefaults", eGDRaiseError_Raise
    
End Sub

Public Function CellSymbols(ByVal bConvertContinuousToIndividual As Boolean, Optional ByVal bJustVisible As Boolean = False) As cGdArray
On Error GoTo ErrSection:

    Dim i&, strSymbol$, bAdd As Boolean
    Dim aSymbols As New cGdArray
    Dim QuoteCell As cQuoteCell

    aSymbols.Create eGDARRAY_Strings
    If Not m.Cells Is Nothing Then
        For i = 1 To m.Cells.Count
            Set QuoteCell = m.Cells(i)
            If Not QuoteCell Is Nothing Then
                If bJustVisible Then
                    bAdd = CellIsVisible(QuoteCell, False)
                Else
                    bAdd = True
                End If
                If bAdd Then
                    strSymbol = QuoteCell.Symbol
                    If Len(strSymbol) > 0 Then
                        If bConvertContinuousToIndividual Then
                            strSymbol = RollSymbolForDate(strSymbol)
                        End If
                        aSymbols.Add strSymbol
                    End If
                End If
            End If
        Next
        aSymbols.Sort eGdSort_DeleteDuplicates
    End If

ErrExit:
    Set CellSymbols = aSymbols
    Exit Function
    
ErrSection:
    RaiseError "cQuoteCellBoard.CellSymbols"
End Function

Public Sub BoxQbAlertInit()
On Error GoTo ErrSection:

    Dim i&
    Dim strMyTabName$, strText$
    
    Dim frm As Form
    Dim Alert As cAlert
    Dim QbCell As cQuoteCell
    Dim bUpdateCell As Boolean
    
    
    gdSetSize m.QuoteWinStruct.gshTabAlerts, 0, 0
    
'JM 02-18-2010 possible implementation for tab alert message on box-style QB
'    If Not m.pbQuotes Is Nothing Then
'        'set QB tab alerts text for grapheng.dll
'        Set frm = m.pbQuotes.Parent
'        If TypeOf frm Is frmQuotes Then
'            strMyTabName = frmQuotes.CurrentTabName(frmQuotes.CurrentTabNum(True))
'        ElseIf TypeOf frm Is frmDetachedQBTab Then
'            strMyTabName = frmQuotes.CurrentTabName(frm.MyTabIndex)
'        End If
'
'        If Len(strMyTabName) > 0 Then
'            For i = 1 To g.Alerts.Count
'                Set Alert = g.Alerts(i)
'                If Not Alert Is Nothing Then
'                    If Alert.TabName = strMyTabName Then
'                        strText = "Convert to grid to view tab alert(s). ;" & Str(Abs(Alert.Active)) & ";" & Alert.AlertKey
'                        'strCondition = strCondition & ")  ;" & Str(Abs(Alert.Active)) & ";" & Alert.AlertKey
'                        gdSetStr m.QuoteWinStruct.gshTabAlerts, 0, strText
'                        Exit For
'                    End If
'                End If
'            Next
'        End If
'    End If
    
    For i = 1 To m.Cells.Count
        Set QbCell = m.Cells(i)
        If Not QbCell Is Nothing Then
            QbCell.QbCellAlertInit
        End If
    Next

'    If gdGetSize(m.QuoteWinStruct.gshTabAlerts) > 0 Then
'        geRecalcCells m.hQuoteObj, m.pbQuotes.hDC
'    End If


ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cQuoteCellBoard.BoxQbAlertInit"

End Sub

Public Sub BoxQbAlertUpdate(oAlert As cAlert, bRemove As Boolean, bAdd As Boolean)
On Error GoTo ErrSection:

    Dim i&, iAddRemove&
    Dim QbCell As cQuoteCell
    
    If Not oAlert Is Nothing Then
        If bRemove Then
            iAddRemove = 0
        ElseIf bAdd Then
            iAddRemove = 1
        Else
            iAddRemove = -1
        End If
        
        If iAddRemove <> -1 Then
            For i = 1 To m.Cells.Count
                Set QbCell = m.Cells(i)
                If Not QbCell Is Nothing Then
                    QbCell.QbCellAlertUpdate oAlert, iAddRemove
                End If
            Next
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cQuoteCellBoard.BoxQbAlertUpdate"

End Sub

Private Function LastDataColForRow(ByVal lRow&) As Long
On Error GoTo ErrSection:
    
    Dim lCol&
    
    If lRow >= 0 And lRow < m.QuoteWinStruct.maxRow Then
    
        For lCol = m.QuoteWinStruct.LastDataCol To 0 Step -1
            If Not Cell(lRow, lCol) Is Nothing Then Exit For
        Next
    
        If lCol >= 0 And lCol < m.QuoteWinStruct.LastDataCol Then
            If Cell(lRow, lCol) Is Nothing Then lCol = 0
        End If
    End If
    
    If lCol < 0 Then
        lCol = 0
    ElseIf lCol > m.QuoteWinStruct.LastDataCol Then
        lCol = m.QuoteWinStruct.LastDataCol
    End If
    
    LastDataColForRow = lCol
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cQuoteCellBoard.LastDataColForRow"

End Function
