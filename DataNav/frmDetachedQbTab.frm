VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmDetachedQBTab 
   ClientHeight    =   6660
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6945
      Top             =   5115
   End
   Begin VB.PictureBox pbQuoteBoard 
      AutoRedraw      =   -1  'True
      Height          =   3930
      Left            =   90
      ScaleHeight     =   3870
      ScaleWidth      =   7830
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   7890
      Begin gdOCX.gdScrollBar gdScrollHz 
         Height          =   255
         Left            =   4320
         TabIndex        =   2
         Top             =   3600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         Horizontal      =   -1  'True
      End
      Begin gdOCX.gdScrollBar gdScrollVt 
         Height          =   1335
         Left            =   7560
         TabIndex        =   3
         Top             =   2250
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2355
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgQuotes 
      Height          =   1290
      Left            =   2475
      TabIndex        =   0
      Top             =   4635
      Width           =   2880
      _cx             =   5080
      _cy             =   2275
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmDetachedQBTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    QB As cQuoteCellBoard
    
    eStyle As eGDQuoteStyle
    
    strSymbols As String
    strRowPeriod As String      'for detecting change in bar period

    iTabIndex As Long           'index into tab info table in form quotes
    iMouseRow As Long
    iMouseCol As Long
    
    iSymbolCol As Long
    iPeriodCol As Long
    
    bRemoveTab As Boolean
End Type

Private m As mPrivate

Public Sub ShowMe(fgSource As VSFlexGrid, qbSource As cQuoteCellBoard, ByVal iTabIdx&)
On Error GoTo ErrSection:

    Dim strText$

    Me.Caption = frmQuotes.TabStr(eGDTabSettings_Name, iTabIdx)
    
    m.iTabIndex = iTabIdx
    m.eStyle = ValOfText(frmQuotes.TabStr(eGDTabSettings_Style, iTabIdx))
    m.strSymbols = frmQuotes.TabStr(eGDTabSettings_Symbols, iTabIdx)
    
    If Not fgSource Is Nothing Then
        ' Initialize the grid...
        fgSource.SaveGrid "TempGrid.txt", flexFileAll
        fgQuotes.LoadGrid "TempGrid.txt", flexFileAll
        KillFile "TempGrid.txt", True
    End If
    
    ' Initialize the box-style board
    QuoteCellBoardReset qbSource, m.strSymbols
            
    strText = GetIniFileProperty(Me.Name & Me.Caption, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If
    
    m.iSymbolCol = frmQuotes.SymbolCol
    m.iPeriodCol = frmQuotes.PeriodCol
    
    ShowForm Me
    
    UpdateStyle True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.ShowMe"

End Sub

Private Sub fgQuotes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strOld As String
    
    If m.strRowPeriod = "InProgress" Then Exit Sub
    
    With fgQuotes
        If Col > 0 And Col < .Cols Then
            If Row >= .FixedRows And Row < .Rows Then
                .Row = Row
                .Col = Col
                If .MergeRow(.Row) Then
                    frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_AddLabel
                ElseIf .Col = m.iPeriodCol Then
                    strOld = Parse(m.strRowPeriod, ";", 2)
                    m.strRowPeriod = "InProgress"
                    frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_ChangePeriod, strOld
                    m.strRowPeriod = ""
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_AfterEdit"

End Sub

Private Sub fgQuotes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strText$

    If m.strRowPeriod = "InProgress" Then Exit Sub
    
    m.strRowPeriod = ""
    
    With fgQuotes
        If Col > 0 And Col < .Cols And Row >= .FixedRows And Row < .Rows Then
            .Row = Row
            .Col = Col
            strText = UCase(.TextMatrix(0, Col))
            If .MergeRow(.Row) = True Then
                .ComboList = ""
                If .Col = m.iSymbolCol Then Cancel = True
            ElseIf .Col = m.iSymbolCol Then
                .ComboList = "..."
            ElseIf .Col = m.iPeriodCol Then
                strText = .TextMatrix(.Row, .Col)
                If Len(strText) > 0 Then
                    m.strRowPeriod = Str(.Row) & ";" & strText
                    .ComboList = "|Daily|60 Minute|30 Minute|15 Minute|10 Minute|5 Minute"
                End If
            Else
                .ComboList = ""
                Cancel = True
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_BeforeEdit"

End Sub

Private Sub fgQuotes_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    ' Keep fixed columns where they are ...
    If Not frmQuotes.CanMoveCol(Col) Then Position = Col        '4391

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_BeforeMoveColumn"
    Resume ErrExit

End Sub

Private Sub fgQuotes_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    With fgQuotes
        If Row >= .FixedRows And Row < .Rows Then
            .Row = Row
            .Col = Col
            HandleSymAddOrChange
        End If
    End With
    
End Sub

Private Sub fgQuotes_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    FinishEdit = True
End Sub

Private Sub fgQuotes_DblClick()
On Error GoTo ErrSection:

    Dim nRow As Long                    ' Row user double clicked on

    ' Make sure that the user double clicked on a non-fixed row
    With fgQuotes
        nRow = .MouseRow
        If nRow >= .FixedRows Then
            If .MergeRow(nRow) = False Then
                .Row = nRow
                .Col = m.iSymbolCol
                SetActiveChartSymbol Parse(.TextMatrix(nRow, m.iSymbolCol), "(", 1)        '4390
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_DblClick"
    Resume ErrExit

End Sub

Private Sub fgQuotes_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If fgQuotes.Col = m.iSymbolCol Then
        Select Case Chr(KeyAscii)
            Case "A" To "Z", "a" To "z", "$" ', "#"         '6079
                HandleSymAddOrChange Chr(KeyAscii)
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_KeyPress"

End Sub

Private Sub fgQuotes_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If fgQuotes.EditWindow = 0 Then
        Select Case KeyCode
            Case vbKeyDelete
                frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_RemoveSymbol
            
            Case vbKeyInsert
                HandleSymAddOrChange ""         '6093
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_KeyUp"

End Sub

Private Sub fgQuotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim Point As POINTAPI
    
    m.iMouseRow = fgQuotes.MouseRow
    m.iMouseCol = fgQuotes.MouseCol
    
    If Button = vbRightButton Then
        
        With fgQuotes
            .Col = m.iMouseCol
            If m.iMouseRow >= .FixedRows And m.iMouseRow < .Rows Then
                .Row = m.iMouseRow
                Point.X = X / Screen.TwipsPerPixelX
                Point.Y = Y / Screen.TwipsPerPixelY
                ClientToScreen fgQuotes.hWnd, Point
                ScreenToClient frmQuotes.fgQuotes.hWnd, Point
                
                frmQuotes.ShowQuotesPopup Point.X * Screen.TwipsPerPixelX, Point.Y * Screen.TwipsPerPixelY, Me
                m.iMouseRow = 0
            End If
        End With
    
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_MouseDown"

End Sub

Private Sub fgQuotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim nRow As Long
    
    With fgQuotes
        If m.iMouseRow <> .MouseRow And m.iMouseRow >= .FixedRows And .MouseRow >= .FixedRows Then
            nRow = m.iMouseRow
            m.iMouseRow = 0
            .DragRow nRow               '5709
        End If
    End With
    
End Sub

Private Sub fgQuotes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m.iMouseRow = 0         '6075
End Sub

Private Sub fgQuotes_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    With fgQuotes       '4365
        If Row >= .FixedRows And Row < .Rows And Col > 0 And Col < .Cols Then
            .Row = Row
            .Col = Col
            If .MergeRow(.Row) = True Then
                If Len(.EditText) = 0 Then
                    .EditText = " "
                ElseIf Not frmQuotes.ValidLabel(fgQuotes.EditText) Then
                    Cancel = True
                End If
            ElseIf .Col = m.iPeriodCol Then
                .EditText = GetPeriodStr(.EditText)
                If GetPeriodicity(.EditText) > ePRD_Days Then
                    .EditText = "Daily"
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.fgQuotes_ValidateEdit"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If Not g.bUnloading Then
        If Not m.bRemoveTab Then frmQuotes.ReAttachTab Me
    End If
    
    If Not m.bRemoveTab Then
        SetIniFileProperty Me.Name & Me.Caption, GetFormPlacement(Me), "Placement", g.strIniFile
        frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_SaveTab
    End If
    
    If Not g.bUnloading Then m.QB.ClearAll
    Set m.QB = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.Form_QueryUnload"

End Sub

Private Sub Form_Resize()
On Error Resume Next

    fgQuotes.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    pbQuoteBoard.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub pbQuoteBoard_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    m.QB.KeyDown KeyCode, Shift
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.pbQuoteBoard_KeyDown"

End Sub

Private Sub pbQuoteBoard_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    HandleSymAddOrChange Chr(KeyAscii)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.pbQuoteBoard_KeyPress"

End Sub

Private Sub pbQuoteBoard_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_RemoveSymbol
    ElseIf KeyCode = vbKeyInsert Then
        HandleSymAddOrChange ""
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.pbQuoteBoard_KeyUp"

End Sub

Private Sub pbQuoteBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim Point As POINTAPI
    
    m.QB.MouseDown Button, Shift, X, Y
   
    If Button = vbRightButton Then
        m.QB.Row = m.QB.MouseRow
        m.QB.Col = m.QB.MouseCol
        
        Point.X = X / Screen.TwipsPerPixelX
        Point.Y = Y / Screen.TwipsPerPixelY
        ClientToScreen pbQuoteBoard.hWnd, Point
        ScreenToClient frmQuotes.fgQuotes.hWnd, Point
        
        frmQuotes.ShowQuotesPopup Point.X * Screen.TwipsPerPixelX, Point.Y * Screen.TwipsPerPixelY, Me
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.pbQuoteBoard.MouseDown", eGDRaiseError_Show

End Sub

Private Sub pbQuoteBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If m.QB.MouseUp(Button, Shift, X, Y) = 1& Then
        m.QB.DrawBoard
        pbQuoteBoard.Refresh
    End If

End Sub

Private Sub pbQuoteBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    m.QB.MouseMove Button, Shift, X, Y

End Sub

Private Sub pbQuoteBoard_Resize()
On Error Resume Next

    m.QB.ResizeBoard False
    m.QB.DrawBoard
    pbQuoteBoard.Refresh

End Sub

Public Function UpdateRT(Bars As cGdBars) As Long
On Error GoTo ErrSection:

    Dim strSym$, strPeriod$
    Dim i&, iFound&

    iFound = -1
    If Bars Is Nothing Then GoTo ErrExit
    
    strSym = Str(Bars.Prop(eBARS_SymbolID))
    strPeriod = Bars.Prop(eBARS_PeriodicityStr)
        
    If m.eStyle = eGDQuoteStyle_Grid Then
        With fgQuotes
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, 0) = strSym And .TextMatrix(i, 3) = strPeriod Then
                    iFound = i
                    Exit For
                End If
            Next
        End With
    ElseIf Not m.QB Is Nothing Then
        m.QB.UpdateSymbol Bars
    End If
    
ErrExit:
    UpdateRT = iFound
    Exit Function


ErrSection:
    RaiseError "frmDetachedQBTab.UpdateRT"
    Resume ErrExit

End Function

Public Property Get MySymbols(ByVal bRebuild As Boolean) As String
On Error GoTo ErrSection:

    Dim i&
    Dim strSymID$, strPeriod$

    If bRebuild Then
        If m.eStyle = eGDQuoteStyle_Grid Then
            m.strSymbols = ","
            With fgQuotes
                For i = .FixedRows To .Rows - 2
                    strSymID = .TextMatrix(i, 0)
                    strPeriod = .TextMatrix(i, 3)
                    If .MergeRow(i) = True Then
                        m.strSymbols = m.strSymbols & "Label;" & .TextMatrix(i, 3) & ","
                    ElseIf Len(strSymID) > 0 And Len(strPeriod) > 0 Then
                        m.strSymbols = m.strSymbols & strSymID & ";" & strPeriod & ","
                    End If
                Next
            End With
        Else
            m.strSymbols = m.QB.BoardToString
        End If
    End If
    
    MySymbols = m.strSymbols

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmDetachedQBTab.MySymbols.Get"

End Property

Public Property Get MyQBStyle() As eGDQuoteStyle
    MyQBStyle = m.eStyle
End Property

Public Property Get MyTabIndex() As Long
    MyTabIndex = m.iTabIndex
End Property

Public Property Let MyTabIndex(ByVal iIndex&)
    m.iTabIndex = iIndex
End Property

Public Property Get QuoteCellBoard() As cQuoteCellBoard
    Set QuoteCellBoard = m.QB
End Property

Public Sub DrawBoxQB()
On Error GoTo ErrSection:

    Dim strCells$, strNewSym$, lIndex&
    Dim BarsColl As cGdTree
    
    strNewSym = frmQuotes.TabStr(eGDTabSettings_Symbols, m.iTabIndex)
    If Len(strNewSym) > 0 And strNewSym <> m.strSymbols Then m.strSymbols = strNewSym
    
    strCells = m.strSymbols
    If Left(strCells, 1) = "," Then strCells = Mid(strCells, 2)
    If Right(strCells, 1) = "," Then strCells = Left(strCells, Len(strCells) - 1)
    m.QB.BoardFromString strCells
    
    Set BarsColl = frmQuotes.GetBarsTree
    
    If Not BarsColl Is Nothing Then
        For lIndex = 1 To BarsColl.Count
            m.QB.UpdateSymbol BarsColl(lIndex), False
        Next lIndex
    End If
    
    pbQuoteBoard_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.DrawBoxQB"
    
End Sub

Private Sub pbQuoteBoard_DblClick()
On Error GoTo ErrSection:

    If Not m.QB.Cell(m.QB.Row, m.QB.Col) Is Nothing Then
        SetActiveChartSymbol Parse(m.QB.Cell(m.QB.Row, m.QB.Col).Symbol, "(", 1)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.pbQuoteBoard.DblClick", eGDRaiseError_Show

End Sub

Private Sub pbQuoteBoard_GotFocus()
On Error GoTo ErrSection:
    
    m.QB.GotFocus
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDetachedQBTab.pbQuoteBoard.GotFocus", eGDRaiseError_Show
    
End Sub

Public Sub UpdateStyle(Optional ByVal bInitialShow As Boolean = False)
On Error GoTo ErrSection:

    Dim i&
    Dim eNewStyle As eGDQuoteStyle

    Dim bLocked As Boolean
    
    eNewStyle = ValOfText(frmQuotes.TabStr(eGDTabSettings_Style, m.iTabIndex))
    
    If m.eStyle <> eNewStyle Or bInitialShow Then
        m.eStyle = eNewStyle
        frmQuotes.TabStr(eGDTabSettings_Style, MyTabIndex) = Str(m.eStyle)
        If m.eStyle = eGDQuoteStyle_Grid Then
            fgQuotes.Visible = True
            pbQuoteBoard.Visible = False
            If g.bStarting Then
                tmr.Enabled = True
            Else
                g.Alerts.DisplayQBAlerts Me
            End If
        Else
            bLocked = LockWindowUpdate(Me.hWnd)
            fgQuotes.Visible = False
            m.QB.QuoteBoardStyle = m.eStyle
            pbQuoteBoard.Visible = True
            DrawBoxQB
            m.QB.BoxQbAlertInit
            m.QB.DrawBoard
            If bLocked Then LockWindowUpdate 0
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.UpdateStyle"

End Sub

Public Function FindInGrid(ByVal strSymbol$, ByVal strPeriod$) As Long
On Error GoTo ErrSection:

    Dim i&, iFound&
    Dim strSymInGrid$
    
    iFound = -1
    With fgQuotes
        For i = .FixedRows To .Rows - 1
            strSymInGrid = Parse(.TextMatrix(i, 2), "(", 1)
            If strSymInGrid = strSymbol Then
                If Len(strPeriod) = 0 Or (.TextMatrix(i, 3) = strPeriod) Then
                    iFound = i
                    Exit For
                End If
            End If
        Next
    End With
    
    FindInGrid = iFound
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmDetachedQBTab.FindInGride"

End Function

Private Sub QuoteCellBoardReset(qbSource As cQuoteCellBoard, ByVal strSymbols As String)
On Error GoTo ErrSection:

    Dim strCells$, i&
    
    Dim aSymbols As New cGdArray
    Dim Bars As cGdBars

    ' Initialize the quote cell board...
    Set m.QB = New cQuoteCellBoard
    m.QB.CopyProperties qbSource                        '4625
    m.QB.Init pbQuoteBoard, gdScrollHz, gdScrollVt
    m.QB.Font = qbSource.Font

    strCells = m.strSymbols
    If Left(strCells, 1) = "," Then strCells = Mid(strCells, 2)
    If Right(strCells, 1) = "," Then strCells = Left(strCells, Len(strCells) - 1)
    m.QB.BoardFromString strCells

    If m.eStyle <> eGDQuoteStyle_Grid Then
        'when call from UpdateSettings, don't switch between box & forex
        If m.eStyle = eGDQuoteStyle_Forex Or qbSource.QuoteBoardStyle = eGDQuoteStyle_Forex Then
            m.QB.QuoteBoardStyle = m.eStyle
        End If
        aSymbols.SplitFields m.strSymbols
        For i = 0 To aSymbols.Size - 1
            Set Bars = frmQuotes.GetBars(Parse(aSymbols(i), ";", 1), "Daily")
            If Not Bars Is Nothing Then
                m.QB.UpdateSymbol Bars, False
            End If
        Next
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.QuoteCellBoardReset"

End Sub

Public Sub UpdateSettings(qbSource As cQuoteCellBoard, fgSource As VSFlexGrid)
On Error GoTo ErrSection:

    Dim strSave$, i&, j&
    
    Dim lUpOld&, lDownOld&, lUnchOld&, lColorSymFlagOld&
    Dim lUpNew&, lDownNew&, lUnchNew&, lColorSymFlagNew&
    
    Dim lSymColumn As Long             'this needs to match eGDCol_Symbol in the quotes form
    Dim lSymColor As Long
                 
    lSymColumn = 2
    
    lUpOld = m.QB.UpColor
    lDownOld = m.QB.DownColor
    lUnchOld = m.QB.UnchColor
    lColorSymFlagOld = m.QB.ColorSymbol
    
    lUpNew = qbSource.UpColor
    lDownNew = qbSource.DownColor
    lUnchNew = qbSource.UnchColor
    lColorSymFlagNew = qbSource.ColorSymbol
    
    m.QB.CopyProperties qbSource
    m.QB.QuoteBoardStyle = m.eStyle     'restore this QB style
    
fgQuotes.Redraw = flexRDNone
    'set format for the quote grid...
    fgSource.SaveGrid "TempGrid.txt", flexFileFormat
    fgQuotes.LoadGrid "TempGrid.txt", flexFileFormat
    KillFile "TempGrid.txt"
    
    'walk through grid and set text color if changed
    If lUpOld <> lUpNew Or lDownOld <> lDownNew Or lColorSymFlagOld <> lColorSymFlagNew Or lUnchOld <> lUnchNew Then
        With fgQuotes
            For i = .FixedRows To .Rows - 1
                For j = lSymColumn + 1 To .Cols - 1
                    If Not .ColHidden(j) Then
                        If .Cell(flexcpForeColor, i, j) = lUpOld Then
                            .Cell(flexcpForeColor, i, j) = lUpNew
                            lSymColor = lUpNew
                        ElseIf .Cell(flexcpForeColor, i, j) = lDownOld Then
                            .Cell(flexcpForeColor, i, j) = lDownNew
                            lSymColor = lDownNew
                        ElseIf .Cell(flexcpForeColor, i, j) = lUnchOld Then
                            .Cell(flexcpForeColor, i, j) = lUnchNew
                            lSymColor = lUnchNew
                        End If
                        'If lColorSymFlagOld <> lColorSymFlagNew Then
                            If lColorSymFlagNew = 0 Or Not g.RealTime.Active Then
                                .Cell(flexcpForeColor, i, lSymColumn) = 0
                            Else
                                .Cell(flexcpForeColor, i, lSymColumn) = lSymColor
                            End If
                        'End If
                    End If
                Next
                DoEvents
            Next
        End With
    End If
fgQuotes.Redraw = flexRDBuffered

    'reset the quote cell board...
    strSave = m.QB.BoardToString
    
'JM 07-27-2012: original code causes detached QB cells to get messed up; leave awhile then remove if all ok
'    QuoteCellBoardReset qbSource, strSave
        
'JM 07-27-2012: just redraw on settings update so cells don't get messed up (as of this date this routine only called from Quotes.frm)
    m.QB.DrawBoard
    pbQuoteBoard.Refresh

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.UpdateSettings"

End Sub

Private Sub HandleSymAddOrChange(Optional ByVal strChar$)
On Error GoTo ErrSection:

    Dim iRowSave&, iRowCountSave&, iRedrawSave&
    Dim bValid As Boolean
    
    Select Case strChar
        Case "A" To "Z", "a" To "z", "$"
            bValid = True
        Case ""
            bValid = True
            If m.eStyle = eGDQuoteStyle_Grid Then fgQuotes.Row = fgQuotes.Rows - 1
    End Select
    
    If bValid Then
        If m.eStyle = eGDQuoteStyle_Grid Then
            With fgQuotes
                iRedrawSave = .Redraw
                .Redraw = flexRDNone
                
                iRowCountSave = .Rows
                iRowSave = -1
                
                If .Row < .Rows - 1 Then
                    'not last row -> user is changing symbol
                    iRowSave = .Row
                End If
                .MergeRow(.Row) = False
                
                frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_AddSymbol, strChar
                
                If iRowSave <> -1 Then
                    If iRowCountSave > 0 And iRowCountSave < .Rows Then
                        .RemoveItem iRowSave + 1      'user changed a symbol
                    End If
                ElseIf iRowCountSave = .Rows Then
                    'user cancelled symbol add - do nothing
                Else
                    .Row = .Rows - 1
                    .Col = m.iSymbolCol
                    .EditCell                       '6094
                End If
                
                .Redraw = iRedrawSave
            End With
        Else
            frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_AddSymbol, strChar
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.HandleSymAddOrChange"

End Sub

Public Sub RemoveAllSymbols()
On Error GoTo ErrSection:

    Dim lRedrawSave As Long

    m.strSymbols = ""
    With fgQuotes
        lRedrawSave = .Redraw
        .Redraw = flexRDNone
        .Rows = .FixedRows
        frmQuotes.TabFuncWrappers Me, eGDTabFuncWrapper_AddLabel
        .Redraw = lRedrawSave
    End With
    m.QB.BoardFromString m.strSymbols

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.RemoveAllSymbols"

End Sub

Public Sub RemoveTab()
On Error Resume Next:

    m.bRemoveTab = True
    Unload Me

End Sub

Public Sub DisplayAlert(ByVal Alert As cAlert, Optional ByVal bRemove As Boolean = False, _
    Optional ByVal bAdd As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex&, lCol&, i&
    Dim strFlexData$
    
    If m.eStyle <> eGDQuoteStyle_Grid Then
        If bAdd Then m.QB.BoxQbAlertUpdate Alert, bRemove, bAdd
        GoTo ErrExit
    End If
    
    If Alert.AlertType <> eGDAlertType_QuoteBoard Then GoTo ErrExit
    
    lCol = -1&
    
    With fgQuotes
        
        ' Find the column in the grid with the given field name...
        For lIndex = 0 To fgQuotes.Cols - 1
            If fgQuotes.TextMatrix(0, lIndex) = Alert.Field Then
                lCol = lIndex
                Exit For
            End If
        Next lIndex
        
        If lCol > 0 And lCol < .Cols Then
            .Redraw = flexRDNone
            If Alert.IsSymbol Then
                'update row with matching symbol & bar period
                For lIndex = .FixedRows To .Rows - 1
                    If Alert.Symbol = Parse(.TextMatrix(lIndex, frmQuotes.SymbolCol), " ", 1) Then  'parse out the (month year) from grid display for continuous contracts of futures
                        If Alert.Period = .TextMatrix(lIndex, frmQuotes.PeriodCol) Then
                            strFlexData = Parse(.Cell(flexcpData, lIndex, lCol), "|", 1)
                            If bRemove = True Then
                                .Cell(flexcpPicture, lIndex, lCol) = Nothing
                                .Cell(flexcpBackColor, lIndex, lCol) = .Cell(flexcpBackColor, lIndex, 0)        'clear out highlight color
                                If Len(strFlexData) > 0 Then .Cell(flexcpData, lIndex, lCol) = strFlexData
                            Else
                                If Alert.Active = True Then
                                    .Cell(flexcpPicture, lIndex, lCol) = Picture16(ToolbarIcon(kActiveAlertIcon))
                                Else
                                    .Cell(flexcpPicture, lIndex, lCol) = Picture16(ToolbarIcon(kInactiveAlertIcon))
                                End If
                                .Cell(flexcpPictureAlignment, lIndex, lCol) = flexAlignLeftTop
                                If Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 1) <> 0 Then
                                    If Alert.LastCheckedFLag = True Then
                                        If Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 2) > 0 Then
                                            .Cell(flexcpBackColor, lIndex, lCol) = ValOfText(Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 2))
                                        End If
                                    End If
                                End If
                                strFlexData = strFlexData & "|" & Alert.AlertKey
                                .Cell(flexcpData, lIndex, lCol) = strFlexData
                            End If
                        End If
                    End If
                Next
            ElseIf frmQuotes.TabStr(eGDTabSettings_Name, MyTabIndex) = Alert.TabName Then
                If bRemove = True Then
                    fgQuotes.Cell(flexcpPicture, 0, lCol) = Nothing
                    If Alert.HasColorAction Then
                        For i = fgQuotes.FixedRows To fgQuotes.Rows - 1
                            fgQuotes.Cell(flexcpBackColor, i, lCol) = fgQuotes.Cell(flexcpBackColor, i, 0)
                        Next
                    End If
                ElseIf Alert.Active = True Then
                    fgQuotes.Cell(flexcpPicture, 0, lCol) = Picture16(ToolbarIcon(kActiveAlertIcon))
                Else
                    fgQuotes.Cell(flexcpPicture, 0, lCol) = Picture16(ToolbarIcon(kInactiveAlertIcon))
                End If
                fgQuotes.Cell(flexcpPictureAlignment, 0, lCol) = flexAlignLeftTop
                fgQuotes.AutoSize lCol, , False, 75
                fgQuotes.Cell(flexcpData, 0, lCol) = "|" & Alert.AlertKey
            End If
            .Redraw = flexRDDirect
        End If
        
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.DisplayAlert"

End Sub

Public Sub ColorCell(ByVal strSymbol As String, ByVal strPeriod As String, ByVal strField As String, _
    ByVal strQbTab As String, Optional ByVal vBackColor As Variant)
On Error GoTo ErrSection:

    Dim lIndex&, lCol&
    
    If m.eStyle <> eGDQuoteStyle_Grid Then Exit Sub
    
    If Len(strQbTab) > 0 Then
        If strQbTab <> frmQuotes.TabStr(eGDTabSettings_Name, m.iTabIndex) Then
            GoTo ErrExit
        End If
    End If
    
    lCol = -1&
    
    With fgQuotes
        ' Find the column in the grid with the given field name...
        For lIndex = 0 To fgQuotes.Cols - 1
            If .TextMatrix(0, lIndex) = strField Then
                lCol = lIndex
                Exit For
            End If
        Next lIndex
        
        If lCol > 0 And lCol < .Cols Then
            .Redraw = flexRDNone
            
            'update row with matching symbol & bar period
            For lIndex = .FixedRows To .Rows - 1
                'parse out the (month year) from grid display for continuous contracts of futures
                If Parse(.TextMatrix(lIndex, frmQuotes.SymbolCol), " ", 1) = strSymbol Then
                    If .TextMatrix(lIndex, frmQuotes.PeriodCol) = strPeriod Then
                        If IsMissing(vBackColor) Then
                            .Cell(flexcpBackColor, lIndex, lCol) = .Cell(flexcpBackColor, lIndex, 0)
                        Else
                            .Cell(flexcpBackColor, lIndex, lCol) = vBackColor
                        End If
                    End If
                End If
            Next
            
            .Redraw = flexRDDirect
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDetachedQBTab.ColorCell"

End Sub

Private Sub tmr_Timer()

    If Not g.bStarting Then
        tmr.Enabled = False
        g.Alerts.DisplayQBAlerts Me
    End If

End Sub

