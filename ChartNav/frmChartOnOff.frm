VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChartOnOff 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Show/Hide Indicators"
   ClientHeight    =   3840
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   2130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3420
      Width           =   2235
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmChartOnOff.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmChartOnOff.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOnOff.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   60
         Width           =   660
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frmChartOnOff.frx":005C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartOnOff.frx":0088
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartOnOff.frx":00FA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditCfg 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   60
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frmChartOnOff.frx":0116
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartOnOff.frx":013E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartOnOff.frx":01A0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frmChartOnOff.frx":01BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmChartOnOff.frx":01E2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmChartOnOff.frx":0236
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1995
      _cx             =   3519
      _cy             =   4789
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
      Rows            =   11
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
      ScrollBars      =   2
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
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   195
      Left            =   60
      Top             =   30
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmChartOnOff.frx":0252
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmChartOnOff.frx":0298
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartOnOff.frx":02B8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Begin VB.Menu mnuShow 
         Caption         =   "Show Indicators"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Indicator/Pane"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmChartOnOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'EVENTS for a Click:
' BeforeMouseDown 26455811
' MouseUp 26455881
' Click 26455891

'EVENTS for a DblClick:
' BeforeMouseDown 26465575
' MouseUp 26465655
' Click 26465655
' DblClick 26465895
' MouseUp 26465895

Private Const kCheckBoxColWidth = 240 '350
Private Const kTradeSystemRow = 1

Private Type mPrivate
    ChartForm As Form
    aExpandPane As New cGdArray
    
    nPricePaneId As Long
    nPriceIndId As Long
    strPrevChart As String
    nMinRowHeight As Long
    nPopupRow As Long
    nTopRow As Long
    bVisible As Boolean
    bMenuShowOverride As Boolean
    bExpandPricePane As Boolean
End Type
Private m As mPrivate

Private Sub cmdAdd_Click()
    KeyPress Asc("A")
End Sub

Private Sub cmdDelete_Click()
    DeleteIndPane fg.Row
End Sub

Private Sub cmdEditCfg_Click()
            
    With fg
        If .Row > .FixedRows Then
            EditSettings .Row
        Else
            KeyPress Asc("E")
        End If
    End With
    
End Sub

Private Sub fg_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error Resume Next:

    m.ChartForm.Chart.geResetPanes
    m.ChartForm.Chart.geForceRecalc
    m.ChartForm.Chart.GenerateChart eRedo3_Settings

End Sub

Private Sub fg_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)

    m.nTopRow = NewTopRow

End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRow&, nID&, i&
    Dim bShow As Boolean
    
    If m.ChartForm Is Nothing Then Exit Sub

    With fg
        nRow = .MouseRow
        m.nPopupRow = -1
        m.nTopRow = fg.TopRow
        If nRow = kTradeSystemRow Then
            If .MouseCol = 0 Then
                If m.ChartForm.Chart.SystemID = 0 Then
                    EditSettings kTradeSystemRow
                Else
                    If m.ChartForm.Chart.ShowTrades Then
                        m.ChartForm.Chart.ShowTrades = False
                        .Cell(flexcpChecked, kTradeSystemRow, 0) = flexUnchecked
                    Else
                        m.ChartForm.Chart.ShowTrades = True
                        .Cell(flexcpChecked, kTradeSystemRow, 0) = flexChecked
                    End If
                    m.ChartForm.Chart.GenerateChart eRedo1_Scrolled
                End If
            End If
        ElseIf nRow >= .FixedRows Then
            ' check for right-click
            If Button = 2 Then
                Cancel = True
                .Row = nRow
                ShowPopup nRow
            ElseIf .MouseCol = 0 Or .MouseCol = 1 Then
                Cancel = True
                i = m.ChartForm.Chart.Tree.NodeLevel(.RowData(nRow))          '0=pane, >0=ind
                If i > 0 Then i = 1
                If .Cell(flexcpChecked, nRow, i) = flexChecked Then
                    .Cell(flexcpChecked, nRow, i) = flexUnchecked
                    ' if unchecking a pane then collapse it
                    nID = .RowData(nRow)
                    If m.ChartForm.Chart.Tree.NodeLevel(nID) = 0 Then
                        ShowIndRows nRow, False
                    End If
                Else
                    .Cell(flexcpChecked, nRow, i) = flexChecked
                End If
                .Cell(flexcpRefresh, nRow, i) = True '(to force checkbox to immediately be repainted)
                ShowHideOnChart nRow, i
            Else
                ' drag row
                .DragRow nRow
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Fg.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_BeforeMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    Dim i&, j&, k&, iMoved&
    Dim Tree As cGdTree
    Dim Pane1 As cPane, Pane2 As cPane
    Dim Ind As cIndicator
        
    If m.ChartForm Is Nothing Then Exit Sub
    If m.ChartForm.Chart Is Nothing Then Exit Sub
    If m.ChartForm.Chart.Tree Is Nothing Then Exit Sub
    If Row = Position Then Exit Sub
    
    Set Tree = m.ChartForm.Chart.Tree
    i = Tree.NodeLevel(fg.RowData(Row))
    j = Tree.NodeLevel(fg.RowData(Position))
    
    If Row < 2 Then
        Position = Row          'disallow moving strategy row
    ElseIf i = 0 And j = 0 Then
        'moving pane
        Set Pane1 = Tree(fg.RowData(Row))
        Set Pane2 = Tree(fg.RowData(Position))
        If Pane1 Is Nothing Or Pane2 Is Nothing Then
            'The default grid behavior will move the row regardless of whether the
            'panes were really moved. Setting position = row just keeps the grid
            'from moving the row back and forth unnecessarily.
            Position = Row
        End If
    ElseIf i = 0 And j > 0 Then
        'moving pane: the 'TO' row is an indicator, get its parent
        Set Ind = Tree(fg.RowData(Position))
        If Not Ind Is Nothing Then
            If Tree.NodeLevel(Ind.geIndpaneId) = 0 Then
                Set Pane1 = Tree(fg.RowData(Row))
                Set Pane2 = Tree(Ind.geIndpaneId)
            End If
            If Pane1 Is Nothing Or Pane2 Is Nothing Then
                Position = Row
            End If
        End If
    ElseIf (i > 0 And j >= 0) Then
        'moving indicators
        Set Ind = Tree(fg.RowData(Row))
        If Not Ind Is Nothing Then Ind.SaveGroupInfo
        If Ind.isPriceInd Then
            Beep 'can't move price from price pane - 4766
        ElseIf (i = j) Then
            'indicator in the FROM row is at same level as indicator in the TO row
            If Row > Position Then
                iMoved = Tree.Move(fg.RowData(Row), fg.RowData(Position), eTREE_PrevSibling)
            Else
                iMoved = Tree.Move(fg.RowData(Row), fg.RowData(Position), eTREE_NextSibling)
            End If
        Else
            'get parent of the indicator in the TO row
            k = Tree.AncestorIndex(fg.RowData(Position), 0)
            If Tree.NodeLevel(k) = 0 Then
                k = Tree.RelativeIndex(k, eTREE_LastChild)
                If Tree.NodeLevel(k) = i Then
                    iMoved = Tree.Move(fg.RowData(Row), k, eTREE_NextSibling)
                End If
            End If
        End If
        If iMoved > 0 And Not Ind Is Nothing Then Ind.CheckGroup
    Else
        Position = Row
    End If
    
    'move the panes
    If Not Pane1 Is Nothing And Not Pane2 Is Nothing Then
        If Row > Position Then
            Tree.Move Pane1.gePaneId, Pane2.gePaneId, eTREE_PrevSibling
        Else
            Tree.Move Pane1.gePaneId, Pane2.gePaneId, eTREE_NextSibling
        End If
    End If
    
    Set Pane1 = Nothing
    Set Pane2 = Nothing
    Set Ind = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".fg_BeforeMoveRow"

End Sub

Private Sub fg_DblClick()
On Error GoTo ErrSection:

    Dim nRow&
    Dim bShow As Boolean

    If m.ChartForm Is Nothing Then Exit Sub
    
    m.strPrevChart = Me.Caption
    EditSettings fg.MouseRow
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Fg.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Static nRow&, nCol&

    With fg
        If nRow <> fg.MouseRow Or nCol <> fg.MouseCol Then
            nRow = fg.MouseRow
            nCol = fg.MouseCol
            .ToolTipText = "" ' (to force moving rows even if unchanged)
            If nRow > kTradeSystemRow And nCol >= 0 Then
                If Not m.ChartForm Is Nothing Then
                    If nCol = 1 Then
                        .ToolTipText = "(can dbl-click to edit or drag to move up/down)"
                    ElseIf m.ChartForm.Chart.Tree.NodeLevel(.RowData(nRow)) = 0 And nCol = 0 Then
                        .ToolTipText = "(click to show/hide or right-click to expand/collapse)"
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Fg.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Activate()
On Error Resume Next:

    fg.Row = -1
    If DockState(Me) = eHidden Then
        m.bExpandPricePane = True
        TextIncDecUnregisterForm Me
    Else
        TextIncDecRegisterForm Me
    End If

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Nothing
    Else
        KeyPress KeyCode, Shift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    mnuPopUp.Visible = False
    m.bExpandPricePane = True
    InitGrid
    Me.Icon = Picture16(ToolbarIcon("ID_ChartOnOff"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    TextIncDecUnregisterForm Me
    
    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_ChartOnOff").State = ssUnchecked
        m.bVisible = False
        m.aExpandPane.Size = 0
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub


Private Sub Form_Resize()

    On Error Resume Next

    With fraButtons
        '.Move (Me.ScaleWidth - .Width) / 2, Me.ScaleHeight - .Height
        .Top = Me.ScaleHeight - .Height
    End With
    With fg
        .Move .Left, .Top, Me.ScaleWidth - .Left * 1, fraButtons.Top - .Top
        '.Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top
    End With
        
    AdjustRowHeight
    AutoSizeChart

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim strFont$

    frmMain.DockPro.RemoveForm Me.Name
    Set m.aExpandPane = Nothing
    
    'save font info
    strFont = fg.Font.Name & "|" & Str(fg.Font.Size)
    SetIniFileProperty "ChartOnOff", strFont, "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ShowData(Optional ByRef frm As Form = Nothing)
On Error GoTo ErrSection:

    Dim lRowSave As Long

    If g.bUnloading Then
        Set m.ChartForm = Nothing
        Exit Sub
    End If
    If DockState(Me) = eHidden Then Exit Sub
    
    If frm Is Nothing Then
        Set m.ChartForm = ActiveChart
    ElseIf IsFrmChart(frm) Then
        Set m.ChartForm = frm
    Else
        Set m.ChartForm = ActiveChart
    End If
    
    If m.ChartForm Is Nothing Then Exit Sub
                          
    ' save index of the price pane and the price indicator
    m.nPricePaneId = m.ChartForm.Chart.Tree.Index("PRICE PANE")
    m.nPriceIndId = m.ChartForm.Chart.Tree.Index("PRICE")
    
    Me.Caption = m.ChartForm.Chart.ChartName
    
    fg.Redraw = flexRDNone
    
    lRowSave = fg.Row
    SetSystemRow
    If m.strPrevChart <> Me.Caption Then m.strPrevChart = ""
    
    fg.Rows = kTradeSystemRow + 1
    LoadGrid
    AdjustRowHeight
    
    fg.Redraw = flexRDBuffered
    
    If lRowSave >= fg.FixedRows And lRowSave < fg.Rows Then
        fg.Select lRowSave, 0, lRowSave, fg.Cols - 1        '6271
    End If
       
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.ShowData", eGDRaiseError_Raise

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:
    
    Dim strFont$, strFontName$
    Dim dFontSize#, i&
    
    strFont = GetIniFileProperty("ChartOnOff", "", "Fonts", g.strIniFile)
    If Len(strFont) > 0 Then
        i = InStr(strFont, "|")
        If i >= 0 Then
            strFontName = Left(strFont, i - 1)
            dFontSize = Val(Right(strFont, Len(strFont) - i))
        End If
    End If
    
    With fg
        SetupGrid fg, eGridMode_Grid
        .HighLight = flexHighlightWithFocus
        '.Editable = flexEDKbdMouse
        '.SelectionMode = flexSelectionFree
        .GridLines = flexGridFlatHorz
        .Cols = 4
        .Rows = kTradeSystemRow + 1
        .FixedCols = 0
        .FixedRows = 1
        'check box for panes
        .ColWidth(0) = kCheckBoxColWidth
        .ColAlignment(0) = flexAlignCenterTop
        'check box for indicators
        .ColWidth(1) = kCheckBoxColWidth
        .ColAlignment(1) = flexAlignCenterTop
        'hidden col to hold show/collapse flag
        .ColHidden(3) = True
        
If 0 Then
' testing this: to look more like buttons
.Left = 60
.ColHidden(0) = True
.Appearance = flexFlat
.BorderStyle = flexBorderNone
.GridLines = flexGridExplorer
.HighLight = flexHighlightNever
.BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspaceMe.BackColor
.BackColor = Me.BackColor
End If
        
        'set font info
        If Len(strFontName) > 0 Then fg.Font.Name = strFontName
        If dFontSize > 0 Then fg.Font.Size = dFontSize
        
        .TextMatrix(0, 0) = "Use"
        .TextMatrix(0, 1) = "Indicator"
        .RowHidden(0) = True
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub SetSystemRow()
On Error GoTo ErrSection:

    If m.ChartForm Is Nothing Then Exit Sub
    
    With fg
        .Cell(flexcpChecked, kTradeSystemRow, 1) = flexNoCheckbox
        If HasGold(False, , False) Then
            If m.ChartForm.Chart.ShowTrades Then
                .Cell(flexcpChecked, kTradeSystemRow, 0) = flexChecked
            Else
                .Cell(flexcpChecked, kTradeSystemRow, 0) = flexUnchecked
            End If
            .RowHidden(kTradeSystemRow) = False
        Else
            .RowHidden(kTradeSystemRow) = True
        End If
        .TextMatrix(kTradeSystemRow, 2) = "Trading Strategy"
        If .GridLines <> flexGridExplorer Then
            If g.nColorTheme = kDarkThemeColor Then
                .Cell(flexcpBackColor, kTradeSystemRow, 0, kTradeSystemRow, .Cols - 1) = kDarkThemeColor
                .Cell(flexcpForeColor, kTradeSystemRow, 0, kTradeSystemRow, .Cols - 1) = vbWhite
            Else
                .Cell(flexcpBackColor, kTradeSystemRow, 0, kTradeSystemRow, .Cols - 1) = vbWhite
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.SetSystemRow", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub ShowIndRows(ByVal nRow&, ByVal bShow As Boolean)
On Error GoTo ErrSection:

    Dim i&, strPane$
    Dim bShowPane As Boolean
    Dim Tree As cGdTree
    Dim Ind As cIndicator

    If m.ChartForm Is Nothing Then Exit Sub

    Set Tree = m.ChartForm.Chart.Tree
    With fg
        i = .RowData(nRow)
        'find parent pane if indicator
        If Tree.NodeLevel(i) <> 0 Then
            For i = nRow - 1 To .FixedRows Step -1
                If Tree.NodeLevel(.RowData(i)) = 0 Then
                    nRow = i
                    Exit For
                End If
            Next
        End If
        If Tree.NodeLevel(.RowData(nRow)) = 0 Then
            'pane's checkbox may not be checked if indicators expanded by popup menu
            If bShow And .Cell(flexcpChecked, nRow, 0) <> flexChecked Then
                .Cell(flexcpChecked, nRow, 0) = flexChecked
                bShowPane = True
            End If
            strPane = .TextMatrix(nRow, 2)
        Else
            Exit Sub        'something very wrong here
        End If
        'step through and toggle all indicator rows
        For i = nRow + 1 To .Rows - 1
            If Tree.NodeLevel(.RowData(i)) <= 0 Then
                Exit For
            Else
                .RowHidden(i) = Not bShow
            End If
        Next
    End With

    If bShowPane Then ShowHideOnChart nRow
    
    If bShow Then
        AdjustRowHeight
    End If
    
    'save or remove pane's name from array
    If InStr(strPane, "Price Pane") Then
        m.bExpandPricePane = bShow
    Else
        m.aExpandPane.Sort
        If m.aExpandPane.BinarySearch(strPane, i) Then
            If Not bShow Then m.aExpandPane.Remove i
        ElseIf bShow Then
            m.aExpandPane.Add strPane
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.ShowIndRows", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub ShowHideOnChart(ByVal nRow&, Optional ByVal nCol& = -1)
On Error GoTo ErrSection:

    Dim Tree As cGdTree
    Dim Pane As cPane
    Dim Ind As cIndicator
    Dim bShow As Boolean
    Dim nIdx&, i&

    If m.ChartForm Is Nothing Then Exit Sub
           
    If nRow = kTradeSystemRow Then
        If fg.Cell(flexcpChecked, kTradeSystemRow, 0) = flexChecked Then
            m.ChartForm.tmr.Tag = "AddSystem"
        Else
            m.ChartForm.Chart.ShowTrades = False
            m.ChartForm.Chart.GenerateChart eRedo1_Scrolled
        End If
        MoveFocus m.ChartForm
        Exit Sub
    End If
                
    m.strPrevChart = Me.Caption
    Set Tree = m.ChartForm.Chart.Tree
    
    With fg
        If nCol = -1 Then
            If .Cell(flexcpChecked, nRow, 0) = flexChecked Then bShow = True
        Else
            If .Cell(flexcpChecked, nRow, nCol) = flexChecked Then bShow = True
        End If
        
        nIdx = .RowData(nRow)
        If nIdx > 0 Then
'            If nIdx = m.nPricePaneId Then
'                Set Pane = Tree(nIdx)
'                If Not Pane Is Nothing Then
'                    If bShow Then
'                        Pane.PricePaneFlag = 1
'                        m.ChartForm.Chart.HidePriceIndicators = False
'                    Else
'                        Pane.PricePaneFlag = -1
'                        m.ChartForm.Chart.HidePriceIndicators = True
'                    End If
'                End If
'            ElseIf Tree.NodeLevel(nIdx) = 0 Then
            If Tree.NodeLevel(nIdx) = 0 Then
                Set Pane = Tree(nIdx)
                If Not Pane Is Nothing Then
                    Pane.Display = bShow
                    If bShow Then CheckNeedExpand nIdx, nRow
                End If
            Else
                Set Ind = m.ChartForm.Chart.Tree(nIdx)
                If Not Ind Is Nothing Then Ind.Display = bShow
                'see if parent pane needs to be turned on too
                If bShow Then
                    Set Pane = Tree(Tree.AncestorIndex(nIdx, 0))
                    If Not Pane Is Nothing Then
                        If Not Pane.Display Then Pane.Display = True
                    End If
                End If
            End If
            m.ChartForm.Chart.GenerateChart eRedo3_Settings
            If Not Ind Is Nothing Then m.ChartForm.pbChart.Refresh
            MoveFocus m.ChartForm
        End If
    End With
            
    Set Pane = Nothing
    Set Ind = Nothing
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.ShowHideOnChart", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Function GetPaneLabel(Pane As cPane) As String
On Error GoTo ErrSection:

    Dim strName$, i&
    Dim Tree As cGdTree
    Dim Ind As cIndicator

    If m.ChartForm Is Nothing Then Exit Function

    strName = Pane.Name
    Set Tree = m.ChartForm.Chart.Tree
    
    If UCase(Tree.Key(Pane.gePaneId)) = "PRICE PANE" Then
        strName = "Price Pane ..."
    ElseIf Len(strName) = 0 Then
        i = Tree.RelativeIndex(Pane.gePaneId, eTREE_FirstChild)
        If Tree.NodeLevel(i) > 0 Then
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                strName = Ind.ChartLabel
                i = InStr(strName, "(")
                If i > 0 Then strName = Mid(strName, 1, i - 1)
            End If
        End If
    End If
    
    GetPaneLabel = strName
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmChartOnOff.GetPaneLabel", eGDRaiseError_Show
    Resume ErrExit

End Function

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim Tree As cGdTree
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    Dim i&, j&, iEnd&
    Dim nColor&, nCheck&, strText$, strPane$
    Dim bAddPane As Boolean, bAddInd As Boolean
    Dim bHideRow As Boolean

    If fg.Rows > kTradeSystemRow + 1 Or m.ChartForm Is Nothing Then
        Exit Sub    'precautionary, should never happen
    End If
    
    Set Tree = m.ChartForm.Chart.Tree
    
    fg.Redraw = flexRDNone
    m.aExpandPane.Sort
    
    iEnd = Tree.Count
    
    For i = 1 To iEnd
        strText = ""
        nCheck = flexUnchecked
        bAddPane = False
        bAddInd = False
        Set Pane = Nothing
        Set Ind = Nothing
        If Tree.NodeLevel(i) = 0 Then
            bHideRow = False
            Set Pane = Tree(i)
            If Not Pane Is Nothing Then
                strText = GetPaneLabel(Pane)
                nColor = ALT_GRID_ROW_COLOR
                bAddPane = True
                If Pane.Display Then
                    nCheck = flexChecked
                End If
                If Tree.Key(i) = kClusterTimeKeyPane Then bHideRow = True
            End If
        Else
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                'JM 10-13-2015 fix for Chart Condition Alerts showing as indicator when should not
                '   issue reported by Heath (not sure how long this has been a bug?)
                If Not Ind.IsAlert Then
                    'set flag for hiding or showing grid row for indicator
                    j = Tree.AncestorIndex(i, 0)
                    Set Pane = Tree(j)
                    If Not Pane Is Nothing Then
                        If Pane.PricePaneFlag Then
                            bHideRow = Not m.bExpandPricePane   'Not Pane.Display
                        ElseIf strPane <> GetPaneLabel(Pane) Then
                            strPane = GetPaneLabel(Pane)
                            bHideRow = Not m.aExpandPane.BinarySearch(strPane)
                        End If
                    End If
                    
                    'set other values for displaying in grid
                    If Len(Ind.ChartLabel) > 0 Then
                        strText = Space(Indent(Tree.NodeLevel(i))) & Ind.ChartLabel
                    Else
                        strText = Space(Indent(Tree.NodeLevel(i))) & Ind.Name
                    End If
                    If g.nColorTheme = kDarkThemeColor Then
                        nColor = kDarkThemeColor
                    Else
                        nColor = vbWhite
                    End If
                    bAddInd = True
                    If Tree.Key(i) = kClusterPriceKey Then
                        bHideRow = True
                    ElseIf Len(Ind.GroupKey) > 0 And Not Ind.IAmGroupLeader Then
                        nCheck = flexNoCheckbox
                    ElseIf Ind.Display Then
                        nCheck = flexChecked
                    End If
                End If
            End If
        End If
        
        If bAddPane Or bAddInd Then
            If Me.Width < 2000 Then strText = Trim(strText)
            With fg
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = i
                .TextMatrix(.Rows - 1, 2) = strText
                If .GridLines <> flexGridExplorer Then
                    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = nColor
                    If g.nColorTheme = kDarkThemeColor Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbWhite
                    End If
                End If
                If bAddPane Then
                    .Cell(flexcpChecked, .Rows - 1, 0) = nCheck
                    .Cell(flexcpChecked, .Rows - 1, 1) = flexNoCheckbox
                Else
                    .Cell(flexcpChecked, .Rows - 1, 0) = flexNoCheckbox
                    .Cell(flexcpChecked, .Rows - 1, 1) = nCheck
                End If
                .RowHidden(.Rows - 1) = bHideRow
            End With
        End If
    Next
    
    fg.Cell(flexcpPictureAlignment, 0, 0, fg.Rows - 1, 0) = flexAlignLeftCenter
    fg.Redraw = flexRDBuffered
    If m.bVisible Then
        If m.nTopRow > 0 And m.nTopRow < fg.Rows Then
            fg.TopRow = m.nTopRow
        Else
            fg.TopRow = 0
        End If
    Else
        m.nTopRow = 0
        m.bVisible = True
    End If
                        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.LoadGrid", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ClearPrevChart()
    m.strPrevChart = ""
End Sub

Private Function RowType(ByVal nRow&) As Integer
On Error GoTo ErrSection:

    Dim nNode&, nLevel&
    
    If m.ChartForm Is Nothing Then Exit Function
    
    With fg
        If nRow > kTradeSystemRow And nRow < .Rows Then
            nNode = .RowData(nRow)
        End If
        If nNode > 0 Then
            nLevel = m.ChartForm.Chart.Tree.NodeLevel(nNode)
            If nLevel = 0 Then
                RowType = 0         'pane
            Else
                RowType = 1         'indicator
            End If
        ElseIf nRow = .Rows - 1 Then
            RowType = 2             'new pane
        Else
            RowType = -1            'don't care
        End If
    End With
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmChartOnOff.RowType", eGDRaiseError_Raise
        
End Function

Private Sub MoveIndRow(ByVal nFromRow&, ByVal nToRow&)
On Error GoTo ErrSection:

    Dim idxFromNode&, idxToNode&, idxMovedTo&
    Dim nParent&, nParentLevel&
    Dim Tree As cGdTree
    Dim IndFrom As cIndicator, IndTo As cIndicator

    If m.ChartForm Is Nothing Then Exit Sub

    Set Tree = m.ChartForm.Chart.Tree
                
    idxFromNode = fg.RowData(nFromRow)
    idxToNode = fg.RowData(nToRow)
    
    With Tree
        Set IndFrom = Tree(idxFromNode)
        ' check if indicator is highlight bars
        If IndFrom.DataType = eINDIC_BooleanArray Then
            If .NodeLevel(idxToNode) > 0 Then   'must move to an indicator level
                Set IndTo = Tree(idxToNode)
                If IndTo.DataType <> eINDIC_Constant Then 'do not highlight horz indicator
                    If IndTo.DataType = eINDIC_BooleanArray Or _
                       .NodeLevel(idxToNode) = .NodeLevel(idxFromNode) Then     'fix for aardvark 689
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
                    Else
                        idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
                        IndFrom.Display = IndTo.Display
                    End If
                End If
                If idxMovedTo > 0 Then IndFrom.geIndId = idxMovedTo
            End If
            Set IndFrom = Nothing
            Set IndTo = Nothing
        ' check if indicator is "unlinked"
        ElseIf .NodeLevel(idxFromNode) = 1 Then
            If .Key(idxFromNode) = "PRICE" And _
                .Key(.RelativeIndex(idxToNode, eTREE_Root)) <> "PRICE PANE" Then
                    Beep 'can't move price from price pane
            ElseIf RowType(nToRow) = 2 Then
                Beep ' need to add new pane - don't do this
            ElseIf idxToNode = 0 Then
                Beep ' above the tree is invalid
            ElseIf RowType(nToRow) = 0 Then
                ' move to be first child of this Pane
                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
            Else
                ' get the level 1 ancestor of the ToNode
                idxToNode = .AncestorIndex(idxToNode, 1)
                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
            End If
        Else
            ' linked indicators MUST stay under their current
            ' parent (and stay at the same level)
            nParent = .RelativeIndex(idxFromNode, eTREE_Parent)
            nParentLevel = .NodeLevel(nParent)
            ' if ToNode is parent, move to be first child
            If idxToNode = nParent Then
                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_FirstChild)
            ' otherwise the ancestor of the ToNode at the parent
            ' level must be the parent of the FromNode
            ElseIf nParent = .AncestorIndex(idxToNode, nParentLevel) Then
                ' then move to be next sibling of ToNode
                ' ancestor at FromNode level
                idxToNode = .AncestorIndex(idxToNode, nParentLevel + 1)
                idxMovedTo = .Move(idxFromNode, idxToNode, eTREE_NextSibling)
            Else
                Beep ' else an invalid move of a linked indicator
            End If
        End If
        
    End With
    
    Set IndFrom = Nothing
    Set IndTo = Nothing
    
    If idxMovedTo > 0 And idxMovedTo <> idxFromNode Then
        m.strPrevChart = ""
        m.ChartForm.Chart.GenerateChart eRedo5_RecalcInd
        MoveFocus m.ChartForm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.MoveIndRow", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub AdjustRowHeight()
On Error GoTo ErrSection:

    Dim i&, nRows&, nRowHeight&
    
    If m.nMinRowHeight = 0 Then
        m.nMinRowHeight = fg.RowHeight(0)
        If m.nMinRowHeight = 0 Then Exit Sub
    End If
        
    With fg
        ' get # of visible rows
        For i = 0 To .Rows - 1
            If Not .RowHidden(i) Then
                nRows = nRows + 1
            End If
        Next
        If nRows > 0 Then       'aardvark 3425 fix
            'row & client height are both in twips (do not convert back & forth)
            nRowHeight = Int(.ClientHeight / nRows)
            'subtract 1 pixel for border (since row height is in twips, 1 pixel = Screen.TwipsPerPixelY)
            nRowHeight = nRowHeight - Screen.TwipsPerPixelY
        End If
        ' don't get less than default row height, and don't get greater than "pretty big"
        If nRowHeight < m.nMinRowHeight Then
            nRowHeight = m.nMinRowHeight
        ElseIf nRowHeight > m.nMinRowHeight * 2 Then
            nRowHeight = m.nMinRowHeight * 2
        End If
        ' set height for all rows (except header)
        .RowHeight(-1) = nRowHeight
        '.RowHeight(0) = m.nMinRowHeight
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.AdjustRowHeight"
    Resume ErrExit
End Sub

Public Sub KeyPress(KeyAscii As Integer, Optional Shift As Integer = -1)
On Error Resume Next

    Dim nRow&
    Dim frm As Form
    Dim bLookForChart As Boolean

    If KeyAscii = 0 Then Exit Sub

    If Shift >= 0 Then ' (came from KeyDown event)
        If KeyAscii >= vbKeyF2 And KeyAscii <= vbKeyF12 Then
            bLookForChart = True
        End If
    Else ' (came from KeyPress event)
        Select Case Asc(UCase(Chr(KeyAscii)))
            Case 32:        ' Space
                nRow = fg.Row
                KeyAscii = 0
                If nRow >= fg.FixedRows Then
                    CheckedCell(fg, nRow, 0) = Not CheckedCell(fg, nRow, 0)
                    ShowHideOnChart nRow
                End If
            
            Case 65 To 90, 48 To 57, 43, 45, 61:
                bLookForChart = True
        End Select
    End If
       
    If bLookForChart Then
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            'MoveFocus frm
            'DoEvents
            frm.KeyPress KeyAscii, Shift
            MoveFocus frm
        End If
        KeyAscii = 0
    End If
       
    Set frm = Nothing

End Sub

Private Sub EditSettings(ByVal nRow)

    Dim i&, idx&                'index of indicator to be edited
    Dim idxFirst&, idxLast&     'index of first & last child of pane node
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    
    With fg
        If nRow = kTradeSystemRow Then
            frmChartCfg.ShowMe m.ChartForm.Chart, -1
        ElseIf nRow > .FixedRows Then
            Set Tree = m.ChartForm.Chart.Tree
            idx = .RowData(nRow)
            'if pane is shown then get first visible indicator in pane
            If .Cell(flexcpChecked, nRow, 0) = flexChecked Then
                If Tree.NodeLevel(idx) = 0 Then
                    idxFirst = Tree.RelativeIndex(idx, eTREE_FirstChild)
                    idxLast = Tree.RelativeIndex(idx, eTREE_LastChild)
                    For i = idxFirst To idxLast
                        Set Ind = Tree(i)
                        If Ind.Display Then
                            idx = i
                            Exit For
                        End If
                    Next
                End If
            End If
            
            frmChartCfg.ShowMe m.ChartForm.Chart, idx
        End If
    End With

End Sub

Private Sub ShowPopup(ByVal nRow&)

    Dim idx&, idxNext&, i&
    Dim Tree As cGdTree

    If nRow <= fg.FixedRows Then Exit Sub
    
    Set Tree = m.ChartForm.Chart.Tree
    
    With fg
        idx = .RowData(nRow)
        If Tree.NodeLevel(idx) = 0 Then
            mnuDelete.Caption = "Delete Pane"
            'determine if next visible row is a pane or indicator
            idxNext = -1
            For i = nRow + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    idxNext = .RowData(i)
                    Exit For
                End If
            Next
            If Tree.NodeLevel(idxNext) > 0 Then
                mnuShow.Checked = True
            Else
                mnuShow.Checked = False
            End If
        Else
            mnuShow.Checked = True
            mnuDelete.Caption = "Delete Indicator"
        End If
    End With
    
    m.nPopupRow = nRow
    Me.PopupMenu mnuPopUp

End Sub

Private Sub DeleteIndPane(ByVal nRow&)
On Error GoTo ErrSection:

    Dim idx&, strKey$, strGroupKey$, nPrevRow&
    
    Dim Chart As cChart
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    Dim IndLeader As cIndicator
    
    If Not m.ChartForm Is Nothing Then
        If Not m.ChartForm.Chart Is Nothing Then
            If Not m.ChartForm.Chart.Tree Is Nothing Then
                Set Chart = m.ChartForm.Chart
                Set Tree = Chart.Tree
                With fg
                    If nRow > .FixedRows Then
                        idx = .RowData(nRow)
                        If idx > 0 Then
                            strKey = Tree.Key(idx)
                            If UCase(Left(strKey, 5)) = "PRICE" Then
                                Beep ' can't remove it
                                InfBox "'Price' cannot be removed.", "!"
                            ElseIf Tree.NodeLevel(strKey) > 0 Then
                                Set Ind = Tree(strKey)
                                If Not Ind Is Nothing Then
                                    strGroupKey = Ind.GroupKey
                                    If Len(strGroupKey) = 0 Then
                                        'not in a group
                                        Tree.Remove (idx)
                                    ElseIf Ind.IAmGroupLeader Then
                                        Ind.RemoveMyGroup
                                    ElseIf InfBox("Remove all indicators in group?", "?", "-Yes|+No", "Confirmation") = "Y" Then
                                        Ind.RemoveMyGroup
                                    Else
                                        Set IndLeader = Tree(strGroupKey)
                                        If Not IndLeader Is Nothing Then
                                            Tree.Remove (idx)
                                            IndLeader.SaveGroupInfo
                                        End If
                                    End If
                                End If
                            Else
                                'this is a pane
                                Tree.Remove (idx)
                            End If
                        End If
                    End If
                End With
                m.strPrevChart = ""
                Chart.GenerateChart eRedo5_RecalcInd
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartOnOff.DeleteIndPane"
    
End Sub

Private Sub mnuDelete_Click()
    DeleteIndPane m.nPopupRow
End Sub

Private Sub mnuFont_Click()
On Error GoTo ErrSection:

    Dim i&

    i = fg.RowHeight(0)     'save for comparison
    fg.AutoSizeMode = flexAutoSizeRowHeight
    If ChangeGridFont(fg, True) Then
        fg.AutoSize 0, 1, True
        If i <> fg.RowHeight(0) And fg.RowHeight(0) > 0 Then
            m.nMinRowHeight = 0
            AdjustRowHeight
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartOnOff.mnuFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuShow_Click()

    With fg
        If m.nPopupRow > .FixedRows Then
            .Redraw = flexRDNone
            m.bMenuShowOverride = Not mnuShow.Checked
            ShowIndRows m.nPopupRow, Not mnuShow.Checked
            .Redraw = flexRDBuffered
        End If
    End With
    
    m.bMenuShowOverride = False
    
End Sub

Private Sub CheckNeedExpand(ByVal nTreeIdx&, nGridRow&)
On Error GoTo ErrSection:

    Dim i&, j&
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    Dim Pane As cPane
    Dim bExpand As Boolean

    Set Tree = m.ChartForm.Chart.Tree
    
    If Tree Is Nothing Then Exit Sub
    
    bExpand = True
    If Not m.bMenuShowOverride Then
        'user just used pop-up menu to show indicators, no need to do this
        For i = nGridRow + 1 To fg.Rows - 1
            j = fg.RowData(i)
            If Tree.NodeLevel(j) = 0 Then
                Exit For
            Else
                Set Ind = Tree(j)
                If Not Ind Is Nothing Then
                    If Ind.Display Then
                        bExpand = False
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    If bExpand Or m.bMenuShowOverride Then
        Set Pane = Tree(nTreeIdx)
        If Not Pane Is Nothing Then
            m.aExpandPane.Add GetPaneLabel(Pane)
        End If
    End If
    
    Set Pane = Nothing
    Set Ind = Nothing
    
    Exit Sub

ErrSection:
     RaiseError Me.Name & ".CheckNeedExpand"

End Sub

Private Function Indent(ByVal nNodeLevel&) As Long
On Error GoTo ErrSection:

    Dim i As Long
    
    If nNodeLevel >= 2 Then
        i = (nNodeLevel - 1) * 4
    End If
    
    Indent = i
    
    Exit Function


ErrSection:
     RaiseError Me.Name & ".Indent"

End Function

