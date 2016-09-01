VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptCustomize 
   Caption         =   "Show/Hide Columns"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmFloatingOnly"
   ScaleHeight     =   4725
   ScaleWidth      =   6180
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3495
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1215
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
      Caption         =   "frmOptCustomize.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptCustomize.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptCustomize.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   0
         TabIndex        =   0
         Top             =   480
         Width           =   1215
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
         Caption         =   "frmOptCustomize.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptCustomize.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptCustomize.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmOptCustomize.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptCustomize.frx":00FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptCustomize.frx":011C
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtDesc 
      Height          =   750
      Left            =   120
      TabIndex        =   2
      Top             =   3660
      Width           =   4440
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOptCustomize.frx":0138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmOptCustomize.frx":0158
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptCustomize.frx":0178
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGrid 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   4455
      _cx             =   7858
      _cy             =   5530
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
   Begin HexUniControls.ctlUniLabelXP lblDesc 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   4455
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
      Caption         =   "frmOptCustomize.frx":0194
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptCustomize.frx":0230
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptCustomize.frx":0250
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmOptCustomize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    OptCols As cOptCols
    bOK As Boolean
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Show = 0
    eGDCol_Stat = 1
    eGDCol_Filters = 2
    eGDCol_Desc = 3
    eGDCol_Oper = 4
    eGDCol_FilterValue = 5
    eGDCol_Format = 6
End Enum
Private Const kGridCols = 7

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    
    With vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .SelectionMode = flexSelectionListBox
    
        .Rows = 1
        .FixedRows = 1
        .Cols = kGridCols
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Show)) = "Show"
        .ColDataType(GDCol(eGDCol_Show)) = flexDTBoolean
        
        .TextMatrix(0, GDCol(eGDCol_Stat)) = "Column"
        
        .TextMatrix(0, GDCol(eGDCol_Filters)) = "Filters"
        .ColComboList(GDCol(eGDCol_Filters)) = "..."

        .ColHidden(GDCol(eGDCol_Desc)) = True
        .ColHidden(GDCol(eGDCol_Oper)) = True
        .ColHidden(GDCol(eGDCol_FilterValue)) = True
        .ColHidden(GDCol(eGDCol_Format)) = True
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lSeqNbr As Long
    Dim OptCol As cOptCol
    Dim lIndex As Integer
    Dim lRedraw As Long
    
    Set m.OptCols = New cOptCols
    m.OptCols.Load
    
    With vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        'Skip first row (it has optimization column)
        For lIndex = 2 To m.OptCols.Count
            Set OptCol = m.OptCols.Item(lIndex)
            .Rows = .Rows + 1
            CheckedCell(vsGrid, .Rows - 1, GDCol(eGDCol_Show)) = Not OptCol.Hide
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Stat)) = OptCol.FieldName
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Desc)) = OptCol.FieldDesc
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Oper)) = OptCol.Operator
            .TextMatrix(.Rows - 1, GDCol(eGDCol_FilterValue)) = OptCol.FilterValue
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Format)) = OptCol.FieldFormat
            FormatFilters .Rows - 1, OptCol.FieldName, OptCol.Operator, OptCol.FilterValue
            lSeqNbr = lSeqNbr + 1000
        Next lIndex
    
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
    Set OptCol = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.LoadGrid", eGDRaiseError_Raise
    
End Sub

'This is also called by frmOptFilters
Public Sub FormatFilters(lRow As Long, pStat As Variant, pOper As Variant, _
    pOperVal As Variant)
On Error GoTo ErrSection:

    Dim strFormattedLine As String
    Dim lRedraw As Long
    
    strFormattedLine = ""
    If pOper <> "N" And pOper <> "" Then
        strFormattedLine = pStat & " " & pOper & " " & pOperVal
    End If
    
    With vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .TextMatrix(lRow, GDCol(eGDCol_Filters)) = strFormattedLine
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.FormatFilters", eGDRaiseError_Raise
    
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim OptCol As cOptCol
    
    ' Update filters collection
    With vsGrid
        For lIndex = .FixedRows To .Rows - 1
            Set OptCol = m.OptCols.Item(.TextMatrix(lIndex, GDCol(eGDCol_Stat)))
            OptCol.Hide = Not CheckedCell(vsGrid, lIndex, GDCol(eGDCol_Show))
            OptCol.Operator = .TextMatrix(lIndex, GDCol(eGDCol_Oper))
            OptCol.FilterValue = ValOfText(.TextMatrix(lIndex, GDCol(eGDCol_FilterValue)))
        Next lIndex
    End With
    m.OptCols.Save

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.Save", eGDRaiseError_Raise
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    txtDesc.Locked = True
    txtDesc.BackColor = cmdCancel.BackColor
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim strResponse As String
    
    If UnloadMode = 0 Then
        Cancel = True
    
        If cmdSave.Enabled Then
            strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", "Confirmation")
            Select Case UCase(strResponse)
                Case "Y"
                    m.bOK = True
                    Me.Hide
                Case "C"
                    Exit Sub
                Case "N"
                    m.bOK = False
                    Me.Hide
            End Select
        Else
            m.bOK = False
            Me.Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long
    Dim lMinHeight As Long
    
    lMinWidth = fraButtons.Width * 5
    lMinHeight = fraButtons.Height
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - lblDesc.Left, lblDesc.Top
    End With
    
    With vsGrid
        .Move lblDesc.Left, lblDesc.Height + lblDesc.Top, _
                ScaleWidth - fraButtons.Width - (lblDesc.Left * 3), _
                ScaleHeight - lblDesc.Height - txtDesc.Height - (lblDesc.Top * 4)
    End With
    
    With txtDesc
        .Move lblDesc.Left, vsGrid.Top + vsGrid.Height + lblDesc.Top, vsGrid.Width
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Set m.OptCols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    cmdSave.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.vsGrid.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Show) And Col <> GDCol(eGDCol_Filters) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.vsGrid.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim pt As POINTAPI
    Dim strOper As String
    Dim strFilterValue As String
    Dim strFormat As String
    
    'get popup window position
    If Col = GDCol(eGDCol_Filters) Then
        With vsGrid
            .Redraw = flexRDNone
            pt.X = .ColPos(Col) / Screen.TwipsPerPixelX
            pt.Y = (.RowPos(Row) + .RowHeight(Row)) / Screen.TwipsPerPixelY
            ClientToScreen .hWnd, pt
            
            strOper = .TextMatrix(Row, GDCol(eGDCol_Oper))
            strFilterValue = .TextMatrix(Row, GDCol(eGDCol_FilterValue))
            strFormat = .TextMatrix(Row, GDCol(eGDCol_Format))
        
            If frmOptFilters.ShowMe(pt.X, pt.Y, .TextMatrix(Row, GDCol(eGDCol_Stat)), _
                                    strOper, strFilterValue, strFormat) Then
                
                .TextMatrix(Row, GDCol(eGDCol_Oper)) = strOper
                .TextMatrix(Row, GDCol(eGDCol_FilterValue)) = strFilterValue
                .TextMatrix(Row, GDCol(eGDCol_Format)) = strFormat
                
                FormatFilters Row, .TextMatrix(Row, GDCol(eGDCol_Stat)), _
                    .TextMatrix(Row, GDCol(eGDCol_Oper)), .TextMatrix(Row, GDCol(eGDCol_FilterValue))
                
                cmdSave.Enabled = True
            End If
            .Redraw = flexRDBuffered
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.vsGrid.CellButtonClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    txtDesc.Text = vsGrid.TextMatrix(NewRow, GDCol(eGDCol_Desc))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptCustomize.vsGrid.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    InitGrid
    LoadGrid
    cmdSave.Enabled = False
    
    ShowForm Me, True, , , ALT_GRID_ROW_COLOR
    
    If m.bOK Then Save
    
    ShowMe = m.bOK
    Unload Me
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptCustomize.ShowMe", eGDRaiseError_Raise
    
End Function

