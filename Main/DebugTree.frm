VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Begin VB.Form frmDebugTree 
   Caption         =   "Tree Display"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   3780
   Begin VSFlex7LCtl.VSFlexGrid fgTree 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _cx             =   5530
      _cy             =   4895
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
Attribute VB_Name = "frmDebugTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
On Error GoTo ErrSection:

    ToolbarSync Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDebugTree.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
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
    RaiseError "frmDebugTree.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Width = fgTree.Width + 500

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDebugTree.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    fgTree.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Public Sub LoadTree(Tree As cGdTree)
On Error GoTo ErrSection:

    Dim idx&, strItem$, strType$, strSep$

    With fgTree
        .Redraw = flexRDNone
        .OutlineBar = flexOutlineBarComplete
        .GridLines = flexGridNone
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = True
        
        .Cols = 1
        '.TextMatrix(0, 0) = "Node"
        '.TextMatrix(0, 1) = "Key"
        '.TextMatrix(0, 2) = "Type"
        '.TextMatrix(0, 3) = "Name"
        .ExtendLastCol = True
        .FixedCols = 1
        .BackColorFixed = vbWhite
        
        strSep = "  -  "
        For idx = 1 To Tree.Count
            .Rows = idx + 1
            .IsSubtotal(idx) = True
            .RowOutlineLevel(idx) = Tree.NodeLevel(idx)
            strItem = ""
            On Error Resume Next
            strItem = Tree(idx).Name & strSep
            On Error GoTo 0
            strItem = strItem & Tree.Key(idx)
            strType = Tree.NodeType(idx)
            If Len(strType) > 0 Then
                strItem = strItem & strSep & strType
            End If
            strItem = strItem & "  (" & Trim(Str(idx)) & ")"
            .TextMatrix(idx, 0) = strItem
        Next
        
        .Outline -1
        '.AutoSize 0, 2
        .Redraw = flexRDBuffered
    End With
    
    'If Not Me.Visible Then Me.Show
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDebugTree.LoadTree", eGDRaiseError_Raise
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    ToolbarSync Me, False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDebugTree.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub
