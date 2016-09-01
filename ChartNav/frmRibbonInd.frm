VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmRibbonInd 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   982
      TabIndex        =   0
      Top             =   4080
      Width           =   750
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
      Caption         =   "frmRibbonInd.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmRibbonInd.frx":0026
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmRibbonInd.frx":0046
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   330
      Left            =   1942
      TabIndex        =   2
      Top             =   4080
      Width           =   750
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
      Caption         =   "frmRibbonInd.frx":0062
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmRibbonInd.frx":0090
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmRibbonInd.frx":00B0
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3015
      Left            =   330
      TabIndex        =   1
      Top             =   840
      Width           =   3015
      _cx             =   5318
      _cy             =   5318
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
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   255
      Left            =   330
      Top             =   600
      Width           =   3015
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
      Caption         =   "frmRibbonInd.frx":00CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmRibbonInd.frx":0140
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRibbonInd.frx":0160
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   495
      Left            =   150
      Top             =   120
      Width           =   3375
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
      Caption         =   "frmRibbonInd.frx":017C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmRibbonInd.frx":0224
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRibbonInd.frx":0244
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmRibbonInd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    nMouseRow As Long
    bOK As Boolean
End Type

Private m As mPrivate

Public Function ShowMe(Chart As cChart, aIndNames As cGdArray) As cIndicator

    Dim i&, idx&, str1$
    Dim Ind As cIndicator

    If aIndNames Is Nothing Then Exit Function
    If aIndNames.Size = 0 Then Exit Function
    
    
    aIndNames.Sort eGdSort_Default

    m.nMouseRow = -1
    
    With fg
        .Sort = flexSortNone
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .FixedRows = 0
        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightWithFocus
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AllowSelection = True
        
        .Cols = 1
        .Rows = 0
        
        .ExtendLastCol = True
        
        For i = 0 To aIndNames.Size - 1
            str1 = Parse(aIndNames(i), "|", 1)
            idx = Val(Parse(aIndNames(i), "|", 2))
            
            If Len(str1) > 0 And idx > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = str1
                .RowData(i) = idx
            End If
        Next
        
    End With
    
    CenterFormOnChart Me, Chart                 '6499
    ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        If fg.Row >= 0 And fg.Row < fg.Rows And Not Chart Is Nothing Then
            i = fg.RowData(fg.Row)
            If i > 0 And i <= Chart.Tree.Count Then
                If Chart.Tree.NodeLevel(i) > 0 Then
                    Set Ind = Chart.Tree(i)
                    If Not Ind Is Nothing Then
                        StatusMsg "Changing " & Ind.Name & " to ribbon style"
                        Ind.DisplayType = eINDIC_Ribbon
                    End If
                End If
            End If
        End If
    
        Unload Me
    End If
    
    Set ShowMe = Ind

End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next

    m.bOK = True
    Me.Hide

End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim idx&

    If fg.Row >= 0 And fg.Row < fg.Rows Then
        m.nMouseRow = fg.Row
        cmdOK_Click
    End If
    
End Sub

Private Sub Form_Load()
    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
    
End Sub

