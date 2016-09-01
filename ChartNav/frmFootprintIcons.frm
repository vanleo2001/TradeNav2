VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFootprintIcons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Icon Palette"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdClose 
      Height          =   375
      Left            =   2145
      TabIndex        =   0
      Top             =   1020
      Width           =   795
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
      Caption         =   "frmFootprintIcons.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmFootprintIcons.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmFootprintIcons.frx":004C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraIcons 
      Height          =   1005
      Left            =   75
      TabIndex        =   2
      Top             =   495
      Width           =   3000
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
      Caption         =   "frmFootprintIcons.frx":0068
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFootprintIcons.frx":0098
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFootprintIcons.frx":00B8
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fg 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   540
         Width           =   1575
         _cx             =   2778
         _cy             =   556
         _ConvInfo       =   1
         Appearance      =   0
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   255
         ColWidthMin     =   255
         ColWidthMax     =   255
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1290
            Picture         =   "frmFootprintIcons.frx":00D4
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   1
            Top             =   18
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            Picture         =   "frmFootprintIcons.frx":01CE
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   12
            Top             =   18
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   258
            Picture         =   "frmFootprintIcons.frx":02C8
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   11
            Top             =   18
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   516
            Picture         =   "frmFootprintIcons.frx":03C2
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   10
            Top             =   18
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   774
            Picture         =   "frmFootprintIcons.frx":04BC
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   9
            Top             =   18
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1032
            Picture         =   "frmFootprintIcons.frx":05B6
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   8
            Top             =   18
            Width           =   255
         End
      End
      Begin HexUniControls.ctlUniRadioXP optAlign 
         Height          =   255
         Index           =   2
         Left            =   2055
         TabIndex        =   6
         Top             =   135
         Width           =   855
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
         Caption         =   "frmFootprintIcons.frx":06B0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmFootprintIcons.frx":06DA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFootprintIcons.frx":06FA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAlign 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   135
         Width           =   855
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
         Caption         =   "frmFootprintIcons.frx":0716
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmFootprintIcons.frx":0742
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFootprintIcons.frx":0762
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAlign 
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   135
         Width           =   855
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
         Caption         =   "frmFootprintIcons.frx":077E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmFootprintIcons.frx":07A6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFootprintIcons.frx":07C6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin VB.PictureBox SelectedPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2895
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   17
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin HexUniControls.ctlUniLabelXP lblInfoPrompt 
      Height          =   435
      Left            =   68
      Top             =   45
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
      Caption         =   "frmFootprintIcons.frx":07E2
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmFootprintIcons.frx":0896
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFootprintIcons.frx":08B6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmFootprintIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kIconFile = "FpIcon.bmp"

Private Enum eFPImage
    eFP_IMG_None = -1
    eFP_IMG_TriUp = 0
    eFP_IMG_TriDown = 1
    eFP_IMG_Dollar = 2
    eFP_IMG_Aster_Red = 3
    eFP_IMG_Aster_Blue = 4
    eFP_IMG_Diamond = 5
End Enum

Private Type mPrivate
    frmBidAsk As frmBidAskDir
    frmPriceVol As frmPriceVol
    
    eImgType As eFPImage
    iColor As Long
    iAlign As Long
End Type

Dim m As mPrivate

Public Sub ShowMe(frm As Form)
On Error GoTo ErrSection:
   
    If TypeOf frm Is frmBidAskDir Then
        Set m.frmBidAsk = frm
    ElseIf TypeOf frm Is frmPriceVol Then
        Set m.frmPriceVol = frm
    End If
    
    If m.eImgType = eFP_IMG_None Then m.eImgType = eFP_IMG_TriUp
    
    'RH commented out fraIcons.BorderStyle = 0
    
    ShowForm Me
    SetFormTopmost Me, True
    
    Exit Sub
    
ErrSection:
     RaiseError Me.Name & ".ShowMe"

End Sub

Public Sub CloseMe(frm As Form)

    Dim bUnload As Boolean
    
    If TypeOf frm Is frmBidAskDir Then
        Set m.frmBidAsk = Nothing
        If m.frmPriceVol Is Nothing Then bUnload = True
    ElseIf TypeOf frm Is frmPriceVol Then
        Set m.frmPriceVol = Nothing
        If m.frmBidAsk Is Nothing Then bUnload = True
    Else
        bUnload = True
    End If
    
    If bUnload Then Unload Me

End Sub

Private Sub clrColor_Changed()
    pic_Click (m.eImgType)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    g.Styler.StyleForm Me
    
    m.iAlign = flexPicAlignLeftCenter
    pic_Click (0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not m.frmBidAsk Is Nothing Then m.frmBidAsk.IconPaletteClose
    If Not m.frmPriceVol Is Nothing Then m.frmPriceVol.IconPaletteClose

    KillFile kIconFile, True
End Sub

Private Sub optAlign_Click(Index As Integer)

    If Index = 1 Then
        m.iAlign = flexPicAlignCenterCenter
    ElseIf Index = 2 Then
        m.iAlign = flexPicAlignRightCenter
    Else
        m.iAlign = flexPicAlignLeftCenter
    End If
    
End Sub

Private Sub pic_Click(Index As Integer)

    pic(m.eImgType).BorderStyle = 0
    
    m.eImgType = Index
    If m.eImgType = eFP_IMG_Aster_Blue Or m.eImgType = eFP_IMG_TriUp Then
        m.iColor = vbBlue
    ElseIf m.eImgType = eFP_IMG_Aster_Red Or m.eImgType = eFP_IMG_TriDown Then
        m.iColor = vbRed
    ElseIf m.eImgType = eFP_IMG_Dollar Then
        m.iColor = RGB(0, 128, 0)
    Else
        m.iColor = 0
    End If
    
    pic(Index).BorderStyle = 1
    
    SelectedPic.Picture = pic(Index)
    
End Sub

Public Property Get IconAlign() As Long
    IconAlign = m.iAlign
End Property

Public Property Get IconColor() As Long
    IconColor = m.iColor
End Property

Public Property Get SelectedPicIndex() As Long
    SelectedPicIndex = m.eImgType
End Property

Public Property Get IconTypeNum(ByVal strIconType$) As Long

    Dim eType As eFPImage
    
    If strIconType = "TRIUP" Then
        eType = eFP_IMG_TriUp
    ElseIf strIconType = "TRIDOWN" Then
        eType = eFP_IMG_TriDown
    ElseIf strIconType = "DOLLAR" Then
        eType = eFP_IMG_Dollar
    ElseIf strIconType = "ASTERISKRED" Then
        eType = eFP_IMG_Aster_Red
    ElseIf strIconType = "ASTERISKBLUE" Then
        eType = eFP_IMG_Aster_Blue
    ElseIf strIconType = "DIAMOND" Then
        eType = eFP_IMG_Diamond
    Else
        eType = eFP_IMG_None
    End If
    
    IconTypeNum = eType

End Property

Public Property Get IconTypeStr(ByVal iIconType&) As String

    Dim strType$
    
    Select Case iIconType
        Case eFP_IMG_TriUp
            strType = "TRIUP"
        Case eFP_IMG_TriDown
            strType = "TRIDOWN"
        Case eFP_IMG_Dollar
            strType = "DOLLAR"
        Case eFP_IMG_Aster_Red
            strType = "ASTERISKRED"
        Case eFP_IMG_Aster_Blue
            strType = "ASTERISKBLUE"
        Case eFP_IMG_Diamond
            strType = "DIAMOND"
    End Select

    IconTypeStr = strType

End Property

