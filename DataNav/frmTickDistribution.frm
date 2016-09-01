VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTickDistribution 
   AutoRedraw      =   -1  'True
   Caption         =   "Tick Distribution"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox IconPicAvgEntry 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6810
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   17
      TabIndex        =   45
      Top             =   1650
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox IconPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6795
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   17
      TabIndex        =   44
      Top             =   1020
      Visible         =   0   'False
      Width           =   255
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrdersInfo 
      Height          =   675
      Left            =   4995
      TabIndex        =   40
      Top             =   90
      Width           =   2010
      _cx             =   3545
      _cy             =   1191
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
   Begin vsOcx6LibCtl.vsElastic vseOrderBar 
      Height          =   4035
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
      Begin HexUniControls.ctlUniFrameWL fraFrontMonth 
         Height          =   1425
         Left            =   0
         TabIndex        =   21
         Top             =   2520
         Width           =   7100
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
         Caption         =   "frmTickDistribution.frx":0000
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistribution.frx":002C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistribution.frx":004C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdFrontMonth 
            Height          =   390
            Left            =   4935
            TabIndex        =   22
            Top             =   427
            Width           =   1250
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
            Caption         =   "frmTickDistribution.frx":0068
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":00A6
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":00C6
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblGoTo 
            Height          =   255
            Left            =   1410
            Top             =   840
            Visible         =   0   'False
            Width           =   1250
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
            Caption         =   "frmTickDistribution.frx":00E2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":010C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":012C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblFrontMonthVert 
            Height          =   1470
            Left            =   4920
            Top             =   0
            Width           =   1200
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
            Caption         =   "frmTickDistribution.frx":0148
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":024C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":026C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblFrontMonthHorz 
            Height          =   420
            Left            =   0
            Top             =   360
            Width           =   4275
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
            Caption         =   "frmTickDistribution.frx":0288
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":038C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":03AC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraOrderBtns 
         Height          =   2430
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   7450
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
         Caption         =   "frmTickDistribution.frx":03C8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistribution.frx":0400
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistribution.frx":0420
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdBrokerConnect 
            Height          =   375
            Left            =   3840
            TabIndex        =   19
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
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
            Caption         =   "frmTickDistribution.frx":043C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":046A
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":048A
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraExitFavorites 
            Height          =   495
            Left            =   4800
            TabIndex        =   20
            Top             =   1320
            Visible         =   0   'False
            Width           =   1365
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
            Caption         =   "frmTickDistribution.frx":04A6
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistribution.frx":04E6
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":0506
            RightToLeft     =   0   'False
            Begin vsOcx6LibCtl.vsElastic vseExitA 
               Height          =   375
               Left            =   30
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "A"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin vsOcx6LibCtl.vsElastic vseExitB 
               Height          =   375
               Left            =   350
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "B"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin vsOcx6LibCtl.vsElastic vseExitC 
               Height          =   375
               Left            =   670
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "C"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
            Begin vsOcx6LibCtl.vsElastic vseExitD 
               Height          =   375
               Left            =   990
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   60
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   600
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "D"
               Align           =   0
               Appearance      =   1
               AutoSizeChildren=   0
               BorderWidth     =   3
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   4
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   0
               TagPosition     =   0
               Style           =   0
               TagSplit        =   0   'False
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   0
               GridCols        =   0
               _GridInfo       =   ""
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraRithmicSmall 
            Height          =   315
            Left            =   4620
            TabIndex        =   37
            Top             =   1980
            Visible         =   0   'False
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":0522
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistribution.frx":0560
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":0580
            RightToLeft     =   0   'False
            Begin VB.Image imgOmne 
               Height          =   105
               Left            =   0
               Picture         =   "frmTickDistribution.frx":059C
               Stretch         =   -1  'True
               Top             =   210
               Width           =   1110
            End
            Begin VB.Image imgRithmic 
               Height          =   180
               Left            =   0
               Picture         =   "frmTickDistribution.frx":08B3
               Top             =   0
               Width           =   1005
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraRithmic 
            Height          =   345
            Left            =   360
            TabIndex        =   39
            Top             =   1830
            Visible         =   0   'False
            Width           =   3975
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
            Caption         =   "frmTickDistribution.frx":0AC9
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistribution.frx":0AF5
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":0B15
            RightToLeft     =   0   'False
            Begin VB.PictureBox picPbo 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   2100
               Picture         =   "frmTickDistribution.frx":0B31
               ScaleHeight     =   210
               ScaleWidth      =   1830
               TabIndex        =   43
               Top             =   120
               Width           =   1830
            End
            Begin VB.PictureBox picRithmic 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   0
               Picture         =   "frmTickDistribution.frx":103F
               ScaleHeight     =   345
               ScaleWidth      =   1995
               TabIndex        =   46
               Top             =   0
               Width           =   1995
            End
         End
         Begin vsOcx6LibCtl.vsElastic vseBracketOrder 
            Height          =   300
            Left            =   3000
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1470
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   1
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   600
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Picture         =   "frmTickDistribution.frx":12CE
            Caption         =   ""
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   0   'False
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
         End
         Begin HexUniControls.ctlUniCheckXP chkAutoJournal 
            Height          =   195
            Left            =   5765
            TabIndex        =   41
            Top             =   705
            Width           =   1275
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
            Caption         =   "frmTickDistribution.frx":2468
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":24A0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":257E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboExchanges 
            Height          =   315
            Left            =   380
            TabIndex        =   38
            Top             =   1140
            Width           =   1290
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
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
            Tip             =   "frmTickDistribution.frx":259A
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":261A
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAutoExit 
            Height          =   255
            Left            =   380
            TabIndex        =   35
            Top             =   1500
            Width           =   1080
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
            Caption         =   "frmTickDistribution.frx":2636
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":266A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":268A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQty1 
            Height          =   300
            Left            =   3000
            TabIndex        =   30
            Top             =   840
            Width           =   480
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
            Caption         =   "frmTickDistribution.frx":26A6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":26C8
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":26E8
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQty2 
            Height          =   300
            Left            =   3480
            TabIndex        =   29
            Top             =   840
            Width           =   480
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
            Caption         =   "frmTickDistribution.frx":2704
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2726
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2746
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdQty3 
            Height          =   300
            Left            =   3960
            TabIndex        =   28
            Top             =   840
            Width           =   480
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
            Caption         =   "frmTickDistribution.frx":2762
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2788
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":27A8
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdReverse 
            Height          =   315
            Left            =   4430
            TabIndex        =   27
            Top             =   900
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":27C4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":27F2
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2812
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboOrderType 
            Height          =   315
            Left            =   380
            TabIndex        =   26
            Top             =   660
            Width           =   1290
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
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
            Tip             =   "frmTickDistribution.frx":282E
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":28AE
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkConfirmOrder 
            Height          =   195
            Left            =   5765
            TabIndex        =   25
            Top             =   1020
            Width           =   1395
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
            Caption         =   "frmTickDistribution.frx":28CA
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2906
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2926
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboAccounts 
            Height          =   315
            Left            =   380
            TabIndex        =   24
            Top             =   180
            Width           =   1290
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ButtonBackColor =   -2147483633
            ButtonForeColor =   -2147483630
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
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
            Tip             =   "frmTickDistribution.frx":2942
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2962
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCancelAll 
            Height          =   315
            Left            =   5765
            TabIndex        =   18
            Top             =   0
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":297E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":29B2
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":29F4
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdClearQty 
            Height          =   360
            Left            =   3035
            TabIndex        =   17
            Top             =   495
            Width           =   235
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
            Caption         =   "frmTickDistribution.frx":2A10
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2A32
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2A52
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBuyBid 
            Height          =   300
            Left            =   1685
            TabIndex        =   16
            Top             =   600
            Visible         =   0   'False
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":2A6E
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2AA2
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2AF6
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBuyMarket 
            Height          =   300
            Left            =   1685
            TabIndex        =   15
            Top             =   0
            Width           =   1335
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
            Caption         =   "frmTickDistribution.frx":2B12
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2B46
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2B66
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBuyAsk 
            Height          =   300
            Left            =   1685
            TabIndex        =   14
            Top             =   300
            Visible         =   0   'False
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":2B82
            BackColor       =   16752800
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2BB6
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2C0A
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSellAsk 
            Height          =   300
            Left            =   4430
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":2C26
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2C5C
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2CB2
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSellMarket 
            Height          =   300
            Left            =   4430
            TabIndex        =   12
            Top             =   0
            Width           =   1335
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
            Caption         =   "frmTickDistribution.frx":2CCE
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2D04
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2D24
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSellBid 
            Height          =   300
            Left            =   4430
            TabIndex        =   11
            Top             =   300
            Visible         =   0   'False
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":2D40
            BackColor       =   10526975
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistribution.frx":2D76
            Style           =   1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2DCC
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtTradeQty 
            Height          =   360
            Left            =   3240
            TabIndex        =   10
            Top             =   495
            Width           =   960
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTickDistribution.frx":2DE8
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
            Alignment       =   2
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmTickDistribution.frx":2E08
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2E28
         End
         Begin gdOCX.gdScrollBar vscrQty 
            Height          =   360
            Left            =   4200
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   480
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   635
         End
         Begin vsOcx6LibCtl.vsElastic cmdBailout 
            Height          =   375
            Left            =   5765
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Cancel all orders and exit current position"
            Top             =   300
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   1
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   600
            BackColor       =   192
            ForeColor       =   65535
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Caption         =   "FLATTEN"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   5
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   0   'False
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
         End
         Begin HexUniControls.ctlUniLabelXP lblBrokerDisconnect 
            Height          =   660
            Left            =   3435
            Top             =   1260
            Width           =   3510
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
            Caption         =   "frmTickDistribution.frx":2E44
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":2EB0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2ED0
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP lblExchange 
            Height          =   255
            Left            =   380
            Top             =   960
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":2EEC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":2F1E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2F3E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAutoExit 
            Height          =   615
            Left            =   1740
            Top             =   1035
            Width           =   1155
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
            Caption         =   "frmTickDistribution.frx":2F5A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":2F82
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":2FA2
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP lblTradePos 
            Height          =   225
            Left            =   3035
            Top             =   60
            Width           =   1395
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
            Caption         =   "frmTickDistribution.frx":2FBE
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   1
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":2FE6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":3006
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblEquity 
            Height          =   225
            Left            =   3035
            Top             =   300
            Width           =   1395
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
            Caption         =   "frmTickDistribution.frx":3022
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   1
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":3042
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":3062
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOrderType 
            Height          =   315
            Left            =   375
            Top             =   480
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":307E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":30A8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":30C8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAccounts 
            Height          =   255
            Left            =   380
            Top             =   0
            Width           =   1290
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
            Caption         =   "frmTickDistribution.frx":30E4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistribution.frx":3124
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistribution.frx":3144
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgDOMPrint 
      Height          =   555
      Left            =   660
      TabIndex        =   6
      Top             =   3660
      Visible         =   0   'False
      Width           =   1275
      _cx             =   2249
      _cy             =   979
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
   Begin VSFlex7LCtl.VSFlexGrid fgQuoteBar 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1515
      _cx             =   2672
      _cy             =   661
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
   Begin VB.PictureBox pbAsk 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4200
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox pbBid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAskDetail 
      Height          =   855
      Left            =   3060
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   1935
      _cx             =   3413
      _cy             =   1508
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
   Begin VSFlex7LCtl.VSFlexGrid fgBidDetail 
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _cx             =   3201
      _cy             =   1085
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
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5640
      Top             =   840
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTickDistribution 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      _cx             =   4683
      _cy             =   5847
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   5980
      Top             =   1900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   9
      DisplayContextMenu=   0   'False
      Tools           =   "frmTickDistribution.frx":3160
      ToolBars        =   "frmTickDistribution.frx":7E97
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAccountBar 
      Height          =   375
      Left            =   5280
      TabIndex        =   34
      Top             =   2700
      Visible         =   0   'False
      Width           =   1515
      _cx             =   2672
      _cy             =   661
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
Attribute VB_Name = "frmTickDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTickDistribution.frm
'' Description: Form to show the price ladder or depth of market views
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'' Source Safe Revisions prior to implementing order bar on Depth of Market View
'' frmTickDistribution.frm      362     02/15/08 12:56p
'' frmTickDistribution.frx       49     11/06/07  2:11p
'' frmTickDistributionCfg.frm    59     02/06/08  4:02p
'' frmMarketDepthCfg.frm          1     02/03/06  3:10p
'' mChartLadderCtrl.bas           5     02/07/08  5:02p
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 08/21/2009   DAJ         Set UserCancel flag on CancelOrder call
'' 09/01/2009   DAJ         Use new Parked order status
'' 03/11/2010   DAJ         Moved global collections off of Trade Console
'' 09/29/2010   DAJ         Changed refrence to global order confirmation flag
'' 04/20/2012   DAJ         Mods for broker view mode
'' 05/01/2012   DAJ         Further mods for broker view mode
'' 05/03/2012   DAJ         Further mods for broker view mode
'' 05/10/2012   DAJ         Further mods for broker view mode
'' 09/27/2012   DAJ         Renamed Ladder_ChangeAccount to ChangeAccount
'' 10/23/2012   DAJ         Fix for infinite FormResize loop when not connected to broker
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 06/24/2013   DAJ         Timer Logging
'' 09/26/2013   DAJ         Added logging for when the position label changes
'' 03/02/2016   DAJ         Made order preset buttons a little bit wider
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kCols = 11
Private Const eCol_BidAskChange = 4         'hidden column stores -1,0,1 for drawing triangle OR percent value for flooding cell
Private Const kBidColHeader = "Maker,Bid,Size,Time"
Private Const kAskColHeader = "Maker,Ask,Size,Time"
Private Const kContinuousCaption = "Trading a Continuous Contract is not currently supported.  Click Active Contract to change to the active contract."
'default depth of market colors
Private Const kFirstColor = 8454143 'vbYellow
Private Const kSecondColor = 12648384 'vbGreen
Private Const kThirdColor = 16744703 '16776960      'bright light blue
Private Const kFourthColor = 16777088 'vbRed
Private Const kFifthColor = 8438015 '12255350       'light purple
Private Const kOtherColor = 16777215 'vbWhite
Private Const kInactiveColor = 16777196             'light, light blue
Private Const kBidAskUpColor = vbGreen
Private Const kBidAskDownColor = vbRed
Private Const kGridFloodColor = 12640511
Private Const kAvgEntryColor = 11665140     'light, light yellow
'default columns for quote bar
Private Const kQBarCols = "Symbol|Bid|Bid Size|Ask|Ask Size|Trade|Trade Size|Open|High|Low|Close"
'fixed/minimum sizes for price ladder columns
Private Const kMinColWidth = 950
Private Const kMinColWidthExt = 1350
Private Const kWidthOrderX = 300
'order bar's height
'Private Const kOBHt = 1840              'none or some bid/ask order buttons not visible
'Private Const kOBHtWithBidAsk = 1840    'all bid/ask order buttons visible
'Private Const kOBHtVert = 2350          'height of buttons frame when form width too narrow (all buttons visible)
Private Const kMinFormSize = 4200
Private Const kSizeToSwitch = 6770      'reposition buttons when width of form is this size
Private Const kAutoExitOnRight = 615    'height of autoexit label when order bar is on right
Private Const kOwnePicLeft = 1020       'left of powered by omne logo
Private Const kOBFrameRight = 1455      'width of orderbar when placed on right side of grid

'columns in tick distribution grid
Private Enum eTDCols
    eTDCols_Volume = 0
    eTDCols_OrderBidX
    eTDCols_OrderBid
    eTDCols_BidSize
    eTDCols_Price
    eTDCols_AskSize
    eTDCols_OrderAsk
    eTDCols_OrderAskX
    eTDCols_Entries
    eTDCols_PL
    eTDCols_HasOrder
End Enum

'fields in data table
Private Enum eDataFields
    eFld_PL = 0
    eFld_BidSize
    eFld_Price
    eFld_AskSize
    eFld_Volume
    eFld_PriceStr
    eFld_BarIdx
End Enum

'action to perform for buy/sell orders at same price
Private Enum eSamePriceAction
    eOrdAction_Unknown = -1         'user was never prompted
    eOrdAction_Consolidate = 0      'consolidate into one order
    eOrdAction_CancelExisting       'cancel existing order and place new one
    eOrdAction_None                 'no change (do nothing)
End Enum

Private Enum eDisplayStyle
    eView_Ladder = 0
    eView_Detail
End Enum

Public Enum eGDOrderBarMode
    eGDOrderBarMode_LastShownBottom = -2
    eGDOrderBarMode_LastShownOnRight = -1
    eGDOrderBarMode_NotShown = 0
    eGDOrderBarMode_BottomWide
    eGDOrderBarMode_BottomNarrow
    eGDOrderBarMode_BottomContinuous
    eGDOrderBarMode_Right
End Enum

Public Enum eBidAskColorMode
    eBidAskColorMode_None = 0
    eBidAskColorMode_ByPrice        'best Bid/Ask
    eBidAskColorMode_BySize         'largest sized Bid/Ask
End Enum

Public Enum eLadderVolStyle         'option for volume display in price column of ladder
    eLadderVol_LastTrade
    eLadderVol_BidAsk
    eLadderVol_None
End Enum

Private Type mPrivate
    WindowLink As New cWindowLink

    TickBars As New cGdBars     ' (all ticks for the day)
    DailyBar As New cGdBars     ' (to get OHLC and bid/ask data)
    Data As New cGdTable
    tbOrders As New cGdTable
    tbOutlineCells As New cGdTable      'holds list of prices that should have their cells outline with user-specified color
    oBidAskDepth As cBidAskDepth
    aBidColHeader As New cGdArray
    aAskColHeader As New cGdArray
    aQBarColHeader As New cGdArray
    aABarColHeader As New cGdArray
    
    dMaxVol As Double
    dMinPrice As Double
    dMaxPrice As Double
    dOpen As Double
    dHigh As Double
    dLow As Double
    dLastPrice As Double
    dLastTradeAtAsk As Double
    dLastTradeAtBid As Double
    dLastPriceVol As Double
    dMinMove As Double
    dLastUpdateTime As Double
    
    'record number in table for last price & zero PL
    nLastPriceRec As Long
    dZeroPLPrice As Double
    nQuantity As Long
    'values for account bar
    dAvgEntry As Double
    strPos As String
    strPosQty As String
    strOpenEq As String
    strAvgEntry As String
    strSessionPL As String
    strSessionQty As String
    
    'row numbers in grid for rows that need updating
    nOpenPriceRow As Long
    nLastPriceRow As Long
    nPrevBidRow As Long
    nPrevAskRow As Long
    nPrevAvgEntryRow As Long
    nLastTradeAtAskRow As Long
    nLastTradeAtBidRow As Long

    nSymID As Long
    nSessionDate As Long
    nFontSize As Long
    nFloodColor As Long         'for volume column
    nBarColor As Long
    nUpColor As Long            'for coloring price column to look like bollinger bar
    nDownColor As Long
    nCurrUpDownColor As Long
    nBidColor As Long
    nAskColor As Long
    nBidTextColor As Long
    nAskTextColor As Long
    nFixedPriceColor As Long    'for one-color background
    nTickLineColor As Long
    nShowVolBar As Long
    nShowVolText As Long
    nShowTickLine As Long
    nShowProfitLoss As Long
    nShowAvgEntry As Long
    nShowOpenEntries As Long
    nTickLineRL As Long             '0=draw left to right, 1=draw right to left
    nBlankRows As Long              'number of rows above/below high/low on price ladder (default=30)
    nSessionBlankRows As Long       'need this for auto-extending without changing saved blank rows value
    bBarColorIsLight As Boolean
    nShowVolMin As Long             'filter for volume
    nShowVolMax As Long
    eVolumeStyle As eLadderVolStyle
    
    nFirstColor As Long             'market depth colors
    nSecondColor As Long
    nThirdColor As Long
    nFourthColor As Long
    nFifthColor As Long
    nOtherColor As Long
    nLargestSizeColor As Long
    nInactiveColor As Long
    nBidAskUpColor As Long
    nBidAskDownColor As Long
    nAvgEntryColor As Long
    nOutlineColor As Long                   'color last used for outlining price cell
    nDrawTriangle As Long                   '0=flood cell, 1=draw triangle
    eFloodMktDepth As eBidAskColorMode      'options for flood price ladder with market depth colors
    nShowSummaryBar As Long
    nVertSummaryBar As Long
    nQBarSumColWidth As Long
    nABarSumColWidth As Long
    
    nShowQuoteBar As Long
    nShowAccountBar As Long
    nShowAcctBarSave As Long        'in brokerview mode nShowAccountBar bar is overridden (=0) to not show
    nOrderColumns As Long           '0=use one column for all orders, 1=separate buy/sell orders into 2 columns
    
    eSamePriceOrdAction As eSamePriceAction
    strOrdBarCtrls As String
    strOrdBarCtrlsSave As String
    nHighlightPos As Long           '(-2)was never set/saved, (< 0):don't highlight, (>= 0):highlight with specified color
    nHighlightEquity As Long
    
    nTradeAcctID As Long
    nDragOrderRow As Long
    nBracketStopID As Long          'bracket orders
    nBracketLimitID As Long
    dBestBid As Double
    dBestAsk As Double
    nBestBidSize As Long            'size of bids (cumulative for sessions & may not be trades)
    nBestAskSize As Long
    nTradeAtAskSize As Long         'size of trades at ask (reset when ask/bid price changes)
    nTradeAtBidSize As Long
    
    strFont As String
    strSym As String
    bIsSpreadSymbol As Boolean
    
    bTimerInProg As Boolean
    bGridRTInProg As Boolean
    bSessionCurrent As Boolean
    bHeaderChanged As Boolean
    bUnloading As Boolean
    bPrinting As Boolean
    bStatusBusy As Boolean
    bLadderHasDOM As Boolean        'true=DOM visible on ladder
    bHideLadderDOM As Boolean
    bMinutized As Boolean
    bDepthOfMarketBad As Boolean
    bAutosizePrice As Boolean
    
    eDisplayStyle As eDisplayStyle
    
    nLadderButton As Long           'save left or right mouse button user clicked in Price Ladder
    nPrevMinute As Long
    geTickObj As Long
    
    bSettingAutoExit As Boolean
    
    bIgnoreClick As Boolean
    bEnableAutoCenter As Boolean
    bUserSetSize As Boolean         'aardvark 4050
    nPriceColWidth As Long          'aardvark 3655
    
    eOrderBarMode As eGDOrderBarMode

    'for bracket order
    oBracketOrdOne As cPtOrder
    oBracketOrdTwo As cPtOrder
    
    nIntervalRT As Long             'for timer speed when streaming is on
    dScrollTickCount As Double      'for not centering ladder within certain number of seconds of user's scroll
    
    bDumpProfileRT As Boolean       'for dumping profile info for price ladder refresh in timer
    frmBroker As frmBrokerView      ' Broker form if running in broker mode
    
    Quantity As cPriceEditor            ' Management for the quantity controls
    lPreset1 As Long                    ' Quantity preset 1
    lPreset2 As Long                    ' Quantity preset 2
    lPreset3 As Long                    ' Quantity preset 3
End Type

Private m As mPrivate

Public Sub ShowMe(ByVal nSymbolID&, ByVal nView&, Optional ByVal bSkipMessage As Boolean = False, Optional frmBroker As frmBrokerView = Nothing)
On Error GoTo ErrSection:
        
    Dim nPointerSave&, nAutoCenter&, nOrderType&, s$
        
    nPointerSave = Screen.MousePointer
    
    m.nSymID = nSymbolID
    m.strSym = GetSymbol(nSymbolID)
    m.bIsSpreadSymbol = IsSpreadSymbol(m.strSym)
    
    m.bLadderHasDOM = False
    m.bDepthOfMarketBad = False
    m.nPrevAvgEntryRow = 0
    m.bAutosizePrice = True
    Set m.frmBroker = frmBroker
    
    lblBrokerDisconnect.Visible = False
    cmdBrokerConnect.Visible = False
    
    ' TLB 2/23/2009: per Glen, auto-center is now ok to do again for everyone
    m.bEnableAutoCenter = True
    'If FileExist("AutoLadder.flg") Then
    '    If DirExist("c:\common") Then m.bEnableAutoCenter = True
    'End If
    
    If nView = 0 Then
        m.eDisplayStyle = eView_Ladder
        Me.Icon = Picture16(ToolbarIcon("ID_TickDistribution"), , True)
        nAutoCenter = GetIniFileProperty("AutoCenter", ssUnchecked, "Price Ladder", g.strIniFile)
        tbToolbar.Tools("ID_CenterPrice").Visible = True
        tbToolbar.Tools("ID_CenterPrice").State = nAutoCenter
        tbToolbar.Tools("ID_VolHistogram").State = ShowVolumeBar
    Else
        m.eDisplayStyle = eView_Detail
        If g.RealTime.Active Then
            ChangeSymbol m.nSymID
            If 0 = IsEnableDOM() Then
                Unload Me
                Exit Sub
            End If
        Else
            InfBox "Real time needs to be on for depth of market.", "I", , "Depth of Market"
            Unload Me
            Exit Sub
        End If
        tbToolbar.Tools("ID_CenterPrice").Visible = False
    End If
    
    s = "Classic"
    If g.nTbIconStyle = 1 Then
        If g.nColorTheme = kDarkThemeColor Then
            s = "Light"
        Else
            s = "Dark"
        End If
        tbToolbar.Tools("ID_CenterPrice").Picture = g.CoreBridge.ImgListToolbarExt(s, "kCenterPrice", "", 16).ExtractIcon
        tbToolbar.Tools("ID_VolHistogram").Picture = g.CoreBridge.ImgListToolbarExt(s, "kLadderVolume", "", 16).ExtractIcon
    End If
    tbToolbar.Tools("ID_Symbol").Picture = g.CoreBridge.ImgListToolbarExt(s, ToolbarIcon("ID_Symbol"), "", 16).ExtractIcon
    tbToolbar.Tools("ID_Settings").Picture = g.CoreBridge.ImgListToolbarExt(s, ToolbarIcon("ID_Settings"), "", 16).ExtractIcon
    tbToolbar.Tools("ID_TextIncrease").Picture = g.CoreBridge.ImgListToolbarExt(s, ToolbarIcon("ID_TextIncrease"), "", 16).ExtractIcon
    tbToolbar.Tools("ID_TextDecrease").Picture = g.CoreBridge.ImgListToolbarExt(s, ToolbarIcon("ID_TextDecrease"), "", 16).ExtractIcon

    InitQuantityEditor
    
    Screen.MousePointer = vbHourglass
    ShowForm Me, eForm_Nonmodal, frmMain
    FormResize Me
    DoEvents
    
    GetData bSkipMessage
    m.WindowLink.Init Me
    If m.eOrderBarMode > eGDOrderBarMode_NotShown Then
        InitCboAccount
        If AllowMIT() Then
            cboOrderType.AddItem "MIT"
        Else
            nOrderType = ValOfText(GetIniFileProperty("PseudoOrderType", 0, "Price Ladder", g.strIniFile))
            If nOrderType > 2 Then nOrderType = 0
        End If
        cboOrderType.ListIndex = nOrderType
        CheckBoxValue(chkConfirmOrder) = g.Broker.ConfirmManual
        
        SetAutoExit
        
        If Not g.RealTime.Active Then
            UpdateEquityPos
            tmr.Enabled = True
            tmr.Interval = 500
        End If
        If m.nShowAccountBar Then FixAcctBarHeader
        
        GetContractInformation
        
        tbToolbar.Tools("ID_OrderBar").State = ssChecked
    Else
        tbToolbar.Tools("ID_OrderBar").State = ssUnchecked
    End If
    
    Screen.MousePointer = nPointerSave

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.ShowMe", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExchanges
'' Description: Load the exchange combo and set the default exchange
'' Inputs:      Exchange List, Default Exchange
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetExchanges(ByVal strExchangeList As String, ByVal strDefaultExchange As String)
On Error GoTo ErrSection:

    Dim astrExchanges As cGdArray       ' Array of exchanges
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Split out the exchange information into an array...
    Set astrExchanges = New cGdArray
    astrExchanges.SplitFields strExchangeList, ","
    
    With cboExchanges
        .Clear
        
        ' Load up combo box with the comma delimited list of exchanges in strExchangeList...
        For lIndex = 0 To astrExchanges.Size - 1
            If Len(astrExchanges(lIndex)) > 0 Then
                cboExchanges.AddItem astrExchanges(lIndex)
            End If
        Next lIndex
        
        ' Set the combo box to the string passed in strDefaultExchange if it exists...
        If Len(strDefaultExchange) > 0 Then
            For lIndex = 0 To cboExchanges.ListCount - 1
                If cboExchanges.List(lIndex) = strDefaultExchange Then
                    cboExchanges.ListIndex = lIndex
                    Exit For
                End If
            Next lIndex
        End If
    End With
    
    InfBox ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.SetExchanges"
    
End Sub

Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    Dim i&, nAccountID&
    
    If Me.Visible Then
        nAccountID = TradeAccountID
        
        If m.frmBroker Is Nothing Then
            With cboAccounts
                If .ListIndex >= 0 Then
                    If TradeAccountID <> .ItemData(.ListIndex) Then
    
'JM 10-04-2011: see note in chart.frm for code change explanation.
'                       If ChangeAccountCombo(cboAccounts.Text) Then
    
                        If True Then
                            TradeAccountID = .ItemData(.ListIndex)
                            m.dAvgEntry = 0
                            m.dZeroPLPrice = -1
                            UpdateEquityPos
                            SetAutoExit
                            GetContractInformation
                        Else
                            For i = 0 To .ListCount - 1
                                If .ItemData(i) = TradeAccountID Then
                                    .ListIndex = i
                                    Exit For
                                End If
                            Next
                        End If
                        
                        ' Call the Form_Resize to hide/show exchange controls appropriately...
                        Form_Resize
                    End If
                End If
            End With
        Else
            TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
            m.frmBroker.ChangeAccount cboAccounts.ItemData(cboAccounts.ListIndex), "price ladder"
        End If
        
        If TradeAccountID <> nAccountID Then
            InitQuantityEditor
        End If
            
        MoveFocus fgTickDistribution
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cboAccounts_Click"
    
End Sub

Private Sub cboOrderType_Click()
On Error GoTo ErrSection:

    SetIniFileProperty "PseudoOrderType", Str(cboOrderType.ListIndex), "Price Ladder", g.strIniFile

    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cboOrderType_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAutoExit_Click
'' Description: When the user changes the auto exit, sync up the world
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAutoExit_Click()
On Error GoTo ErrSection:

    Dim strAutoExit As String           ' Auto Exit selected
    
    If Visible And (Not m.bSettingAutoExit) Then
        If chkAutoExit.Value = vbChecked Then
            strAutoExit = ActivateAutoExit(TradeAccountID, SymbolID, "Ladder")
            If Len(strAutoExit) > 0 Then
                SetAutoExitCaptions FileBase(strAutoExit)
            Else
                chkAutoExit.Value = vbUnchecked
                SetAutoExitCaptions "None"
            End If
        Else
            g.Broker.BrokerDebug Broker, "Auto Exit Deactivate from Ladder (" & m.strSym & ", " & AccountNumber & "): Done"
            g.OrderStrategies.DeactivateExit TradeAccountID, SymbolID, True, "Turned off on Price Ladder"
            SetAutoExitCaptions "None"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.chkAutoExit_Click"
    
End Sub

Private Sub chkAutoJournal_Click()
On Error GoTo ErrSection:

    If Visible Then
        g.Broker.AutoJournalPopUp = (chkAutoJournal.Value = vbChecked)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.chkAutoJournal_Click"

End Sub

Private Sub chkConfirmOrder_Click()
On Error GoTo ErrSection:

    g.Broker.ConfirmManual = CheckBoxValue(chkConfirmOrder)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.chkConfirmOrder_Click"
    
End Sub

Private Sub cmdBrokerConnect_Click()
On Error GoTo ErrSection:

    g.Broker.Connect Broker
    If ConnectionStatus = eGDConnectionStatus_Connected Then
        Form_Resize
    End If

    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cmdBrokerConnect_Click"
    
End Sub

Private Sub cmdBuyAsk_Click()
On Error GoTo ErrSection:

    ClearBuySellButtons True
    OneClickOrder m.dBestAsk, True, eTT_OrderType_Limit

    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cmdBuyAsk_Click"
    
End Sub

Private Sub cmdBuyBid_Click()
On Error GoTo ErrSection:

    ClearBuySellButtons True
    OneClickOrder m.dBestBid, True, eTT_OrderType_Limit

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdBuyBid_Click"
    
End Sub

Private Sub cmdBuyMarket_Click()
On Error GoTo ErrSection:

    ClearBuySellButtons True
    OneClickOrder 1, True, eTT_OrderType_Market

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdBuyMarket_Click"
    
End Sub

Private Sub cmdBuyMarket_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancelAll_Click
'' Description: Cancel all working orders for this symbol and account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancelAll_Click()
On Error GoTo ErrSection:

    CancelAll

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdCancelAll_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClearQty_Click
'' Description: Allow the user to start over with a new order quantity
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClearQty_Click()
On Error GoTo ErrSection:

    'txtTradeQty.Text = ""       '10-06-2006 (reverse Harry's request for now)
    m.Quantity.Price = m.Quantity.Min
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdClearQty_Click"
    
End Sub

Private Sub cmdBailOut_Click()
On Error GoTo ErrSection:
    
    cmdBailout.Enabled = False
    cmdBailout.BackColor = Me.BackColor
    
    g.Broker.BrokerDebug Broker, "Flattening Position for " & m.strSym & " in account " & AccountNumber & " from the Price Ladder", True
    If m.frmBroker Is Nothing Then
        FlattenForSymbol TradeAccountID, SymbolID, 0&
    Else
        m.frmBroker.Ladder_FlattenPosition m.strSym
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdBailout_Click"

End Sub

Private Sub ExitPosition()
On Error GoTo ErrSection:
      
    Dim strPos$, nQty&
    
    strPos = Parse(lblTradePos.Caption, " ", 1)
    nQty = Int(ValOfText(Parse(lblTradePos.Caption, " ", 2)))
    
    If nQty > 0 Then
        If strPos = "Long" Then
            OneClickOrder 1, False, eTT_OrderType_Market, nQty
        ElseIf strPos = "Short" Then
            OneClickOrder 1, True, eTT_OrderType_Market, nQty
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.ExitPosition"

End Sub

Private Sub cmdFrontMonth_Click()
On Error GoTo ErrSection:

    Dim nSymbolID As Long

    If cmdFrontMonth.Caption = "Active Contract" Then
        ' for a continuous contract, look up current contract
        If Not m.TickBars Is Nothing Then
            nSymbolID = g.SymbolPool.SymbolIDforSymbol(RollSymbolForDate(m.TickBars.Prop(eBARS_Symbol), m.nSessionDate))
        End If
    Else
        nSymbolID = g.SymbolPool.SymbolIDforSymbol(cmdFrontMonth.Caption)
    End If
    
    If nSymbolID > 0 Then ChangeSymbol nSymbolID
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cmdFrontMonth_Click"
    
End Sub

Private Sub cmdQty1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    MoveFocus cmdQty1

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdQty1_MouseDown"
    
End Sub

Private Sub cmdQty1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    UpdateTradeQuantity Button, m.lPreset1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cmdQty1_MouseUp"
    
End Sub

Private Sub cmdQty2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    MoveFocus cmdQty2
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.cmdQty2_MouseDown"
    
End Sub

Private Sub cmdQty2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    UpdateTradeQuantity Button, m.lPreset2
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdQty2_MouseUp"
    
End Sub

Private Sub cmdQty3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    MoveFocus cmdQty3

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdQty3_MouseDown"
    
End Sub

Private Sub cmdQty3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    UpdateTradeQuantity Button, m.lPreset3
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdQty3_MouseUp"
    
End Sub

Private Sub cmdReverse_Click()
On Error GoTo ErrSection:
    
    cmdReverse.Enabled = False
    g.Broker.BrokerDebug Broker, "Reversing Position for " & m.strSym & " in account " & AccountNumber & " from the Price Ladder", True
    ReverseForSymbol TradeAccountID, SymbolID, 0&
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdReverse_Click"

End Sub

Private Sub cmdSellAsk_Click()
On Error GoTo ErrSection:
    
    ClearBuySellButtons True
    OneClickOrder m.dBestAsk, False, eTT_OrderType_Limit
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdSellAsk_Click"
    
End Sub

Private Sub cmdSellBid_Click()
On Error GoTo ErrSection:

    ClearBuySellButtons True
    OneClickOrder m.dBestBid, False, eTT_OrderType_Limit

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdSellBid_Click"

End Sub

Private Sub cmdSellMarket_Click()
On Error GoTo ErrSection:
    
    ClearBuySellButtons True
    OneClickOrder 1, False, eTT_OrderType_Market

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.cmdSellMarket_Click"
    
End Sub

Private Sub cmdSellMarket_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If

End Sub

Private Sub fgAccountBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar, True
    End If

End Sub

Private Sub fgAskDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    Dim strText$

    With fgAskDetail
        If .Row = 0 Then
            strText = .TextMatrix(0, 0) & ","
            strText = strText & .TextMatrix(0, 1) & ","
            strText = strText & .TextMatrix(0, 2) & ","
            strText = strText & .TextMatrix(0, 3)
            If strText <> AskHeader Then
                AskHeader = strText
            End If
        Else
            Position = Col
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.fgAskDetail.AfterMoveColumn", eGDRaiseError_Show

End Sub

Private Sub fgAskDetail_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
On Error GoTo ErrSection:

    Dim i&, nColor&, nAskSizeCol&
    
    If m.nDrawTriangle = 1 Then
        nAskSizeCol = HeaderToColIndex(m.aAskColHeader, "Size")
        If nAskSizeCol >= 0 And nAskSizeCol < 4 Then
            With fgAskDetail
                If (Col = nAskSizeCol Or Col = nAskSizeCol + 4) And Row >= .FixedRows Then
                    'Triangle UpDown parameter: 1 = up, -1 = down
                    i = Val(.TextMatrix(Row, eCol_BidAskChange))
                    If i = 1 Then
                        nColor = m.nBidAskUpColor
                    ElseIf i = -1 Then
                        nColor = m.nBidAskDownColor
                    End If
                    i = geDrawTickTriangle(m.geTickObj, hDC, .ForeColor, nColor, i, Top, Left, Bottom, Right)
                End If
            End With
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgAskDetail_DrawCell"
    
End Sub

Private Sub fgAskDetail_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    ShowHelp KeyCode
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgAskDetail_KeyDown"
    
End Sub

Private Sub fgAskDetail_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If UCase(Chr(KeyAscii)) = "S" Then ChangeSymbol 0

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgAskDetail_KeyPress"
    
End Sub

Private Sub fgBidDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    Dim strText$
    
    With fgBidDetail
        If .Row = 0 Then
            strText = .TextMatrix(0, 0) & ","
            strText = strText & .TextMatrix(0, 1) & ","
            strText = strText & .TextMatrix(0, 2) & ","
            strText = strText & .TextMatrix(0, 3)
            If strText <> BidHeader Then
                BidHeader = strText
            End If
        Else
            Position = Col
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.fgBidDetail.AfterMoveColumn", eGDRaiseError_Show

End Sub

Private Sub fgBidDetail_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
On Error GoTo ErrSection:

    Dim i&, nColor&, nBidColSize&
        
    If m.nDrawTriangle = 1 Then
        nBidColSize = HeaderToColIndex(m.aBidColHeader, "Size")
        If nBidColSize >= 0 And nBidColSize < 4 Then
            With fgBidDetail
                If Col = nBidColSize And Row >= .FixedRows Then
                    'Triangle UpDown parameter: 1 = up, -1 = down
                    i = Val(.TextMatrix(Row, eCol_BidAskChange))
                    If i = 1 Then
                        nColor = m.nBidAskUpColor
                    ElseIf i = -1 Then
                        nColor = m.nBidAskDownColor
                    End If
                    i = geDrawTickTriangle(m.geTickObj, hDC, .ForeColor, nColor, i, Top, Left, Bottom, Right)
                End If
            End With
        End If
    End If

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgBidDetail_DrawCell"

End Sub

Private Sub fgBidDetail_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ShowHelp KeyCode

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgBidDetail_KeyDown"
    
End Sub

Private Sub fgBidDetail_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If UCase(Chr(KeyAscii)) = "S" Then ChangeSymbol 0

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgBidDetail_KeyPress"
    
End Sub

Private Sub fgDOMPrint_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
On Error GoTo ErrSection:

    Dim nSizeCol&
    
    nSizeCol = HeaderToColIndex(m.aBidColHeader, "Size")
    If Col = nSizeCol Then
        fgBidDetail_DrawCell hDC, Row, Col, Left, Top, Right, Bottom, Done
    End If
    
    nSizeCol = HeaderToColIndex(m.aAskColHeader, "Size")
    If Col = nSizeCol + 4 Then
        fgAskDetail_DrawCell hDC, Row, Col, Left, Top, Right, Bottom, Done
    End If

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgDOMPrint_DrawCell"

End Sub

Private Sub fgOrdersInfo_Click()
On Error GoTo ErrSection:

    Dim PtOrder As cPtOrder

    With fgOrdersInfo
        If .Row >= .FixedRows And .Row < .Rows Then
            If .Col = 1 Then
                Set PtOrder = LoadOrderFromGrid(.Row, 2)
                If Not PtOrder Is Nothing Then
                    g.Broker.BrokerDebug Broker, "Cancelling Order from Ladder: " & PtOrder.OrderText, True
                    CancelThisOrder PtOrder, False
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgOrdersInfo_Click"

End Sub

Private Sub fgQuoteBar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ShowHelp KeyCode

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgQuoteBar_KeyDown"
    
End Sub

Private Sub fgQuoteBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_QbBar
    End If
    
End Sub

Private Sub fgTickDistribution_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error Resume Next

    If tmr.Enabled Then m.dScrollTickCount = gdTickCount()

End Sub

Private Sub fgTickDistribution_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    With fgTickDistribution
        Select Case Col
            Case eTDCols_OrderAskX, eTDCols_OrderBidX
                .ColWidth(Col) = kWidthOrderX
            Case eTDCols_OrderBid, eTDCols_OrderAsk, eTDCols_Volume         '5536
                If .ColWidth(Col) < kMinColWidthExt Then .ColWidth(Col) = kMinColWidthExt
                m.bUserSetSize = True
            Case eTDCols_Price
                If .ColWidth(Col) < m.nPriceColWidth Then
                    .AutoSize eTDCols_Price
                Else
                    m.bUserSetSize = True    'allow user to resize price col bigger
                End If
                m.nPriceColWidth = .ColWidth(eTDCols_Price)
            Case Else
                If Col >= 0 And Col < .Cols Then
                    If Col <> eTDCols_BidSize And Col <> eTDCols_AskSize Then
                        If .ColWidth(Col) < kMinColWidth Then .ColWidth(Col) = kMinColWidth
                    End If
                    m.bUserSetSize = True
                End If
        End Select
    End With
    
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgTickDistribution_AfterUserResize"
    
End Sub

Private Sub fgTickDistribution_Click()
On Error GoTo ErrSection:

    Dim lCurrentPosition&
    
    Dim PtOrder As cPtOrder
    Dim bNewOrder As Boolean

    If m.bIgnoreClick Then
        Exit Sub
    End If
    
    With fgTickDistribution
        If .Row >= .FixedRows Then
            Select Case .Col
            Case eTDCols_Price
                If .Row >= .FixedRows And .Row < .Rows Then
                    frmLadderColorSelector.ShowMe Me, .Row, m.nOutlineColor, m.Data(eFld_PriceStr, .RowData(.Row))
                End If
            Case eTDCols_PL
                If m.dAvgEntry = 0 Or vseOrderBar.Visible = False Then
                    'only do this if position is Flat or order bar not shown (i.e. position is unknown)
                    If .Row = m.nLastPriceRow Then
                        m.dZeroPLPrice = m.dLastPrice
                    Else
                        m.dZeroPLPrice = ValOfColPrice(.Row)
                    End If
                    lCurrentPosition = .TopRow
                    CalcPLData True
                    If lCurrentPosition <> .TopRow And lCurrentPosition < .Rows Then
                        .TopRow = lCurrentPosition
                        m.dScrollTickCount = 0          'to distinguish program scroll from user's scrol
                    End If
                End If
            Case eTDCols_OrderAsk, eTDCols_OrderBid
                If InStr(m.strSym, "-0") <> 0 Then
                    'do nothing (continuous contract not supported)
                Else
                    Set PtOrder = LoadOrderFromGrid(.Row, .Col)
                    If Not PtOrder Is Nothing Then
                        If OrderIsPending(PtOrder) Then
                            InfBox "This order cannot be modified because it is in a pending state.  Please wait for order confirmation.", "!", , "Price Ladder Order Error"
                        Else
                            g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & PtOrder.OrderText, True
                            ModifyThisOrder PtOrder
                        End If
                    End If
                End If
            Case eTDCols_OrderAskX, eTDCols_OrderBidX
                If InStr(m.strSym, "-0") <> 0 Then
                    'do nothing (continuous contract not supported)
                Else
                    Set PtOrder = LoadOrderFromGrid(.Row, .Col)
                    If Not PtOrder Is Nothing Then
                        g.Broker.BrokerDebug Broker, "Cancelling Order from Ladder: " & PtOrder.OrderText, True
                        CancelThisOrder PtOrder, False
                        If .Col = eTDCols_OrderAskX Then
                            .TextMatrix(.Row, eTDCols_OrderAsk) = ""
                            .TextMatrix(.Row, eTDCols_OrderAskX) = ""
                            .Cell(flexcpBackColor, .Row, eTDCols_OrderAsk) = .BackColor
                            .Cell(flexcpBackColor, .Row, eTDCols_OrderAskX) = .BackColor
                        Else
                            .TextMatrix(.Row, eTDCols_OrderBid) = ""
                            .TextMatrix(.Row, eTDCols_OrderBidX) = ""
                            .Cell(flexcpBackColor, .Row, eTDCols_OrderBid) = .BackColor
                            .Cell(flexcpBackColor, .Row, eTDCols_OrderBidX) = .BackColor
                        End If
                    End If
                End If
            Case eTDCols_BidSize, eTDCols_AskSize
                bNewOrder = False
                If InStr(m.strSym, "-0") <> 0 Then
                    'do nothing (continuous contract not supported)
                ElseIf Len(.TextMatrix(.Row, eTDCols_HasOrder)) <> 0 Then
                    AmmendOrder .Col
                ElseIf m.nLadderButton = vbLeftButton Then
                    bNewOrder = True
                End If
                If bNewOrder Then NewOrderOnClick
            End Select
        End If
    End With
    
    Set PtOrder = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.fgTickDistribution_Click"
    
End Sub

Private Sub fgTickDistribution_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
On Error GoTo ErrSection:

    Dim dPrice#, strIndex$
    Dim rc&, nIdx&, i&
    Dim aIndex As New cGdArray
    
    Exit Sub        '09-18-2006 (cannot draw tick line until performance issue is resolved)
                                        
    If m.bUnloading Or m.bPrinting Then
        Exit Sub
    End If
    
    If m.nShowTickLine And m.TickBars.Size > 0 Then
        With fgTickDistribution
            If Col = eTDCols_Volume Then
                'let graphics engine know direction to draw tick line
                geTickLineDirection m.geTickObj, m.nTickLineRL
                If Row = 0 Then
                    rc = geDrawTickTimeScale(m.geTickObj, hDC, Me.Font.Name, Me.Font.Size, Top, Left, Bottom, Right)
                ElseIf Row >= .FixedRows Then
                    strIndex = m.Data(eFld_BarIdx, .RowData(Row))
                    aIndex.SplitFields strIndex, ","
                    For i = 0 To aIndex.Size - 1
                        nIdx = Val(aIndex(i))
                        If nIdx >= 0 Then
                            dPrice = m.TickBars(eBARS_Close, nIdx)
                            rc = geDrawTicks(m.geTickObj, hDC, dPrice, m.TickBars.BarsHandle, m.nTickLineColor, Top, Left, Bottom, Right - 2)
'                            If rc <> 0 Then
'                                If FileExist("c:\common\files32.exe") Then
'                                    StatusMsg "geDrawTicks failed: rc = " & Str(rc)
'                                End If
'                            End If
                        End If
                    Next
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "fgTickDistribution.DrawCell", eGDRaiseError_Show

End Sub

Private Sub fgTickDistribution_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ShowHelp KeyCode

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgTickDistribution_KeyDown"

End Sub

Private Sub fgTickDistribution_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If UCase(Chr(KeyAscii)) = "S" Then ChangeSymbol 0

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgTickDistribution_KeyPress"

End Sub

Private Sub fgTickDistribution_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim Order As cPtOrder
    Dim nID&
    
    With fgTickDistribution
        m.nDragOrderRow = 0
        If Button = vbRightButton Then
            m.nLadderButton = Button
            .Col = .MouseCol
            .Row = .MouseRow
        ElseIf .Col = eTDCols_OrderAsk Or .Col = eTDCols_OrderBid Then
            nID = ValOfText(.TextMatrix(.Row, eTDCols_HasOrder))
            If nID > 0 Then
'                Set Order = New cPtOrder
'                If Order.Load(nID) Then
                Set Order = LoadOrderFromGrid(.Row, .Col)
                If Not Order Is Nothing Then
                    If Not OrderIsPending(Order) Then
                        m.nDragOrderRow = .Row
'                        StatusMsg Str(m.nDragOrderRow)
                    End If
                End If
            End If
            Set Order = Nothing
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgTickDistribution_MouseDown"

End Sub

Private Sub fgTickDistribution_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
      
    HandleMouseMove X, Y
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.MouseMove", eGDRaiseError_Show

End Sub

Private Sub fgTickDistribution_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim OrderDragged As cPtOrder        ' Order that user was dragging
    Dim lRow As Long                    ' Current Mouse Row in the grid
    Dim lCol As Long                    ' Current Mouse Column in the grid
    
    m.nLadderButton = Button
    m.bIgnoreClick = False
    
    If Button <> vbRightButton Then
        With fgTickDistribution
            lRow = .MouseRow
            lCol = .MouseCol
            
            If (lCol = eTDCols_OrderBid) Or (lCol = eTDCols_OrderAsk) Then
                If (m.nDragOrderRow >= .FixedRows) And (m.nDragOrderRow < .Rows) And (lRow >= .FixedRows) And (lRow < .Rows) Then
                    If (lRow <> m.nDragOrderRow) Then
                        Set OrderDragged = LoadOrderFromGrid(m.nDragOrderRow, lCol)
                        If Not OrderDragged Is Nothing Then
                            g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & OrderDragged.OrderText, True
                            ModifyThisOrder OrderDragged, ValOfColPrice(lRow), , ConfirmOrder
                            'This flag is set here because fgTickDistribution_Click is called immediately upon exit of this function.
                            'The status of this ammended order is not yet known and grid has not been redrawn with ammended information.
                            'Setting this flag here prevents code in fgTickDistribution_Click from executing this once.
                            m.bIgnoreClick = True       'cause of aardvark bug 4554
                        End If
                    End If
                End If
            End If
        End With
    End If

    m.nDragOrderRow = 0
    
    Set OrderDragged = Nothing
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.fgTickDistribution_MouseUp"

End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    TextIncDecRegisterForm Me, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.Form_Activate"

End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:
    
    TextIncDecUnregisterForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.Form_Deactivate"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText$, nLastUsedType&
    Dim strQBarText$, strABarText$
    
    g.Styler.StyleForm Me
    
    
    'RH populate list
    With cboExchanges
        .AddItem "Auto"
        .AddItem "Limit"
        .AddItem "Stop"
    End With
    
    
    strText = g.strAppPath & "\LadderInterval.flg"
    If FileExist(strText) Then
        m.nIntervalRT = GetIniFileProperty("interval", 125, "Timer", strText)
    Else
        m.nIntervalRT = 125
    End If


    fgTickDistribution.FocusRect = flexFocusNone
    fgTickDistribution.HighLight = flexHighlightNever
    fgQuoteBar.HighLight = flexHighlightNever
    fgAccountBar.HighLight = flexHighlightNever
    
    cmdBailout.BackColor = cmdReverse.BackColor
    cmdBailout.Enabled = False
    With fraOrderBtns
        fraFrontMonth.Move .Left, .Top, .Width, .Height
    End With
    
    With tbToolbar
        .Tools("ID_Symbol").Picture = Picture16(ToolbarIcon("ID_Symbol"))
        .Tools("ID_Settings").Picture = Picture16(ToolbarIcon("ID_Settings"))   'want new toolbar to use kSettings for consistency
        .Tools("ID_DumpFile").Visible = FileExist(g.strAppPath & "\LadderDebug.flg")
        .Tools("ID_TextIncrease").Picture = Picture16(ToolbarIcon("ID_TextIncrease"))
        .Tools("ID_TextDecrease").Picture = Picture16(ToolbarIcon("ID_TextDecrease"))
    End With
    
    m.nSessionDate = 0
    m.bSessionCurrent = True
    m.aBidColHeader.SplitFields kBidColHeader
    m.aAskColHeader.SplitFields kAskColHeader
    m.aQBarColHeader.Size = 0
    m.aABarColHeader.Size = 0
    'get saved settings
    m.nFloodColor = GetIniFileProperty("GridFloodColor", kGridFloodColor, "Price Ladder", g.strIniFile)
    m.nUpColor = GetIniFileProperty("GridUpColor", RGB(0, 192, 0), "Price Ladder", g.strIniFile)
    m.nDownColor = GetIniFileProperty("GridDownColor", RGB(255, 0, 0), "Price Ladder", g.strIniFile)
    m.nBarColor = GetIniFileProperty("GridBarColor", RGB(0, 0, 128), "Price Ladder", g.strIniFile)
    m.nTickLineColor = GetIniFileProperty("TickLineColor", vbBlue, "Price Ladder", g.strIniFile)
    m.nBidColor = GetIniFileProperty("BidColor", kBidColor, "Price Ladder", g.strIniFile)
    m.nAskColor = GetIniFileProperty("AskColor", kAskColor, "Price Ladder", g.strIniFile)
    m.nBidTextColor = GetIniFileProperty("BidTextColor", vbBlack, "Price Ladder", g.strIniFile)
    m.nAskTextColor = GetIniFileProperty("AskTextColor", vbBlack, "Price Ladder", g.strIniFile)
    m.nFixedPriceColor = GetIniFileProperty("FixedPriceColor", vbBlack, "Price Ladder", g.strIniFile)
    m.nFontSize = GetIniFileProperty("GridFontSize", 8, "Price Ladder", g.strIniFile)
    If m.nFontSize <= 0 Then m.nFontSize = 8
    m.strFont = GetIniFileProperty("GridFontName", "MS Sans Serif", "Price Ladder", g.strIniFile)
    m.nShowVolBar = GetIniFileProperty("ShowVolumeBar", 1, "Price Ladder", g.strIniFile)
    m.nShowVolText = GetIniFileProperty("ShowVolumeText", 1, "Price Ladder", g.strIniFile)
    m.nShowVolMin = GetIniFileProperty("ShowVolMin", 0, "Price Ladder", g.strIniFile)
    m.nShowVolMax = GetIniFileProperty("ShowVolMax", 0, "Price Ladder", g.strIniFile)
    m.nShowTickLine = GetIniFileProperty("ShowTickLine", 1, "Price Ladder", g.strIniFile)
    m.nShowProfitLoss = GetIniFileProperty("ShowProfitLoss", 0, "Price Ladder", g.strIniFile)
    m.nShowAvgEntry = GetIniFileProperty("ShowAvgEntry", 0, "Price Ladder", g.strIniFile)
    m.nShowOpenEntries = GetIniFileProperty("ShowOpenEntries", 1, "Price Ladder", g.strIniFile)
    m.nTickLineRL = GetIniFileProperty("TickLineRightToLeft", 0, "Price Ladder", g.strIniFile)
    m.eVolumeStyle = GetIniFileProperty("LadderVolStyle", eLadderVol_LastTrade, "Price Ladder", g.strIniFile)
    
    m.nFirstColor = GetIniFileProperty("FirstColor", kFirstColor, "Price Ladder", g.strIniFile)
    m.nSecondColor = GetIniFileProperty("SecondColor", kSecondColor, "Price Ladder", g.strIniFile)
    m.nThirdColor = GetIniFileProperty("ThirdColor", kThirdColor, "Price Ladder", g.strIniFile)
    m.nFourthColor = GetIniFileProperty("FourthColor", kFourthColor, "Price Ladder", g.strIniFile)
    m.nFifthColor = GetIniFileProperty("FifthColor", kFifthColor, "Price Ladder", g.strIniFile)
    m.nOtherColor = GetIniFileProperty("OtherColor", kOtherColor, "Price Ladder", g.strIniFile)
    m.nLargestSizeColor = GetIniFileProperty("LargestSizeColor", kFirstColor, "Price Ladder", g.strIniFile)
    m.nOutlineColor = GetIniFileProperty("OutlineColor", -1, "Price Ladder", g.strIniFile)
    
    m.nAvgEntryColor = GetIniFileProperty("AvgEntryColor", kAvgEntryColor, "Price Ladder", g.strIniFile)
    m.nInactiveColor = GetIniFileProperty("InactiveColor", kInactiveColor, "Price Ladder", g.strIniFile)
    m.nDrawTriangle = GetIniFileProperty("DrawTriangle", 1, "Price Ladder", g.strIniFile)
    m.nBidAskUpColor = GetIniFileProperty("BidAskUpColor", kBidAskUpColor, "Price Ladder", g.strIniFile)
    m.nBidAskDownColor = GetIniFileProperty("BidAskDownColor", kBidAskDownColor, "Price Ladder", g.strIniFile)
    m.nShowSummaryBar = GetIniFileProperty("BidAskSummaryBar", 1, "Price Ladder", g.strIniFile)
    m.nVertSummaryBar = GetIniFileProperty("BidAskSummaryVert", 1, "Price Ladder", g.strIniFile)
    m.nOrderColumns = GetIniFileProperty("OrderColumns", 0, "Price Ladder", g.strIniFile)
    m.nTradeAcctID = GetIniFileProperty("AccountID", 0, "Price Ladder", g.strIniFile)
    
    If m.eDisplayStyle = eView_Ladder Then
        m.nShowQuoteBar = GetIniFileProperty("QuoteBar", 1, "Price Ladder", g.strIniFile)
        m.nShowAccountBar = GetIniFileProperty("AccountBar", 1, "Price Ladder", g.strIniFile)
        m.eOrderBarMode = GetIniFileProperty("OrderBar", eGDOrderBarMode_Right, "Price Ladder", g.strIniFile)
    Else
        m.nShowQuoteBar = GetIniFileProperty("QuoteBar", 1, "Market Depth", g.strIniFile)
        m.nShowAccountBar = GetIniFileProperty("AccountBar", 0, "Market Depth", g.strIniFile)
        m.eOrderBarMode = GetIniFileProperty("OrderBar", eGDOrderBarMode_LastShownOnRight, "Market Depth", g.strIniFile)
    End If
    If BrokerViewMode Then
        m.nShowAcctBarSave = m.nShowAccountBar      'save what was loaded
        m.nShowAccountBar = 0                       'set to zero so account bar will not show
    End If
    
    'order bar
    strText = GetIniFileProperty("OrdBarCtrls", kOrdBarDefaults, "Price Ladder", g.strIniFile)
    m.strOrdBarCtrls = ConvertOrdBarButtons(strText, Me)
    If Len(m.strOrdBarCtrls) = 0 Then m.strOrdBarCtrls = kOrdBarDefaults    'something went wrong just use new defaults
    
    FixOrderBarCtrlString
    
    m.strOrdBarCtrlsSave = m.strOrdBarCtrls
    fraExitFavorites.Visible = False        'frame for exit favorites buttons
    'RH commented out fraExitFavorites.BorderStyle = 0

    'save new format
    If InStr(strText, "OE") = 0 Then SetIniFileProperty "OrdBarCtrls", m.strOrdBarCtrls, "Price Ladder", g.strIniFile
    m.nHighlightPos = GetIniFileProperty("HighlightPos", -2, "Price Ladder", g.strIniFile)
    m.nHighlightEquity = GetIniFileProperty("HighlightEquity", -2, "Price Ladder", g.strIniFile)
    
    m.nBlankRows = GetIniFileProperty("BlankRows", 30, "Price Ladder", g.strIniFile)
    m.eFloodMktDepth = GetIniFileProperty("FloodMktDepth", eBidAskColorMode_ByPrice, "Price Ladder", g.strIniFile)
    pbBid.Height = GetIniFileProperty("BidAskSummaryHeight", 600, "Price Ladder", g.strIniFile)
    strText = GetIniFileProperty("BidColumns", "", "Price Ladder", g.strIniFile)
    BidHeader = strText
    strText = GetIniFileProperty("AskColumns", "", "Price Ladder", g.strIniFile)
    AskHeader = strText
    strText = GetIniFileProperty("QuoteBarColumns", kQBarCols, "Price Ladder", g.strIniFile)
    QuoteBarHeader strText
    strText = GetIniFileProperty("AccountBarColumns", kABarCols, "Price Ladder", g.strIniFile)
    strText = Replace(strText, "# of", "#")
    AccountBarHeader strText
    ' create table for price data
    m.Data.CreateField eGDARRAY_Doubles, eFld_PL, "Profit/Loss"
    m.Data.CreateField eGDARRAY_Longs, eFld_BidSize, "BidSize"
    m.Data.CreateField eGDARRAY_Doubles, eFld_Price, "Price"
    m.Data.CreateField eGDARRAY_Longs, eFld_AskSize, "AskSize"
    m.Data.CreateField eGDARRAY_Doubles, eFld_Volume, "Volume"
    m.Data.CreateField eGDARRAY_Strings, eFld_PriceStr, "PriceStr"
    m.Data.CreateField eGDARRAY_Strings, eFld_BarIdx, "BarIndex"
    'create table for orders
    m.tbOrders.CreateField eGDARRAY_Longs, 0, "OrderID"
    m.tbOrders.CreateField eGDARRAY_Longs, 1, "OrderBuy"
    m.tbOrders.CreateField eGDARRAY_Longs, 2, "OrderQty"
    m.tbOrders.CreateField eGDARRAY_Longs, 3, "OrderType"
    m.tbOrders.CreateField eGDARRAY_Doubles, 4, "OrderPrice"
    m.tbOrders.CreateField eGDARRAY_Longs, 5, "OrderPending"
    m.tbOrders.CreateField eGDARRAY_Longs, 6, "OrderStatus"
    'create table for prices that should have their cells outlined with specific color
    m.tbOutlineCells.CreateField eGDARRAY_Doubles, 0, "PriceToOutline"
    m.tbOutlineCells.CreateField eGDARRAY_Longs, 1, "Color"
                
    InitLadderGrid
    ResetGridBar fgQuoteBar, m.aQBarColHeader, m.nQBarSumColWidth
    ResetGridBar fgAccountBar, m.aABarColHeader, m.nABarSumColWidth
    ResetDetailGrid fgBidDetail, "Bid"
    ResetDetailGrid fgAskDetail, "Ask"
    InitPrintGrid
        
    'initialize tick object for graphics engine
    m.geTickObj = geInitTickObj()
    
    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If
    
    chkAutoExit.Visible = True
    
    Set m.Quantity = New cPriceEditor
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.Form.Load", eGDRaiseError_Show

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If Cancel = 0 Then m.WindowLink.Unhook
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.Form_QueryUnload"

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long
    
    Select Case m.eOrderBarMode
'        Case eGDOrderBarMode_NotShown, eGDOrderBarMode_LastShownOnRight
'            lMinScaleWidth = kMinFormSize
'        Case eGDOrderBarMode_BottomWide
'            lMinScaleWidth = kMinFormSize
'        Case eGDOrderBarMode_BottomNarrow
'            lMinScaleWidth = kMinFormSize
'        Case eGDOrderBarMode_BottomContinuous
'            lMinScaleWidth = lblFrontMonthHorz.Width + cmdFrontMonth.Width + 180
'        Case eGDOrderBarMode_Right
'            lMinScaleWidth = lblFrontMonthHorz.Width + cmdFrontMonth.Width + 180
        
        ' TLB: if order bar shown on bottom, make min width the order bar
        Case eGDOrderBarMode_BottomWide, eGDOrderBarMode_BottomNarrow, eGDOrderBarMode_BottomContinuous
            lMinScaleWidth = kMinFormSize
        Case Else ' otherwise allow it to go quite a bit smaller
            lMinScaleWidth = kMinFormSize / 2
    End Select
                        
    If LimitFormSize(Me, lMinScaleWidth, kMinFormSize) Then Exit Sub
    
    'determine height so can correctly position order bar if necessary
    If m.nShowQuoteBar = 1 Then
        With fgQuoteBar
            .Redraw = flexRDNone
            If m.nQBarSumColWidth + 50 >= Me.ScaleWidth Then
                .Move 0, 0, Me.ScaleWidth, .RowHeightMax * 3 + 50   'scroll bar is visible
            Else
                .Move 0, 0, Me.ScaleWidth, .RowHeightMax * 2
            End If
            .Visible = True
            .Redraw = flexRDBuffered
        End With
    Else
        fgQuoteBar.Visible = False
    End If
    
    'determine height so can correctly position order bar if necessary
    With fgAccountBar
        If m.nABarSumColWidth + 50 >= Me.ScaleWidth Then
            .Height = .RowHeight(0) * 3 + 50 'scroll bar is visible
        Else
            .Height = .RowHeight(0) * 2
        End If
    End With
    
    If m.eDisplayStyle = eView_Ladder Then
        HandleLadderResize
    Else
        HandleDOMResize
    End If
    
End Sub

Private Sub ResetDetailGrid(fgGrid As VSFlexGrid, ByVal strBidAskLabel$)
On Error GoTo ErrSection:

    Dim aHeader As cGdArray, nCol&

    If strBidAskLabel = "Bid" Then
        Set aHeader = m.aBidColHeader
    Else
        Set aHeader = m.aAskColHeader
    End If
    
    With fgGrid
        .Redraw = flexRDNone
        SetupGrid fgGrid, eGridMode_Grid
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDNone
        .FixedRows = 1
        .FrozenCols = 5
        .Rows = 1
        .Cols = 5
        .Font.Name = m.strFont
        .Font.Size = m.nFontSize
        .OwnerDraw = flexODOver
        .ExtendLastCol = False
        'hide last col
        .ColHidden(eCol_BidAskChange) = True
        'header tittles
        .TextMatrix(0, 0) = aHeader(0)
        .TextMatrix(0, 1) = aHeader(1)
        .TextMatrix(0, 2) = aHeader(2)
        .TextMatrix(0, 3) = aHeader(3)
        'header alignment
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        'alignment
        nCol = HeaderToColIndex(aHeader, "Maker")
        If nCol >= 0 And nCol < 4 Then .ColAlignment(nCol) = flexAlignLeftCenter
        nCol = HeaderToColIndex(aHeader, strBidAskLabel)
        If nCol >= 0 And nCol < 4 Then .ColAlignment(nCol) = flexAlignRightCenter
        nCol = HeaderToColIndex(aHeader, "Size")
        If nCol >= 0 And nCol < 4 Then .ColAlignment(nCol) = flexAlignRightCenter
        nCol = HeaderToColIndex(aHeader, "Time")
        If nCol >= 0 And nCol < 4 Then .ColAlignment(nCol) = flexAlignCenterCenter
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.ResetDetailGrid", eGDRaiseError_Raise
    
End Sub

Private Sub InitLadderGrid()
On Error GoTo ErrSection:

    With fgTickDistribution
        .Redraw = flexRDNone
        SetupGrid Me.fgTickDistribution, eGridMode_Grid
        .AutoSizeMode = flexAutoSizeColWidth
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDNone
        .FixedRows = 1
        .Rows = 1
        .Cols = kCols
        .OwnerDraw = flexODOver
        .Font.Name = m.strFont
        .Font.Size = m.nFontSize
        .ExtendLastCol = False
        If g.nColorTheme = kDarkThemeColor Then
            .BackColor = kDarkThemeColor
            .BackColorFixed = kDarkThemeColor
            .ForeColorFixed = vbWhite
        End If
        'alignment
        .ColAlignment(eTDCols_Price) = flexAlignRightCenter
        .ColAlignment(eTDCols_Volume) = flexAlignRightCenter
        .ColAlignment(eTDCols_OrderBid) = flexAlignCenterCenter
        .ColAlignment(eTDCols_OrderAsk) = flexAlignCenterCenter
        .ColAlignment(eTDCols_AskSize) = flexAlignLeftCenter
        .ColAlignment(eTDCols_OrderBidX) = flexAlignCenterCenter
        .ColAlignment(eTDCols_OrderAskX) = flexAlignCenterCenter
        .ColAlignment(eTDCols_PL) = flexAlignRightCenter
        'sort
        .ColSort(eTDCols_Volume) = flexSortNone
        .ColSort(eTDCols_OrderBid) = flexSortNone
        .ColSort(eTDCols_BidSize) = flexSortNone
        .ColSort(eTDCols_Price) = flexSortNone
        .ColSort(eTDCols_AskSize) = flexSortNone
        .ColSort(eTDCols_OrderAsk) = flexSortNone
        .ColSort(eTDCols_PL) = flexSortNone
        'column headers
        .TextMatrix(0, eTDCols_OrderBidX) = "X"
        .TextMatrix(0, eTDCols_OrderBid) = "Orders"
        .TextMatrix(0, eTDCols_BidSize) = "Bid Size"
        .TextMatrix(0, eTDCols_Price) = "Price"
        .TextMatrix(0, eTDCols_AskSize) = "Ask Size"
        .TextMatrix(0, eTDCols_OrderAsk) = "Orders"
        .TextMatrix(0, eTDCols_OrderAskX) = "X"
        .TextMatrix(0, eTDCols_Entries) = "E"
        .TextMatrix(0, eTDCols_PL) = "P/L"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        'hidden columns
        .ColHidden(eTDCols_OrderBidX) = True
        .ColHidden(eTDCols_OrderBid) = True
        .ColHidden(eTDCols_OrderAsk) = True
        .ColHidden(eTDCols_OrderAskX) = True
        .ColHidden(eTDCols_HasOrder) = True
        .Redraw = flexRDBuffered
        
        Me.Font = .Font
        Me.Font.Size = .Font.Size
        Me.Font.Bold = True
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.InitLadderGrid", eGDRaiseError_Raise

End Sub

Private Sub LoadGrid(Optional ByVal strErr$ = "", Optional ByVal bCenter As Boolean = True)
On Error GoTo ErrSection:
    
    'if error then display error message on first row and exit
    If Len(strErr) > 0 Then
        fgBidDetail.Visible = False
        fgAskDetail.Visible = False
        fgQuoteBar.Visible = False
        pbBid.Visible = False
        pbAsk.Visible = False
            
        With fgTickDistribution
            .Rows = .FixedRows + 1
            .Cell(flexcpAlignment, .FixedRows, 0, .FixedRows, .Cols - 1) = flexAlignLeftCenter
            .Cell(flexcpText, .FixedRows, 0, .FixedRows, .Cols - 1) = strErr
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.FixedRows) = True
            .Visible = True
        End With
        
        SetCaption
        FormResize Me
        Exit Sub
    End If
    
    If m.eDisplayStyle = eView_Ladder Then
        fgTickDistribution.Visible = True
        fgBidDetail.Visible = False
        fgAskDetail.Visible = False
        fgQuoteBar.Visible = False
        LoadLadderGrid bCenter
    Else
        fgTickDistribution.Visible = False
        fgBidDetail.Visible = True
        fgAskDetail.Visible = True
        If m.bHeaderChanged Then
            ResetDetailGrid fgBidDetail, "Bid"
            ResetDetailGrid fgAskDetail, "Ask"
            m.bHeaderChanged = False
        End If
        With fgOrdersInfo
            .Redraw = flexRDNone
            SetupGrid Me.fgTickDistribution, eGridMode_Grid
            .ExplorerBar = flexExNone
            .BackColorAlternate = ALT_GRID_ROW_COLOR
            .Editable = flexEDNone
            .ExtendLastCol = False
            .Font.Bold = True
            .ScrollBars = flexScrollBarVertical
            
            .Rows = 1
            .Cols = 3
            .FixedRows = 1
            .FixedCols = 0
            .ColHidden(2) = True
            
            .TextMatrix(0, 0) = "Orders Information"
            .TextMatrix(0, 1) = "X"
            .ColAlignment(1) = flexAlignCenterCenter
            
            .Height = .RowHeight(0) * 5
            .Redraw = flexRDBuffered
        End With
        LoadDetailGrid
    End If
    
    FormResize Me
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.LoadGrid", eGDRaiseError_Raise

End Sub

Private Sub GetData(Optional ByVal bFromRefresh As Boolean = False)
On Error GoTo ErrSection:

    Dim rc As Byte, i&
    Dim strMsg$, strLastTickTime$
    Dim bFullTicksAvail As Boolean
    Dim bUnloadNow As Boolean
    
    If m.bUnloading Then Exit Sub

    If Left(m.strSym, 1) = "$" And Not IsForex(m.strSym) And Not bFromRefresh Then
        strMsg = "This feature does not support index symbols."
        InfBox strMsg, "I", , "Price Ladder"
        LoadGrid strMsg
        Exit Sub
    End If

'   if "$" or minutized then NO volume
         
    chkConfirmOrder.Visible = g.RealTime.Active
         
    ResetData
    m.bUnloading = False
                    
    If bFromRefresh And m.eDisplayStyle = eView_Detail And Not g.RealTime.Active Then
        Exit Sub
    End If
    
    With fgTickDistribution
        .Redraw = flexRDNone
        .Rows = .FixedRows '(to clear all rows)
        .Rows = .FixedRows + 1
        .TextMatrix(.FixedRows, 0) = "Loading intraday data for " & m.strSym & " ..."
        .MergeCells = flexMergeSpill
        .Redraw = flexRDBuffered
        .Refresh
    End With
    
    SetBarProperties m.DailyBar, m.nSymID
    If m.bSessionCurrent = True Then      'current session
        GetAvailTickData m.TickBars, m.nSessionDate, m.strSym, m.nSymID, 0, -6
    ElseIf m.nSessionDate > 0 Then
        GetAvailTickData m.TickBars, i, m.strSym, m.nSymID, m.nSessionDate, 0
    End If
        
    If g.nReplaySession = 0 Then
        'don't do this check if streaming replay is running
        m.bMinutized = IsMinutized(m.TickBars, bFullTicksAvail)
    End If
    
    'show update volume button when all of following are true:
    'a)symbol does not start with $
    'b)user wants volume bar or volume text shown
    'c)data is minutized
    'd)full tick data is available for download
    If m.TickBars.Size > 0 And m.bMinutized And Left(m.strSym, 1) <> "$" Then
        If m.nShowVolBar Or m.nShowVolText Then
            tbToolbar.Tools("ID_Volume").Visible = True
        Else
            tbToolbar.Tools("ID_Volume").Visible = False
        End If
    Else
        tbToolbar.Tools("ID_Volume").Visible = False
    End If
        
    ' if no data ...
    If m.TickBars.Size = 0 Then
        strMsg = "No intraday data available for " & m.strSym
        If InStr(m.strSym, "-") > 0 And Not HasModule("FT") Then
            strMsg = "You must be enabled for intraday data for Futures."
        ElseIf InStr(m.strSym, "-") = 0 And Not HasModule("ST") Then
            strMsg = "You must be enabled for intraday data for Stocks."
        ElseIf g.RealTime.SalmonIsRunning Then
            If g.RealTime.SymbolInfo(m.strSym).GetDataRequestStatus(ePRD_EachTick) = eSalmonPending Then
                strMsg = "Retrieving intraday data for " & m.strSym & " ..."
            End If
        End If
        'tmr.Enabled = False
        SetCaption
        With fgTickDistribution
            .Redraw = flexRDNone
            .Rows = .FixedRows '(to clear all rows)
            .Rows = .FixedRows + 1
            .TextMatrix(.FixedRows, 0) = strMsg
            .MergeCells = flexMergeSpill
            .Redraw = flexRDBuffered
        End With
        If g.RealTime.Active Then tmr.Enabled = True
        Exit Sub
    End If
    
    'check minimum move
    m.dMinMove = m.TickBars.MinMove(m.nSessionDate)
    
    If m.dMinMove <= 0 Then
        m.bUnloading = True
        bUnloadNow = True
        strMsg = "Invalid minimum move for " & m.strSym & vbCrLf & _
            "MinMove = " & Str(m.dMinMove) & vbCrLf & _
            ". Session date = " & DateFormat(m.nSessionDate, MM_DD_YYYY) & _
            vbCrLf & ". Unloading price ladder."
            InfBox strMsg, "E", , "Price Ladder"
            
        Me.Hide
        'do this in the timer otherwise the form will not get removed from the collection and the mMain.SaveVisibleForms routine will pick it up
        tmr.Enabled = False
        tmr.Interval = 1000
        tmr.Tag = "Unload Now"
        tmr.Enabled = True
        
        Exit Sub
    End If
    
    'build daily bar from ticks
    m.DailyBar.ArrayMask = eBARS_Eod Or eBARS_BidAsk
    m.DailyBar.BuildBars "D", m.TickBars.BarsHandle
            
    LoadTable
    LoadGrid
    
    'get minute of last tick
    strLastTickTime = DateFormat(m.TickBars(eBARS_DateTime, m.TickBars.Size - 1), NO_DATE, HH_MM)
    If InStr(strLastTickTime, ":") Then     'precautionary - should always be true
        m.nPrevMinute = Val(Right(strLastTickTime, 2))
    Else
        m.nPrevMinute = -1              'to indicate something is wrong
    End If
        
    'set window title
    SetCaption m.nSessionDate
    
    'let graphics engine know what time zone to use for display
    If g.bShowInLocalTimeZone Then
        geTickObjDisplayTimeZone m.geTickObj, ""
    Else
        geTickObjDisplayTimeZone m.geTickObj, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf)
    End If
    
    If g.RealTime.Active Or m.eOrderBarMode > eGDOrderBarMode_NotShown Then tmr.Enabled = True
                    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.GetData", eGDRaiseError_Raise

End Sub

Private Sub LoadTable()
On Error GoTo ErrSection:

    Dim i&, j&, nQty&, strIndex$
    Dim dPrice#, dPrevPrice#, dVol#, dSum#
    Dim dMinPrice#, dMaxPrice#
    
    Dim aIndex As New cGdArray
    Dim aEquityPos As New cGdArray
    
    Dim bFound As Boolean
    Dim bUseVol As Boolean
        
    aIndex.Create eGDARRAY_Longs, m.TickBars.Size
    i = gdSortAsIndex(aIndex.ArrayHandle, m.TickBars.ArrayHandle(eBARS_Close), 1, eGdSort_Descending, 0, -1)
        
'    If SecurityType(m.TickBars) = "S" Then bUseVol = True
    If m.TickBars(eBARS_Vol, 0) <> kNullData Then
        bUseVol = True
    End If
        
    'add records from 15 values above and below price to accomodate bid/ask depth of market
'    m.dMinMove = m.TickBars.MinMove(m.nSessionDate)            'moved to GetData so can exit if zero
                  
    m.dOpen = RoundToMinMove(m.DailyBar(eBARS_Open, m.DailyBar.Size - 1), m.dMinMove)
    m.dHigh = RoundToMinMove(m.TickBars(eBARS_Close, aIndex(0)), m.dMinMove)
    m.dLow = RoundToMinMove(m.TickBars(eBARS_Close, aIndex(aIndex.Size - 1)), m.dMinMove)
    
    If m.nBlankRows <= 0 Then m.nBlankRows = 30
    If m.nSessionBlankRows <= 0 Then m.nSessionBlankRows = m.nBlankRows
    
    dMaxPrice = m.TickBars(eBARS_Close, aIndex(0)) + m.dMinMove * m.nSessionBlankRows
    dMinPrice = m.TickBars(eBARS_Close, aIndex(aIndex.Size - 1)) - m.dMinMove * m.nSessionBlankRows
        
    ' TLB 10/18/2014: evidently it's possible for the prices to be null (esp. on Saturdays)
    If dMinPrice <= kNullData Then
        ' so not sure what's best, but we'll just set to 0 for now
        dMaxPrice = 0
        dMinPrice = 0
    Else
        dMaxPrice = RoundToMinMove(dMaxPrice, m.dMinMove)
        dMinPrice = RoundToMinMove(dMinPrice, m.dMinMove)
    End If
    m.dMinPrice = dMinPrice
    m.dMaxPrice = dMaxPrice
    
    'get last price & volume
    m.dLastPrice = RoundToMinMove(m.TickBars(eBARS_Close, m.TickBars.Size - 1), m.dMinMove)
    If bUseVol And m.bSessionCurrent = True Then
        'last trade volume is available only for current session
        m.dLastPriceVol = m.TickBars(eBARS_Vol, m.TickBars.Size - 1)
    Else
        m.dLastPriceVol = 0#
    End If
    
    'set price of zero PL to open price
    'aEquityPos.SplitFields g.Broker.PositionString(TradeAccountID, m.nSymID, 0&), "|"
    aEquityPos.SplitFields GetPositionString, "|"
    If aEquityPos.Size > 0 Then
        m.nQuantity = ValOfText(aEquityPos(1))
        If aEquityPos.Size > 6 Then
            'use Val because Dave uses Str in PositionString - 4430
            m.dAvgEntry = Val(aEquityPos(6))
            If m.dAvgEntry < 0 Or m.dAvgEntry > 9999999 Then
                DebugLog m.TickBars.Prop(eBARS_Symbol) & " Avg Entry out of range: " & aEquityPos(3) & " (" & aEquityPos(6) & ")"
                m.dAvgEntry = ValOfText(aEquityPos(3))
            End If
        Else
            m.dAvgEntry = ValOfText(aEquityPos(3))
        End If
    End If
    
    If m.dAvgEntry <> 0 And nQty <> 0 Then
        m.dZeroPLPrice = m.dAvgEntry
        m.nQuantity = nQty
    Else
        m.dZeroPLPrice = -1             '5712
        If SecurityType(m.nSymID) = "S" Then
            m.nQuantity = 100
        Else
            m.nQuantity = 1
        End If
    End If
        
    m.Data.NumRecords = 0
    'fill table with all possible values that are dMinMove apart from dMinPrice to dMaxPrice
    'Do While dMinPrice <= dMaxPrice
    Do While dMaxPrice >= dMinPrice
        AddToTable 0, 0, dMaxPrice, 0, 0
        If bUseVol And m.bSessionCurrent = True And dMaxPrice = m.dLastPrice Then
            m.nLastPriceRec = m.Data.NumRecords - 1
        End If
        dMaxPrice = dMaxPrice - m.dMinMove
        dMaxPrice = RoundToMinMove(dMaxPrice, m.dMinMove)
    Loop
                
    If m.nShowVolMin < 0 Then m.nShowVolMin = 0     'precautionary
    If m.nShowVolMax < 0 Then m.nShowVolMax = 0
    'walk through bars, sum up all volume for a given price
    'then put sum in volume field of matching price in table
    m.dMaxVol = 1
    For i = 0 To aIndex.Size - 1
        dPrice = RoundToMinMove(m.TickBars(eBARS_Close, aIndex(i)), m.dMinMove)
        If Len(strIndex) = 0 Then
            strIndex = Str(aIndex(i))
        Else
            If m.TickBars(eBARS_Close, aIndex(i)) <> m.TickBars(eBARS_Close, aIndex(i - 1)) And _
                dPrice = dPrevPrice Then
                strIndex = strIndex & "," & Str(aIndex(i))
            End If
        End If
        If dPrevPrice = 0 Then
            dPrevPrice = dPrice
        ElseIf dPrice <> dPrevPrice Then
            If gdBinarySearch(m.Data.FieldArrayHandle(eFld_Price), dPrevPrice, j, eGdSort_Descending, 0, -1) Then
                If dSum < 0 Then '1 Then
                    dSum = 0        'precautionary
                End If
                m.Data(eFld_Volume, j) = dSum
                m.Data(eFld_BarIdx, j) = strIndex
            Else
                'theoretically should never happen
                AddToTable 0, 0, dPrevPrice, 0, dSum
            End If
            If dSum > m.dMaxVol Then m.dMaxVol = dSum
            dPrevPrice = dPrice
            dSum = 0
            strIndex = Str(aIndex(i))
        End If
        
        If bUseVol Then
            dVol = m.TickBars(eBARS_Vol, aIndex(i))
            If m.nShowVolMin > 0 Then
                If dVol >= m.nShowVolMin Then
                    If m.nShowVolMax = 0 Or dVol <= m.nShowVolMax Then dSum = dSum + dVol
                End If
            ElseIf m.nShowVolMax > 0 Then
                If dVol <= m.nShowVolMax Then dSum = dSum + dVol
            ElseIf dVol > 0 Then                'precautionary, should always be true
                dSum = dSum + dVol
            End If
        End If
    Next
                    
    'set volume for the last price in the for loop
    If gdBinarySearch(m.Data.FieldArrayHandle(eFld_Price), dPrice, j, eGdSort_Descending, 0, -1) Then
        m.Data(eFld_Volume, j) = dSum
        m.Data(eFld_BarIdx, j) = strIndex
    Else
        'theoretically should never happen
        AddToTable 0, 0, dPrice, 0, dSum
    End If
        
    m.dMaxVol = m.dMaxVol * 1.1    'make max volume 10% higher than actual
    
    CalcPLData
    geInitTickData m.geTickObj, m.TickBars.BarsHandle
                    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.LoadTable", eGDRaiseError_Raise

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Set m.WindowLink = Nothing
    
    tmr.Enabled = False
    
    ResetData
    m.bUnloading = True
    
    TextIncDecUnregisterForm Me
    
    ClearBuySellButtons True
    Set m.TickBars = Nothing    'reset clears data and handles buffers for bars correctly
    Set m.DailyBar = Nothing
    Set m.Data = Nothing
    Set m.tbOutlineCells = Nothing

    'make sure histogram at top not showing
    pbBid.Visible = False
    pbAsk.Visible = False

    SaveSettings
    
    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile
    
    If Not m.frmBroker Is Nothing Then
        m.frmBroker.Ladder_Unloaded
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.Form_Unload"

End Sub

Private Sub fraOrderBtns_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If
    
End Sub

Private Sub fraRithmic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If

End Sub

Private Sub fraRithmicSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If


End Sub

Private Sub lblAutoExit_Click()
On Error GoTo ErrSection:

'commented out as fix for aardvark 4451
'    If chkAutoExit.Value = vbChecked Then
'        chkAutoExit.Value = vbUnchecked
'    Else
'        chkAutoExit.Value = vbChecked
'    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.lblAutoExit_Click"

End Sub

Private Sub lblAutoExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If


End Sub

Private Sub lblTradePos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If

End Sub

Private Sub pbAsk_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    ShowHelp KeyCode

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.pbAsk_KeyDown"

End Sub

Private Sub pbBid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    ShowHelp KeyCode

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.pbBid_KeyDown"

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim i&, strMsg$
    Dim bMinutized As Boolean, bFullTicksAvail As Boolean

    Select Case Tool.ID
        Case "ID_Symbol"
            ChangeSymbol 0
        Case "ID_Settings"
            If m.eDisplayStyle = eView_Ladder Then
                If lblBrokerDisconnect.Visible Then
                    m.strOrdBarCtrls = m.strOrdBarCtrlsSave
                    FixOrderBarCtrlString
                    frmTickDistributionCfg.ShowMe Me
                    lblBrokerDisconnect.Visible = False
                    cmdBrokerConnect.Visible = False
                Else
                    frmTickDistributionCfg.ShowMe Me
                End If
            Else
                frmMarketDepthCfg.ShowMe Me
            End If
        Case "ID_TextIncrease"
            GridTextIncrease
        Case "ID_TextDecrease"
            GridTextDecrease
        Case "ID_Volume"
            bMinutized = IsMinutized(m.TickBars, bFullTicksAvail)
            If bFullTicksAvail Then
                strMsg = "Data for " & DateFormat(m.nSessionDate) & " is currently compressed data. To view tick distribution for this date, you will need to download all of the ticks.||Do you wish to download now?|"
                If InfBox(strMsg, "?", "+Yes|-No", "Confirmation") = "Y" Then
                    If DownloadTicks(m.TickBars, m.nSessionDate) Then
                        DM_GetBars m.TickBars, m.nSymID, ePRD_EachTick, m.nSessionDate, m.nSessionDate, , , , , False
                    End If
                End If
            Else
                strMsg = "Full tick data for " & DateFormat(m.nSessionDate) & " is not available for download at this time."
                InfBox strMsg, "I"
            End If
        Case "ID_CenterPrice"
            CenterLadderOnCurrPrice
            SetIniFileProperty "AutoCenter", Str(Tool.State), "Price Ladder", g.strIniFile
            If Not m.bEnableAutoCenter Then Tool.State = ssUnchecked

'JM 03-01-2011: keep awhile then remove if not needed
'        Case "ID_OrderBar"
'            If Tool.State = ssUnchecked And ShowOrderBar > eGDOrderBarMode_NotShown Then
'                If ShowOrderBar = eGDOrderBarMode_Right Then
'                    ShowOrderBar = eGDOrderBarMode_LastShownOnRight
'                Else
'                    ShowOrderBar = eGDOrderBarMode_NotShown
'                End If
'                RefreshGrid , , , , True
'            ElseIf Tool.State = ssChecked And ShowOrderBar <= eGDOrderBarMode_NotShown Then
'                If ShowOrderBar = eGDOrderBarMode_LastShownOnRight Then
'                    ShowOrderBar = eGDOrderBarMode_Right
'                Else
'                    ShowOrderBar = eGDOrderBarMode_BottomWide
'                End If
'                RefreshGrid , , , , True
'            End If
        
        Case "ID_VolHistogram"
            ' TLB: the LockWindowUpdate and me.Refresh just makes the transition look a little cleaner
            LockWindowUpdate Me.hWnd
            If Tool.State = ssUnchecked And ShowVolumeBar = 1 Then
                ShowVolumeBar = 0
                RefreshGrid , , , , , True
            ElseIf Tool.State = ssChecked And ShowVolumeBar = 0 Then
                ShowVolumeBar = 1
                RefreshGrid , , , , , True
            End If
            LockWindowUpdate 0
            Me.Refresh
        
        Case "ID_DumpFile"
            If Not m.oBidAskDepth Is Nothing Then
                m.oBidAskDepth.ToggleSalmonFile
                m.bDumpProfileRT = True
            End If

    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.tbToolbar.ToolClick", eGDRaiseError_Show

End Sub

Public Sub ChangeSymbol(ByVal lSymbolID&)
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray     ' Symbol(s) back from the symbol selector
    Dim bTmrSave As Boolean
    Dim strMsg$
    Dim i&
       
'Aardvark 4613 (09-19-2008) - all code for PauseBidAsk commented out (not needed) to fix this issue
'
'    If m.eDisplayStyle = eView_Detail Then
'        g.ChartGlobals.nPauseBidAsk = 1
'        For i = 0 To Forms.Count - 1
'            If TypeOf Forms(i) Is frmTickDistribution Then
'                Forms(i).PauseBidAskDepth
'                DoEvents
'            End If
'        Next
'    End If
        
    If lSymbolID = 0 Then
        If m.eDisplayStyle = eView_Detail Then
            Set astrSymbols = frmSymbolSelector.ShowMe(m.strSym, False, True, "Symbol for Market Depth", True, , , , True)
        Else
            Set astrSymbols = frmSymbolSelector.ShowMe(m.strSym, False, True, "Symbol for the Price Ladder", True, , , , True)
        End If
        If astrSymbols.Size > 0 Then
            lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
        End If
        If lSymbolID = 0 Then Beep
    End If
    
    If lSymbolID <> 0 Then
        bTmrSave = tmr.Enabled
        tmr.Enabled = False
        
        ResetData
        m.strSym = GetSymbol(lSymbolID)
        m.nSymID = lSymbolID
        m.bIsSpreadSymbol = IsSpreadSymbol(m.strSym)
        
        SetAutoExit
        GetContractInformation
        
        While m.bTimerInProg
            DoEvents
        Wend
        
        If m.eDisplayStyle = eView_Ladder Then
            'clear timescale
            fgTickDistribution.OwnerDraw = flexODNone
            fgTickDistribution.TextMatrix(0, 0) = ""
            DoEvents
        ElseIf g.RealTime.Active Then
            i = IsEnableDOM(True)
            If i = 1 Then
                m.bStatusBusy = False
                tmr.Interval = 125
                If g.RealTime.IsMarketDepthMax(m.strSym) Then
                    strMsg = "Market depth information cannot be displayed for " & m.strSym & _
                             " because your account only allows market depth information for " & _
                             Str(MaxSymbolsAllowed(True)) & " symbols at a time."
                    
                    InfBox strMsg, "I", , "Depth of Market"
                    LoadGrid strMsg
                    tmr.Enabled = bTmrSave
                    Exit Sub
                Else
                    Me.Icon = Picture16(ToolbarIcon("ID_MarketDepth"), , True)
                    m.eDisplayStyle = eView_Detail
                End If
            ElseIf i = 2 Then
                If m.bStatusBusy Then
                    tmr.Enabled = True
                Else
'                    strMsg = "Data download is in progress. The symbol " & m.strSym
'                    strMsg = strMsg & " cannot be added to the real time stream at this time."
'                    InfBox strMsg, "I", , "Depth of Market"
                    strMsg = "Waiting data download complete to continue..."
                    LoadGrid strMsg
                    m.bStatusBusy = True
                    tmr.Interval = 3000
                    tmr.Enabled = True
                End If
'                g.ChartGlobals.nPauseBidAsk = 0
                Exit Sub
            Else
                m.bStatusBusy = False
                'not enabled for market depth
                strMsg = "You are not enabled for depth of market information for " & m.strSym & "."
                InfBox strMsg, "I", , "Depth of Market"
                If m.eDisplayStyle = eView_Detail Then
                    Unload Me
                    Exit Sub
                Else
                    LoadGrid strMsg
                    tmr.Interval = 125
                    tmr.Enabled = bTmrSave
                End If
'                g.ChartGlobals.nPauseBidAsk = 0
                Exit Sub
            End If
        Else
            InfBox "Real time needs to be on for depth of market.", "I", , "Depth of Market"
            pbBid.Cls
            pbAsk.Cls
            Unload Me
'            g.ChartGlobals.nPauseBidAsk = 0
            Exit Sub
        End If
                
        GetData
        If m.nShowAccountBar Then FixAcctBarHeader
        
        If m.eDisplayStyle = eView_Ladder Then
            'trigger timescale redraw
            fgTickDistribution.OwnerDraw = flexODOver
            DoEvents
            fgTickDistribution.TextMatrix(0, 0) = ""
'            If g.ChartGlobals.nPauseBidAsk <> 0 Then PauseBidAskDepth
        End If
        
        tmr.Interval = 125
                
        InitQuantityEditor
                
        frmMain.SetWindowLink Me
        Form_Resize
        UpdateEquityPos
        
    End If
        
ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbols = Nothing
    RaiseError "frmTickDistribution.ChangeSymbol", eGDRaiseError_Raise
    
End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim bExtend As Boolean
    Dim strText As String
    Dim lRow As Long
    Dim lCol As Long
    
    m.bPrinting = True
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = taCenterMiddle
        If m.eDisplayStyle = eView_Ladder Then
            .Text = "Tick Distribution (" & m.strSym & ": " & m.TickBars.Prop(eBARS_Desc) & ")" & vbCrLf & DateFormat(Date)
        Else
            .MarginLeft = "0.5in"
            .MarginRight = "0.5in"
            .Text = "Market Depth (" & m.strSym & ": " & m.TickBars.Prop(eBARS_Desc) & ")" & vbCrLf & DateFormat(Date)
        End If
        .TextAlign = taLeftMiddle
        .Font.Bold = False
        
        .Paragraph = ""
        .Paragraph = ""
                
                        
        If m.eDisplayStyle = eView_Ladder Then
            If frmPrintPreview.GoingToFile Then
                frmPrintPreview.GridToTable fgTickDistribution
            Else
                .RenderControl = fgTickDistribution.hWnd
            End If
        Else
            If m.nShowQuoteBar Then .RenderControl = fgQuoteBar.hWnd
            .RenderControl = fgDOMPrint.hWnd
        End If
        
        .EndDoc
    End With
    m.bPrinting = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.GenerateReport", eGDRaiseError_Raise

End Sub

Private Sub ResetData()
On Error GoTo ErrSection:

    Dim i&
    Static bPrevReplayVisible As Boolean
    
    tmr.Enabled = False
    Set m.TickBars = New cGdBars
    Set m.DailyBar = New cGdBars
    m.DailyBar.ArrayMask = eBARS_Eod Or eBARS_BidAsk
    
    m.Data.NumRecords = 0
    m.nLastPriceRec = -1
    m.nLastPriceRow = 0
    m.dLastUpdateTime = 0
    m.bUserSetSize = False
    m.nPriceColWidth = 0
    
    m.dOpen = kNullData
    m.dHigh = kNullData
    m.dLow = kNullData
    m.dLastPrice = kNullData
    m.dLastTradeAtAsk = kNullData
    m.dLastTradeAtBid = kNullData
    
    m.dBestAsk = kNullData
    m.dBestBid = kNullData
    m.nBestBidSize = 0
    m.nBestAskSize = 0
    m.nTradeAtAskSize = 0
    m.nTradeAtBidSize = 0
    m.nLastTradeAtAskRow = 0
    m.nLastTradeAtBidRow = 0
    m.dScrollTickCount = 0
    
    fgTickDistribution.Rows = fgTickDistribution.FixedRows
    fgAskDetail.Rows = fgAskDetail.FixedRows
    fgBidDetail.Rows = fgBidDetail.FixedRows
    pbBid.Cls
    pbAsk.Cls
    
    If m.frmBroker Is Nothing Then
        If bPrevReplayVisible <> frmReplay.Visible Then
            cboAccounts.Clear       'trigger reload of accounts
        End If
    End If
    
    ClearBidAskCells
    If Not m.oBidAskDepth Is Nothing Then
        g.RealTime.RemoveMarketDepthSymbol m.strSym
    End If
    Set m.oBidAskDepth = Nothing
    
    bPrevReplayVisible = frmReplay.Visible
        
    DoEvents
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.ResetData", eGDRaiseError_Raise

End Sub

Public Property Get FloodColor() As Long
On Error GoTo ErrSection:

    FloodColor = m.nFloodColor
    
    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FloodColorGet"

End Property

Public Property Let FloodColor(ByVal nColor&)
On Error GoTo ErrSection:
    
    m.nFloodColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FloodColorLet"
    
End Property

Public Property Get BarColor() As Long
On Error GoTo ErrSection:
    
    If m.nBarColor = RGB(1, 1, 1) Then
        BarColor = 0
    Else
        BarColor = m.nBarColor
    End If

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BarColorGet"
    
End Property

Public Property Let BarColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nBarColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BarColorLet"

End Property

Public Property Get UpColor() As Long
On Error GoTo ErrSection:
    
    UpColor = m.nUpColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.UpColorGet"

End Property

Public Property Let UpColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nUpColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.UpColorLet"

End Property

Public Property Get DownColor() As Long
On Error GoTo ErrSection:
    
    DownColor = m.nDownColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.DownColorGet"

End Property

Public Property Let DownColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nDownColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.DownColorLet"
    
End Property

Public Property Get GridFontSize() As Long
On Error GoTo ErrSection:

    GridFontSize = m.nFontSize

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.GridFontSizeGet"

End Property

Public Property Let GridFontSize(ByVal nSize&)
On Error GoTo ErrSection:
    
    m.nFontSize = nSize
    If m.nFontSize <= 0 Then m.nFontSize = 8

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.GridFontSizeLet"
    
End Property

Public Property Get GridFontName() As String
On Error GoTo ErrSection:

    GridFontName = m.strFont

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.GridFontNameGet"

End Property

Public Property Let GridFontName(ByVal strName$)
On Error GoTo ErrSection:
    
    m.strFont = strName

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.GridFontNameLet"
    
End Property

Public Property Get SessionDate() As Long
On Error GoTo ErrSection:

    SessionDate = m.nSessionDate

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SessionDateGet"

End Property

Public Property Get IsCurrentSession() As Boolean
On Error GoTo ErrSection:

    IsCurrentSession = m.bSessionCurrent

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.IsCurrentSessionGet"
    
End Property

Public Property Let IsCurrentSession(ByVal bCurrent As Boolean)
On Error GoTo ErrSection:

    m.bSessionCurrent = bCurrent

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.IsCurrentSessionLet"

End Property

Public Property Let SessionDate(ByVal nDate As Long)
On Error GoTo ErrSection:

    m.nSessionDate = nDate
    If m.nSessionDate = 0 Then
        m.bSessionCurrent = True
    Else
        m.bSessionCurrent = False
    End If

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SessionDateLet"
    
End Property

Public Sub RefreshGrid(Optional ByVal bDateChanged As Boolean = False, _
    Optional ByVal bColorChanged As Boolean = False, _
    Optional ByVal nSummaryBarHeight As Long = 0, _
    Optional ByVal bMinMaxVolChanged As Boolean = False, _
    Optional ByVal bOrderBarOnOff As Boolean = False, _
    Optional ByVal bVolOnOff As Boolean = False)
On Error GoTo ErrSection:
        
    Dim strCaptionSave$, i&, iWidth&
    
    If m.frmBroker Is Nothing Then
        cboAccounts.Clear       'to trigger repopulattion of accounts combo box
    End If
    
    iWidth = Me.Width
    
    If bOrderBarOnOff Then
        With fgTickDistribution
            i = .ColWidth(eTDCols_Price) + .ColWidth(eTDCols_BidSize) * 3 + .ColWidth(eTDCols_OrderBidX)
            If m.eOrderBarMode = eGDOrderBarMode_NotShown Or m.eOrderBarMode = eGDOrderBarMode_LastShownOnRight Then
                iWidth = i
            Else
                'trigger symbol change if user was on a continuous
                If InStr(m.strSym, "-0") <> 0 Then
                    PositionOrderBar
                    bDateChanged = True
                End If
                iWidth = i + .ColWidth(eTDCols_OrderBid) * 2 + .ColWidth(eTDCols_OrderBidX)
            End If
            
            If Not bVolOnOff And m.nShowVolBar = 1 Then
                iWidth = iWidth + kMinColWidthExt
            End If
        End With
    End If
    
    If bVolOnOff Then        '6192
        If m.nShowVolBar = 1 Then
            iWidth = iWidth + kMinColWidthExt
        Else
            iWidth = iWidth - fgTickDistribution.ColWidth(eTDCols_Volume)
        End If
    End If
    
    ' #6320: can only change size if "normal" (not maximized or minimized)
    If Me.Width <> iWidth And Me.WindowState = vbNormal Then
        On Error Resume Next
        Me.Width = iWidth
        On Error GoTo ErrSection
    End If
    
    If lblAutoExit.Visible Then
        strCaptionSave = lblAutoExit.Caption
    Else
        i = InStr(chkAutoExit.Caption, "Exit:")
        If i > 0 And Len(chkAutoExit.Caption) > (i + 5) Then
            strCaptionSave = Right(chkAutoExit.Caption, Len(chkAutoExit.Caption) - (i + 5))
        Else
            strCaptionSave = "None"
        End If
    End If
    
    If bDateChanged Then
        GetData
    ElseIf bMinMaxVolChanged Then
        LoadTable
        LoadGrid
    Else
        With fgTickDistribution
            .Font.Name = m.strFont
            .Font.Size = m.nFontSize
        End With
        LoadGrid
    End If

    SaveSettings
    
    If pbBid.Height <> nSummaryBarHeight And nSummaryBarHeight >= 200 Then
        pbBid.Height = nSummaryBarHeight
    End If
   
    If tbToolbar.Tools("ID_VolHistogram").State <> ShowVolumeBar Then
        tbToolbar.Tools("ID_VolHistogram").State = ShowVolumeBar
    End If
    
    SetAutoExitCaptions strCaptionSave
    
    m.bUserSetSize = False
    FormResize Me
    If Not g.RealTime.Active Then UpdateEquityPos

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.RefreshGrid", eGDRaiseError_Raise

End Sub

Public Sub RefreshData()
On Error GoTo ErrSection:

    If m.bGridRTInProg Then Exit Sub
    
    GetData True
    If m.eDisplayStyle = eView_Ladder Then Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.RefreshData", eGDRaiseError_Raise

End Sub

Private Sub DisplayPriceLadderRT(ByVal nOldSize&)
On Error GoTo ErrSection:
   
    Dim i&, j&, k&
    Dim dVol#, dPrice#, dNewHigh#, dNewLow#
    Dim dDiffH#, dDiffL#, dDiff#                'for extending blank rows
    
    Dim bUseVol As Boolean
    Dim bShowVol As Boolean
    Dim bReload As Boolean
    Dim bFound As Boolean
    Dim bOneColor As Boolean
    
    Dim nPrevPriceRow&, hPriceArray&
    Dim nLastPriceTextColor&
    Dim strLastTickTime$, nLastTickMinute&
    Dim strText$, n&
    
    Dim aIndex As New cGdArray
            
    If m.bGridRTInProg Then Exit Sub
    m.bGridRTInProg = True
    
    pbBid.Visible = False
    pbAsk.Visible = False
    
    hPriceArray = m.Data.FieldArrayHandle(eFld_Price)
    If hPriceArray = 0 Then
        m.bGridRTInProg = False     'theoretically should never happen
        Exit Sub
    End If
              
    'set these to determine whether to extend blank rows
    dNewHigh = m.dHigh
    dNewLow = m.dLow
    'walk through new data to see if there are any prices not already in the table
    For i = nOldSize To m.TickBars.Size - 1
        dPrice = RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)
        If dPrice >= m.dMaxPrice Or dPrice <= m.dMinPrice Then
            bReload = True
            Exit For
        ElseIf dPrice > dNewHigh Or dPrice < dNewLow Then
            If dPrice > dNewHigh Then
                dNewHigh = dPrice
            ElseIf dPrice < dNewLow Then
                dNewLow = dPrice
            End If
        End If
    Next

    If bReload Then
        LoadTable
        LoadGrid
        m.bGridRTInProg = False
        Exit Sub
    ElseIf dNewHigh > m.dHigh Or dNewLow < m.dLow Then        'check extend blank rows
        If m.dMinMove > 0 Then
            'calculate how many rows would be above/below based on new high or low
            dDiffH = Abs(m.dMaxPrice - dNewHigh) / m.dMinMove
            dDiffL = Abs(m.dMinPrice - dNewLow) / m.dMinMove
            'if either calculated number of rows above/below is less than user requested
            'blank rows then bump the number of blank rows for this session
            If (Int(dDiffH) < m.nBlankRows Or Int(dDiffL) < m.nBlankRows) Then
                If Int(m.nBlankRows / 2) > 10 Then
                    m.nSessionBlankRows = m.nSessionBlankRows + m.nBlankRows / 2
                Else
                    m.nSessionBlankRows = m.nSessionBlankRows + 10
                End If
                LoadTable
                LoadGrid , False
                m.bGridRTInProg = False
                Exit Sub
            End If
            If dNewHigh > m.dHigh Then m.dHigh = dNewHigh
            If dNewLow < m.dLow Then m.dLow = dNewLow
            'color all rows between high & low prices with bar color (aardvark 3797)
            With fgTickDistribution
                gdBinarySearch hPriceArray, m.dHigh, i, eGdSort_Descending, 0, -1
                If ValOfColPrice(i + .FixedRows) = m.Data(eFld_Price, i) Then
                    gdBinarySearch hPriceArray, m.dLow, j, eGdSort_Descending, 0, -1
                    If ValOfColPrice(j + .FixedRows) = m.Data(eFld_Price, j) Then
                        If m.nBarColor = vbWhite Or (bOneColor And m.bBarColorIsLight) Then
                            .Cell(flexcpForeColor, i + .FixedRows, eTDCols_Price, j + .FixedRows, eTDCols_Price) = vbBlack
                        Else
                            .Cell(flexcpForeColor, i + .FixedRows, eTDCols_Price, j + .FixedRows, eTDCols_Price) = vbWhite
                        End If
                        .Cell(flexcpBackColor, i + .FixedRows, eTDCols_Price, j + .FixedRows, eTDCols_Price) = m.nBarColor
                    End If
                End If
            End With
        End If
    End If
    
    'bUseVol determines whether to use the volume data in the tick bars
    '- only stocks have valid volume data in the tick bars object
    'bShowVol determines whether to show volume text/bar as requested by user
    'nShowVolBar is whether user requested volume bar to be shown
    'nShowVolText is whether user requested volume text to be shown
    '- for security type I, volume data is meaningless, don't show at all
    If m.TickBars(eBARS_Vol, 0) <> kNullData Then
        bUseVol = True
        If Left(m.strSym, 1) <> "$" Then
            If m.nShowVolBar Or m.nShowVolText Then bShowVol = True
        End If
    End If
                
    'get minute of last tick
    strLastTickTime = DateFormat(m.TickBars(eBARS_DateTime, m.TickBars.Size - 1), NO_DATE, HH_MM)
    If InStr(strLastTickTime, ":") Then     'precautionary - should always be true
        nLastTickMinute = Val(Right(strLastTickTime, 2))
    Else
        nLastTickMinute = -1
    End If
                    
    nPrevPriceRow = m.nLastPriceRow
    dPrice = RoundToMinMove(m.TickBars(eBARS_Close, m.TickBars.Size - 1), m.dMinMove)
    
    If gdBinarySearch(hPriceArray, dPrice, j, eGdSort_Descending, 0, -1) Then
        If bUseVol And Len(m.Data(eFld_PriceStr, m.nLastPriceRec)) > 0 Then
            If m.eVolumeStyle = eLadderVol_LastTrade Then
                'reset grid price text in case the text had volume with it
                fgTickDistribution.TextMatrix(m.nLastPriceRow, eTDCols_Price) = m.Data(eFld_PriceStr, m.nLastPriceRec)
            ElseIf m.eVolumeStyle = eLadderVol_BidAsk Then
                'reset grid price text in case the text had volume with it
                If m.nLastTradeAtAskRow >= fgTickDistribution.FixedRows And m.nLastTradeAtAskRow < fgTickDistribution.Rows Then
                    If gdBinarySearch(hPriceArray, m.dLastTradeAtAsk, i, eGdSort_Descending, 0, -1) Then
                        fgTickDistribution.TextMatrix(m.nLastTradeAtAskRow, eTDCols_Price) = m.Data(eFld_PriceStr, i)
                    End If
                End If
                'reset grid price text in case the text had volume with it
                If m.nLastTradeAtBidRow >= fgTickDistribution.FixedRows And m.nLastTradeAtBidRow < fgTickDistribution.Rows Then
                    If gdBinarySearch(hPriceArray, m.dLastTradeAtBid, i, eGdSort_Descending, 0, -1) Then
                        fgTickDistribution.TextMatrix(m.nLastTradeAtBidRow, eTDCols_Price) = m.Data(eFld_PriceStr, i)
                    End If
                End If
            End If
        End If
        'update last price and its record number in table
        If j > 0 Then m.nLastPriceRec = j
    Else
        'theoretically should never happen
        'act like there's no change and just quit
        m.bGridRTInProg = False     'theoretically should never happen
        Exit Sub
    End If
            
    Dim bVolFilterOk As Boolean
    
    For i = nOldSize To m.TickBars.Size - 1
        If bUseVol Then
            dVol = m.TickBars(eBARS_Vol, i)
        Else
            dVol = 1
        End If
        
        'reset volume filter flag
        If m.nShowVolMin > 0 Or m.nShowVolMax > 0 Then
            bVolFilterOk = False
        Else
            bVolFilterOk = True
        End If
        
        dPrice = RoundToMinMove(m.TickBars(eBARS_Close, i), m.dMinMove)
        nLastPriceTextColor = GetLastPriceTextColor(dPrice)
        
        ' accumulate volume while price is the same     - 6209
        If dPrice = m.dLastPrice Then
            m.dLastPriceVol = m.dLastPriceVol + dVol
        Else ' otherwise reset the volume
            m.dLastPriceVol = dVol
            m.dLastPrice = dPrice
        End If
        ' accumulate volume for trades at ask/bid
        If m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
            If dPrice = m.dLastTradeAtAsk Then
                m.nTradeAtAskSize = m.nTradeAtAskSize + dVol
            Else
                m.nTradeAtAskSize = dVol
                m.dLastTradeAtAsk = dPrice
                ' if new TradeAtAsk is same price as old TradeAtBid, then clear the TradeAtBid volume
                If m.dLastTradeAtAsk <= m.dLastTradeAtBid Then
                    m.nTradeAtBidSize = 0
                    m.dLastTradeAtBid = 0
                End If
            End If
        ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
            If dPrice = m.dLastTradeAtBid Then
                m.nTradeAtBidSize = m.nTradeAtBidSize + dVol
            Else
                m.nTradeAtBidSize = dVol
                m.dLastTradeAtBid = dPrice
                ' if new TradeAtBid is same price as old TradeAtAsk, then clear the TradeAtAsk volume
                If m.dLastTradeAtBid >= m.dLastTradeAtAsk Then
                    m.nTradeAtAskSize = 0
                    m.dLastTradeAtAsk = 0
                End If
            End If
        End If
        
        'find price in table
        If gdBinarySearch(hPriceArray, dPrice, j, eGdSort_Descending, 0, -1) And _
            Len(m.Data(eFld_PriceStr, j)) > 0 Then
            'update volume field in table
            If m.nShowVolMin > 0 Then
                If dVol >= m.nShowVolMin Then
                    If m.nShowVolMax = 0 Or dVol <= m.nShowVolMax Then
                        m.Data(eFld_Volume, j) = m.Data(eFld_Volume, j) + dVol
                        bVolFilterOk = True
                    End If
                End If
            ElseIf m.nShowVolMax > 0 Then
                If dVol <= m.nShowVolMax Then
                    m.Data(eFld_Volume, j) = m.Data(eFld_Volume, j) + dVol
                    bVolFilterOk = True
                End If
            Else
                m.Data(eFld_Volume, j) = m.Data(eFld_Volume, j) + dVol
            End If
            'update grid rows
            With fgTickDistribution
                k = j + .FixedRows
                If k >= 0 And k < .Rows Then
                    If m.TickBars(eBARS_Flags, i) = eTICK_AtAsk Then
                        m.nLastTradeAtAskRow = k
                    ElseIf m.TickBars(eBARS_Flags, i) = eTICK_AtBid Then
                        m.nLastTradeAtBidRow = k
                    End If
                    If dPrice = m.dLastPrice Then m.nLastPriceRow = k
                    
                    'update grid price data (only do this for the very last tick)
                    If i = m.TickBars.Size - 1 Then
                        If bUseVol Then
                            If m.eVolumeStyle = eLadderVol_LastTrade Then
                                .TextMatrix(k, eTDCols_Price) = Str(m.dLastPriceVol) & Space(3) & m.Data(eFld_PriceStr, j)
                            ElseIf m.eVolumeStyle = eLadderVol_BidAsk Then
                                If m.nTradeAtAskSize > 0 Then
                                    strText = .TextMatrix(m.nLastTradeAtAskRow, eTDCols_Price)
                                    .TextMatrix(m.nLastTradeAtAskRow, eTDCols_Price) = Str(m.nTradeAtAskSize) & Space(3) & strText
                                End If
                                If m.nTradeAtBidSize > 0 Then
                                    strText = .TextMatrix(m.nLastTradeAtBidRow, eTDCols_Price)
                                    .TextMatrix(m.nLastTradeAtBidRow, eTDCols_Price) = Str(m.nTradeAtBidSize) & Space(3) & strText
                                End If
                            End If
                        End If
                    End If
                    'update grid volume data
                    'volume for security type I is meaningless, don't show even if user asked for it
                    dVol = m.Data(eFld_Volume, j)
                    If bShowVol And bVolFilterOk Then
                        If m.nShowVolText And dVol > 0 Then
                            If bUseVol Then
                                .TextMatrix(k, eTDCols_Volume) = Format(dVol, "#,##0")
                            Else
                                .TextMatrix(k, eTDCols_Volume) = Format(dVol, "#,##0") & " trades"
                            End If
                        ElseIf .TextMatrix(k, eTDCols_Volume) <> "" Then
                            .TextMatrix(k, eTDCols_Volume) = ""
                        End If
                        'flood volume bar with flood color
                        If dVol > 0 And m.dMaxVol > 0 And m.nShowVolBar Then
                            If dVol > m.dMaxVol Then
                                m.dMaxVol = dVol * 1.1      'make max volume 10% higher than actual
                            End If
                            .Cell(flexcpFloodColor, k, eTDCols_Volume) = m.nFloodColor
                            .Cell(flexcpFloodPercent, k, eTDCols_Volume) = (dVol / m.dMaxVol * 100) * -1
                        Else
                            .Cell(flexcpFloodPercent, k, eTDCols_Volume) = 0
                        End If
                    End If
                    'update bars index in table
                    bFound = False
                    strText = m.Data(eFld_BarIdx, j)
                    aIndex.SplitFields strText, ","
                    For n = 0 To aIndex.Size - 1
                        If m.TickBars(eBARS_Close, i) = m.TickBars(eBARS_Close, Val(aIndex(n))) Then
                            bFound = True
                            Exit For
                        End If
                    Next
                    If Not bFound Then
                        If Len(m.Data(eFld_BarIdx, j)) > 0 Then
                            m.Data(eFld_BarIdx, j) = m.Data(eFld_BarIdx, j) & "," & Str(i)
                        Else
                            m.Data(eFld_BarIdx, j) = Str(i)
                        End If
                    End If
                End If
            End With
            
            DisplayBestBidAsk False
        End If
    Next    'end loop through bars
                            
    If nLastTickMinute > 0 Then
        If m.nPrevMinute <> nLastTickMinute Then
            m.nPrevMinute = nLastTickMinute
            
            If bShowVol And m.nShowVolBar Then
            
                With fgTickDistribution
                    .Redraw = flexRDNone
                    'trigger redraw of all rows in volume column
                    For i = .FixedRows To .Rows - 1
                        If m.nShowVolText Then
                            strText = .TextMatrix(i, eTDCols_Volume)
                            If InStr(UCase(strText), "TRADES") > 0 Then
                                dVol = ValOfText(Left(strText, Len(strText) - 7))
                            Else
                                dVol = ValOfText(strText)
                            End If
                        Else
                            dPrice = ValOfColPrice(i)
                            'find price in table        - aardvark 4557
                            If gdBinarySearch(hPriceArray, dPrice, j, eGdSort_Descending, 0, -1) Then
                                If Len(m.Data(eFld_PriceStr, j)) > 0 Then
                                    dVol = m.Data(eFld_Volume, j)
                                End If
                            End If
                        End If
                    
                        If dVol > 0 And m.dMaxVol > 0 Then
                            .Cell(flexcpFloodColor, k, eTDCols_Volume) = m.nFloodColor
                            .Cell(flexcpFloodPercent, i, eTDCols_Volume) = (dVol / m.dMaxVol * 100) * -1
                        Else
                            .Cell(flexcpFloodPercent, i, eTDCols_Volume) = 0
                        End If
                    Next
                    .Redraw = flexRDBuffered
                End With
                
            End If
        End If
    End If
    
    ' color rows between open and close price with up or down color
    If m.dLastPrice > m.dOpen Then
        m.nCurrUpDownColor = m.nUpColor
    ElseIf m.dLastPrice < m.dOpen Then
        m.nCurrUpDownColor = m.nDownColor
    Else
        m.nCurrUpDownColor = m.nBarColor
    End If

    If m.nBarColor = m.nUpColor And m.nBarColor = m.nDownColor Then bOneColor = True
    If m.nBarColor = 0 Then m.nBarColor = RGB(1, 1, 1) 'see flexgrid documentation
    If m.nCurrUpDownColor = 0 Then m.nCurrUpDownColor = RGB(1, 1, 1)
        
    With fgTickDistribution
        If m.nLastPriceRow < .Rows Then
            .Col = eTDCols_Price
            'use bar color for rows where last price used to be to where last price is now
            If m.nBarColor = vbWhite Or (bOneColor And m.bBarColorIsLight) Then
                .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price, nPrevPriceRow, eTDCols_Price) = vbBlack
            Else
                .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price, nPrevPriceRow, eTDCols_Price) = vbWhite
            End If
            .Cell(flexcpBackColor, m.nLastPriceRow, eTDCols_Price, nPrevPriceRow, eTDCols_Price) = m.nBarColor
            .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price, nPrevPriceRow, eTDCols_Price) = m.nBarColor
            .Cell(flexcpFloodPercent, m.nLastPriceRow, eTDCols_Price, nPrevPriceRow, eTDCols_Price) = 100

            'use up/down color for rows between open and current last price
            If m.nCurrUpDownColor = vbWhite Or (bOneColor And IsLightColor(m.nCurrUpDownColor)) Then
                .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = vbBlack
            Else
                .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = vbWhite
            End If
            .Cell(flexcpBackColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = m.nCurrUpDownColor
            .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = m.nCurrUpDownColor
            .Cell(flexcpFloodPercent, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = 100

            ' invert last price cell
            If nLastPriceTextColor = vbWhite Then
                .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price) = RGB(1, 1, 1)
            ElseIf bOneColor Then
                If nLastPriceTextColor = vbBlack Then nLastPriceTextColor = RGB(1, 1, 1)
                .Cell(flexcpBackColor, m.nLastPriceRow, eTDCols_Price) = nLastPriceTextColor
                .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price) = .BackColor
            Else
                .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price) = .BackColor
            End If
            .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price) = nLastPriceTextColor
            .Cell(flexcpFloodPercent, m.nLastPriceRow, eTDCols_Price) = 100
            
        End If
        'trigger redraw of timescale
        '.TextMatrix(0, eTDCols_Volume) = ""
    End With

ErrExit:
    m.bGridRTInProg = False
    Exit Sub
    
ErrSection:
    m.bGridRTInProg = False
    RaiseError "frmTickDistribution.DisplayPriceLadderRT", eGDRaiseError_Raise

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:
       
    Static bLogDOMUpdate As Boolean
    Dim bNewDay As Boolean
    Dim nOldSize&, dTick#, i&, strText$
    Dim bHasDepthData As Boolean
    Dim bHasNewTicks As Boolean, bNewBestBidAsk As Boolean
    
    TimerStart "frmTickDistribution.tmr." & m.strSym & "." & Str(hWnd)
    If tmr.Tag = "Unload Now" Then
        tmr.Tag = ""
        tmr.Enabled = False
        Unload Me
        Exit Sub
    End If
    
    If m.bUnloading Then Exit Sub
    
    If vseOrderBar.Visible Then
        If vseBracketOrder.Appearance = apInset Then
            If Not Me Is Screen.ActiveForm Then
                If Not TypeOf Screen.ActiveForm Is frmAsk Then
                    'user left the ladder before completing the bracket order (eg - user could cancel order from trade console)
                    ClearBuySellButtons True
                End If
            End If
        End If
        'this tag is set to toggle exit favorite buttons show/hide (need to redraw order bar)
        If tmr.Tag = "OrderbarReset" Then
            PositionOrderBar
            tmr.Tag = ""
        End If
    End If
    
    If m.bTimerInProg Or m.bGridRTInProg Then Exit Sub
    
    If Screen.ActiveForm Is Me And MouseIsPressed Then Exit Sub     '5597
    
    'sync auto journal flag
    If g.Broker.AutoJournalPopUp Then
        chkAutoJournal.Value = vbChecked
    Else
        chkAutoJournal.Value = vbUnchecked
    End If
    
    'If Not g.RealTime.Active Or m.bSessionCurrent = False Then
    If Not m.bSessionCurrent Or (Not g.RealTime.Active And m.eOrderBarMode > eGDOrderBarMode_NotShown) Then
        m.bTimerInProg = True
        GoTo ErrExit
    End If
    
    If m.eDisplayStyle = eView_Detail And m.bStatusBusy Then
        ChangeSymbol m.nSymID
        If m.bStatusBusy Then Exit Sub
    End If
    
    m.bTimerInProg = True
    
    If fgTickDistribution.MouseCol < 0 Then
        ' if mouse has moved off form, then call HandleMouseMove to clear stuff
        HandleMouseMove
    End If
    
    If g.RealTime.Active And tmr.Interval <> m.nIntervalRT Then
        tmr.Interval = m.nIntervalRT                         '125
    End If
    
    ' update the ticks
    nOldSize = m.TickBars.Size
    If g.RealTime.UpdateBars(m.TickBars, bNewDay) Then
        If bNewDay Then
            m.bTimerInProg = False
DebugLog m.strSym & " new day on ladder."
            RefreshData
DebugLog m.strSym & " new day on ladder done."
            bLogDOMUpdate = True
            Exit Sub
        ElseIf m.TickBars.Size > nOldSize Then
            bHasNewTicks = True
        End If
    End If

If m.bDumpProfileRT Then
    gdResetProfiles 900, 909
    gdStartProfile 900
    gdStartProfile 901
End If
           
    ' update the "daily bar" (best bid and ask along with new price)
    bNewBestBidAsk = g.RealTime.UpdateBidAsk(m.DailyBar)
    If m.DailyBar(eBARS_Bid, m.DailyBar.Size - 1) <> kNullData Then
        m.dBestBid = RoundToMinMove(m.DailyBar(eBARS_Bid, m.DailyBar.Size - 1), m.dMinMove)
    End If
    If m.DailyBar(eBARS_Ask, m.DailyBar.Size - 1) <> kNullData Then
        m.dBestAsk = RoundToMinMove(m.DailyBar(eBARS_Ask, m.DailyBar.Size - 1), m.dMinMove)
    End If

    If bNewBestBidAsk And Not m.oBidAskDepth Is Nothing Then
        ' remove market depth bids > best bid and market depth asks < best ask
        m.oBidAskDepth.CheckBestBidAsk RoundToMinMove(m.DailyBar(eBARS_Bid, m.DailyBar.Size - 1), m.dMinMove), _
            RoundToMinMove(m.DailyBar(eBARS_Ask, m.DailyBar.Size - 1), m.dMinMove)
    End If
If m.bDumpProfileRT Then gdStopProfile 901

    ' update market depth
    If m.oBidAskDepth Is Nothing Then
'        If m.eDisplayStyle = eView_Detail Or g.ChartGlobals.nPauseBidAsk = 0 Then
            Set m.oBidAskDepth = g.RealTime.AddMarketDepthSymbol(m.strSym, True)
'        End If
        If Not m.oBidAskDepth Is Nothing Then
'            g.ChartGlobals.nPauseBidAsk = 0
            If m.oBidAskDepth.ClientInit(m.strSym, m.dMinMove, m.nSessionDate) Then
                bHasDepthData = True
            Else
                g.RealTime.RemoveMarketDepthSymbol m.strSym
                bHasDepthData = False
                Set m.oBidAskDepth = Nothing
            End If
        Else
            'Note: Intentionally not resetting the global pause bid/ask flag.
            '   We will only get here if request to add market depth symbol fails.
            '   If request fails for tick ladder, we don't care, but we don't want
            '   to reset the flag as a DOM display may be trying to get data.
            '   Request should NEVER fail for DOM because we do all the checks
            '   for enablement & max symbols allowed PRIOR to getting here.
            '   If request fails for DOM, something is seriously broken and
            '   resetting the flag will only hide the bug & compound the problem.
            '   Customers should restart at this point and Genesis employees
            '   should report the bug so it can be fixed.
        End If
'    ElseIf g.ChartGlobals.nPauseBidAsk <> 0 And m.eDisplayStyle = eView_Ladder Then
'        PauseBidAskDepth
    End If
    
    If m.oBidAskDepth Is Nothing Then
        bHasDepthData = False
    ElseIf Not bHasDepthData Then
If m.bDumpProfileRT Then gdStartProfile 902

If bLogDOMUpdate Then DebugLog m.strSym & " requesting DOM for new day."
        
        bHasDepthData = m.oBidAskDepth.Update(m.dLastUpdateTime)

If bLogDOMUpdate And bHasDepthData Then
    DebugLog m.strSym & " DOM for new day received."
    bLogDOMUpdate = False
End If

If m.bDumpProfileRT Then gdStopProfile 902
    End If

    ' display ladder
    'LoadLadderGrid displays data for non-current sessions
    'DisplayPriceLadderRT is only for real time
    If m.eDisplayStyle = eView_Ladder Then
        If bHasNewTicks Then
        
If m.bDumpProfileRT Then gdStartProfile 903
            DisplayPriceLadderRT nOldSize
If m.bDumpProfileRT Then gdStopProfile 903

If m.bDumpProfileRT Then gdStartProfile 904
            If m.bEnableAutoCenter Then CheckAutoCenter
If m.bDumpProfileRT Then gdStopProfile 904

        ElseIf bNewBestBidAsk And Not bHasDepthData Then
            DisplayBestBidAsk True
        Else
            If m.bEnableAutoCenter Then CheckAutoCenter
        End If
        
If m.bDumpProfileRT Then gdStartProfile 905
        If Me.WindowState = vbNormal Then   ' aardvark 6776
            With fgTickDistribution
                'resize price column if needed      - aardvark 3655
                If m.nLastPriceRow > .FixedRows And m.nLastPriceRow < .Rows Then
                    strText = .TextMatrix(m.nLastPriceRow, eTDCols_Price) & "  "
                    i = Me.TextWidth(strText)       '- aardvark 6736
                    If i > m.nPriceColWidth Then
                        .AutoSize eTDCols_Price
                        m.nPriceColWidth = .ColWidth(eTDCols_Price)
                        Me.Width = Me.Width + (i - m.nPriceColWidth)
                    End If
                End If
            End With
        End If
If m.bDumpProfileRT Then gdStopProfile 905
    End If
    
    ' display market depth
    If bHasDepthData Then
        If m.eDisplayStyle = eView_Ladder Then
If m.bDumpProfileRT Then gdStartProfile 906
            DisplayMarketDepthOnLadder
If m.bDumpProfileRT Then gdStopProfile 906
        Else
            m.bGridRTInProg = True
            LoadDetailGrid
            m.bGridRTInProg = False
        End If
    ElseIf m.eDisplayStyle = eView_Ladder Then
        CheckLadderDOM
    End If
    
If m.bDumpProfileRT Then gdStartProfile 907
    If m.nShowQuoteBar = 1 Then DisplayQBar
    If m.nShowAccountBar = 1 Then DisplayAcctBar
If m.bDumpProfileRT Then gdStopProfile 907
    
If m.bDumpProfileRT Then gdStartProfile 908
    If vseOrderBar.Visible Then UpdateEquityPos
If m.bDumpProfileRT Then gdStopProfile 908
                                                                                              
ErrExit:
    'check disconnect
    If m.eOrderBarMode <> eGDOrderBarMode_NotShown Then

If m.bDumpProfileRT Then gdStartProfile 909
        
'JM 10-04-2011: With 6.1 release, all accounts are treated as broker accounts and must be connected to trade.
'   Leave commented out code awhile then remove if all okay.
'        If g.Broker.AccountTypeForID(TradeAccountID) = eTT_AccountType_SimTrade Then
'            If lblBrokerDisconnect.Visible Then
'                m.strOrdBarCtrls = m.strOrdBarCtrlsSave
'                PositionOrderBar
'                lblBrokerDisconnect.Visible = False
'            End If
'        Else
            If ConnectionStatus = eGDConnectionStatus_Connected Then
                If lblBrokerDisconnect.Visible Then
                    m.strOrdBarCtrls = m.strOrdBarCtrlsSave
                    FixOrderBarCtrlString
                    PositionOrderBar
                    lblBrokerDisconnect.Visible = False
                    cmdBrokerConnect.Visible = False
                End If
            ElseIf Not lblBrokerDisconnect.Visible Or chkConfirmOrder.Visible Then
                m.strOrdBarCtrls = kOrdBarDisconnect
                FixOrderBarCtrlString
                PositionOrderBar
                cmdClearQty.Visible = False
                txtTradeQty.Visible = False
                vscrQty.Visible = False
                lblBrokerDisconnect.Move cmdClearQty.Left, cmdClearQty.Top, fraOrderBtns.Width
                If m.eOrderBarMode = eGDOrderBarMode_Right Then
                    cmdBrokerConnect.Move Me.fraOrderBtns.Left + (fraOrderBtns.Width - cmdBrokerConnect.Width) / 2, lblBrokerDisconnect.Top + lblBrokerDisconnect.Height + 45
                Else
                    fraOrderBtns.Width = lblBrokerDisconnect.Width * 2
                    cmdBrokerConnect.Move lblBrokerDisconnect.Left + lblBrokerDisconnect.Width * 0.75, lblBrokerDisconnect.Top - lblBrokerDisconnect.Height * 0.1
                End If
                lblBrokerDisconnect.Visible = True
                cmdBrokerConnect.Visible = True
                chkConfirmOrder.Visible = False
            End If
'        End If
        If vseOrderBar.Visible Then ExitFavoritesCheck Me, m.strPos

If m.bDumpProfileRT Then gdStopProfile 909

    End If
    
If m.bDumpProfileRT Then
    gdStopProfile 900
    DebugLog "Ladder DOM " & m.strSym & ": (" & Str(m.TickBars.Size - nOldSize) _
        & ") (" & Str(bHasDepthData) & ")" & vbCrLf & gdGetProfiles(900, 909, vbCrLf)
    If bHasNewTicks And bHasDepthData Then m.bDumpProfileRT = False
End If
    TimerEnd "frmTickDistribution.tmr." & m.strSym & "." & Str(hWnd), tmr.Interval
    
    m.bTimerInProg = False
    Exit Sub
    
ErrSection:
    m.bTimerInProg = False
    RaiseError "frmTickDistribution.Timer", eGDRaiseError_Show

End Sub

Private Sub CalcPLData(Optional ByVal bRedraw As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, lRedraw&
    Dim dPrice#, dChangeAmt#
    Dim strPos$
    
    If m.dZeroPLPrice <= 0 Then     '5172
        With fgTickDistribution
            If .Rows > .FixedRows + 1 Then
                If Len(fgTickDistribution.TextMatrix(fgTickDistribution.FixedRows + 1, eTDCols_PL)) > 0 Then
                    'true if user just flattened a position
                    lRedraw = .Redraw
                    .Redraw = flexRDNone
                    For i = .FixedRows To .Rows - 1
                        .TextMatrix(i, eTDCols_PL) = ""
                    Next
                    .Redraw = lRedraw
                End If
            End If
        End With
        Exit Sub
    End If
    
    If m.nQuantity = 0 Then
        'precautionary to prevent division by zero (theoretically should never get here)
        If SecurityType(m.nSymID) = "S" Then
            m.nQuantity = 100
        Else
            m.nQuantity = 1
        End If
    End If
    
    'strPos = Parse(g.Broker.PositionString(TradeAccountID, m.nSymID, 0&), "|", 1)
    strPos = Parse(GetPositionString, "|", 1)
    
    dChangeAmt = m.TickBars.Prop(eBARS_TickValue) / m.TickBars.Prop(eBARS_TickMove) * m.nQuantity
    For i = 0 To m.Data.NumRecords - 1
        If UCase(strPos) = "SHORT" Then
            dPrice = (m.dZeroPLPrice - m.Data(eFld_Price, i)) * dChangeAmt
        Else
            dPrice = (m.Data(eFld_Price, i) - m.dZeroPLPrice) * dChangeAmt
        End If
        m.Data(eFld_PL, i) = dPrice 'Format(dPrice, "$#,##0;($#,##0)")
        If bRedraw Then
            With fgTickDistribution
                If i + 1 < .Rows Then
                    .TextMatrix(i + 1, eTDCols_PL) = Format(m.Data(eFld_PL, i), "$#,##0;($#,##0)")
                End If
            End With
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.CalcPLData", eGDRaiseError_Raise
End Sub

Private Sub DisplayBestBidAsk(ByVal bRedraw As Boolean)
On Error GoTo ErrSection:
       
    'display best bid/ask size if available else display best bid/ask prices
    Dim strBidInfo$, strAskInfo$
    Dim nBidRow&, nAskRow&, i&
    Dim dBidPrice#, dAskPrice#
    Dim hPriceArray&
        
    'don't want to do this if depth data is available
    If Not m.oBidAskDepth Is Nothing Then
        If Not m.oBidAskDepth.BidAskData Is Nothing Then
            If m.oBidAskDepth.BidAskData.NumRecords > 0 Then
                Exit Sub
            End If
        End If
    End If
        
    hPriceArray = m.Data.FieldArrayHandle(eFld_Price)
    If hPriceArray = 0 Then Exit Sub
            
    GetDailyBarsBidAsk
    
    With fgTickDistribution
        'find price row in the price ladder that matches the best bid/ask values
        If m.dBestBid <> kNullData Then
            m.dBestBid = RoundToMinMove(m.dBestBid, m.dMinMove)
            If gdBinarySearch(m.Data.FieldArrayHandle(eFld_Price), m.dBestBid, i, eGdSort_Descending, 0, -1) Then
                nBidRow = i + .FixedRows
                If m.nBestBidSize > 0 Then
                    strBidInfo = Str(m.nBestBidSize)
                Else
                    strBidInfo = ""             'm.DailyBar.PriceDisplay(m.dBestBid)
                End If
            End If
        End If
        If m.dBestAsk <> kNullData Then
            m.dBestAsk = RoundToMinMove(m.dBestAsk, m.dMinMove)
            If gdBinarySearch(m.Data.FieldArrayHandle(eFld_Price), m.dBestAsk, i, eGdSort_Descending, 0, -1) Then
                nAskRow = i + .FixedRows
                If m.nBestAskSize > 0 Then
                    strAskInfo = Str(m.nBestAskSize)
                Else
                    strAskInfo = ""             'm.DailyBar.PriceDisplay(m.dBestAsk)
                End If
            End If
        End If

        'clear out any previous best bid/ask data
        If m.nPrevBidRow > .FixedRows And m.nPrevBidRow < .Rows - 1 Then
            .TextMatrix(m.nPrevBidRow, eTDCols_BidSize) = ""
        End If
        If m.nPrevAskRow > .FixedRows And m.nPrevAskRow < .Rows - 1 Then
            .TextMatrix(m.nPrevAskRow, eTDCols_AskSize) = ""
        End If
                
        If nBidRow > .FixedRows And nBidRow < .Rows - 1 Then
            .TextMatrix(nBidRow, eTDCols_BidSize) = strBidInfo
        End If
        If nAskRow > .FixedRows And nAskRow < .Rows - 1 Then
            .TextMatrix(nAskRow, eTDCols_AskSize) = strAskInfo
        End If
        
        m.nPrevBidRow = nBidRow
        m.nPrevAskRow = nAskRow
    End With
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.DisplayBestBidAsk", eGDRaiseError_Raise
End Sub

Private Sub AddToTable(ByVal dPL#, ByVal nBidSize&, ByVal dPrice#, _
    ByVal nAskSize&, ByVal dVolume#, Optional ByVal bAddToTop As Boolean = False)
On Error GoTo ErrSection:
    
    Dim i&, strPrice$
    
    strPrice = m.TickBars.PriceDisplay(dPrice)
    If bAddToTop Then
        i = 0
        m.Data.AddRecord "", 0
    Else
        i = m.Data.NumRecords
    End If
    
    m.Data(eFld_PL, i) = dPL
    m.Data(eFld_BidSize, i) = nBidSize
    m.Data(eFld_Price, i) = dPrice
    m.Data(eFld_AskSize, i) = nAskSize
    m.Data(eFld_Volume, i) = dVolume
        
    m.Data(eFld_PriceStr, i) = strPrice
    m.Data(eFld_BarIdx, i) = -1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.AddToTable", eGDRaiseError_Raise

End Sub

Private Function GetLastPriceTextColor(Optional ByVal dNewPrice As Double = kNullData) As Long
On Error GoTo ErrSection:

    Static nPrevTextColor&
    Dim nTextColor&, iBarColor&, i&
    Dim dPrevPrice#

    If dNewPrice = kNullData Then
        'initial load
        dNewPrice = m.TickBars(eBARS_Close, m.TickBars.Size - 1)
        For i = m.TickBars.Size - 1 To 0 Step -1
            dPrevPrice = m.TickBars(eBARS_Close, i)
            If dPrevPrice <> dNewPrice Then Exit For
        Next
    Else
        dPrevPrice = m.dLastPrice
    End If
    If m.nBarColor = RGB(1, 1, 1) Then
        iBarColor = 0
    Else
        iBarColor = m.nBarColor
    End If
    If iBarColor = m.nUpColor And iBarColor = m.nDownColor Then
        If m.nFixedPriceColor >= 0 Then
            nTextColor = m.nFixedPriceColor
        Else
            If dNewPrice > dPrevPrice Then
                 nTextColor = RGB(0, 128, 0)
            ElseIf dNewPrice < dPrevPrice Then
                 nTextColor = vbRed
            Else
                 nTextColor = nPrevTextColor
            End If
        End If
    Else
        If dNewPrice > dPrevPrice Then
             nTextColor = m.nUpColor
        ElseIf dNewPrice < dPrevPrice Then
             nTextColor = m.nDownColor
        Else
             nTextColor = nPrevTextColor
        End If
    End If
        
    nPrevTextColor = nTextColor
    GetLastPriceTextColor = nTextColor
    
    Exit Function

ErrSection:
    RaiseError "frmTickDistribution.GetLastPriceTextColor"

End Function

Public Property Get ShowVolumeBar() As Long
On Error GoTo ErrSection:

    ShowVolumeBar = m.nShowVolBar

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowVolumeBarGet"

End Property

Public Property Let ShowVolumeBar(ByVal nShow&)
On Error GoTo ErrSection:
    
    m.nShowVolBar = nShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowVolumeBarLet"

End Property

Public Property Get ShowVolumeText() As Long
On Error GoTo ErrSection:

    ShowVolumeText = m.nShowVolText

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowVolumeTextGet"

End Property

Public Property Let ShowVolumeText(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowVolText = nShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowVolumeTextLet"

End Property

Public Property Get ShowTickline() As Long
On Error GoTo ErrSection:

    ShowTickline = m.nShowTickLine

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowTickLineGet"

End Property

Public Property Let ShowTickline(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowTickLine = nShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowTickLineLet"

End Property

Public Property Get TickLineColor() As Long
On Error GoTo ErrSection:
    
    TickLineColor = m.nTickLineColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.TickLineColorGet"

End Property

Public Property Let TickLineColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nTickLineColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.TickLineColorLet"

End Property

Public Property Get ShowProfitLoss() As Long
On Error GoTo ErrSection:

    ShowProfitLoss = m.nShowProfitLoss

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowProfitLossGet"

End Property

Public Property Let ShowProfitLoss(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowProfitLoss = nShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowProfitLossLet"

End Property

Private Function IsNewMinute(ByVal nOldSize&) As Boolean
On Error GoTo ErrSection:

    Dim k&, dTime#, dTime1#
    
    k = m.TickBars.Size()
    IsNewMinute = False
    
    If k > nOldSize Then
        dTime = m.TickBars(eBARS_DateTime, nOldSize - 1)
        dTime1 = m.TickBars(eBARS_DateTime, k - 1)
        If dTime <> dTime1 Then
            IsNewMinute = True
        End If
    End If

    Exit Function

ErrSection:
    RaiseError "frmTickDistribution.IsNewMinute"

End Function

Public Property Get BidColor() As Long
On Error GoTo ErrSection:

    BidColor = m.nBidColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidColorGet"

End Property

Public Property Let BidColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nBidColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidColorLet"

End Property

Public Property Get AskColor() As Long
On Error GoTo ErrSection:

    AskColor = m.nAskColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.AskColorGet"

End Property

Public Property Let AskColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nAskColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.AskColorLet"

End Property

Public Property Get SecType() As String
On Error GoTo ErrSection:

    SecType = SecurityType(m.DailyBar)

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SecTypeGet"

End Property

Private Sub DisplayMarketDepthOnLadder()
On Error GoTo ErrSection:
    
    Dim BidAskTable As New cGdTable
    Dim aBestBids As cGdArray
    Dim aBestAsks As cGdArray
    Dim aIdx As cGdArray
    
    Dim dPriceInGrid#, dPriceFound#, dDontCare#
    Dim nBidSize&, nAskSize&, nLargestSize&
    Dim i&, j&, k&

    Dim nLargestBidSize&, nLargestAskSize&

    If m.bHideLadderDOM Then
        m.bLadderHasDOM = False
        Exit Sub
    End If
        
    If m.oBidAskDepth Is Nothing Then
        m.bLadderHasDOM = False
        Exit Sub
    End If
    
    CheckLadderDOM
    If m.bDepthOfMarketBad Then Exit Sub

    Set BidAskTable = m.oBidAskDepth.BidAskData
    If BidAskTable Is Nothing Then
        m.bLadderHasDOM = False
        Exit Sub
    End If
    If BidAskTable.NumRecords <= 0 Then
        m.bLadderHasDOM = False
        Exit Sub
    End If
                                    
    Set aBestBids = m.oBidAskDepth.BestBids
    Set aBestAsks = m.oBidAskDepth.BestAsks
    
    m.oBidAskDepth.LargestBidAskSize nLargestSize, nLargestBidSize, nLargestAskSize, dDontCare, dDontCare

    Set aIdx = BidAskTable.CreateSortedIndex(0)

    'walk through grid updating bid/ask sizes with data in table
    'if price is found, else set bid/ask sizes to blanks
    With fgTickDistribution
        'For i = .FixedRows To .Rows - 1

'JM 08-11-2011: optimization test - go from top to bottom row only
        For i = .TopRow To .BottomRow
            If i = m.nLastPriceRow Then
                'current price has volume number as well as price in the cell
                dPriceInGrid = RoundToMinMove(m.Data(eFld_Price, m.nLastPriceRec), m.dMinMove)
                'j = m.oBidAskDepth.SearchForPrice(dPriceInGrid)
                BidAskTable.SearchAsIndex aIdx, 0, dPriceInGrid, j
                If j >= 0 Then
                    dPriceFound = RoundToMinMove(BidAskTable(0, aIdx(j)), m.dMinMove)
                    If dPriceFound <> dPriceInGrid Then
                        j = -1
                    End If
                End If
            ElseIf i - 1 >= 0 And i - 1 < m.Data.NumRecords Then
                dPriceInGrid = RoundToMinMove(m.Data(eFld_Price, i - 1), m.dMinMove)
                BidAskTable.SearchAsIndex aIdx, 0, dPriceInGrid, j
                If j >= 0 Then
                    dPriceFound = RoundToMinMove(BidAskTable(0, aIdx(j)), m.dMinMove)
                    If dPriceFound <> dPriceInGrid Then
                        j = -1
                    End If
                End If
            Else
                j = -1       'theoretically should never get here
            End If
            
            If j >= 0 Then
                nBidSize = BidAskTable(1, aIdx(j))
                nAskSize = BidAskTable(2, aIdx(j))
                If nBidSize > 0 Then
                    If nLargestSize > 0 And Not aBestBids Is Nothing Then
                        If dPriceInGrid = aBestBids(0) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = (nBidSize / nLargestSize * 100) * -1
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nFirstColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestBidSize = nBidSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_BidSize) = Str(nBidSize)
                            .Cell(flexcpForeColor, i, eTDCols_BidSize) = m.nBidTextColor
                            m.dBestBid = aBestBids(0)
                            m.nBestBidSize = nBidSize
                        ElseIf dPriceInGrid = aBestBids(1) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = (nBidSize / nLargestSize * 100) * -1
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nSecondColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestBidSize = nBidSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = -100
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_BidSize) = Str(nBidSize)
                            .Cell(flexcpForeColor, i, eTDCols_BidSize) = m.nBidTextColor
                        ElseIf dPriceInGrid = aBestBids(2) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = (nBidSize / nLargestSize * 100) * -1
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nThirdColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestBidSize = nBidSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = -100
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_BidSize) = Str(nBidSize)
                            .Cell(flexcpForeColor, i, eTDCols_BidSize) = m.nBidTextColor
                        ElseIf dPriceInGrid = aBestBids(3) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = (nBidSize / nLargestSize * 100) * -1
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nFourthColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestBidSize = nBidSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = -100
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_BidSize) = Str(nBidSize)
                            .Cell(flexcpForeColor, i, eTDCols_BidSize) = m.nBidTextColor
                        ElseIf dPriceInGrid = aBestBids(4) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = (nBidSize / nLargestSize * 100) * -1
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nFifthColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestBidSize = nBidSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = -100
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_BidSize) = Str(nBidSize)
                            .Cell(flexcpForeColor, i, eTDCols_BidSize) = m.nBidTextColor
                        ElseIf dPriceInGrid = dPriceFound Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = (nBidSize / nLargestSize * 100) * -1
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nOtherColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestBidSize = nBidSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = -100
                                .Cell(flexcpFloodColor, i, eTDCols_BidSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_BidSize) = Str(nBidSize)
                            .Cell(flexcpForeColor, i, eTDCols_BidSize) = m.nBidTextColor
                        Else
                            .TextMatrix(i, eTDCols_BidSize) = ""
                            .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                        End If
                    End If
                Else
                    .TextMatrix(i, eTDCols_BidSize) = ""
                    .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                End If
                If nAskSize > 0 Then
                    If nLargestSize > 0 And Not aBestAsks Is Nothing Then
                        If dPriceInGrid = aBestAsks(0) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = nAskSize / nLargestSize * 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nFirstColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestAskSize = nAskSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_AskSize) = Str(nAskSize)
                            .Cell(flexcpForeColor, i, eTDCols_AskSize) = m.nAskTextColor
                            m.dBestAsk = aBestAsks(0)
                            m.nBestAskSize = nAskSize
                        ElseIf dPriceInGrid = aBestAsks(1) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = nAskSize / nLargestSize * 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nSecondColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestAskSize = nAskSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_AskSize) = Str(nAskSize)
                            .Cell(flexcpForeColor, i, eTDCols_AskSize) = m.nAskTextColor
                        ElseIf dPriceInGrid = aBestAsks(2) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = nAskSize / nLargestSize * 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nThirdColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestAskSize = nAskSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_AskSize) = Str(nAskSize)
                            .Cell(flexcpForeColor, i, eTDCols_AskSize) = m.nAskTextColor
                        ElseIf dPriceInGrid = aBestAsks(3) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = nAskSize / nLargestSize * 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nFourthColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestAskSize = nAskSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_AskSize) = Str(nAskSize)
                            .Cell(flexcpForeColor, i, eTDCols_AskSize) = m.nAskTextColor
                        ElseIf dPriceInGrid = aBestAsks(4) Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = nAskSize / nLargestSize * 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nFifthColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestAskSize = nAskSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_AskSize) = Str(nAskSize)
                            .Cell(flexcpForeColor, i, eTDCols_AskSize) = m.nAskTextColor
                        ElseIf dPriceInGrid = dPriceFound Then
                            If m.eFloodMktDepth = eBidAskColorMode_ByPrice Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = nAskSize / nLargestSize * 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nOtherColor
                            ElseIf m.eFloodMktDepth = eBidAskColorMode_BySize And nLargestAskSize = nAskSize Then
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 100
                                .Cell(flexcpFloodColor, i, eTDCols_AskSize) = m.nLargestSizeColor
                            Else
                                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                            End If
                            .TextMatrix(i, eTDCols_AskSize) = Str(nAskSize)
                            .Cell(flexcpForeColor, i, eTDCols_AskSize) = m.nAskTextColor
                        Else
                            .TextMatrix(i, eTDCols_AskSize) = ""
                            .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                        End If
                    End If
                Else
                     .TextMatrix(i, eTDCols_AskSize) = ""
                     .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                End If
            Else
                .TextMatrix(i, eTDCols_BidSize) = ""
                .TextMatrix(i, eTDCols_AskSize) = ""
                .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
            End If
        Next
    End With

    m.bLadderHasDOM = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.DisplayMarketDepthOnLadder", eGDRaiseError_Raise

End Sub

Private Sub LoadLadderGrid(Optional ByVal bCenter As Boolean = True)
On Error GoTo ErrSection:

    Dim strPrice$, strPL$, strFgTopPrice$, strErr$
    Dim dPrice#, dVol#
    Dim bUseVol As Boolean, bShowVol As Boolean
    Dim i&, j&, iRec&, nLastPriceColor&
    Dim bOneColor As Boolean
    
    Dim aOutlineIdx As cGdArray
    
    pbBid.Visible = False
    pbAsk.Visible = False
        
    If m.Data.NumRecords < 1 Then
        'this can happen when user changes settings, but there is no intraday data
        Exit Sub
    End If
    
    If m.dLastPrice <= 0 And Not m.bIsSpreadSymbol Then       'aardvark 5198
        strErr = "Error: invalid price " & m.TickBars.PriceDisplay(m.dLastPrice)
        With fgTickDistribution
            .Rows = .FixedRows + 1
            .MergeCells = flexMergeSpill
            .MergeRow(.Rows - 1) = True
            For i = 0 To .Cols - 1
                If Not .ColHidden(i) Then
                    .TextMatrix(.Rows - 1, i) = strErr
                    Exit For
                End If
            Next
        End With
        Exit Sub
    End If
              
    m.nLastPriceRow = -1
    m.nOpenPriceRow = -1
        
    'for now: show bid/ask prices in bid/ask size columns for Forex (06-10-2005)
    With fgTickDistribution
        .MergeCells = flexMergeNever
        If SecurityType(m.TickBars) = "I" And InStr(.TextMatrix(0, eTDCols_BidSize), "Size") > 0 Then
            .TextMatrix(0, eTDCols_BidSize) = "  Bid   "    'pad with space to preserve column size
            .TextMatrix(0, eTDCols_AskSize) = "  Ask   "
        ElseIf InStr(.TextMatrix(0, eTDCols_BidSize), "Size") = 0 Then
            .TextMatrix(0, eTDCols_BidSize) = "Bid Size"
            .TextMatrix(0, eTDCols_AskSize) = "Ask Size"
        End If
    End With
                       
    If m.TickBars(eBARS_Vol, 0) <> kNullData Then
        bUseVol = True          'see note in DisplayPriceLadderRT for additional details
        If Left(m.strSym, 1) <> "$" Then
            If m.nShowVolBar Or m.nShowVolText Then
                bShowVol = Not m.bMinutized         'True
            End If
        End If
    End If
        
    If m.nBarColor = m.nUpColor And m.nBarColor = m.nDownColor Then
        bOneColor = True
        m.bBarColorIsLight = IsLightColor(m.nBarColor)
    End If
    If m.nBarColor = 0 Then m.nBarColor = RGB(1, 1, 1) 'see flexgrid documentation
    
    With fgTickDistribution
        strFgTopPrice = ""
        'save top row value to keep grid in place
        If .Rows > .FixedRows Then
            If .TopRow > 0 And .TopRow < .Rows - 1 Then
                strFgTopPrice = .TextMatrix(.TopRow, eTDCols_Price)
            End If
        End If
        fgTickDistribution.Rows = fgTickDistribution.FixedRows
    End With
    
    If m.tbOutlineCells.NumRecords > 0 Then
        'this happens when user outline cells then turn streaming on or off
        Set aOutlineIdx = m.tbOutlineCells.CreateSortedIndex(0, eGdSort_Default)
    End If
    
    For i = 0 To m.Data.NumRecords - 1
    
        If m.dZeroPLPrice > 0 Then
            strPL = Format(m.Data(eFld_PL, i), "$#,##0;($#,##0)")
        End If
        
        dVol = m.Data(eFld_Volume, i)
        dPrice = RoundToMinMove(m.Data(eFld_Price, i), m.dMinMove)
               
        'appends the volume to the last price
        If bUseVol And dPrice = m.dLastPrice And m.dLastPriceVol > 0 And m.eVolumeStyle = eLadderVol_LastTrade Then
            strPrice = Str(m.dLastPriceVol) & Space(3) & m.Data(eFld_PriceStr, m.nLastPriceRec)
        Else
            strPrice = m.Data(eFld_PriceStr, i)
        End If
        
        With fgTickDistribution
            .AddItem ""
            j = .Rows - 1
            .TextMatrix(j, eTDCols_Price) = strPrice
            If m.nShowVolText And bShowVol And dVol > 0 Then
                If bUseVol Then
                    .TextMatrix(j, eTDCols_Volume) = Format(dVol, "#,###0")
                Else
                    .TextMatrix(j, eTDCols_Volume) = Format(dVol, "#,###0") & " trades"
                End If
            Else
                .TextMatrix(j, eTDCols_Volume) = ""
            End If
            If m.nShowProfitLoss Then
                .TextMatrix(j, eTDCols_PL) = strPL
            End If
            
            If Not aOutlineIdx Is Nothing Then
                If aOutlineIdx.Size > 0 Then
                    If m.tbOutlineCells.SearchAsIndex(aOutlineIdx, 0, dPrice, iRec) Then
                        .Select j, eTDCols_Price
                        .CellBorder m.tbOutlineCells(1, aOutlineIdx(iRec)), 2, 2, 2, 2, 0, 0
                    End If
                End If
            End If
            
            'color volume cells
            If bShowVol Then
                .Cell(flexcpFloodColor, .Rows - 1, eTDCols_Volume) = m.nFloodColor
                If dVol > 0 And m.dMaxVol > 0 And m.nShowVolBar Then
                    .Cell(flexcpFloodPercent, .Rows - 1, eTDCols_Volume) = (dVol / m.dMaxVol * 100) * -1
                Else
                    .Cell(flexcpFloodPercent, .Rows - 1, eTDCols_Volume) = 0
                End If
            End If
            'save row containing open price
            If m.dOpen = dPrice Then
                m.nOpenPriceRow = .Rows - 1
            End If
            If m.dLastPrice = dPrice Then
                m.nLastPriceRow = .Rows - 1     'save row containing last price
            ElseIf dPrice > m.dHigh Or dPrice < m.dLow Then
                'out of range rows
                .Cell(flexcpBackColor, .Rows - 1, eTDCols_Price) = .BackColor
                If .BackColor = kDarkThemeColor Then
                    .Cell(flexcpForeColor, .Rows - 1, eTDCols_Price) = vbWhite
                Else
                    .Cell(flexcpForeColor, .Rows - 1, eTDCols_Price) = .ForeColor
                End If
            Else
                'color all rows between high & low prices with bar color
                'rows between open & close prices are painted over with up/down color below
                If m.nBarColor = vbWhite Or (bOneColor And m.bBarColorIsLight) Then
                    .Cell(flexcpForeColor, .Rows - 1, eTDCols_Price) = vbBlack
                Else
                    .Cell(flexcpForeColor, .Rows - 1, eTDCols_Price) = vbWhite
                End If
                .Cell(flexcpBackColor, .Rows - 1, eTDCols_Price) = m.nBarColor
            End If
            
            'set index into table
            .RowData(.Rows - 1) = i
        End With
    Next
    
    If m.nOpenPriceRow < 0 Or m.nLastPriceRow < 0 Then
        DebugLog "Invalid price row: OpenPriceRow = " & Str(m.nOpenPriceRow) & ", Open = " & Str(m.dOpen) & ", LastPriceRow = " & Str(m.nLastPriceRow) & ", LastPrice = " & Str(m.dLastPrice)
    End If
                   
    With fgTickDistribution
        .Redraw = flexRDNone
        If .Rows > .FixedRows Then
            .Cell(flexcpFontBold, .FixedRows, eTDCols_Price, .Rows - 1) = True
        End If
        
        ' color rows between open and close price with up or down color
        If m.dLastPrice > m.dOpen Then
            m.nCurrUpDownColor = m.nUpColor
        ElseIf m.dLastPrice < m.dOpen Then
            m.nCurrUpDownColor = m.nDownColor
        Else
            m.nCurrUpDownColor = m.nBarColor
        End If
        If m.nCurrUpDownColor = 0 Then m.nCurrUpDownColor = RGB(1, 1, 1) 'see flexgrid documentation
        If m.nLastPriceRow >= .FixedRows And m.nLastPriceRow < .Rows Then
            If m.nOpenPriceRow >= .FixedRows And m.nOpenPriceRow < .Rows Then
                If m.nCurrUpDownColor = vbWhite Or (bOneColor And IsLightColor(m.nCurrUpDownColor)) Then
                    .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = vbBlack
                Else
                    .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = vbWhite
                End If
                .Cell(flexcpBackColor, m.nLastPriceRow, eTDCols_Price, m.nOpenPriceRow, eTDCols_Price) = m.nCurrUpDownColor
            End If
            
            ' invert last price cell
            nLastPriceColor = GetLastPriceTextColor
            If nLastPriceColor = vbWhite Then
                .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price) = RGB(1, 1, 1)
            ElseIf bOneColor Then
                'single color background
                If nLastPriceColor = vbBlack Then nLastPriceColor = RGB(1, 1, 1)
                .Cell(flexcpBackColor, m.nLastPriceRow, eTDCols_Price) = nLastPriceColor
                .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price) = .BackColor
            Else
                .Cell(flexcpFloodColor, m.nLastPriceRow, eTDCols_Price) = .BackColor
            End If
            .Cell(flexcpForeColor, m.nLastPriceRow, eTDCols_Price) = GetLastPriceTextColor
            .Cell(flexcpFloodPercent, m.nLastPriceRow, eTDCols_Price) = 100
        End If
        
        ' color bid/ask columns
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, eTDCols_BidSize, .Rows - 1, eTDCols_BidSize) = m.nBidColor
            .Cell(flexcpBackColor, .FixedRows, eTDCols_AskSize, .Rows - 1, eTDCols_AskSize) = m.nAskColor
        End If
        .Row = -1
        .Redraw = flexRDBuffered
        '.TopRow = (.Rows - .BottomRow) / 2 + 1      'top row is always 1 on initial show
        If m.nLastPriceRow >= .FixedRows And m.nLastPriceRow < .Rows Then
            .ShowCell m.nLastPriceRow, eTDCols_Price
            'flag this cell's width to fix scenario where price spills into next column
            If .CellWidth > .ColWidth(eTDCols_Price) Then .ColWidth(eTDCols_Price) = .CellWidth
        End If
        
        If bCenter Then
            m.dScrollTickCount = 0
            CenterLadderOnCurrPrice
        Else
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, eTDCols_Price) = strFgTopPrice Then
                    .TopRow = i
                    m.dScrollTickCount = 0          'to distinguish program scroll from user's scrol
                    Exit For
                End If
            Next
            .Redraw = flexRDBuffered
        End If
    End With
        
    DisplayMarketDepthOnLadder
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.LoadLadderGrid", eGDRaiseError_Raise

End Sub

Private Sub LoadDetailGrid()
On Error GoTo ErrSection:

    Dim i&, j&
    Dim aIdxBid As cGdArray
    Dim aIdxAsk As cGdArray
    
    Dim BidTable As cGdTable
    Dim AskTable As cGdTable
    
    Dim nBidMakerCol&, nBidPriceCol&, nBidSizeCol&, nBidTimeCol&
    Dim nAskMakerCol&, nAskPriceCol&, nAskSizeCol&, nAskTimeCol&
    
    Dim dDateTime#
    
    If m.oBidAskDepth Is Nothing Then Exit Sub
    
    nBidMakerCol = HeaderToColIndex(m.aBidColHeader, "Maker")
    nBidPriceCol = HeaderToColIndex(m.aBidColHeader, "Bid")
    nBidSizeCol = HeaderToColIndex(m.aBidColHeader, "Size")
    nBidTimeCol = HeaderToColIndex(m.aBidColHeader, "Time")
    
    nAskMakerCol = HeaderToColIndex(m.aAskColHeader, "Maker")
    nAskPriceCol = HeaderToColIndex(m.aAskColHeader, "Ask")
    nAskSizeCol = HeaderToColIndex(m.aAskColHeader, "Size")
    nAskTimeCol = HeaderToColIndex(m.aAskColHeader, "Time")
    
    'holds data to pass to graphics engine
    Dim aColors As New cGdArray
    Dim aBidSizes As New cGdArray
    Dim aAskSizes As New cGdArray
    Dim nTop&, nLeft&, nBottom&, nRight&
    Dim nSumBidSizes&, nSumAskSizes&, nSum&
    Dim nIdx&, nFlag&
    
    Dim dPrevPrice#, nLevel&, nColor&
    Dim nBidSizeMax&, nAskSizeMax&, nBidSize&, nAskSize&
    Dim bNoBidData As Boolean, bNoAskData As Boolean
    Dim bDataIsValid As Boolean, bActive As Boolean
    Dim strFlag$
                       
    Set BidTable = m.oBidAskDepth.BidData
    Set AskTable = m.oBidAskDepth.AskData
    
    'sort by size to get largest bid and ask sizes
    nBidSizeMax = 0
    If BidTable Is Nothing Then
        bNoBidData = True
    ElseIf BidTable.NumRecords > 0 Then
        Set aIdxBid = BidTable.CreateSortedIndex(2, eGdSort_Descending)
        For i = 0 To aIdxBid.Size - 1
            bDataIsValid = m.oBidAskDepth.AllDataValid(0, aIdxBid(i), bActive, strFlag)
            If bDataIsValid And bActive Then
                nBidSizeMax = BidTable(2, aIdxBid(i))
                Exit For
            End If
        Next
    Else
        bNoBidData = True
    End If
                
    nAskSizeMax = 0
    If AskTable Is Nothing Then
        bNoAskData = True
    ElseIf AskTable.NumRecords > 0 Then
        Set aIdxAsk = AskTable.CreateSortedIndex(2, eGdSort_Descending)
        For i = 0 To aIdxAsk.Size - 1
            bDataIsValid = m.oBidAskDepth.AllDataValid(1, aIdxAsk(i), bActive, strFlag)
            If bDataIsValid And bActive Then
                nAskSizeMax = AskTable(2, aIdxAsk(i))
                Exit For
            End If
        Next
    Else
        bNoAskData = True
    End If
        
    Set aIdxBid = Nothing
    Set aIdxAsk = Nothing
    
    If bNoBidData And bNoAskData Then
        With fgBidDetail
            .Redraw = flexRDNone
            .Rows = 1
            .Redraw = flexRDBuffered
        End With
        With fgAskDetail
            .Redraw = flexRDNone
            .Rows = 1
            .Redraw = flexRDBuffered
        End With
        With fgDOMPrint
            .Redraw = flexRDNone
            .Rows = 1
            .Redraw = flexRDBuffered
        End With
        Exit Sub
    End If
    
    'sort by active flag, price, size, time
    Set aIdxBid = BidTable.CreateSortedIndex(4, eGdSort_Descending, 1, eGdSort_Descending, 2, eGdSort_Descending, 3, eGdSort_Descending)
    
    If m.nShowSummaryBar = 1 Then
        If Not pbBid.Visible Then
            pbBid.Visible = True
            pbAsk.Visible = True
        End If
    Else
        If pbBid.Visible Then
            pbBid.Visible = False
            pbAsk.Visible = False
        End If
    End If
        
    i = aColors.Create(eGDARRAY_Longs, 6, 0)
    i = aBidSizes.Create(eGDARRAY_Longs, 6, 0)
    i = aAskSizes.Create(eGDARRAY_Longs, 6, 0)
    
    aColors(0) = m.nFirstColor
    aColors(1) = m.nSecondColor
    aColors(2) = m.nThirdColor
    aColors(3) = m.nFourthColor
    aColors(4) = m.nFifthColor
    aColors(5) = m.nOtherColor
        
    nSumBidSizes = 0
    nSumAskSizes = 0
    nSum = 0
    
    'header tittles for print grid (in case user moved things around)
    With fgDOMPrint
        .TextMatrix(0, 0) = fgBidDetail.TextMatrix(0, 0)
        .TextMatrix(0, 1) = fgBidDetail.TextMatrix(0, 1)
        .TextMatrix(0, 2) = fgBidDetail.TextMatrix(0, 2)
        .TextMatrix(0, 3) = fgBidDetail.TextMatrix(0, 3)
        
        .TextMatrix(0, 4) = fgAskDetail.TextMatrix(0, 0)
        .TextMatrix(0, 5) = fgAskDetail.TextMatrix(0, 1)
        .TextMatrix(0, 6) = fgAskDetail.TextMatrix(0, 2)
        .TextMatrix(0, 7) = fgAskDetail.TextMatrix(0, 3)
    End With
    
    With fgBidDetail
        .Redraw = flexRDNone
        .Rows = .FixedRows
        fgDOMPrint.Redraw = flexRDNone
        fgDOMPrint.Rows = .FixedRows
        For i = 0 To aIdxBid.Size - 1
            j = aIdxBid(i)
            bDataIsValid = m.oBidAskDepth.AllDataValid(0, j, bActive, strFlag)
            If bDataIsValid Then
                .Rows = .Rows + 1
                fgDOMPrint.Rows = .Rows
                nBidSize = BidTable(2, j)
                If bActive Then nSumBidSizes = nSumBidSizes + nBidSize
                If nBidMakerCol >= 0 And nBidMakerCol < 4 Then
                    .TextMatrix(.Rows - 1, nBidMakerCol) = BidTable(0, j)
                End If
                If nBidPriceCol >= 0 And nBidPriceCol < 4 Then
                    .TextMatrix(.Rows - 1, nBidPriceCol) = m.TickBars.PriceDisplay(BidTable(1, j))
                End If
                If nBidSizeCol >= 0 And nBidSizeCol < 4 Then
                    .TextMatrix(.Rows - 1, nBidSizeCol) = Str(nBidSize) & BidTable(5, j)
                End If
                If nBidTimeCol >= 0 And nBidTimeCol < 4 Then
                    dDateTime = BidTable(3, j)
                    If g.bShowInLocalTimeZone Then
                        dDateTime = ConvertTimeZone(dDateTime, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
                    End If
                    .TextMatrix(.Rows - 1, nBidTimeCol) = DateFormat(dDateTime, NO_DATE, HH_MM_SS)
                End If
                'save pluses or minuses to hidden column
                If BidTable(5, j) = "+" Then
                    .TextMatrix(.Rows - 1, eCol_BidAskChange) = 1
                ElseIf BidTable(5, j) = "-" Then
                    .TextMatrix(.Rows - 1, eCol_BidAskChange) = -1
                Else
                    .TextMatrix(.Rows - 1, eCol_BidAskChange) = 0
                End If
                'flood cell if not drawing triangles
                If m.nDrawTriangle = 1 Then
                    If nBidSizeCol >= 0 And nBidSizeCol < 4 Then
                        .Cell(flexcpFloodPercent, .Rows - 1, nBidSizeCol) = 0
                        fgDOMPrint.Cell(flexcpFloodPercent, .Rows - 1, nBidSizeCol) = 0
                    End If
                ElseIf nBidSizeMax > 0 Then
                    If nBidSizeCol >= 0 And nBidSizeCol < 4 Then
                        .Cell(flexcpFloodPercent, .Rows - 1, nBidSizeCol) = (nBidSize / nBidSizeMax * 100) * -1
                        .Cell(flexcpFloodColor, .Rows - 1, nBidSizeCol) = m.nBidAskUpColor
                        
                        fgDOMPrint.Cell(flexcpFloodPercent, .Rows - 1, nBidSizeCol) = .Cell(flexcpFloodPercent, .Rows - 1, nBidSizeCol)
                        fgDOMPrint.Cell(flexcpFloodColor, .Rows - 1, nBidSizeCol) = m.nBidAskUpColor
                    End If
                End If
                'set color for different price levels
                If Not bActive Then
                    nColor = m.nInactiveColor
                ElseIf dPrevPrice = 0 Then
                    nLevel = 1
                    nColor = m.nFirstColor
                    aBidSizes(0) = aBidSizes(0) + nBidSize
                    nIdx = 0
                ElseIf dPrevPrice <> BidTable(1, j) Then
                    nLevel = nLevel + 1
                    Select Case nLevel
                        Case 2:
                            nColor = m.nSecondColor
                            aBidSizes(1) = aBidSizes(1) + nBidSize
                            nIdx = 1
                        Case 3:
                            nColor = m.nThirdColor
                            aBidSizes(2) = aBidSizes(2) + nBidSize
                            nIdx = 2
                        Case 4:
                            nColor = m.nFourthColor
                            aBidSizes(3) = aBidSizes(3) + nBidSize
                            nIdx = 3
                        Case 5:
                            nColor = m.nFifthColor
                            aBidSizes(4) = aBidSizes(4) + nBidSize
                            nIdx = 4
                        Case Else
                            nColor = m.nOtherColor
                            aBidSizes(5) = aBidSizes(5) + nBidSize
                            nIdx = 5
                    End Select
                ElseIf dPrevPrice = BidTable(1, j) Then
                    aBidSizes(nIdx) = aBidSizes(nIdx) + nBidSize
                End If
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 3) = nColor
                dPrevPrice = BidTable(1, j)
                'place data into print grid
                fgDOMPrint.TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 1, 0)
                fgDOMPrint.TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
                fgDOMPrint.TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2)
                fgDOMPrint.TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3)
                fgDOMPrint.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 3) = nColor
            End If
        Next
        .Redraw = flexRDBuffered
        fgDOMPrint.Redraw = flexRDBuffered
    End With
                
    Set aIdxAsk = AskTable.CreateSortedIndex(4, eGdSort_Descending, 1, eGdSort_Default, 2, eGdSort_Descending, 3, eGdSort_Descending)
    
    dPrevPrice = 0
    nLevel = 0
    nColor = vbYellow
    With fgAskDetail
        .Redraw = flexRDNone
        .Rows = .FixedRows
        fgDOMPrint.Redraw = flexRDNone
        For i = 0 To aIdxAsk.Size - 1
            j = aIdxAsk(i)
            bDataIsValid = m.oBidAskDepth.AllDataValid(1, j, bActive, strFlag)
            'If AskTable(1, j) > 0 And AskTable(2, j) > 0 And AskTable(3, j) > 0 Then
            If bDataIsValid Then
                .Rows = .Rows + 1
                nAskSize = AskTable(2, j)
                If fgDOMPrint.Rows < .Rows Then fgDOMPrint.Rows = .Rows
                If bActive Then nSumAskSizes = nSumAskSizes + nAskSize
                If nAskMakerCol >= 0 And nAskMakerCol < 4 Then
                    .TextMatrix(.Rows - 1, nAskMakerCol) = AskTable(0, j)
                End If
                If nAskPriceCol >= 0 And nAskPriceCol < 4 Then
                    .TextMatrix(.Rows - 1, nAskPriceCol) = m.TickBars.PriceDisplay(AskTable(1, j))
                End If
                If nAskSizeCol >= 0 And nAskSizeCol < 4 Then
                    .TextMatrix(.Rows - 1, nAskSizeCol) = Str(nAskSize) & AskTable(5, j)
                End If
                If nAskTimeCol >= 0 And nAskTimeCol < 4 Then
                    dDateTime = AskTable(3, j)
                    If g.bShowInLocalTimeZone Then
                        dDateTime = ConvertTimeZone(dDateTime, m.TickBars.Prop(eBARS_ExchangeTimeZoneInf), "")
                    End If
                    .TextMatrix(.Rows - 1, nAskTimeCol) = DateFormat(dDateTime, NO_DATE, HH_MM_SS)
                End If
                'save pluses & minuses to hidden column
                If AskTable(5, j) = "+" Then
                    .TextMatrix(.Rows - 1, eCol_BidAskChange) = 1
                ElseIf AskTable(5, j) = "-" Then
                    .TextMatrix(.Rows - 1, eCol_BidAskChange) = -1
                Else
                    .TextMatrix(.Rows - 1, eCol_BidAskChange) = 0
                End If
                'flood cell if not drawing triangles
                If m.nDrawTriangle = 1 Then
                    If nAskSizeCol >= 0 And nAskSizeCol < 4 Then
                        .Cell(flexcpFloodPercent, .Rows - 1, nAskSizeCol) = 0
                        fgDOMPrint.Cell(flexcpFloodPercent, .Rows - 1, nAskSizeCol + 4) = 0
                    End If
                ElseIf nAskSizeMax > 0 Then
                    If nAskSizeCol >= 0 And nAskSizeCol < 4 Then
                        .Cell(flexcpFloodPercent, .Rows - 1, nAskSizeCol) = (nAskSize / nAskSizeMax * 100)
                        .Cell(flexcpFloodColor, .Rows - 1, nAskSizeCol) = m.nBidAskDownColor
                        
                        fgDOMPrint.Cell(flexcpFloodPercent, .Rows - 1, nAskSizeCol + 4) = .Cell(flexcpFloodPercent, .Rows - 1, nAskSizeCol)
                        fgDOMPrint.Cell(flexcpFloodColor, .Rows - 1, nAskSizeCol + 4) = m.nBidAskDownColor
                    End If
                End If
                'set color for different price levels
                If Not bActive Then
                    nColor = m.nInactiveColor
                ElseIf dPrevPrice = 0 Then
                    nLevel = 1
                    nColor = m.nFirstColor
                    aAskSizes(0) = aAskSizes(0) + nAskSize
                    nIdx = 0
                ElseIf dPrevPrice <> AskTable(1, j) Then
                    nLevel = nLevel + 1
                    Select Case nLevel
                        Case 2:
                            nColor = m.nSecondColor
                            aAskSizes(1) = aAskSizes(1) + nAskSize
                            nIdx = 1
                        Case 3:
                            nColor = m.nThirdColor
                            aAskSizes(2) = aAskSizes(2) + nAskSize
                            nIdx = 2
                        Case 4:
                            nColor = m.nFourthColor
                            aAskSizes(3) = aAskSizes(3) + nAskSize
                            nIdx = 3
                        Case 5:
                            nColor = m.nFifthColor
                            aAskSizes(4) = aAskSizes(4) + nAskSize
                            nIdx = 4
                        Case Else
                            nColor = m.nOtherColor
                            aAskSizes(5) = aAskSizes(5) + nAskSize
                            nIdx = 5
                    End Select
                ElseIf dPrevPrice = AskTable(1, j) Then
                    aAskSizes(nIdx) = aAskSizes(nIdx) + nAskSize
                End If
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 3) = nColor
                dPrevPrice = AskTable(1, j)
                'place data into print grid
                fgDOMPrint.TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 0)
                fgDOMPrint.TextMatrix(.Rows - 1, 5) = .TextMatrix(.Rows - 1, 1)
                fgDOMPrint.TextMatrix(.Rows - 1, 6) = .TextMatrix(.Rows - 1, 2)
                fgDOMPrint.TextMatrix(.Rows - 1, 7) = .TextMatrix(.Rows - 1, 3)
                fgDOMPrint.Cell(flexcpBackColor, .Rows - 1, 4, .Rows - 1, 7) = nColor
            End If
        Next
        .Redraw = flexRDBuffered
        fgDOMPrint.Redraw = flexRDBuffered
    End With
    
    'for horizontal stacked bar use larger of ask or bid sum of sizes
    'for vertical bar use sum of sizes at individual price level that is largest
    If m.nVertSummaryBar = 1 Then
        nSum = 0
        For i = 0 To 5
            If aBidSizes(i) > nSum Then nSum = aBidSizes(i)
            If aAskSizes(i) > nSum Then nSum = aAskSizes(i)
        Next
    Else
        If nSumBidSizes > nSumAskSizes Then
            nSum = nSumBidSizes
        Else
            nSum = nSumAskSizes
        End If
    End If
        
    'Justify Flag (bounding rectangle is top,left,bottom,right):
    '   1=vert bars starting on right side of bounding rectangle going left
    '   2=vert bars starting on left side of bounding rectangle going right
    '   3=horz bars starting on right side of bounding rectangle going left
    '   4=horz bars starting on left side of bounding rectangle going right
    If m.nShowSummaryBar = 1 And nSum > 0 And fgBidDetail.Rows > fgBidDetail.FixedRows Then
        nTop = pbBid.Top / Screen.TwipsPerPixelY
        nBottom = (pbBid.Top + pbBid.Height) / Screen.TwipsPerPixelY
        'draw summary histogram on bid side
        nLeft = pbBid.Left / Screen.TwipsPerPixelX
        nRight = pbBid.Width / Screen.TwipsPerPixelX
        If m.nVertSummaryBar = 1 Then
            nFlag = 2
        Else
            nFlag = 3
        End If
        pbBid.Cls
        i = geDrawTickHistogram(m.geTickObj, pbBid.hWnd, pbBid.hDC, aColors.ArrayHandle, aBidSizes.ArrayHandle, nSum, pbBid.BackColor, nTop, nLeft, nBottom, nRight, nFlag)

        'draw summary histogram on ask size
        nLeft = pbAsk.Left / Screen.TwipsPerPixelX
        nRight = (pbAsk.Left + pbAsk.Width) / Screen.TwipsPerPixelX
        If m.nVertSummaryBar = 1 Then
            nFlag = 1
        Else
            nFlag = 4
        End If
        pbAsk.Cls
        i = geDrawTickHistogram(m.geTickObj, pbAsk.hWnd, pbAsk.hDC, aColors.ArrayHandle, aAskSizes.ArrayHandle, nSum, pbAsk.BackColor, nTop, nLeft, nBottom, nRight, nFlag)
    End If
    
    aIdxBid.Destroy
    aIdxAsk.Destroy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.LoadDetailGrid", eGDRaiseError_Raise

End Sub

Public Property Get DisplayStyle() As Long
On Error GoTo ErrSection:

    DisplayStyle = m.eDisplayStyle

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.DisplayStyleGet"

End Property

Public Property Let DisplayStyle(ByVal nStyle&)
On Error GoTo ErrSection:

    m.eDisplayStyle = nStyle

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.DisplayStyleLet"

End Property

Public Property Get FirstColor() As Long
On Error GoTo ErrSection:

    FirstColor = m.nFirstColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FirstColorGet"

End Property

Public Property Get SecondColor() As Long
On Error GoTo ErrSection:

    SecondColor = m.nSecondColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SecondColorGet"

End Property

Public Property Get ThirdColor() As Long
On Error GoTo ErrSection:

    ThirdColor = m.nThirdColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ThirdColorGet"

End Property

Public Property Get FourthColor() As Long
On Error GoTo ErrSection:

    FourthColor = m.nFourthColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FourthColorGet"

End Property

Public Property Get FifthColor() As Long
On Error GoTo ErrSection:

    FifthColor = m.nFifthColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FifthColorGet"

End Property

Public Property Get OtherColor() As Long
On Error GoTo ErrSection:

    OtherColor = m.nOtherColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.OtherColorGet"

End Property

Public Property Get InactiveColor() As Long
On Error GoTo ErrSection:

    InactiveColor = m.nInactiveColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.InactiveColorGet"

End Property

Public Property Get LargestSizeColor() As Long
On Error GoTo ErrSection:

    LargestSizeColor = m.nLargestSizeColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.LargestSizeColorGet"

End Property

Public Property Get BidAskUpColor() As Long
On Error GoTo ErrSection:

    BidAskUpColor = m.nBidAskUpColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidAskUpColorGet"

End Property

Public Property Get BidAskDownColor() As Long
On Error GoTo ErrSection:

    BidAskDownColor = m.nBidAskDownColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidAskDownColorGet"

End Property

Public Property Get DrawTriangles() As Long
On Error GoTo ErrSection:

    DrawTriangles = m.nDrawTriangle

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.DrawTrianglesGet"

End Property

Public Property Get ShowSummaryBar() As Long
On Error GoTo ErrSection:

    ShowSummaryBar = m.nShowSummaryBar

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowSummaryBarGet"

End Property

Public Property Get VerticalSummaryBar() As Long
On Error GoTo ErrSection:

    VerticalSummaryBar = m.nVertSummaryBar

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.VerticalSummaryBarGet"

End Property

Public Property Get SummaryBarHeight() As Long
On Error GoTo ErrSection:

    SummaryBarHeight = pbBid.Height

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SummaryBarHeightGet"

End Property

Public Property Let FirstColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nFirstColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FirstColorLet"

End Property

Public Property Let SecondColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nSecondColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SecondColorLet"

End Property

Public Property Let ThirdColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nThirdColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ThirdColorLet"

End Property

Public Property Let FourthColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nFourthColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FourthColorLet"

End Property

Public Property Let FifthColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nFifthColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.FifthColorLet"

End Property

Public Property Let OtherColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nOtherColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.OtherColorLet"

End Property

Public Property Let InactiveColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nInactiveColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.InactiveColorLet"

End Property

Public Property Let LargestSizeColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nLargestSizeColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.LargestSizeColorLet"

End Property

Public Property Let BidAskUpColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nBidAskUpColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidAskUpColorLet"

End Property

Public Property Let BidAskDownColor(ByVal nColor&)
On Error GoTo ErrSection:

    m.nBidAskDownColor = nColor

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidAskDownColorLet"

End Property

Public Property Let DrawTriangles(ByVal nDraw&)
On Error GoTo ErrSection:

    m.nDrawTriangle = nDraw

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.DrawTrianglesLet"

End Property

Public Property Let ShowSummaryBar(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowSummaryBar = nShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowSummaryBarLet"

End Property

Public Property Let VerticalSummaryBar(ByVal nVertical&)
On Error GoTo ErrSection:

    m.nVertSummaryBar = nVertical

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.VerticalSummaryBarLet"

End Property

Private Function HeaderToColIndex(aHeader As cGdArray, ByVal strHeader) As Long
On Error GoTo ErrSection:

    Dim i&, nCol&
    
    nCol = -1
    If Not aHeader Is Nothing Then
        For i = 0 To aHeader.Size - 1
            If aHeader(i) = strHeader Then
                nCol = i
                Exit For
            End If
        Next
    End If
    
    HeaderToColIndex = nCol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.HeaderToColIndex", eGDRaiseError_Raise

End Function

Private Property Get BidHeader() As String
On Error GoTo ErrSection:

    BidHeader = m.aBidColHeader(0) & "," & m.aBidColHeader(1) & "," & m.aBidColHeader(2) & "," & m.aBidColHeader(3)

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.BidHeaderGet"

End Property

Private Property Get AskHeader() As String
On Error GoTo ErrSection:

    AskHeader = m.aAskColHeader(0) & "," & m.aAskColHeader(1) & "," & m.aAskColHeader(2) & "," & m.aAskColHeader(3)

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.AskHeaderGet"

End Property

Private Property Let BidHeader(ByVal strText$)
On Error GoTo ErrSection:

    Dim aNewHeader As New cGdArray
    
    aNewHeader.SplitFields strText
    
    If aNewHeader.Size = 4 Then
        If InStr(strText, "Maker") > 0 And _
           InStr(strText, "Bid") > 0 And _
           InStr(strText, "Size") > 0 And _
           InStr(strText, "Time") > 0 Then
           
           CheckHeaderChange m.aBidColHeader, aNewHeader, 4
           
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.BidHeader.Let", eGDRaiseError_Raise

End Property

Private Property Let AskHeader(ByVal strText$)
On Error GoTo ErrSection:

    Dim aNewHeader As New cGdArray
    
    aNewHeader.SplitFields strText
    
    If aNewHeader.Size = 4 Then
        If InStr(strText, "Maker") > 0 And _
           InStr(strText, "Ask") > 0 And _
           InStr(strText, "Size") > 0 And _
           InStr(strText, "Time") > 0 Then
           
           CheckHeaderChange m.aAskColHeader, aNewHeader, 4
           
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.AskHeader.Let", eGDRaiseError_Raise

End Property

Private Sub CheckHeaderChange(aOldHeader As cGdArray, aNewHeader As cGdArray, ByVal nSize&)
On Error GoTo ErrSection:

    Dim i&
        
    If aOldHeader.Size = nSize And aNewHeader.Size = nSize Then
        For i = 0 To nSize - 1
            If aOldHeader(i) <> aNewHeader(i) Then
                aOldHeader(i) = aNewHeader(i)
                m.bHeaderChanged = True
            End If
        Next
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.CheckHeaderChange", eGDRaiseError_Raise

End Sub

Public Property Get ShowQuoteBar() As Long
On Error GoTo ErrSection:

    ShowQuoteBar = m.nShowQuoteBar

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowQuoteBarGet"

End Property

Public Property Let ShowQuoteBar(ByVal nShow&)
On Error GoTo ErrSection:

    m.nShowQuoteBar = nShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowQuoteBarLet"

End Property

Public Property Get ShowOrderBar() As eGDOrderBarMode
On Error GoTo ErrSection:

    ShowOrderBar = m.eOrderBarMode

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowOrderBarGet"

End Property

Public Property Let ShowOrderBar(ByVal eShow As eGDOrderBarMode)
On Error GoTo ErrSection:

    m.eOrderBarMode = eShow

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ShowOrderBarLet"

End Property

Public Property Get QBarColArray() As cGdArray
On Error GoTo ErrSection:

    Set QBarColArray = m.aQBarColHeader

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.QBarColArrayGet"

End Property

Public Sub QuoteBarHeader(ByVal strText$)
On Error GoTo ErrSection:

    GridBarHeader fgQuoteBar, m.aQBarColHeader, m.nQBarSumColWidth, strText
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.QuoteBarHeader", eGDRaiseError_Raise

End Sub

Public Property Get ABarColArray() As cGdArray
On Error GoTo ErrSection:

    Set ABarColArray = m.aABarColHeader

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.ABarColArray"

End Property

Public Sub AccountBarHeader(ByVal strText$)
On Error GoTo ErrSection:

    GridBarHeader fgAccountBar, m.aABarColHeader, m.nABarSumColWidth, strText
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.AccountBarHeader", eGDRaiseError_Raise

End Sub

Private Sub DisplayQBar()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    Static dTickCount#, strSym$
    Static dPrevTrade#, nPrevTradeSize&
    
    Dim strText$, dTrade#, nTradeSize&, i&
        
    If bInProgress Then Exit Sub
    
    bInProgress = True
        
    If m.eDisplayStyle = eView_Ladder And m.oBidAskDepth Is Nothing Then
        GetDailyBarsBidAsk
    ElseIf m.eDisplayStyle = eView_Detail Then
        With fgBidDetail
            If .Rows > .FixedRows Then
                i = HeaderToColIndex(m.aBidColHeader, "Size")
                If i > 0 And i < .Cols Then
                    strText = .TextMatrix(.FixedRows, i)
                    If InStr(strText, "+") <> 0 Or InStr(strText, "-") <> 0 Then
                        strText = Left(strText, Len(strText) - 1)
                    End If
                    m.nBestBidSize = ValOfText(strText)
                End If
                If m.dBestBid <> kNullData Then
                    i = HeaderToColIndex(m.aBidColHeader, "Bid")
                    If i > 0 And i < .Cols Then
                        strText = .TextMatrix(.FixedRows, i)
                        m.dBestBid = m.TickBars.PriceFromString(strText)
                    End If
                End If
            End If
        End With
        With fgAskDetail
            If .Rows > .FixedRows Then
                i = HeaderToColIndex(m.aAskColHeader, "Size")
                If i > 0 And i < .Cols Then
                    strText = .TextMatrix(.FixedRows, i)
                    If InStr(strText, "+") <> 0 Or InStr(strText, "-") <> 0 Then
                        strText = Left(strText, Len(strText) - 1)
                    End If
                    m.nBestAskSize = ValOfText(strText)
                End If
                If m.dBestAsk <> kNullData Then
                    i = HeaderToColIndex(m.aAskColHeader, "Ask")
                    If i > 0 And i < .Cols Then
                        strText = .TextMatrix(.FixedRows, i)
                        m.dBestAsk = m.TickBars.PriceFromString(strText)
                    End If
                End If
            End If
        End With
    End If
        
    If strSym <> m.strSym Then
        dPrevTrade = 0
        nPrevTradeSize = 0
        strSym = m.strSym
    End If
    
    With fgQuoteBar
        .Redraw = flexRDNone
        dTickCount = gdTickCount - dTickCount
        If dTickCount > g.nUpdatedColorDuration Then
            i = HeaderToColIndex(m.aQBarColHeader, "Trade")
            If i > .FixedRows Then
                .Cell(flexcpForeColor, 1, i, 1, i) = vbBlack
                .Cell(flexcpFontBold, 1, i, 1, i) = False
            End If
            i = HeaderToColIndex(m.aQBarColHeader, "Trade Size")
            If i > .FixedRows Then
                .Cell(flexcpForeColor, 1, i, 1, i) = vbBlack
                .Cell(flexcpFontBold, 1, i, 1, i) = False
            End If
        End If
        For i = 0 To .Cols - 1
            Select Case .TextMatrix(0, i)
                Case "Symbol"
                    .TextMatrix(1, i) = m.strSym
                Case "Bid"
                    If m.dBestBid <> kNullData Then
                        .TextMatrix(1, i) = m.TickBars.PriceDisplay(m.dBestBid)
                    Else
                        .TextMatrix(1, i) = ""
                    End If
                Case "Bid Size"
                    If m.nBestBidSize > 0 Then
                        .TextMatrix(1, i) = Str(m.nBestBidSize)
                    Else
                        .TextMatrix(1, i) = ""
                    End If
                Case "Ask"
                    If m.dBestAsk <> kNullData Then
                        .TextMatrix(1, i) = m.TickBars.PriceDisplay(m.dBestAsk)
                    Else
                        .TextMatrix(1, i) = ""
                    End If
                Case "Ask Size"
                    If m.nBestAskSize > 0 Then
                        .TextMatrix(1, i) = Str(m.nBestAskSize)
                    Else
                        .TextMatrix(1, i) = ""
                    End If
                Case "Trade"
                    dTrade = m.TickBars(eBARS_Close, m.TickBars.Size - 1)
                    .TextMatrix(1, i) = m.TickBars.PriceDisplay(dTrade)
                    If dTrade = dPrevTrade Or (dTrade <= 0 And Not m.bIsSpreadSymbol) Then
                        'do nothing
                    ElseIf dTrade > dPrevTrade Then
                        .Cell(flexcpForeColor, 1, i, 1, i) = m.nBidAskUpColor
                        .Cell(flexcpFontBold, 1, i, 1, i) = True
                    ElseIf dTrade < dPrevTrade Then
                        .Cell(flexcpForeColor, 1, i, 1, i) = m.nBidAskDownColor
                        .Cell(flexcpFontBold, 1, i, 1, i) = True
                    End If
                    If dTrade > 0 Then dPrevTrade = dTrade
                Case "Trade Size"
                    nTradeSize = m.TickBars(eBARS_Vol, m.TickBars.Size - 1)
'original code ( we used to do volume for only stocks ) - leave awhile then remove 01-05-2005
'If SecurityType(m.TickBars) = "S" Then
                    If nTradeSize > 0 Then
                        .TextMatrix(1, i) = Format(nTradeSize, "#,##0")
                        If nTradeSize = nPrevTradeSize Or nPrevTradeSize = 0 Then
                            'do nothing
                        ElseIf nTradeSize > nPrevTradeSize Then
                            .Cell(flexcpForeColor, 1, i, 1, i) = m.nBidAskUpColor
                            .Cell(flexcpFontBold, 1, i, 1, i) = True
                        ElseIf nTradeSize < nPrevTradeSize Then
                            .Cell(flexcpForeColor, 1, i, 1, i) = m.nBidAskDownColor
                            .Cell(flexcpFontBold, 1, i, 1, i) = True
                        End If
                        nPrevTradeSize = nTradeSize
                    Else
                        .TextMatrix(1, i) = ""
                    End If
                Case "Open"
                    .TextMatrix(1, i) = m.DailyBar.PriceDisplay(m.DailyBar(eBARS_Open, m.DailyBar.Size - 1))
                Case "High"
                    .TextMatrix(1, i) = m.DailyBar.PriceDisplay(m.dHigh)
                Case "Low"
                    .TextMatrix(1, i) = m.DailyBar.PriceDisplay(m.dLow)
                Case "Close"
                    .TextMatrix(1, i) = m.DailyBar.PriceDisplay(m.DailyBar(eBARS_Close, m.DailyBar.Size - 1))
            End Select
        Next
                
        .Redraw = flexRDBuffered
    End With

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmTickDistribution.DisplayQBar", eGDRaiseError_Raise

End Sub

Public Sub ReleaseMarketDepth()
On Error GoTo ErrSection:

    If Not m.oBidAskDepth Is Nothing Then
        g.RealTime.RemoveMarketDepthSymbol m.strSym
        Set m.oBidAskDepth = Nothing
    End If
    
    tmr.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.ReleaseMarketDepth", eGDRaiseError_Raise

End Sub

Public Property Get TickLineRightToLeft() As Long
On Error GoTo ErrSection:

    TickLineRightToLeft = m.nTickLineRL

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.TickLineRightToLeftGet"

End Property

Public Property Let TickLineRightToLeft(ByVal nFlag&)
On Error GoTo ErrSection:

    m.nTickLineRL = nFlag

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.TickLineRightToLeftLet"

End Property

Private Sub SetCaption(Optional ByVal nDate As Long = 0)
On Error GoTo ErrSection:

    Dim strName$, strType$

    If m.eDisplayStyle = eView_Detail Then
        strType = "Market Depth"
    Else
        strType = "Price Ladder"
    End If
    
    If nDate > 0 Then
        strName = m.strSym & " on " & DateFormat(nDate)
    Else
        strName = m.strSym
    End If
    
    SetEditorCaption Me, strType, strName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.SetCaption", eGDRaiseError_Raise
End Sub

Public Property Get WindowLink() As cWindowLink
On Error GoTo ErrSection:

    Set WindowLink = m.WindowLink

    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.WindowLinkGet"

End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.nSymID = 0& Then
        SymbolOrSymbolID = m.strSym
    Else
        SymbolOrSymbolID = m.nSymID
    End If
End Property

Public Property Get SymbolID() As Long
On Error GoTo ErrSection:

    SymbolID = m.nSymID

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.SymbolID.Get"

End Property

Public Property Let SymbolID(ByVal nSymbolID As Long)
On Error GoTo ErrSection:
    
    ChangeSymbol nSymbolID

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.SymbolID.Let"

End Property

Private Sub ShowHelp(KeyCode As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Nothing
    End If


    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.ShowHelp"

End Sub

Private Function IsEnableDOM(Optional ByVal bAddToStream As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim TempBars As New cGdBars
    Dim nOkay As Long
    Dim nEndDate&, i&
                              
    If g.nReplaySession > 0 Or frmReplay.Visible Then
        IsEnableDOM = 1
        Exit Function
    ElseIf frmStatus.IsBusy Or frmStatus.Visible Then
        IsEnableDOM = 2
        Exit Function
    End If
    
    'JM:09-04-2009 - If the symbol is not already streaming, the call to CanHaveMarketDepth will fail.
    'Adding the symbol to the stream is intended to have the same result as adding it to QB or putting
    'it on a chart. For some reason, as of 09-04-2009, this attempt to add the symbol to the stream
    'no longer ensures a correct response from CanHaveMarketDepth. Steps to reproduce: bring up a
    'symbol not currently in stream, CanHaveMarketDepth will return 0 even if you are enabled. Add
    'the symbol to QB or bring it up in a chart, CanHaveMarketDepth will return 1 ef you are enabled.
    
    'try adding the symbol to RT and ask again for depth of market (DOM)
    SetBarProperties TempBars, m.nSymID
    TempBars.Prop(eBARS_Periodicity) = ePRD_EachTick
    g.RealTime.AddTickBuffer TempBars
    
    For nEndDate = Date + 1 To LastDailyDownload + 1 Step -1
        g.RealTime.SpliceBars TempBars, nEndDate
    Next
    
    'if symbol not currently in real time stream then add it
    If TempBars.Size < 1 Then
        If frmStatus.IsBusy Or frmStatus.Visible Then
            nOkay = 2
            g.RealTime.RemoveTickBuffer TempBars
            Set TempBars = Nothing
            IsEnableDOM = nOkay
            Exit Function
        End If
        g.RealTime.RefreshSymbolList (Not g.RealTime.Active)
    End If
                
'    If m.eDisplayStyle = eView_Detail Then
'        g.ChartGlobals.nPauseBidAsk = 1
'        For i = 0 To Forms.Count - 1
'            If TypeOf Forms(i) Is frmTickDistribution Then
'                Forms(i).PauseBidAskDepth
'                DoEvents
'            End If
'        Next
'    End If
    
    'ask again
    nOkay = Abs(g.RealTime.CanHaveMarketDepth(m.strSym))
    g.RealTime.RemoveTickBuffer TempBars
    Set TempBars = Nothing
    
    IsEnableDOM = nOkay

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.IsEnableDOM", eGDRaiseError_Raise
    
End Function

Private Sub InitPrintGrid()
On Error GoTo ErrSection:

    Dim i&

    With fgDOMPrint
        .Redraw = flexRDNone
        SetupGrid fgDOMPrint, eGridMode_Grid
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDNone
        .FixedRows = 1
        .FrozenCols = 8
        .Rows = 1
        .Cols = 8
        .Font.Name = m.strFont
        .Font.Size = m.nFontSize
        .OwnerDraw = flexODOver
        .ExtendLastCol = False
        
        'header alignment
        For i = 0 To 7
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        'alignment
        .ColAlignment(0) = fgBidDetail.ColAlignment(0)
        .ColAlignment(1) = fgBidDetail.ColAlignment(1)
        .ColAlignment(2) = fgBidDetail.ColAlignment(2)
        .ColAlignment(3) = fgBidDetail.ColAlignment(3)
    
        .ColAlignment(4) = fgAskDetail.ColAlignment(0)
        .ColAlignment(5) = fgAskDetail.ColAlignment(1)
        .ColAlignment(6) = fgAskDetail.ColAlignment(2)
        .ColAlignment(7) = fgAskDetail.ColAlignment(3)
    End With

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.IsPrintGrid"

End Sub

Private Sub ClearBidAskCells()
On Error GoTo ErrSection:

    Dim i&

    'clear out depth of market data on ladder grid
    With fgTickDistribution
        .Cell(flexcpFloodPercent, 0, eTDCols_BidSize, .Rows - 1, eTDCols_BidSize) = 0
        .Cell(flexcpFloodPercent, 0, eTDCols_AskSize, .Rows - 1, eTDCols_AskSize) = 0
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, eTDCols_BidSize) = ""
            .TextMatrix(i, eTDCols_AskSize) = ""
        Next
    End With

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.ClearBidAskCells"

End Sub

Public Sub PauseBidAskDepth()
On Error GoTo ErrSection:

    If m.eDisplayStyle = eView_Ladder Then
        If Not m.oBidAskDepth Is Nothing Then
            g.RealTime.RemoveMarketDepthSymbol m.strSym
            Set m.oBidAskDepth = Nothing
            ClearBidAskCells
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.PauseBidaskDepth"
    
End Sub

Private Sub InitCboAccount()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If m.frmBroker Is Nothing Then
        PopulateAccountsCbo cboAccounts, TradeAccountID
        
        If cboAccounts.ListCount > 0 Then
            If cboAccounts.ItemData(cboAccounts.ListIndex) <> TradeAccountID Then
                TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
            End If
                
            If g.nReplaySession > 0 Or frmReplay.Visible Then
                cboAccounts.Enabled = False
                If cboAccounts.ItemData(cboAccounts.ListIndex) <> g.nReplayAccountID Then
                    InfBox "Account " & Str(cboAccounts.ItemData(cboAccounts.ListIndex)) & " does not match Replay Account " & Str(g.nReplayAccountID), "E", "", "Account Error"
                    If m.eOrderBarMode = eGDOrderBarMode_Right Then
                        m.eOrderBarMode = eGDOrderBarMode_LastShownOnRight
                    Else
                        m.eOrderBarMode = eGDOrderBarMode_NotShown
                    End If
                    FormResize Me
                End If
            Else
                cboAccounts.Enabled = True
                If cboAccounts.ItemData(cboAccounts.ListIndex) <> TradeAccountID Then
                    TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
                End If
            End If
        Else
            InfBox "Unable to obtain trading account information.", "!"
            If m.eOrderBarMode = eGDOrderBarMode_Right Then
                m.eOrderBarMode = eGDOrderBarMode_LastShownOnRight
            Else
                m.eOrderBarMode = eGDOrderBarMode_NotShown
            End If
            FormResize Me
        End If
    Else
        cboAccounts.Clear
        
        If m.frmBroker.cboAccounts.ListCount > 0 Then
            For lIndex = 0 To m.frmBroker.cboAccounts.ListCount - 1
                cboAccounts.AddItem Parse(m.frmBroker.cboAccounts.List(lIndex), "(", 1)
                cboAccounts.ItemData(lIndex) = m.frmBroker.cboAccounts.ItemData(lIndex)
            Next lIndex
            
            cboAccounts.ListIndex = m.frmBroker.cboAccounts.ListIndex
            If cboAccounts.ListIndex <> -1& Then
                If cboAccounts.ItemData(cboAccounts.ListIndex) <> TradeAccountID Then
                    TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
                End If
            End If
            
            FormResize Me
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.InitCboAccount"

End Sub

Private Sub UpdateEquityPos()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    Static strPrevPos As String
    
    Dim strText$
    Dim nQuantity&, nBackColor&, i&, iTopRow&
    Dim bCalcPL As Boolean
    Dim bInPosition As Boolean
    Dim bMismatch As Boolean            ' Is the position currently in a position mismatch?
    
    Dim nReverseColor&
    Dim bReverseEnable As Boolean
    
    Dim OrderTree As cGdTree
    Dim Order As cPtOrder
    
    Dim aEquityPos As New cGdArray
    Dim strPrevCaption As String        ' Previous text of the string
    Dim nBroker As eTT_AccountType      ' Broker for the given account ID
                        
    If bInProgress Or CancellingAll Or m.eOrderBarMode <= eGDOrderBarMode_NotShown Then
        Exit Sub
    End If
    
    bInProgress = True
            
    'sync up confirm order flag
    CheckBoxValue(chkConfirmOrder) = g.Broker.ConfirmManual
    'update current position & equity
    'strText = g.Broker.PositionString(TradeAccountID, m.nSymID, 0&)
    strText = GetPositionString
    
    nBroker = g.Broker.AccountTypeForID(TradeAccountID)
    strPrevCaption = lblTradePos.Caption
    
    If Len(strText) > 0 Then
        'get current position info
        m.strPos = Parse(strText, "|", 1)
        m.strPosQty = Parse(strText, "|", 2)
        m.strOpenEq = Parse(strText, "|", 3)
        m.strAvgEntry = Parse(strText, "|", 4)
        m.strSessionQty = Parse(strText, "|", 5)
        m.strSessionPL = Parse(strText, "|", 6)
        
        bMismatch = False
        If InStr(UCase(m.strPos), "LONG") > 0 Then
            lblTradePos.ForeColor = ChartGlbClrForCtl(lblTradePos, g.ChartGlobals.nLongColor, "LongColor")
            lblTradePos.Caption = m.strPos & " " & m.strPosQty
            bInPosition = True
        ElseIf InStr(UCase(m.strPos), "SHORT") > 0 Then
            lblTradePos.ForeColor = g.ChartGlobals.nShortColor
            lblTradePos.Caption = m.strPos & " " & m.strPosQty
            bInPosition = True
        ElseIf UCase(m.strPos) = "MISMATCH" Then
            lblTradePos.ForeColor = 0
            lblTradePos.Caption = "Mismatch"
            m.strPos = "Mismatch"
            bInPosition = True
            bMismatch = True
        Else
            lblTradePos.ForeColor = 0
            lblTradePos.Caption = "Flat"
            m.strPos = "Flat"
            bInPosition = False
        End If
        
'JM: 04-07-2008: some times this double comes in at a non-tradeable price like 1378.125 for ES
'   it is simpler just to convert the display string back into a price that is show-able on ladder
'        aEquityPos.SplitFields strText, "|"
'        If aEquityPos.Size > 6 Then
'            'use Val because Dave uses Str in PositionString    - 4430
'            m.dAvgEntry = Val(aEquityPos(6))
'            m.dAvgEntry = RoundToMinMove(m.dAvgEntry, m.dMinMove)
'            If m.dAvgEntry < 0 Or m.dAvgEntry > 9999999 Then
'                DebugLog m.TickBars.Prop(eBARS_Symbol) & " Avg Entry out of range: " & aEquityPos(3) & " (" & aEquityPos(6) & ")"
'                m.dAvgEntry = ValOfText(aEquityPos(3))
'            End If
        If bMismatch = True Then
            lblTradePos.Caption = "Mismatch"
            lblTradePos.ForeColor = 0
            lblEquity.Caption = ""
            bInPosition = True
            m.strPos = "Flat"
            m.strPosQty = ""
            m.strOpenEq = ""
            m.strAvgEntry = ""
            m.strSessionQty = ""
            m.strSessionPL = ""
        Else
            If InStr(m.strAvgEntry, "^") Then
                m.dAvgEntry = m.TickBars.PriceFromString(m.strAvgEntry)
            Else
                m.dAvgEntry = ValOfText(m.strAvgEntry)
            End If
            
            nQuantity = ValOfText(m.strPosQty)
            
            m.strSessionQty = Parse(strText, "|", 5)
            m.strSessionPL = Parse(strText, "|", 6)
        
            If Len(m.strOpenEq) > 0 Then
                lblEquity.Caption = m.strOpenEq
                If InStr(m.strOpenEq, "-") Or InStr(m.strOpenEq, "(") Then
                    lblEquity.ForeColor = g.ChartGlobals.nLossColor
                Else
                    lblEquity.ForeColor = ChartGlbClrForCtl(lblEquity, g.ChartGlobals.nWinColor, "WinColor")
                End If
            Else
                lblEquity.Caption = ""
            End If
        End If
    Else
        lblTradePos.Caption = "Flat"
        lblTradePos.ForeColor = 0
        lblEquity.Caption = ""
        bInPosition = False
        m.strPos = "Flat"
        m.strPosQty = ""
        m.strOpenEq = ""
        m.strAvgEntry = ""
        m.strSessionQty = ""
        m.strSessionPL = ""
        m.dAvgEntry = 0#            '4049
        nQuantity = 0
    End If
            
    ' DAJ 09/26/2013: If the position label changes, dump it to the log...
    If (lblTradePos.Caption <> strPrevCaption) And (nBroker > 0) Then
        g.Broker.BrokerDebug nBroker, vbTab & vbTab & "Ladder ( '" & m.strSym & "', '" & g.Broker.AccountNameForID(TradeAccountID) & "' ): Position changed to '" & lblTradePos.Caption & "' ( " & strText & " )"
    End If
    
    If Not fgTickDistribution.ColHidden(eTDCols_PL) Then
        If ((m.dAvgEntry <> 0 And ((m.dAvgEntry <> m.dZeroPLPrice) Or (m.nQuantity <> nQuantity))) Or _
            (strPrevPos <> m.strPos)) Then
    
            m.dZeroPLPrice = m.dAvgEntry
            m.nQuantity = nQuantity
            bCalcPL = True
        ElseIf m.dAvgEntry = 0 And m.dZeroPLPrice = -1 Then
            bCalcPL = True      'user switched account and both accounts have same position string (ie long,short,flat)
            m.dZeroPLPrice = 0
        End If
    End If
    
    'save off previous position info
    strPrevPos = m.strPos
    
    'clear out all existing orders from grid
    If m.eDisplayStyle = eView_Ladder Then
        With fgTickDistribution
            .Redraw = flexRDNone
            ClearAllGridOrders
            .Redraw = flexRDBuffered
        End With
        
        'update order info
        m.tbOrders.NumRecords = 0
                
        'Set OrderTree = g.Broker.PrimaryOrdersForSymbol(TradeAccountID, m.nSymID, 0&)
        Set OrderTree = GetWorkingOrders
        If OrderTree Is Nothing Then
            If Not m.oBracketOrdOne Is Nothing Then
                UpdateOrderTable OrderTree, bCalcPL, m.strPos
                OrderTblToGrid
            End If
        ElseIf OrderTree.Count > 0 Or Not m.oBracketOrdOne Is Nothing Then
            UpdateOrderTable OrderTree, bCalcPL, m.strPos
            OrderTblToGrid
        End If
                
        UpdateOpenEntries

    Else
        Set OrderTree = g.Broker.PrimaryOrdersForSymbol(TradeAccountID, m.nSymID, 0&)
        
        With fgOrdersInfo
            .Redraw = flexRDNone
            iTopRow = .TopRow
            If Not OrderTree Is Nothing Then
                .Rows = OrderTree.Count + 1
                For i = 1 To OrderTree.Count
                    Set Order = OrderTree(i)
                    If Not Order Is Nothing Then
                        .TextMatrix(i, 0) = Order.OrderText & " (" & OrderStatus(Order.Status) & ")"
                        .TextMatrix(i, 1) = "X"
                        .TextMatrix(i, 2) = Order.OrderID
                    End If
                Next
            Else
                .Rows = 1
            End If
            If .Rows > 4 And .ColWidth(0) > .Width - (kWidthOrderX * 2) Then
                .ColWidth(0) = .Width - (kWidthOrderX * 2)      'scroll bar visible
            ElseIf .Rows < 5 And .ColWidth(0) < .Width - kWidthOrderX Then
                .ColWidth(0) = .Width - kWidthOrderX
            End If
            If iTopRow >= .FixedRows And iTopRow < .Rows Then
                .TopRow = iTopRow
                m.dScrollTickCount = 0          'to distinguish program scroll from user's scrol
            End If
            
            .Redraw = flexRDBuffered
        End With
    End If
    
    If bCalcPL Then
        CalcPLData True
    End If
    
    'set frame color
    nBackColor = -1
    If TypeOfAccount <> eGDTypeOfAccount_BrokerLive Or g.nReplaySession > 0 Or frmReplay.Visible Then
        nBackColor = fraFrontMonth.BackColor           '5962  Do NOT use Me.BackColor as this gets changed (see code below)
        'need to do this here because account drop down does not get re-populated while user is downloading data
        If frmReplay.Visible Then cboAccounts.Enabled = False
    End If
    
    Select Case UCase(m.strPos)
        Case "LONG"
            If nBackColor = -1 Then nBackColor = kFrameLong
            nReverseColor = kFrameShort
            bReverseEnable = True
        Case "SHORT"
            If nBackColor = -1 Then nBackColor = kFrameShort
            nReverseColor = kFrameLong
            bReverseEnable = True
        Case Else
            If nBackColor = -1 Then nBackColor = kFrameLive
            nReverseColor = Me.BackColor
            bReverseEnable = False
    End Select
    
    BackColor = nBackColor
    vseOrderBar.BackColor = nBackColor
    fraOrderBtns.BackColor = nBackColor
    fraRithmic.BackColor = nBackColor
    chkConfirmOrder.BackColor = nBackColor
    chkAutoExit.BackColor = nBackColor
    lblAutoExit.BackColor = nBackColor
    chkAutoJournal.BackColor = nBackColor
    lblBrokerDisconnect.BackColor = nBackColor
    fraExitFavorites.BackColor = nBackColor
    
    'highlight colors
    If m.nHighlightPos > 0 Then
        lblTradePos.BackColor = m.nHighlightPos
    Else
        lblTradePos.BackColor = Me.BackColor
    End If
    If m.nHighlightEquity > 0 Then
        lblEquity.BackColor = m.nHighlightEquity
    Else
        lblEquity.BackColor = Me.BackColor
    End If
    
    If (CancellingAll = True) Or _
       (m.eDisplayStyle = eView_Ladder And m.tbOrders.NumRecords = 0) Or _
       (m.eDisplayStyle = eView_Detail And fgOrdersInfo.Rows <= fgOrdersInfo.FixedRows) Then
       
        cmdCancelAll.Enabled = False
       
    Else
        cmdCancelAll.Enabled = True
    End If
    
    If Flattening Or Reversing Then
        If cmdCancelAll.Enabled = True Then
            cmdCancelAll.Enabled = False
        End If
        If cmdReverse.Enabled = True Then
            cmdReverse.Enabled = False
        End If
        If cmdBailout.Enabled = True Then
            cmdBailout.Enabled = False
            cmdBailout.BackColor = Me.BackColor
        End If
    ElseIf bInPosition Then
'JM 09-01-2010: original code commented out to implement 5900 (leave awhile then remove if all okay)
'        If cmdReverse.Enabled = False Then
'            cmdReverse.Enabled = True
'        End If
        If cmdReverse.BackColor <> nReverseColor Then           '5900
            cmdReverse.BackColor = nReverseColor
            cmdReverse.Enabled = bReverseEnable
        End If
        If cmdBailout.Enabled = False Then
            cmdBailout.Enabled = True
            cmdBailout.BackColor = RGB(192, 0, 0) ' &HFFFF&
        End If
    Else
        If cmdReverse.Enabled = True Then
            cmdReverse.Enabled = False
        End If
        cmdReverse.BackColor = Me.BackColor
        If cmdBailout.Enabled = True Then
            cmdBailout.Enabled = False
            cmdBailout.BackColor = Me.BackColor
        End If
    End If
    
    'color average entry row if turned on
    If m.nShowAvgEntry Then
        i = AvgEntryGridRow()
        With fgTickDistribution
            If i > .FixedRows And i < .Rows Then
                .Cell(flexcpBackColor, i, 0, i, eTDCols_Price - 1) = m.nAvgEntryColor
                .Cell(flexcpBackColor, i, eTDCols_Price + 1, i, .Cols - 1) = m.nAvgEntryColor
                If Not fgTickDistribution.ColHidden(eTDCols_PL) Then .TextMatrix(i, eTDCols_PL) = "Avg Entry"
            End If
        End With
    ElseIf m.nPrevAvgEntryRow > 0 Then
        i = AvgEntryGridRow()       'call to clear
    End If
        
    bInProgress = False
    Exit Sub

ErrSection:
    bInProgress = False
    RaiseError "frmTickDistribution.UpdateEquityPos"
    
End Sub

Private Sub OneClickOrder(ByVal dPrice#, ByVal bBuy As Boolean, _
    ByVal eOrderType As eTT_OrderType, Optional ByVal nQty& = 0, _
    Optional ByVal bShowInOrderCol As Boolean = False)
On Error GoTo ErrSection:

    Dim Order As cPtOrder
    Dim nQuantity As Long
    Dim strType As String
    
    Dim bOK As Boolean
       
'    If m.bIgnoreClick Then Exit Sub        'aardvark 4554 - always allow one-click buy/sell orders from buy market, sell market etc.

If InStr(m.strSym, "-0") <> 0 Then
    'something very, very wrong - do not continue!!!
    InfBox "Internal Order bar error: " & m.strSym & " cannot be traded. Please close the ladder and try again.", "E", "Ok", "Price Ladder"
    m.eOrderBarMode = eGDOrderBarMode_NotShown
    Form_Resize
    Exit Sub
End If

    
    If nQty > 0 Then
        nQuantity = nQty
    Else
        nQuantity = m.Quantity.Price
    End If
    If nQuantity <= 0 Then
        InfBox "Order quantity must be greater than zero", "!", , "Price Ladder Order Error"
        Exit Sub
    ElseIf eOrderType = eTT_OrderType_Market Then
        dPrice = 0#
    ElseIf dPrice <= 0 And Not m.bIsSpreadSymbol Then
        InfBox "Invalid price.", "E", , Me.Caption
        Exit Sub
    End If
    
    Set Order = New cPtOrder
    Order.AccountID = TradeAccountID

'JM 12-15-2011 - should not need to roll symbol here (symbol should already be correct for order bar to be visible)
'    Order.SymbolOrSymbolID = RollSymbolForDate(m.strSym, m.nSessionDate)
    
    Order.SymbolOrSymbolID = m.strSym
    Order.Buy = bBuy
    Order.OrderType = eOrderType
    Order.OrderPrice(False) = dPrice
    Order.Quantity = nQuantity
    Order.OrderID = -1
    
    If vseBracketOrder.Appearance = apInset Then
        bOK = OkayToExecute(Order, RoundToMinMove(m.TickBars(eBARS_Close, m.TickBars.Size - 1), m.dMinMove), True, Me)
    Else
        bOK = OkayToExecute(Order, RoundToMinMove(m.TickBars(eBARS_Close, m.TickBars.Size - 1), m.dMinMove))
    End If
    
    If Not bOK Then Exit Sub
    
    If cboExchanges.ListIndex >= 0 Then
        Order.Exchange = cboExchanges.Text
    End If
    
    g.Broker.BrokerDebug Broker, "Creating Order from Ladder: " & Order.OrderText, True
    If vseBracketOrder.Appearance = apInset Then
        Order.GenesisOrderID = NextGenesisOrderID(AccountNumber)
        If m.oBracketOrdOne Is Nothing Then
            Set m.oBracketOrdOne = Order
            ParkOrder m.oBracketOrdOne
        Else
            Dim bBracketOk As Boolean
            Dim strOrdOne$, strOrd2$
            Dim strOrdTwo$
            Dim X&, Y&
            
            bBracketOk = True
            
            If ConfirmOrder Then
                strOrdOne = m.oBracketOrdOne.OrderText
                strOrdTwo = Order.OrderText
                
                If InfBox(strOrdOne & vbCrLf & strOrdTwo, "?", "+Ok|-Cancel", "Bracket order confirm", , , , , , , , , , , X, Y) = "C" Then
                    bBracketOk = False
                    Set Order = Nothing
                End If
            End If
            
            If bBracketOk Then
                Set m.oBracketOrdTwo = Order
                ParkOrder Order
                
                If g.Broker.HoldOcoAtBroker(Order.AccountID) Then
                    m.oBracketOrdOne.BrokerCancelOrderID = -(Order.OrderID)
                    Order.BrokerCancelOrderID = -(m.oBracketOrdOne.OrderID)
                Else
                    ' JM 09/02/2009: Per Dave - Must set the cancel order ID on both orders
                    m.oBracketOrdOne.CancelOrderID = Order.OrderID
                    Order.CancelOrderID = m.oBracketOrdOne.OrderID
                End If
                                
                ' JM 09/02/2009: Per Dave - Allow the SubmitOrder routine to handle submitting the
                ' other side of the OCO, but tell it not to ask the user about it...
                SubmitOrder m.oBracketOrdOne, , , , False
            End If
            ClearBuySellButtons True
        
        End If
    Else
        CreateThisOrder Order
    End If
    
    Set Order = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.OneClickOrder"

End Sub

Private Sub UpdateTradeQuantity(Button As Integer, ByVal lQuantity As Long)
On Error GoTo ErrSection:
    
    If Button = vbRightButton Then
        m.Quantity.Price = m.Quantity.Price - lQuantity
    Else
        m.Quantity.Price = m.Quantity.Price + lQuantity
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.UpdateTradeQuantity"

End Sub

Private Sub txtTradeQty_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTradeQty

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.txtTradeQty_GotFocus"

End Sub

Private Sub txtTradeQty_LostFocus()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.txtTradeQty_LostFocus"

End Sub

Private Sub ClearAllGridOrders()
On Error GoTo ErrSection:

    Dim i&
            
    With fgTickDistribution
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, eTDCols_HasOrder) <> "" Then
                .TextMatrix(i, eTDCols_OrderBidX) = ""
                .TextMatrix(i, eTDCols_OrderBid) = ""
                .TextMatrix(i, eTDCols_OrderAsk) = ""
                .TextMatrix(i, eTDCols_OrderAskX) = ""
                .TextMatrix(i, eTDCols_HasOrder) = ""
                .Cell(flexcpFontBold, i, eTDCols_OrderBidX) = False
                .Cell(flexcpFontBold, i, eTDCols_OrderBid) = False
                .Cell(flexcpFontBold, i, eTDCols_OrderAsk) = False
                .Cell(flexcpFontBold, i, eTDCols_OrderAskX) = False
                .Cell(flexcpBackColor, i, eTDCols_OrderAsk) = .BackColor
                .Cell(flexcpBackColor, i, eTDCols_OrderAskX) = .BackColor
                .Cell(flexcpBackColor, i, eTDCols_OrderBid) = .BackColor
                .Cell(flexcpBackColor, i, eTDCols_OrderBidX) = .BackColor
                .Cell(flexcpForeColor, i, eTDCols_OrderAsk) = vbBlack
                .Cell(flexcpForeColor, i, eTDCols_OrderAskX) = vbBlack
                .Cell(flexcpForeColor, i, eTDCols_OrderBid) = vbBlack
                .Cell(flexcpForeColor, i, eTDCols_OrderBidX) = vbBlack
            End If
            If Len(.TextMatrix(i, eTDCols_Entries)) > 0 Then
                .Cell(flexcpPicture, i, eTDCols_Entries) = Nothing
                .TextMatrix(i, eTDCols_Entries) = ""
            End If
        Next
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.ClearAllGridOrders"
    
End Sub

Private Sub vseBracketOrder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    Else
        If vseBracketOrder.Appearance = ap3D Then
            vseBracketOrder.Appearance = apInset
        Else
            ClearBuySellButtons
        End If
    
    End If

End Sub

Private Sub vseExitA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitA, Button, "A", 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.vseExitA_MouseUp"

End Sub

Private Sub vseExitB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitB, Button, "B", 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.vseExitB_MouseUp"

End Sub

Private Sub vseExitC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitC, Button, "C", 2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.vseExitC_MouseUp"

End Sub

Private Sub vseExitD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    ExitFavoriteBtnClick Me, vseExitD, Button, "D", 3

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.vseExitD_MouseUp"

End Sub

Private Sub vseOrderBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
                
    HandleMouseMove
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.vseOrderBar_MouseMove"

End Sub

' Note: X and Y should only be passed in from the mousemove event of the fgTickDist grid
Private Sub HandleMouseMove(Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1)
On Error GoTo ErrSection:

    Dim i&, strType$
    Static nRow&, nCol&, nPrevOrderRow&, strTip$
    
    'If fgTickDistribution.MouseCol < 0 Or fgTickDistribution.MouseRow < 0 Then
    '    Exit Sub        'continuing will give invalid array index error
    'End If
    
    With fgTickDistribution
        ' We only need to do something when the mouse row or column has changed
        ' (when mouse is not over the grid, MouseRow and MouseCol will be -1)
        If .MouseRow <> nRow Or .MouseCol <> nCol Then
            If .MouseCol < 0 Or .MouseRow < 0 Then
                nRow = -1
                nCol = -1
                strTip = ""
            Else
                nRow = .MouseRow
                nCol = .MouseCol
            End If
            .ToolTipText = "" '(to force the tip to move whenever Row or Col has changed)
            
            ' check if market depth data should be shown
            Select Case nCol
            Case eTDCols_OrderBidX, eTDCols_OrderBid, eTDCols_BidSize, eTDCols_AskSize, eTDCols_OrderAsk, eTDCols_OrderAskX
                m.bHideLadderDOM = True
                ' turn off if is now shown
                If m.bLadderHasDOM Then
                    For i = .TopRow To .BottomRow
                        .TextMatrix(i, eTDCols_BidSize) = ""
                        .TextMatrix(i, eTDCols_AskSize) = ""
                        .Cell(flexcpFloodPercent, i, eTDCols_BidSize) = 0
                        .Cell(flexcpFloodPercent, i, eTDCols_AskSize) = 0
                    Next
                    m.bLadderHasDOM = False
                End If
            Case Else
                m.bHideLadderDOM = False
                ' turn on if not now shown
                If Not m.bLadderHasDOM Then DisplayMarketDepthOnLadder
            End Select
            
            'clear out any order info from last mousemove
            If nPrevOrderRow >= .FixedRows And nPrevOrderRow < .Rows Then
                If Len(.TextMatrix(nPrevOrderRow, eTDCols_HasOrder)) = 0 Then
                    .TextMatrix(nPrevOrderRow, eTDCols_OrderBid) = ""
                    .TextMatrix(nPrevOrderRow, eTDCols_OrderAsk) = ""
                    .Cell(flexcpBackColor, nPrevOrderRow, eTDCols_OrderAsk) = .BackColor
                    .Cell(flexcpBackColor, nPrevOrderRow, eTDCols_OrderBid) = .BackColor
                End If
                nPrevOrderRow = 0
            End If
            
            ' clear dragging order if move out of same column
            If m.nDragOrderRow > 0 Then
                If nCol <> eTDCols_OrderAsk And nCol <> eTDCols_OrderBid And nCol >= 0 Then
                    m.nDragOrderRow = 0
                ElseIf nRow < .FixedRows Or nRow > .Rows Then
                    'm.nDragOrderRow = 0
                End If
            End If
            
            If nRow >= .FixedRows Then
                
                If InStr(m.strSym, "-0") <> 0 Then
                    'continuous contract not supported for orders
                    If nCol = eTDCols_BidSize Or nCol = eTDCols_AskSize Or _
                        nCol = eTDCols_OrderAsk Or nCol = eTDCols_OrderBid Or _
                        nCol = eTDCols_OrderAskX Or nCol = eTDCols_OrderBidX Then
                        
                        nCol = eTDCols_Volume
                        strTip = ""
                        
                    End If
                End If
                
                Select Case nCol
                Case eTDCols_PL
                    If m.dAvgEntry <= 0 Or Not vseOrderBar.Visible Then
                        strTip = "Can click in cell to set the 0-level for profit/loss"
                    End If
                Case eTDCols_OrderAskX, eTDCols_OrderBidX
                    If Len(.TextMatrix(nRow, nCol)) > 0 Then
                        strTip = "Click here to cancel order"
                    Else
                        strTip = ""
                    End If
                Case eTDCols_BidSize
                    If .ColHidden(eTDCols_OrderAsk) Then
                        strTip = "Click in cell to BUY at specified price"
                    ElseIf Len(.TextMatrix(nRow, eTDCols_HasOrder)) = 0 Then
                        Select Case cboOrderType.ListIndex
                        Case 1
                            strType = "Limit"
                        Case 2
                            strType = "Stop"
                        Case 3
                            strType = "MIT"
                        Case Else
                            If nRow >= m.nLastPriceRow Then
                                strType = "Limit"
                            Else
                                strType = "Stop"
                            End If
                        End Select
                        If .ColHidden(eTDCols_OrderBid) Then
                            .TextMatrix(nRow, eTDCols_OrderAsk) = "Buy " & txtTradeQty.Text & " " & strType
                            .Cell(flexcpBackColor, nRow, eTDCols_OrderAsk) = m.nBidColor
                            If m.nBidColor = 0 Then .Cell(flexcpForeColor, nRow, eTDCols_OrderAsk) = .ForeColorFixed
                        Else
                            .TextMatrix(nRow, eTDCols_OrderBid) = "Buy " & txtTradeQty.Text & " " & strType
                            .Cell(flexcpBackColor, nRow, eTDCols_OrderBid) = m.nBidColor
                            If m.nBidColor = 0 Then .Cell(flexcpForeColor, nRow, eTDCols_OrderBid) = .ForeColorFixed
                        End If
                        nPrevOrderRow = nRow
                        strTip = ""
                    Else
                        strTip = ""
                    End If
                Case eTDCols_AskSize
                    If .ColHidden(eTDCols_OrderAsk) Then
                        strTip = "Click in cell to SELL at specified price"
                    ElseIf Len(.TextMatrix(nRow, eTDCols_HasOrder)) = 0 Then
                        Select Case cboOrderType.ListIndex
                        Case 1
                            strType = "Limit"
                        Case 2
                            strType = "Stop"
                        Case 3
                            strType = "MIT"
                        Case Else
                            If nRow <= m.nLastPriceRow Then
                                strType = "Limit"
                            Else
                                strType = "Stop"
                            End If
                        End Select
                        .TextMatrix(nRow, eTDCols_OrderAsk) = "Sell " & txtTradeQty.Text & " " & strType
                        .Cell(flexcpBackColor, nRow, eTDCols_OrderAsk) = m.nAskColor
                        If m.nAskColor = 0 Then .Cell(flexcpForeColor, nRow, eTDCols_OrderAsk) = .ForeColorFixed
                        nPrevOrderRow = nRow
                        strTip = ""
                    Else
                        strTip = ""
                    End If
                Case eTDCols_OrderAsk
                    If m.nDragOrderRow > 0 And Len(.TextMatrix(nRow, eTDCols_HasOrder)) = 0 Then
                        .TextMatrix(nRow, eTDCols_OrderAsk) = .TextMatrix(m.nDragOrderRow, eTDCols_OrderAsk)
                        nPrevOrderRow = nRow
                    End If
                Case eTDCols_OrderBid
                    If m.nDragOrderRow > 0 And Len(.TextMatrix(nRow, eTDCols_HasOrder)) = 0 Then
                        .TextMatrix(nRow, eTDCols_OrderBid) = .TextMatrix(m.nDragOrderRow, eTDCols_OrderBid)
                        nPrevOrderRow = nRow
                    End If
                Case Else
                    strTip = ""
                End Select
            End If
            
            ' set mouse pointer
            If nCol = eTDCols_BidSize And nRow >= .FixedRows Then
                .MousePointer = flexCustom
                If vseBracketOrder.Appearance = apInset Then
                    .MouseIcon = Picture16(ToolbarIcon("kOrderBuySell"))
                Else
                    .MouseIcon = Picture16(ToolbarIcon("kOrderBuy"))
                End If
            ElseIf nCol = eTDCols_AskSize And nRow >= .FixedRows Then
                .MousePointer = flexCustom
                If vseBracketOrder.Appearance = apInset Then
                    .MouseIcon = Picture16(ToolbarIcon("kOrderBuySell"))
                Else
                    .MouseIcon = Picture16(ToolbarIcon("kOrderSell"))
                End If
            Else
                .MousePointer = flexArrow
            End If
        End If
        
        ' but always refresh Tooltip when moving over the volume column
        If m.nShowTickLine And (.MouseCol = eTDCols_Volume) And X >= 0 And Y >= 0 Then
            On Error Resume Next
            strTip = Space(15)
            i = geTickTime(m.geTickObj, strTip, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
            If i <> 0 Then
                strTip = ""
            Else
                FixNullTermStr strTip
            End If
            On Error GoTo ErrSection
        End If
        
        If strTip <> .ToolTipText Then
            .ToolTipText = "" '(to force it to move to correct spot after updating
            .ToolTipText = strTip
        End If
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.HandleMouseMove"
End Sub

Private Sub CenterLadderOnCurrPrice()
On Error GoTo ErrSection:

    Dim i&, j&
    
    If m.dScrollTickCount > 0 And gdTickCount() - m.dScrollTickCount < 1500 Then Exit Sub
    
    'center on current price
    With fgTickDistribution
        If .Rows > .FixedRows + 3 Then
            i = .TopRow
            j = .BottomRow
            i = (j - i) / 2
            If m.nLastPriceRow - i > .FixedRows Then
                .TopRow = m.nLastPriceRow - i
            Else
                .TopRow = (.Rows - .BottomRow) / 2 + 1
            End If
            If .Row < .TopRow Then
                .Row = .TopRow       '3853
            ElseIf .Row > .BottomRow Then
                .Row = .BottomRow
            End If
            'trigger timescale redraw
            '.TextMatrix(0, 0) = ""
        End If
    End With
    
    'grid scroll event is called every time the grid's top row is changed
    m.dScrollTickCount = 0

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.CenterLadderOnCurrPrice"

End Sub

Public Property Get TradeAccountID() As Long
On Error GoTo ErrSection:

    If g.nReplaySession > 0 Or frmReplay.Visible Then
        TradeAccountID = g.nReplayAccountID
    Else
        TradeAccountID = m.nTradeAcctID
    End If
    
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.TradeAccountID.Get"
    
End Property

Public Property Let TradeAccountID(ByVal nAccountID&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    If g.nReplaySession = 0 And Not frmReplay.Visible Then
        m.nTradeAcctID = nAccountID
        
        If Not m.frmBroker Is Nothing Then
            For lIndex = 0 To cboAccounts.ListCount - 1
                If cboAccounts.ItemData(lIndex) = nAccountID Then
                    If cboAccounts.ListIndex <> lIndex Then
                        cboAccounts.ListIndex = lIndex
                        If Not g.RealTime.Active Then
                            UpdateEquityPos
                            tmr.Enabled = True
                            tmr.Interval = 500
                        End If
                    End If
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.TradeAccountID.Let"
    
End Property

Private Function HandleLadderResize() As Boolean
On Error Resume Next        'to be called from resize event so don't use raiseerror

    Dim bMergeRow As Boolean
    Dim bLockedWindow As Boolean

m.bIgnoreClick = True
bLockedWindow = LockWindowUpdate(Me.hWnd)

    If OrdBarVisible() Then PositionOrderBar
    
    With fgTickDistribution
        If .Visible And .Rows = .FixedRows + 1 Then
            bMergeRow = .MergeRow(.FixedRows)
        End If
        If Not vseOrderBar.Visible Then     '4289
            If fgQuoteBar.Visible Then
                .Move 0, fgQuoteBar.Height, Me.ScaleWidth, Me.ScaleHeight - fgQuoteBar.Height
            Else
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
        End If
    End With
    
    PositionLadderCols
    
    If Not g.RealTime.Active Then
        DisplayBestBidAsk False
        If fgQuoteBar.Visible Then DisplayQBar
        If fgAccountBar.Visible Then DisplayAcctBar
    End If
    
    HandleLadderResize = bMergeRow

If bLockedWindow Then LockWindowUpdate 0
m.bIgnoreClick = False

End Function

Private Sub PositionLadderCols()
On Error Resume Next            'to be called from form_resize so don't use raiseerror

    Dim nX&, i&
        
    If m.bUserSetSize Then Exit Sub
        
    With fgTickDistribution
        .Redraw = flexRDNone
        .ScrollBars = flexScrollBarVertical
        
        .ColWidth(eTDCols_Volume) = kMinColWidthExt         '5536
        .ColWidth(eTDCols_OrderBid) = kMinColWidthExt
        .ColWidth(eTDCols_OrderAsk) = kMinColWidthExt
        .ColWidth(eTDCols_PL) = kMinColWidthExt
        
        If m.nLastPriceRow > 0 And m.nLastPriceRow >= .FixedRows And m.nLastPriceRow < .Rows Then
            .Select m.nLastPriceRow, eTDCols_Price
            If .CellWidth > kMinColWidthExt Or m.bAutosizePrice Then
                .AutoSize eTDCols_Price     'fixes scenario where price spills into next column
                m.bAutosizePrice = False
            End If
        Else
            .ColWidth(eTDCols_Price) = kMinColWidthExt
        End If
        
        .ColWidth(eTDCols_OrderAskX) = kWidthOrderX
        .ColWidth(eTDCols_OrderBidX) = kWidthOrderX
        .ColWidth(eTDCols_HasOrder) = kWidthOrderX
    
        .ColWidth(eTDCols_Entries) = kWidthOrderX * 2
        
        .AutoSize eTDCols_AskSize
        .AutoSize eTDCols_BidSize
        '.ColWidth(eTDCols_BidSize) = kMinColWidth
        '.ColWidth(eTDCols_AskSize) = kMinColWidth
        
        If m.eOrderBarMode <= eGDOrderBarMode_NotShown Or Not m.bSessionCurrent Then
            .ColHidden(eTDCols_OrderAsk) = True
            .ColHidden(eTDCols_OrderAskX) = True
            .ColHidden(eTDCols_OrderBid) = True
            .ColHidden(eTDCols_OrderBidX) = True
        Else
            .ColHidden(eTDCols_OrderAsk) = False
            .ColHidden(eTDCols_OrderAskX) = False
            
            If m.nOrderColumns = 1 Then
                .ColHidden(eTDCols_OrderBid) = False        'orders shown in 2 columns
                .ColHidden(eTDCols_OrderBidX) = False
            Else
                .ColHidden(eTDCols_OrderBid) = True
                .ColHidden(eTDCols_OrderBidX) = True
            End If
        End If
                
        .ColHidden(eTDCols_BidSize) = Not m.bSessionCurrent
        .ColHidden(eTDCols_AskSize) = Not m.bSessionCurrent
        If m.nShowOpenEntries = 1 Then
            .ColHidden(eTDCols_Entries) = False
        Else
            .ColHidden(eTDCols_Entries) = True
        End If
        
        If SecurityType(m.DailyBar) = "I" Then
            .ColHidden(eTDCols_PL) = True
            .ColHidden(eTDCols_Volume) = True
        Else
            If m.nShowVolBar = 1 Then
                .ColHidden(eTDCols_Volume) = False
            Else
                .ColHidden(eTDCols_Volume) = True
            End If
            If m.nShowProfitLoss Then
                .ColHidden(eTDCols_PL) = False
            Else
                .ColHidden(eTDCols_PL) = True
            End If
        End If
        
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then nX = nX + .ColWidth(i)
        Next
        
        If nX < .ClientWidth Then
            .ColWidth(eTDCols_Volume) = .ColWidth(eTDCols_Volume) + (.ClientWidth - nX)
        End If
        
        m.nPriceColWidth = .ColWidth(eTDCols_Price)
        
        .Redraw = flexRDBuffered
    End With

End Sub

Private Sub ShowFrontMonthLadder()
On Error Resume Next        'to be called from resize event so don't use raiseerror


    Dim nY&, nTop&, lWidth&
    
    If m.nShowQuoteBar = 1 Then
        nY = Me.ScaleHeight - fgQuoteBar.Height
        nTop = fgQuoteBar.Top + fgQuoteBar.Height
    Else
        nY = Me.ScaleHeight
        nTop = 0
    End If
    
    If m.eOrderBarMode = eGDOrderBarMode_Right Then
        ' Show the "Front Month" frame on the right hand side of the dialog...
        lblFrontMonthHorz.Visible = False
        lblFrontMonthVert.Visible = True
        
        With fgTickDistribution
            .Move 0, nTop, Me.ScaleWidth - fraFrontMonth.Width, nY
            
            vseOrderBar.Move .Left + .Width, nTop, cmdFrontMonth.Width + 120, nY
            fraFrontMonth.Move 0, 0, cmdFrontMonth.Width + 120, nY
        End With
        
        lblFrontMonthVert.Move 60, 120, cmdFrontMonth.Width, 2000
        cmdFrontMonth.Move 60, lblFrontMonthVert.Top + lblFrontMonthVert.Height
    ElseIf m.eOrderBarMode <> eGDOrderBarMode_NotShown Then
        ' Show the "Front Month" frame at the bottom of the dialog...
        m.eOrderBarMode = eGDOrderBarMode_BottomContinuous
        lblFrontMonthHorz.Visible = True
        lblFrontMonthVert.Visible = False
        
        lWidth = lblFrontMonthHorz.Width + cmdFrontMonth.Width + 60
        
        lblFrontMonthHorz.Move (ScaleWidth - lWidth) / 2, 60
        cmdFrontMonth.Move lblFrontMonthHorz.Left + lblFrontMonthHorz.Width + 60, lblFrontMonthHorz.Top
        
        With fgTickDistribution
            .Move 0, nTop, Me.ScaleWidth, nY - fraFrontMonth.Height
            
            vseOrderBar.Move 0, .Top + .Height, ScaleWidth, lblFrontMonthHorz.Height + 120
            fraFrontMonth.Move 0, 0, ScaleWidth, lblFrontMonthHorz.Height + 120
        End With
    End If

End Sub

Private Sub ShowFrontMonthDOM()
On Error Resume Next        'to be called from resize event so don't use raiseerror

    Dim nTop&, nY&, nX&
    Dim bLockedWindow As Boolean
    
    If m.nShowQuoteBar = 1 Then
        nY = Me.ScaleHeight - fgQuoteBar.Height
        nTop = fgQuoteBar.Top + fgQuoteBar.Height
    Else
        nY = Me.ScaleHeight
        nTop = 0
    End If
    
    If m.eOrderBarMode = eGDOrderBarMode_Right Then
        ' Show the "Front Month" frame on the right hand side of the dialog...
        lblFrontMonthHorz.Visible = False
        lblFrontMonthVert.Visible = True
        
        With fraFrontMonth
            .Width = cmdFrontMonth.Width + 50
            vseOrderBar.Move Me.ScaleWidth - .Width, nTop, .Width, nY
            .Move 0, 0, .Width, nY
            
            lblFrontMonthVert.Move 60, 120, cmdFrontMonth.Width, 2000
            cmdFrontMonth.Move 60, lblFrontMonthVert.Top + lblFrontMonthVert.Height
        
            nX = (Me.ScaleWidth - .Width) / 2
        End With
    ElseIf m.eOrderBarMode <> eGDOrderBarMode_NotShown Then
        ' Show the "Front Month" frame at the bottom of the dialog...
        m.eOrderBarMode = eGDOrderBarMode_BottomContinuous
        lblFrontMonthHorz.Visible = True
        lblFrontMonthVert.Visible = False
        
        With fraFrontMonth
            .Width = Me.ScaleWidth
            .Height = lblFrontMonthHorz.Height + 50
            vseOrderBar.Move 0, Me.ScaleHeight - .Height, .Width, .Height
            .Move 0, 0
            
            nX = lblFrontMonthHorz.Width + cmdFrontMonth.Width + 60
            lblFrontMonthHorz.Move (ScaleWidth - nX) / 2, 60
            cmdFrontMonth.Move lblFrontMonthHorz.Left + lblFrontMonthHorz.Width + 60, lblFrontMonthHorz.Top
        End With
        
        nX = Me.ScaleWidth / 2
    End If
    
    If m.nShowSummaryBar = 1 Then
        With pbBid
            .Move 0, nTop, nX, pbBid.Height
        End With
        With pbAsk
            .Move nX, nTop, nX, pbBid.Height
        End With
        pbBid.Visible = True
        pbAsk.Visible = True
        nTop = pbBid.Top + pbBid.Height
        nY = nY - pbBid.Height
    Else
        pbBid.Visible = False
        pbAsk.Visible = False
    End If
            
    With fgBidDetail
        .Redraw = flexRDNone
            .Move 0, nTop, nX, nY
            .ExtendLastCol = True
            .ColWidth(-1) = nX / .Cols
        .Redraw = flexRDBuffered
    End With
    
    With fgAskDetail
        .Redraw = flexRDNone
            .Move fgBidDetail.Width, nTop, nX, nY
            .ExtendLastCol = True
            .ColWidth(-1) = nX / .Cols
        .Redraw = flexRDBuffered
    End With

End Sub

Private Sub ShowFrontMonthFrame()
On Error Resume Next        'to be called from resize event so don't use raiseerror

    If m.eDisplayStyle = eView_Ladder Then
        ShowFrontMonthLadder
    Else
        ShowFrontMonthDOM
    End If

End Sub

Private Sub CenterOrdBarFrame()
On Error Resume Next        'to be called from resize event so don't use raiseerror

    Dim nTop&, nWidth&, nY&
    
    If m.nShowQuoteBar = 1 And m.nShowAccountBar = 1 Then
        nY = Me.ScaleHeight - vseOrderBar.Height - fgQuoteBar.Height - fgAccountBar.Height
        nTop = fgQuoteBar.Top + fgQuoteBar.Height
    ElseIf m.nShowQuoteBar = 1 And m.nShowAccountBar = 0 Then
        nY = Me.ScaleHeight - vseOrderBar.Height - fgQuoteBar.Height
        nTop = fgQuoteBar.Top + fgQuoteBar.Height
    ElseIf m.nShowQuoteBar = 0 And m.nShowAccountBar = 1 Then
        nY = Me.ScaleHeight - vseOrderBar.Height - fgAccountBar.Height
        nTop = 0
    Else
        nY = Me.ScaleHeight - vseOrderBar.Height
        nTop = 0
    End If
        
    With fgTickDistribution
        .Move 0, nTop, Me.ScaleWidth, nY
    End With
    
    If fgAccountBar.Visible Then
        With fgAccountBar
            .Redraw = flexRDNone
            .Move 0, fgTickDistribution.Top + fgTickDistribution.Height, Me.ScaleWidth, .Height
            .Redraw = flexRDBuffered
        End With
        vseOrderBar.Move 0, fgAccountBar.Top + fgAccountBar.Height, Me.ScaleWidth, vseOrderBar.Height
    Else
        vseOrderBar.Move 0, fgTickDistribution.Top + fgTickDistribution.Height, Me.ScaleWidth, vseOrderBar.Height
    End If
    
    nWidth = cmdCancelAll.Left + cmdCancelAll.Width
    
    If m.eOrderBarMode = eGDOrderBarMode_Right Or ConnectionStatus = eGDConnectionStatus_Connected Or Not Me.cmdBrokerConnect.Visible Then
        fraOrderBtns.Move (ScaleWidth - nWidth) / 2, 0, nWidth, vseOrderBar.Height
    End If
    
End Sub

Private Sub PositionOrderBar()
On Error Resume Next        'to be called from resize event so don't use raiseerror

    Dim bShowExchanges As Boolean       ' Do we want to show the exchange controls?
    Dim bOkayToTrade As Boolean
    
    Dim eSymbolPitType As eFutureSymbolType
        
    Dim strElectronic As String
    Dim strMsg As String
    
    If g.bUnloading Then Exit Sub
    
    eSymbolPitType = GetSymbolPitType(m.strSym)
    
    If eSymbolPitType = eCombinedSymbol Then
        strElectronic = ConvertFutureSymbol(m.strSym, eElectronicSymbol)
        If InStr(strElectronic, "-0") = 0 Then
            strMsg = m.strSym & " cannot be traded on this bar."        'cannot trade individual contract of a combined symbol
        Else
            'this is continuous contract of a combined symbol
            'since ladder cannot trade a continuous contract anyhow, give user choice to go to active individual contract
            strElectronic = RollSymbolForDate(strElectronic)
            If m.eOrderBarMode = eGDOrderBarMode_Right Then
                strMsg = m.strSym & " cannot be traded on this bar. The active electronic contract is " & strElectronic & "."
            Else
                strMsg = m.strSym & " cannot be traded on this bar. The" & vbCrLf & " active electronic contract is " & strElectronic & "."
            End If
        End If
        cmdFrontMonth.Caption = strElectronic
    ElseIf eSymbolPitType = ePitSymbol Then
        If TypeOfAccount = eGDTypeOfAccount_Simulated Then
            'pit symbols can only be traded using a Geneisis Simulated account
            bOkayToTrade = True
            strElectronic = m.strSym
        Else
            'since only individual contracts can be traded in a broker's account, give user choice to go to the active individual contract
            strElectronic = ConvertFutureSymbol(m.strSym, eElectronicSymbol)
            If InStr(strElectronic, "-0") <> 0 Then strElectronic = RollSymbolForDate(strElectronic)
            If m.eOrderBarMode = eGDOrderBarMode_Right Then
                strMsg = m.strSym & " cannot be traded on this bar. The active electronic contract is " & strElectronic & "."
            Else
                strMsg = m.strSym & " cannot be traded on this bar. The" & vbCrLf & " active electronic contract is " & strElectronic & "."
            End If
            cmdFrontMonth.Caption = strElectronic
        End If
    Else
        bOkayToTrade = True
        strElectronic = m.strSym
    End If
    
'JM 10-04-2011: Show order bar only if symbol is trade-able as checked above AND connected to broker.
    If ConnectionStatus <> eGDConnectionStatus_Connected Then
        bOkayToTrade = False
    Else
        bShowExchanges = g.Broker.ShowExchangeControls(m.nTradeAcctID, m.nSymID)
    End If
    lblExchange.Visible = bShowExchanges
    cboExchanges.Visible = bShowExchanges
    
    If InStr(strElectronic, "-0") = 0 And bOkayToTrade Then
        fraOrderBtns.Visible = True
        fraFrontMonth.Visible = False
        If cboAccounts.ListCount = 0 Then InitCboAccount
    ElseIf Len(strMsg) = 0 Then
        If InStr(m.strSym, "-0") <> 0 Then
            'user does not have open positions in any individual contract, automatically go to current contract
            m.strSym = RollSymbolForDate(m.strSym)
            m.bIsSpreadSymbol = IsSpreadSymbol(m.strSym)
            'fixes Ken S. issue with individual contract rolling when should not
            m.nSymID = g.SymbolPool.SymbolIDforSymbol(m.strSym)
        End If
        fraOrderBtns.Visible = True
        fraFrontMonth.Visible = False
        If cboAccounts.ListCount = 0 Then InitCboAccount
    Else
        lblFrontMonthHorz.Alignment = 0
        lblFrontMonthHorz.Caption = strMsg
        lblFrontMonthVert.Caption = strMsg
        lblGoTo.Visible = True
        If m.eOrderBarMode = eGDOrderBarMode_Right Then
            lblGoTo.Move cmdFrontMonth.Left, cmdFrontMonth.Top - lblGoTo.Height
        Else
            lblGoTo.Move (cmdFrontMonth.Left - lblGoTo.Width) + 300, cmdFrontMonth.Top + 100
        End If
    
        fgAccountBar.Visible = False        'no order bar, no accout
        fraOrderBtns.Visible = False
        fgOrdersInfo.Visible = False
        fraFrontMonth.Visible = True
        ShowFrontMonthFrame
        Exit Sub
    End If
    
    If m.eOrderBarMode = eGDOrderBarMode_Right Then
        ' Show the order bar on the right hand side of the form...
        OrdBarOnRight bShowExchanges
    ElseIf m.eOrderBarMode > eGDOrderBarMode_NotShown Then
        ' Show the order bar at the bottom of the form...
        cmdBuyMarket.Top = 10
        cmdSellMarket.Top = 10
        
        cmdBuyMarket.Width = 1290
        cmdSellMarket.Width = 1290
        
        cboOrderType.Width = cboAccounts.Width
        
        If ShowControl("AE") And Len(ExitFavoritesAssigned(m.strSym)) > 0 Then
            fraExitFavorites.Visible = True
        Else
            fraExitFavorites.Visible = False
        End If
            
        
        ' Show order bar at the bottom (narrow - 3 columns)
        If Me.ScaleWidth <= kSizeToSwitch Then
            m.eOrderBarMode = eGDOrderBarMode_BottomNarrow
            OrderBarNarrow bShowExchanges
        ' Show order bar at the bottom (wide - up to 5 columns)
        Else
            m.eOrderBarMode = eGDOrderBarMode_BottomWide
            OrderBarWide bShowExchanges
        End If
        
        cmdClearQty.Visible = bOkayToTrade
        txtTradeQty.Visible = bOkayToTrade
        vscrQty.Visible = bOkayToTrade
        
        'center order bar frame/elastic control
        If m.eDisplayStyle = eView_Ladder Then CenterOrdBarFrame
    End If

End Sub

Private Sub HandleDOMResize()
On Error Resume Next            'to be called from form_resize so don't use raiseerror
    
    Dim nX&, nWidth&, nGridHeight&
    Dim bOrdBarVisible As Boolean
    Dim bLockedWindow As Boolean
    
    Dim nTop&, nBidAskGridHeight&
    
m.bIgnoreClick = True
    
    nX = Me.ScaleWidth / 2
    
    nBidAskGridHeight = Me.ScaleHeight
    
    bOrdBarVisible = OrdBarVisible()

bLockedWindow = LockWindowUpdate(Me.hWnd)
    
    If bOrdBarVisible Then
        PositionOrderBar
        If fraOrderBtns.Visible Then
            If m.eOrderBarMode = eGDOrderBarMode_Right Then
                nX = (Me.ScaleWidth - fraOrderBtns.Width) / 2
            End If
        Else
            GoTo ErrExit
        End If
    End If
        
    If m.nShowQuoteBar = 1 Then
        nTop = fgQuoteBar.Height
        nBidAskGridHeight = Me.ScaleHeight - fgQuoteBar.Height
    End If

    If m.nShowSummaryBar = 1 Then
        With pbBid
            .Move 0, nTop, nX, pbBid.Height
        End With
        With pbAsk
            .Move nX, nTop, nX, pbBid.Height
        End With
        pbBid.Visible = True
        pbAsk.Visible = True
        nTop = pbBid.Top + pbBid.Height
        nBidAskGridHeight = nBidAskGridHeight - pbBid.Height
    Else
        pbBid.Visible = False
        pbAsk.Visible = False
    End If
    
    If bOrdBarVisible Then
        With fgOrdersInfo
            .ColWidth(0) = .Width - kWidthOrderX
            .ColWidth(1) = kWidthOrderX
            .Height = .RowHeight(0) * 5
            nBidAskGridHeight = nBidAskGridHeight - .Height
            If m.nShowAccountBar = 1 Then
                fgAccountBar.Move 0, .Top + .Height, Me.Width
                nBidAskGridHeight = nBidAskGridHeight - fgAccountBar.Height
            End If
        End With
        
        With vseOrderBar
            If m.eOrderBarMode = eGDOrderBarMode_Right Then
                fgOrdersInfo.Move 0, nTop + nBidAskGridHeight, Me.ScaleWidth - .Width, fgOrdersInfo.Height
                If m.nShowQuoteBar Then
                    .Move fgOrdersInfo.Width, fgQuoteBar.Height, .Width, Me.ScaleHeight - fgQuoteBar.Height
                Else
                    .Move fgOrdersInfo.Width, 0, .Width, Me.ScaleHeight
                End If
            Else
                nBidAskGridHeight = nBidAskGridHeight - .Height
                fgOrdersInfo.Move 0, nTop + nBidAskGridHeight, Me.ScaleWidth, fgOrdersInfo.Height
                If m.nShowAccountBar = 1 Then
                    .Move 0, fgAccountBar.Top + fgAccountBar.Height, Me.ScaleWidth, .Height
                Else
                    vseOrderBar.Move 0, fgOrdersInfo.Top + fgOrdersInfo.Height, Me.ScaleWidth, .Height
                End If
                nWidth = cmdCancelAll.Left + cmdCancelAll.Width
                fraOrderBtns.Move (ScaleWidth - nWidth) / 2, 0, nWidth, vseOrderBar.Height
            End If
        End With
        
    End If
    
    With fgBidDetail
        .Redraw = flexRDNone
            .Move 0, nTop, nX, nBidAskGridHeight
            .ExtendLastCol = True
            .ColWidth(-1) = nX / .Cols
        .Redraw = flexRDBuffered
    End With
    
    With fgAskDetail
        .Redraw = flexRDNone
            .Move fgBidDetail.Width, nTop, nX, nBidAskGridHeight
            .ExtendLastCol = True
            .ColWidth(-1) = nX / .Cols
        .Redraw = flexRDBuffered
    End With

ErrExit:
    If bLockedWindow Then LockWindowUpdate 0
    m.bIgnoreClick = False
                   
End Sub

Private Sub GetDailyBarsBidAsk()
On Error GoTo ErrSection:

    Dim i&

    If m.TickBars Is Nothing Then
        Exit Sub
    ElseIf m.TickBars.Size = 0 Then
        Exit Sub
    End If
    
    i = m.DailyBar.Size - 1
    
    If m.DailyBar(eBARS_Bid, i) = kNullData Or m.DailyBar(eBARS_Ask, i) = kNullData Then
        If m.nSessionDate > 0 Then
            Set m.DailyBar = New cGdBars
            m.DailyBar.ArrayMask = eBARS_Eod Or eBARS_BidAsk
            If DM_GetBars(m.DailyBar, m.nSymID, , Int(m.TickBars(eBARS_DateTime, m.TickBars.Size - 1)), , , False) Then
                i = m.DailyBar.Size - 1
                m.dBestBid = m.DailyBar(eBARS_Bid, i)
                m.nBestBidSize = m.DailyBar(eBARS_BidSize, i)
                m.dBestAsk = m.DailyBar(eBARS_Ask, i)
                m.nBestAskSize = m.DailyBar(eBARS_AskSize, i)
            Else
                m.dBestBid = kNullData
                m.dBestAsk = kNullData
                m.nBestAskSize = 0
                m.nBestBidSize = 0
            End If
        End If
    Else
        m.dBestBid = m.DailyBar(eBARS_Bid, i)
        m.nBestBidSize = m.DailyBar(eBARS_BidSize, i)
        m.dBestAsk = m.DailyBar(eBARS_Ask, i)
        m.nBestAskSize = m.DailyBar(eBARS_AskSize, i)
    End If
        
    If m.nBestAskSize < 0 Then m.nBestAskSize = 0
    If m.nBestBidSize < 0 Then m.nBestBidSize = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.GetDailyBarsBidAsk"
    
End Sub

Public Property Get OrderColumns() As Long
On Error GoTo ErrSection:

    OrderColumns = m.nOrderColumns

    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.OrderColumns.Get"
    
End Property

Public Property Let OrderColumns(ByVal nCols&)
On Error GoTo ErrSection:

    m.nOrderColumns = nCols

    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.OrderColumns.Let"
    
End Property

Private Sub NewOrderOnClick()
On Error GoTo ErrSection:

    Dim strAction$, strType$, lCurrentPosition&, dPrice#
    Dim PtOrder As New cPtOrder
        
    Dim bConfirm As Boolean
        
If InStr(m.strSym, "-0") <> 0 Then
    'something very, very wrong - do not continue!!!
    InfBox "Internal Order bar error: " & m.strSym & " cannot be traded. Please close the ladder and try again.", "E", "Ok", "Price Ladder"
    m.eOrderBarMode = eGDOrderBarMode_NotShown
    Form_Resize
    Exit Sub
End If
    
    With fgTickDistribution
        dPrice = ValOfColPrice(.Row)
        PtOrder.StopPrice = dPrice
        PtOrder.LimitPrice = dPrice
        If .Col = eTDCols_BidSize Then
            PtOrder.Buy = True
            strAction = "Buy "
            If cboOrderType.ListIndex = 1 Then
                PtOrder.OrderType = eTT_OrderType_Limit
                strType = " Limit"
            ElseIf cboOrderType.ListIndex = 2 Then
                PtOrder.OrderType = eTT_OrderType_Stop
                strType = " Stop"
            ElseIf cboOrderType.ListIndex = 3 Then
                PtOrder.OrderType = eTT_OrderType_MIT
            Else
                If dPrice > m.dLastPrice Then                   'default to auto order type (4182)
                    PtOrder.OrderType = eTT_OrderType_Stop
                    strType = " Stop"
                Else
                    PtOrder.OrderType = eTT_OrderType_Limit
                    strType = " Limit"
                End If
            End If
        Else
            PtOrder.Buy = False
            strAction = "Sell "
            If cboOrderType.ListIndex = 1 Then
                PtOrder.OrderType = eTT_OrderType_Limit
                strType = " Limit"
            ElseIf cboOrderType.ListIndex = 2 Then
                PtOrder.OrderType = eTT_OrderType_Stop
                strType = " Stop"
            ElseIf cboOrderType.ListIndex = 3 Then
                PtOrder.OrderType = eTT_OrderType_MIT
            Else
                If dPrice < m.dLastPrice Then
                    PtOrder.OrderType = eTT_OrderType_Stop
                    strType = " Stop"
                Else
                    PtOrder.OrderType = eTT_OrderType_Limit
                    strType = " Limit"
                End If
            End If
        End If
                
        PtOrder.OrderID = -1
        
'JM 12-15-2011 - should not need to roll symbol here (symbol should already be correct for order bar to be visible)
'        PtOrder.SymbolOrSymbolID = RollSymbolForDate(m.strSym, m.nSessionDate)
        
        PtOrder.SymbolOrSymbolID = m.strSym
        If m.eOrderBarMode > eGDOrderBarMode_NotShown Then
            PtOrder.AccountID = TradeAccountID
            PtOrder.Quantity = m.Quantity.Price
        Else
            PtOrder.AccountID = DefaultAccount
            lCurrentPosition = g.Broker.CurrentPosition(PtOrder.AccountID, PtOrder.Symbol, 0&)
            If (PtOrder.Buy = True And lCurrentPosition < 0) Or (PtOrder.Buy = False And lCurrentPosition > 0) Then
                PtOrder.Quantity = Abs(lCurrentPosition)
            ElseIf SecurityType(m.strSym) = "S" Then
                PtOrder.Quantity = 100
            Else
                PtOrder.Quantity = 1
            End If
        End If
        
        If cboExchanges.ListIndex >= 0 Then
            PtOrder.Exchange = cboExchanges.Text
        End If

        If vseBracketOrder.Appearance = apInset Then
            bConfirm = False
        Else
            bConfirm = ConfirmOrder
        End If

        If vseOrderBar.Visible And Not bConfirm Then
            OneClickOrder PtOrder.OrderPrice(False), PtOrder.Buy, PtOrder.OrderType, PtOrder.Quantity, True
            .Redraw = flexRDNone
            If .Col = eTDCols_BidSize And Not .ColHidden(eTDCols_OrderBid) Then
                .TextMatrix(.Row, eTDCols_OrderBidX) = "X"
                .TextMatrix(.Row, eTDCols_OrderBid) = strAction & " " & Str(PtOrder.Quantity) & strType
                .Cell(flexcpFontBold, .Row, eTDCols_OrderBidX, .Row, eTDCols_OrderBid) = True
                .Cell(flexcpBackColor, .Row, eTDCols_OrderBidX, .Row, eTDCols_OrderBid) = m.nBidColor
                If m.nBidColor = 0 Then
                    .Cell(flexcpForeColor, .Row, eTDCols_OrderBidX, .Row, eTDCols_OrderBid) = .ForeColorFixed
                End If
            Else
                .TextMatrix(.Row, eTDCols_OrderAskX) = "X"
                .TextMatrix(.Row, eTDCols_OrderAsk) = strAction & Str(PtOrder.Quantity) & strType
                .Cell(flexcpFontBold, .Row, eTDCols_OrderAsk, .Row, eTDCols_OrderAskX) = True
                If PtOrder.Buy Then
                    .Cell(flexcpBackColor, .Row, eTDCols_OrderAsk, .Row, eTDCols_OrderAskX) = m.nBidColor
                Else
                    .Cell(flexcpBackColor, .Row, eTDCols_OrderAsk, .Row, eTDCols_OrderAskX) = m.nAskColor
                End If
                If m.nBidColor = 0 Or m.nAskColor = 0 Then
                    .Cell(flexcpForeColor, .Row, eTDCols_OrderAsk, .Row, eTDCols_OrderAskX) = .ForeColorFixed
                End If
            End If
            .TextMatrix(.Row, eTDCols_HasOrder) = -1    'to prevent clear out in mousemove
            .Redraw = flexRDBuffered
        Else
            g.Broker.BrokerDebug Broker, "Creating Order from Ladder: " & PtOrder.OrderText, True
            PlaceOrder PtOrder
        End If
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.NewOrderOnClick"
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'this routine ammends an existing order based on clicking in bid or ask size columns
'there could be:
'   a)sell order at this row
'   b)buy order at this row
'   c)a buy AND sell order at this row
'
'if only one order column is showing then we load and ammend the visible order
'if two order columns are showing then:
'   a)if only one order exists at this row then ammend that order
'   b)if buy AND sell orders exist at this row then ammend the order matching the column clicked
'     i.e. if ask size column clicked then ammend sell order and vice versa
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AmmendOrder(ByVal eColClicked As eTDCols)
On Error GoTo ErrSection:

    Dim Order As cPtOrder               ' Order object to work with
    Dim lQuantity As Long               ' Quantity from the edit box
    Dim bConfirm As Boolean             ' Confirm order cancels?
    
    If m.bIgnoreClick Then Exit Sub

    bConfirm = ConfirmOrder

    If (eColClicked = eTDCols_BidSize) Or (eColClicked = eTDCols_AskSize) Then
        If (m.nLadderButton <> vbRightButton) Then
            With fgTickDistribution
                Set Order = LoadOrderFromGrid(.Row, eColClicked)
                If (Order Is Nothing) And (eColClicked = eTDCols_AskSize) Then
                    Set Order = LoadOrderFromGrid(.Row, eTDCols_BidSize)
                ElseIf (Order Is Nothing) And (eColClicked = eTDCols_BidSize) Then
                    Set Order = LoadOrderFromGrid(.Row, eTDCols_AskSize)
                End If
            End With
            
            If Not Order Is Nothing Then
                If OrderIsPending(Order) Then
                    InfBox "This order cannot be modified because it is in a pending state.  Please wait for order confirmation.", "!", , "Price Ladder Order Error"
                Else
                    lQuantity = m.Quantity.Price
                    If lQuantity < 1 Then
                        InfBox "Order quantity must be greater than zero", "!", , "Price Ladder Order Error"
                    Else
                        If (eColClicked = eTDCols_BidSize) Then
                            If (Order.Buy = True) Then
                                g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & Order.OrderText, True
                                ModifyThisOrder Order, , Order.RemainingQuantity + lQuantity, bConfirm
                            Else
                                If Order.RemainingQuantity > lQuantity Then
                                    g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & Order.OrderText, True
                                    ModifyThisOrder Order, , Order.RemainingQuantity - lQuantity, bConfirm
                                Else
                                    g.Broker.BrokerDebug Broker, "Cancelling Order from Ladder: " & Order.OrderText, True
                                    CancelThisOrder Order, bConfirm
                                End If
                            End If
                        ElseIf (eColClicked = eTDCols_AskSize) Then
                            If (Order.Buy = False) Then
                                g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & Order.OrderText, True
                                ModifyThisOrder Order, , Order.RemainingQuantity + lQuantity, bConfirm
                            Else
                                If Order.RemainingQuantity > lQuantity Then
                                    g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & Order.OrderText, True
                                    ModifyThisOrder Order, , Order.RemainingQuantity - lQuantity, bConfirm
                                Else
                                    g.Broker.BrokerDebug Broker, "Cancelling Order from Ladder: " & Order.OrderText, True
                                    CancelThisOrder Order, bConfirm
                                End If
                            End If
                        Else
                            g.Broker.BrokerDebug Broker, "Modifying Order from Ladder: " & Order.OrderText, True
                            ModifyThisOrder Order
                        End If
                    End If
                End If
            End If
        End If
    End If
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.AmmendOrder"
    
End Sub

#If 0 Then
Private Sub AmmendOrder(ByVal eColClicked As eTDCols)
On Error GoTo ErrSection:

    Dim PtOrder As cPtOrder
    Dim nQty&, dOrderPrice#
    
    Dim bNewBuyFlag As Boolean
    Dim bAmmendOrder As Boolean
    
    Dim eNewOrdAction As eSamePriceAction

    If eColClicked <> eTDCols_BidSize And eColClicked <> eTDCols_AskSize Then Exit Sub
    If m.nLadderButton = vbRightButton Then Exit Sub            'per Tim: do nothing on right click

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'this routine ammends an existing order based on clicking in bid or ask size columns
'there could be:
'   a)sell order at this row
'   b)buy order at this row
'   c)a buy AND sell order at this row
'
'if only one order column is showing then we load and ammend the visible order
'if two order columns are showing then:
'   a)if only one order exists at this row then ammend that order
'   b)if buy AND sell orders exist at this row then ammend the order matching the column clicked
'     i.e. if ask size column clicked then ammend sell order and vice versa
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With fgTickDistribution
        Set PtOrder = LoadOrderFromGrid(.Row, eColClicked)
        If PtOrder Is Nothing And eColClicked = eTDCols_AskSize Then
            Set PtOrder = LoadOrderFromGrid(.Row, eTDCols_BidSize)
        ElseIf PtOrder Is Nothing And eColClicked = eTDCols_BidSize Then
            Set PtOrder = LoadOrderFromGrid(.Row, eTDCols_AskSize)
        End If
    End With
    
    If PtOrder Is Nothing Then Exit Sub
    
    If OrderIsPending(PtOrder) Then
        InfBox "This order cannot be modified because it is in a pending state.  Please wait for order confirmation.", "!", , "Price Ladder Order Error"
        Exit Sub
    End If
    
    dOrderPrice = PtOrder.OrderPrice(False)
    nQty = m.Quantity.Price
    If nQty < 1 Then
        InfBox "Order quantity must be greater than zero", "!", , "Price Ladder Order Error"
        Exit Sub
    End If
    
    eNewOrdAction = eOrdAction_Unknown
    bAmmendOrder = False
    
    'check existing order and prompt user if necessary
    If eColClicked = eTDCols_BidSize Then
        If PtOrder.Buy Then
            bAmmendOrder = True        'ammend quantity existing buy order
        Else
            'user wants to buy when sell order exists
            If PtOrder.RemainingQuantity > nQty Then
                '01-12-2006 (per Tim): just consolidate
                eNewOrdAction = eOrdAction_Consolidate
                bNewBuyFlag = False
                nQty = PtOrder.RemainingQuantity - nQty
'            ElseIf PtOrder.Quantity < nQty Then
'                nQty = nQty - PtOrder.Quantity
'                strConsolidate = "Consolidate into one BUY " & Str(nQty) & " order."
'                bNewBuyFlag = True
            Else
                'same price, same quantity -- cannot cosolidate
                '01-11-2006 (per Tim): cancel order
                eNewOrdAction = eOrdAction_CancelExisting
            End If
        End If
    ElseIf eColClicked = eTDCols_AskSize Then
        If PtOrder.Buy Then
            'user wants to sell when buy order exists
            If PtOrder.RemainingQuantity > nQty Then
                nQty = PtOrder.RemainingQuantity - nQty
                bNewBuyFlag = True
                eNewOrdAction = eOrdAction_Consolidate
'            ElseIf PtOrder.Quantity < nQty Then
'                nQty = nQty - PtOrder.Quantity
'                bNewBuyFlag = False
'                strConsolidate = "Consolidate into a single SELL " & Str(nQty) & " order."
            Else
                '01-11-2006 (Per Tim): cancel order
                eNewOrdAction = eOrdAction_CancelExisting
            End If
        Else
            bAmmendOrder = True        'ammend quantity of existing sell order
        End If
    Else
        EditOrder PtOrder               'precautionary (theoretically should never get here)
    End If
    
    If bAmmendOrder Then
        PtOrder.Quantity = PtOrder.RemainingQuantity + nQty
    ElseIf eNewOrdAction = eOrdAction_Consolidate Then
        bAmmendOrder = True
        PtOrder.Quantity = nQty
        If PtOrder.Buy <> bNewBuyFlag Then
            If PtOrder.OrderType = eTT_OrderType_Limit Or PtOrder.OrderType = eTT_OrderType_LimitCloseOnly Then
                PtOrder.OrderType = eTT_OrderType_Stop
                PtOrder.OrderPrice(False) = dOrderPrice
            Else
                PtOrder.OrderType = eTT_OrderType_Limit
                PtOrder.OrderPrice(False) = dOrderPrice
            End If
            PtOrder.Buy = bNewBuyFlag
        End If
    ElseIf eNewOrdAction = eOrdAction_CancelExisting Then
        If ConfirmOrder Then
            CancelThisOrder PtOrder, True
        Else
            CancelThisOrder PtOrder, False
        End If
    End If
        
    If bAmmendOrder Then
        If ConfirmOrder Then
            EditOrder PtOrder
        Else
            EditOrder PtOrder, , False, eGDEditOrderReturn_Submit
        End If
    End If
        
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.AmmendOrder"
    
End Sub
#End If

Public Property Get BlankRows() As Long
On Error GoTo ErrSection:

    BlankRows = m.nBlankRows
    
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.BlankRows.Get"
    
End Property

Public Property Let BlankRows(ByVal nRows&)
On Error GoTo ErrSection:

    m.nBlankRows = nRows
    m.nSessionBlankRows = nRows
    
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.BlankRows.Let"
    
End Property

Private Function UpdateOrderTable(OrderTree As cGdTree, bCalPL As Boolean, strPos As String) As Boolean
On Error GoTo ErrSection:

    Dim dPrice#, dNewMax#, dNewMin#, i&
    Dim nRowsForNewMax&, nRowsForNewMin&
        
    Dim bMaxExceeded As Boolean
    Dim bMinExceeded As Boolean
    Dim bExitOrder As Boolean

    Dim Order As cPtOrder
    
    'initialize min/max prices
    dNewMax = m.dMaxPrice
    dNewMin = m.dMinPrice
    
    If Not m.oBracketOrdOne Is Nothing Then
        If m.oBracketOrdOne.OrderType = eTT_OrderType_StopWithLimit Then
            dPrice = RoundToMinMove(m.oBracketOrdOne.OrderPrice(True), m.dMinMove)
        Else
            dPrice = RoundToMinMove(m.oBracketOrdOne.OrderPrice(False), m.dMinMove)
        End If
        If dPrice = 0 And m.oBracketOrdOne.OrderType = eTT_OrderType_Market Then
            dPrice = m.dLastPrice
        End If
        If dPrice > 0 Then
            m.tbOrders.AddRecord " "
            m.tbOrders(0, m.tbOrders.NumRecords - 1) = m.oBracketOrdOne.OrderID
            m.tbOrders(1, m.tbOrders.NumRecords - 1) = Abs(m.oBracketOrdOne.Buy)
            m.tbOrders(2, m.tbOrders.NumRecords - 1) = m.oBracketOrdOne.RemainingQuantity
            m.tbOrders(3, m.tbOrders.NumRecords - 1) = m.oBracketOrdOne.OrderType
            m.tbOrders(4, m.tbOrders.NumRecords - 1) = dPrice
            m.tbOrders(5, m.tbOrders.NumRecords - 1) = Abs(OrderIsPending(m.oBracketOrdOne))
            m.tbOrders(6, m.tbOrders.NumRecords - 1) = m.oBracketOrdOne.Status
            If dPrice > dNewMax Then
                bMaxExceeded = True
                dNewMax = dPrice
            ElseIf dPrice < dNewMin Then
                bMinExceeded = True
                dNewMin = dPrice
            End If
        End If
    End If
    
    OrdersToTableNormal OrderTree, dNewMax, dNewMin
    bMaxExceeded = (dNewMax > m.dMaxPrice)
    bMinExceeded = (dNewMin < m.dMinPrice)
    
    If m.dMinMove > 0 Then
        If bMaxExceeded Then
            dNewMax = dNewMax + m.dMinMove * 3      'add a couple extras
            nRowsForNewMax = (dNewMax - m.dMaxPrice) / m.dMinMove
        End If
        If bMinExceeded Then
            dNewMin = dNewMin - m.dMinMove * 3      'add a couple extras
            nRowsForNewMin = (m.dMinPrice - dNewMin) / m.dMinMove
        End If
        
        If bMaxExceeded Or bMinExceeded Then
            If nRowsForNewMax > nRowsForNewMin Then
                m.nSessionBlankRows = m.nSessionBlankRows + nRowsForNewMax
            Else
                m.nSessionBlankRows = m.nSessionBlankRows + nRowsForNewMin
            End If
            
            LoadTable
            LoadGrid , False
            bCalPL = False   'set to false so won't recalculate P/L on exit
        End If
        
        UpdateOrderTable = bExitOrder
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTickDistribution.UpdateOrderTable"
    
End Function

Private Sub OrderTblToGrid()
On Error GoTo ErrSection:

    Dim i&, j&, strText$
    Dim dPrice#, dPriceGrid#
    
    Dim aIdx As cGdArray
    
    Dim eOrderCol As eTDCols
    Dim eOrderColX As eTDCols
    Dim eOrderType As eTT_OrderType

    
    Set aIdx = m.tbOrders.CreateSortedIndex(4, eGdSort_Default)
    j = 0
    With fgTickDistribution
        .Redraw = flexRDNone
        'strPrice = m.TickBars.PriceDisplay(m.tbOrders(4, aIdx(j)))
        dPrice = m.tbOrders(4, aIdx(j))
        For i = .Rows - 1 To .FixedRows Step -1
            If i - 1 > 0 And i - 1 < m.Data.NumRecords Then
                dPriceGrid = m.Data(eFld_Price, i - 1)
            Else
                dPriceGrid = -1#
            End If
            If dPrice = dPriceGrid Then
                'place order id in hidden column
                .TextMatrix(i, eTDCols_HasOrder) = Str(m.tbOrders(0, aIdx(j)))
                ' === build string to place in grid [begin] ===
                'set buy or sell text
                If m.tbOrders(1, aIdx(j)) = 1 Then
                    strText = "Buy "
                    If m.nOrderColumns = 0 Then
                        eOrderCol = eTDCols_OrderAsk
                        eOrderColX = eTDCols_OrderAskX
                    Else
                        eOrderCol = eTDCols_OrderBid
                        eOrderColX = eTDCols_OrderBidX
                    End If
                    .Cell(flexcpBackColor, i, eOrderCol) = m.nBidColor
                    If m.nBidColor = 0 Then .Cell(flexcpForeColor, i, eOrderCol) = .ForeColorFixed
                Else
                    strText = "Sell "
                    eOrderCol = eTDCols_OrderAsk
                    eOrderColX = eTDCols_OrderAskX
                    .Cell(flexcpBackColor, i, eTDCols_OrderAsk) = m.nAskColor
                    If m.nAskColor = 0 Then .Cell(flexcpForeColor, i, eTDCols_OrderAsk) = .ForeColorFixed
                End If
                'quantity
                strText = strText & Str(m.tbOrders(2, aIdx(j)))
                'order type
                eOrderType = m.tbOrders(3, aIdx(j))
                If eOrderType = eTT_OrderType_Limit Or eOrderType = eTT_OrderType_LimitCloseOnly Then
                    strText = strText & " Limit"
                ElseIf eOrderType = eTT_OrderType_Stop Or _
                       eOrderType = eTT_OrderType_StopCloseOnly Or _
                       eOrderType = eTT_OrderType_StopWithLimit Or _
                       eOrderType = eTT_OrderType_StopWithLimit Or _
                       eOrderType = eTT_OrderType_StopWithLimitCloseOnly Then
                    strText = strText & " Stop"
                End If
                ' === build string to place in grid [end] ===
                .TextMatrix(i, eOrderCol) = strText
                .Cell(flexcpFontBold, i, eOrderCol) = True
                'set text color to gray for pending orders
                If m.tbOrders(5, aIdx(j)) = 0 Then
                    If .Cell(flexcpBackColor, i, eOrderCol) = 0 Then
                        .Cell(flexcpForeColor, i, eOrderCol) = .ForeColorFixed
                    Else
                        .Cell(flexcpForeColor, i, eOrderCol) = vbBlack
                    End If
                Else
                    .Cell(flexcpForeColor, i, eOrderCol) = vbGrayText
                End If
                'set X in order cancel column
                .TextMatrix(i, eOrderColX) = "X"
                .Cell(flexcpFontBold, i, eOrderColX) = True
                .Cell(flexcpBackColor, i, eOrderColX) = .Cell(flexcpBackColor, i, eOrderCol)
                If m.tbOrders(6, aIdx(j)) = eTT_OrderStatus_CancelPending Then
                    .Cell(flexcpForeColor, i, eOrderColX) = vbGrayText
                ElseIf .Cell(flexcpBackColor, i, eOrderColX) = 0 Then
                    .Cell(flexcpForeColor, i, eOrderColX) = .ForeColorFixed
                Else
                    .Cell(flexcpForeColor, i, eOrderColX) = vbBlack
                End If
                'get next price in table
                j = j + 1
                If dPrice = m.tbOrders(4, aIdx(j)) Then
                    i = i + 1       'fix for orders disappearing when AutoExits exists at same price
                End If
                dPrice = m.tbOrders(4, aIdx(j))
            End If
            If j >= aIdx.Size Then Exit For
        Next
        .Redraw = flexRDBuffered
    End With

    Set aIdx = Nothing
        
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.OrderTblToGrid"
    
End Sub

Private Function ShowControl(ByVal strCode$) As Boolean

    If strCode = "BA" Or strCode = "BB" Or strCode = "SA" Or strCode = "SB" Then
        If g.RealTime.Active And InStr(m.strOrdBarCtrls, strCode & ";1") <> 0 Then ShowControl = True
    ElseIf InStr(m.strOrdBarCtrls, strCode & ";1") <> 0 Then
        ShowControl = True
    End If

End Function

Private Sub OrderBarNarrow(ByVal bShowExchanges As Boolean)
On Error Resume Next            'called from resize

    Dim nTop&, nHeightLeft&, nHeightMiddle&, nHeightRight&
    Dim bShow As Boolean
    
    Dim aTemp As New cGdArray
    
    cmdBuyMarket.Left = 0
    cmdBuyBid.Left = 0
    cmdBuyAsk.Left = 0
    lblAccounts.Left = 0
    cboAccounts.Left = 0
    lblOrderType.Left = 0
    cboOrderType.Left = 0
    lblExchange.Left = 0
    cboExchanges.Left = 0
    
    lblTradePos.Left = cmdBuyMarket.Width + 80
    lblEquity.Left = lblTradePos.Left
    cmdClearQty.Left = lblTradePos.Left
    txtTradeQty.Left = cmdClearQty.Left + cmdClearQty.Width
    vscrQty.Left = txtTradeQty.Left + txtTradeQty.Width
    cmdQty1.Left = cmdClearQty.Left
    cmdQty2.Left = cmdQty1.Left + cmdQty1.Width
    cmdQty3.Left = cmdQty2.Left + cmdQty2.Width
    chkConfirmOrder.Left = cmdClearQty.Left
    
    cmdSellMarket.Left = lblTradePos.Left + lblTradePos.Width + 80
    cmdSellBid.Left = cmdSellMarket.Left
    cmdSellAsk.Left = cmdSellMarket.Left
    cmdReverse.Left = cmdSellMarket.Left
    cmdCancelAll.Left = cmdSellMarket.Left
    cmdBailout.Left = cmdSellMarket.Left
    
    If cmdBrokerConnect.Visible Then
        'do nothing
    Else
        fraOrderBtns.Width = cmdBuyMarket.Width * 2
    End If
    
    'left-most controls: [buy market][buy ask][buy bid] trading account [order type][exchange]
    nTop = 10
    bShow = ShowControl("BM")
    If bShow Then
        cmdBuyMarket.Top = nTop
        nTop = nTop + cmdBuyMarket.Height
        nHeightLeft = nTop
    End If
    cmdBuyMarket.Visible = bShow
    
    bShow = ShowControl("BA")
    If bShow Then
        cmdBuyAsk.Top = nTop
        nTop = nTop + cmdBuyAsk.Height
        nHeightLeft = nTop
    End If
    cmdBuyAsk.Visible = bShow
    
    bShow = ShowControl("BB")
    If bShow Then
        cmdBuyBid.Top = nTop
        nTop = nTop + cmdBuyBid.Height
        nHeightLeft = nTop
    End If
    cmdBuyBid.Visible = bShow
    
    If bShowExchanges Then
        lblExchange.Top = nTop + 15
        cboExchanges.Top = nTop + 200
        nTop = cboExchanges.Top + cboExchanges.Height + 75
        nHeightLeft = nTop
    End If
    
    lblAccounts.Top = nTop                  'these controls are not optional
    nTop = nTop + lblAccounts.Height - 60
    cboAccounts.Top = nTop
    nTop = nTop + cboAccounts.Height
    nHeightLeft = nTop
    
    If m.eDisplayStyle = eView_Ladder Then
        bShow = ShowControl("OT")
    Else
        bShow = False       'order type does not apply to depth of market view
    End If
    If bShow Then
        lblOrderType.Top = nTop
        nTop = nTop + lblOrderType.Height / 2 + 50
        cboOrderType.Top = nTop
        nTop = nTop + cboOrderType.Height
        nHeightLeft = nTop
    End If
    lblOrderType.Visible = bShow
    cboOrderType.Visible = bShow

    'middle controls: [position] quantity [preset quantities][exchanges] confirm
    nTop = 10
    
    If m.eDisplayStyle = eView_Ladder Then
        bShow = ShowControl("BRORD")
        If bShow Then
            vseBracketOrder.Left = lblTradePos.Left
            vseBracketOrder.Top = nTop
            nTop = nTop + vseBracketOrder.Height
            nHeightMiddle = nTop
        End If
        vseBracketOrder.Visible = bShow
    Else
        vseBracketOrder.Visible = False
    End If
    
    bShow = ShowControl("POS")
    If bShow Then
        lblTradePos.Top = nTop
        nTop = nTop + lblTradePos.Height
        nHeightMiddle = nTop
    End If
    lblTradePos.Visible = bShow
    
    bShow = ShowControl("OE")           '5617
    If bShow Then
        lblEquity.Top = nTop
        nTop = lblEquity.Top + lblEquity.Height
        nHeightMiddle = nTop
    End If
    lblEquity.Visible = bShow
    
    cmdClearQty.Top = nTop              'these controls are not optional
    txtTradeQty.Top = nTop
    vscrQty.Top = nTop
    nTop = nTop + vscrQty.Height
    nHeightMiddle = nHeightMiddle + vscrQty.Height
    
    bShow = ShowControl("QTY")
    If bShow Then
        cmdQty1.Top = nTop
        cmdQty2.Top = nTop
        cmdQty3.Top = nTop
        nTop = nTop + cmdQty3.Height
        nHeightMiddle = nTop
    End If
    cmdQty1.Visible = bShow
    cmdQty2.Visible = bShow
    cmdQty3.Visible = bShow
    
    If g.RealTime.Active Then
        nTop = nHeightMiddle + 15
    ElseIf nHeightLeft > nHeightMiddle Then
        nTop = nHeightLeft + 15
    Else
        nTop = nHeightMiddle + 15
    End If
    
    bShow = ShowControl("AJ")
    If bShow Then
        chkAutoJournal.Left = chkConfirmOrder.Left
        chkAutoJournal.Top = nTop
        nTop = nTop + chkAutoJournal.Height
    End If
    chkAutoJournal.Visible = bShow
        
    If g.RealTime.Active Then
        chkConfirmOrder.Top = nTop
        chkConfirmOrder.Visible = True
        nHeightMiddle = nHeightMiddle + chkConfirmOrder.Height
    Else
        chkConfirmOrder.Visible = False
        chkAutoJournal.Left = 0
        If chkAutoJournal.Visible Then nHeightLeft = nTop
    End If
    
    'right-most controls: [sell market][sell bid][sell ask][reverse][cancel all][flatten]
    nTop = 10
    
    bShow = ShowControl("SM")
    If bShow Then
        cmdSellMarket.Top = nTop
        nTop = nTop + cmdSellMarket.Height
        nHeightRight = nHeightRight + cmdSellMarket.Height
    End If
    cmdSellMarket.Visible = bShow

    bShow = ShowControl("SB")
    If bShow Then
        cmdSellBid.Top = nTop
        nTop = nTop + cmdSellBid.Height
        nHeightRight = nHeightRight + cmdSellBid.Height
    End If
    cmdSellBid.Visible = bShow

    bShow = ShowControl("SA")
    If bShow Then
        cmdSellAsk.Top = nTop
        nTop = nTop + cmdSellAsk.Height
        nHeightRight = nHeightRight + cmdSellAsk.Height
    End If
    cmdSellAsk.Visible = bShow

    bShow = ShowControl("RV")
    If bShow Then
        cmdReverse.Top = nTop
        nTop = nTop + cmdReverse.Height
        nHeightRight = nHeightRight + cmdReverse.Height
    End If
    cmdReverse.Visible = bShow

    bShow = ShowControl("CA")
    If bShow Then
        cmdCancelAll.Top = nTop
        nTop = nTop + cmdCancelAll.Height
        nHeightRight = nHeightRight + cmdCancelAll.Height
    End If
    cmdCancelAll.Visible = bShow

    bShow = ShowControl("FL")
    If bShow Then
        cmdBailout.Top = nTop
        nTop = nTop + cmdBailout.Height
        nHeightRight = nHeightRight + cmdBailout.Height
    End If
    cmdBailout.Visible = bShow
    
    'find which height is highest
    aTemp.Add nHeightLeft
    aTemp.Add nHeightMiddle
    aTemp.Add nHeightRight
    
    aTemp.Sort
    nHeightLeft = aTemp(2)
    aTemp.Size = 0
    
    'auto exit always at bottom below all other controls due to long label
    nTop = nHeightLeft
    
    If fraExitFavorites.Visible Then
        nTop = nTop + 90
        fraExitFavorites.Move -45, nTop
        nTop = nTop + fraExitFavorites.Height
        nHeightLeft = nTop
    End If
    
    bShow = ShowControl("AE")
    If bShow Then
        chkAutoExit.Move 0, nTop, cboAccounts.Width - 200
        lblAutoExit.Move chkAutoExit.Left + chkAutoExit.Width, nTop + 20, fraOrderBtns.Width, chkAutoExit.Height * 2
        nHeightLeft = nHeightLeft + lblAutoExit.Height
        nTop = nTop + chkAutoExit.Height
    End If
    chkAutoExit.Visible = bShow
    lblAutoExit.Visible = bShow

    If ShowRithmic(TradeAccountID) Then
        With fraRithmic
            .Move 0, nTop + 105, cmdCancelAll.Left + cmdCancelAll.Width
        End With
        With picPbo
            .Move fraRithmic.Width - .Width
        End With
        fraRithmic.Visible = True
        fraRithmicSmall.Visible = False
        nHeightLeft = nHeightLeft + fraRithmic.Height - 105
    Else
        fraRithmic.Visible = False
        fraRithmicSmall.Visible = False
    End If
    
    vseOrderBar.Height = nHeightLeft
    
    Set aTemp = Nothing

End Sub

Private Sub OrderBarWide(ByVal bShowExchanges As Boolean)
On Error Resume Next            'called from resize

    Dim nLeft&, nTop&, nHeight&
    Dim bShow As Boolean
    
    Dim aHeights As New cGdArray
        
    'column 1
    lblAccounts.Left = 0
    cboAccounts.Left = 0
    lblOrderType.Left = 0
    cboOrderType.Left = 0
    lblExchange.Left = 0
    cboExchanges.Left = 0
    nLeft = lblAccounts.Width + 80
    
    If ShowControl("BM") Or ShowControl("BB") Or ShowControl("BA") Then
        cmdBuyMarket.Left = nLeft
        cmdBuyBid.Left = nLeft
        cmdBuyAsk.Left = nLeft
        nLeft = nLeft + cmdBuyMarket.Width + 80
    End If
    
    lblTradePos.Left = nLeft
    lblEquity.Left = nLeft
    cmdClearQty.Left = nLeft
    txtTradeQty.Left = nLeft + cmdClearQty.Width
    vscrQty.Left = txtTradeQty.Left + txtTradeQty.Width
    cmdQty1.Left = nLeft
    cmdQty2.Left = nLeft + cmdQty1.Width
    cmdQty3.Left = cmdQty2.Left + cmdQty2.Width
    nLeft = lblTradePos.Left + lblTradePos.Width + 80
       
    If ShowControl("SM") Or ShowControl("SB") Or ShowControl("SA") Then
        cmdSellMarket.Left = nLeft
        cmdSellBid.Left = nLeft
        cmdSellAsk.Left = nLeft
        nLeft = cmdSellMarket.Left + cmdSellMarket.Width
    End If
    
    cmdReverse.Left = nLeft
    cmdCancelAll.Left = nLeft
    cmdBailout.Left = nLeft
    chkConfirmOrder.Left = nLeft
        
    'column 1: trading account [order type][exchanges]
    'column 2: [buy market][buy ask][buy bid]
    'column 3: [position] quantity [preset quantities]
    'column 4: [sell market] [sell bid] [sell ask]
    'column 5: [reverse][cancel all][flatten] confirm

    aHeights.Create eGDARRAY_Longs, 5, 0

    'column 1
    If bShowExchanges Then
        lblExchange.Top = nTop + 15
        cboExchanges.Top = nTop + 200
        nTop = cboExchanges.Top + cboExchanges.Height + 60
        aHeights(0) = aHeights(0) + lblExchange.Height + cboExchanges.Height
    Else
        nTop = 10
    End If
    
    lblAccounts.Top = nTop
    cboAccounts.Top = lblAccounts.Top + lblAccounts.Height - 60
    nTop = cboAccounts.Top + cboAccounts.Height - 15
    aHeights(0) = aHeights(0) + lblAccounts.Height + cboAccounts.Height
    
    If m.eDisplayStyle = eView_Ladder Then
        bShow = ShowControl("OT")
    Else
        bShow = False   'order type does not apply to depth of market view
    End If
    If bShow Then
        lblOrderType.Top = nTop
        cboOrderType.Top = nTop + lblOrderType.Height / 2 + 50
        nTop = cboOrderType.Top + cboOrderType.Height
        aHeights(0) = aHeights(0) + lblOrderType.Height + cboOrderType.Height
    End If
    lblOrderType.Visible = bShow
    cboOrderType.Visible = bShow
    
    'column2
    nTop = 10
    bShow = ShowControl("BM")
    If bShow Then
        cmdBuyMarket.Top = nTop
        nTop = nTop + cmdBuyMarket.Height
        aHeights(1) = aHeights(1) + cmdBuyMarket.Height
    End If
    cmdBuyMarket.Visible = bShow

    bShow = ShowControl("BA")
    If bShow Then
        cmdBuyAsk.Top = nTop
        nTop = nTop + cmdBuyAsk.Height
        aHeights(1) = aHeights(1) + cmdBuyAsk.Height
    End If
    cmdBuyAsk.Visible = bShow

    bShow = ShowControl("BB")
    If bShow Then
        cmdBuyBid.Top = nTop
        nTop = nTop + cmdBuyBid.Height
        aHeights(1) = aHeights(1) + cmdBuyBid.Height
    End If
    cmdBuyBid.Visible = bShow
    
    'column 3 (has controls that are not optional)
    nTop = 10
    
    If m.eDisplayStyle = eView_Ladder Then
        bShow = ShowControl("BRORD")
        If bShow Then
            vseBracketOrder.Left = lblTradePos.Left
            vseBracketOrder.Top = nTop
            nTop = nTop + vseBracketOrder.Height
            aHeights(2) = aHeights(2) + vseBracketOrder.Height
        End If
        vseBracketOrder.Visible = bShow
    Else
        vseBracketOrder.Visible = False
    End If
    
    bShow = ShowControl("POS")
    If bShow Then
        lblTradePos.Top = nTop
        nTop = nTop + lblTradePos.Height
        aHeights(2) = aHeights(2) + lblTradePos.Height
    End If
    lblTradePos.Visible = bShow
    
    bShow = ShowControl("OE")               '5617
    If bShow Then
        lblEquity.Top = nTop
        nTop = nTop + lblEquity.Height
        aHeights(2) = aHeights(2) + lblEquity.Height
    End If
    lblEquity.Visible = bShow
    
    cmdClearQty.Top = nTop             'these controls are not optional
    txtTradeQty.Top = nTop
    vscrQty.Top = nTop
    nTop = nTop + vscrQty.Height
    aHeights(2) = aHeights(2) + cmdClearQty.Height
    
    bShow = ShowControl("QTY")
    If bShow Then
        cmdQty1.Top = nTop
        cmdQty2.Top = nTop
        cmdQty3.Top = nTop
        nTop = nTop + cmdQty3.Height
        aHeights(2) = aHeights(2) + cmdQty3.Height
    End If
    cmdQty1.Visible = bShow
    cmdQty2.Visible = bShow
    cmdQty3.Visible = bShow
        
    'column 4
    nTop = 10
    bShow = ShowControl("SM")
    If bShow Then
        cmdSellMarket.Top = nTop
        nTop = nTop + cmdSellMarket.Height
        aHeights(3) = aHeights(3) + cmdSellMarket.Height
    End If
    cmdSellMarket.Visible = bShow
    
    bShow = ShowControl("SB")
    If bShow Then
        cmdSellBid.Top = nTop
        nTop = nTop + cmdSellBid.Height
        aHeights(3) = aHeights(3) + cmdSellBid.Height
    End If
    cmdSellBid.Visible = bShow
    
    bShow = ShowControl("SA")
    If bShow Then
        cmdSellAsk.Top = nTop
        nTop = nTop + cmdSellAsk.Height
        aHeights(3) = aHeights(3) + cmdSellAsk.Height
    End If
    cmdSellAsk.Visible = bShow
        
    'column 5
    nTop = 10
    bShow = ShowControl("RV")
    If bShow Then
        cmdReverse.Top = nTop
        nTop = nTop + cmdReverse.Height
        aHeights(4) = aHeights(4) + cmdReverse.Height
    End If
    cmdReverse.Visible = bShow

    bShow = ShowControl("CA")
    If bShow Then
        cmdCancelAll.Top = nTop
        nTop = nTop + cmdCancelAll.Height
        aHeights(4) = aHeights(4) + cmdCancelAll.Height
    End If
    cmdCancelAll.Visible = bShow

    bShow = ShowControl("FL")
    If bShow Then
        cmdBailout.Top = nTop
        nTop = nTop + cmdBailout.Height
        aHeights(4) = aHeights(4) + cmdBailout.Height
    End If
    cmdBailout.Visible = bShow
    
    'find which column height is highest
    aHeights.Sort
    nHeight = aHeights(4)
    nTop = nHeight
    
    If g.RealTime.Active Then
        chkConfirmOrder.Left = 0
        chkConfirmOrder.Top = nTop
        chkConfirmOrder.Visible = True
        nHeight = nHeight + chkConfirmOrder.Height
    Else
        chkConfirmOrder.Visible = False
    End If
    
    nTop = nHeight
    bShow = ShowControl("AJ")
    If bShow Then
        chkAutoJournal.Left = 0
        chkAutoJournal.Top = nTop
        nHeight = nHeight + chkAutoJournal.Height
    End If
    chkAutoJournal.Visible = bShow
    
    'auto exit always at bottom below all other controls due to long label
    nTop = nHeight
    bShow = ShowControl("AE")
    If bShow Then
    
        If fraExitFavorites.Visible Then
            fraExitFavorites.Move 0, nTop
            nTop = nTop + fraExitFavorites.Height
            nHeight = nHeight + fraExitFavorites.Height
        End If
    
        chkAutoExit.Move 0, nTop, cboAccounts.Width - 200
        lblAutoExit.Move chkAutoExit.Left + chkAutoExit.Width, nTop + 20, fraOrderBtns.Width, chkAutoExit.Height
        nHeight = nHeight + lblAutoExit.Height
        nTop = nTop + chkAutoExit.Height
    End If
    chkAutoExit.Visible = bShow
    lblAutoExit.Visible = bShow

    If ShowRithmic(TradeAccountID) Then
        With fraRithmic
            .Move 0, nTop + 105, fraOrderBtns.Width
        End With
        With picPbo
            .Move fraRithmic.Width - .Width
        End With
        fraRithmic.Visible = True
        fraRithmicSmall.Visible = False
        nHeight = nHeight + fraRithmic.Height + 120
    Else
        fraRithmic.Visible = False
        fraRithmicSmall.Visible = False
    End If
    
    vseOrderBar.Height = nHeight
    
    Set aHeights = Nothing

End Sub

Private Sub DisplayAcctBar()
On Error GoTo ErrSection:
    
    Static bInProgress As Boolean
    
    If m.nShowAccountBar = 0 Or bInProgress Then
        Exit Sub
    End If
        
    bInProgress = True
    
    UpdateAccountBar fgAccountBar, m.strPos, m.strPosQty, m.strOpenEq, m.strAvgEntry, _
            m.strSessionPL, m.strSessionQty, TradeAccountID
        
    bInProgress = False
    Exit Sub

ErrSection:
    bInProgress = False
    RaiseError "frmTickDistribution.DisplayAcctBar"
    
End Sub

Public Property Get FloodMktDepth() As eBidAskColorMode
On Error Resume Next

    FloodMktDepth = m.eFloodMktDepth
    
End Property

Public Property Let FloodMktDepth(ByVal eFlood As eBidAskColorMode)
On Error Resume Next

    m.eFloodMktDepth = eFlood

End Property

Public Property Get ShowAccountBar() As Long
On Error Resume Next

    If BrokerViewMode Then
        ShowAccountBar = m.nShowAcctBarSave
    Else
        ShowAccountBar = m.nShowAccountBar
    End If
    
End Property

Public Property Let ShowAccountBar(ByVal nShow&)
On Error Resume Next

    If BrokerViewMode Then
        m.nShowAcctBarSave = nShow
    Else
        m.nShowAccountBar = nShow
    End If
    
End Property

Public Property Get BidTextColor() As Long
On Error Resume Next

    BidTextColor = m.nBidTextColor

End Property

Public Property Let BidTextColor(ByVal nColor&)
On Error Resume Next

    m.nBidTextColor = nColor
    
End Property

Public Property Get AskTextColor() As Long
On Error Resume Next

    AskTextColor = m.nAskTextColor
    
End Property

Public Property Let AskTextColor(ByVal nColor&)
On Error Resume Next

    m.nAskTextColor = nColor
    
End Property

Public Property Get FixedPriceColor() As Long
On Error Resume Next

    FixedPriceColor = m.nFixedPriceColor
    
End Property

Public Property Let FixedPriceColor(ByVal nColor&)
On Error Resume Next

    m.nFixedPriceColor = nColor
    
End Property

Public Property Get DefaultBidColor() As Long
    DefaultBidColor = kBidColor
End Property

Public Property Get DefaultAskColor() As Long
    DefaultAskColor = kAskColor
End Property

Private Sub FixAcctBarHeader()
On Error GoTo ErrSection:

    Dim i&, strText$, strSec$
    Dim bFix As Boolean
    
    strSec = SecType
    For i = 0 To m.aABarColHeader.Size - 1
        strText = m.aABarColHeader(i)
        If InStr(strText, "#") Then
            If strSec = "S" And InStr(strText, "Contracts") Then
                m.aABarColHeader(i) = "# Shares"
                bFix = True
                Exit For
            ElseIf strSec <> "S" And InStr(strText, "Shares") Then
                m.aABarColHeader(i) = "# Contracts"
                bFix = True
                Exit For
            End If
        End If
    Next
    
    If bFix Then
        ResetGridBar fgAccountBar, m.aABarColHeader, m.nABarSumColWidth
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.FixAcctBarHeader"
    
End Sub

Private Sub CheckLadderDOM()
On Error GoTo ErrSection:

    'this subroutine check depth of market data is okay to display on ladder
    Dim dBidAskTime#, dTradeTime#
    Dim aBestBids As cGdArray
    Dim aBestAsks As cGdArray

    If m.oBidAskDepth Is Nothing Then
        m.bDepthOfMarketBad = False     'reset
        Exit Sub
    End If
        
    dBidAskTime = m.oBidAskDepth.NewestData
    dTradeTime = m.TickBars(eBARS_DateTime, m.TickBars.Size - 1)
    
    If dBidAskTime = 0 Then Exit Sub     'first time through
    
    If m.bDepthOfMarketBad Then
        If dBidAskTime > dTradeTime Then m.bDepthOfMarketBad = False
    Else
        Set aBestBids = m.oBidAskDepth.BestBids
        Set aBestAsks = m.oBidAskDepth.BestAsks
        'don't show if last trade time > last bid/ask time + 2 seconds
        'AND (last trade > lowest ask or last trade < highest bid)
        If Not aBestBids Is Nothing And Not aBestAsks Is Nothing Then
            If dTradeTime > dBidAskTime + 0.000025 And (m.dLastPrice > aBestAsks(0) Or m.dLastPrice < aBestBids(0)) Then
''StatusMsg "time diff = " & Str(dTradeTime - dBidAskTime)
                m.bDepthOfMarketBad = True
                ClearBidAskCells
            End If
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.CheckLadderDOM"
    
End Sub

Public Property Get ShowVolMin() As Long
    ShowVolMin = m.nShowVolMin
End Property

Public Property Let ShowVolMin(ByVal nVol&)
    m.nShowVolMin = nVol
End Property

Public Property Get ShowVolMax() As Long
    ShowVolMax = m.nShowVolMax
End Property

Public Property Let ShowVolMax(ByVal nVol&)
    m.nShowVolMax = nVol
End Property

Private Sub CheckAutoCenter()
On Error Resume Next

    Dim nPercent25&, nLimitTop&, nLimitBottom&
        
    If tbToolbar.Tools("ID_CenterPrice").State = ssChecked Then
        With fgTickDistribution
            If .MouseCol < 0 Or .MouseRow < 0 Then
                nPercent25 = (.BottomRow - .TopRow) * 0.25
                nLimitTop = .TopRow + nPercent25
                nLimitBottom = .BottomRow - nPercent25
                
                If m.nLastPriceRow < nLimitTop Or m.nLastPriceRow > nLimitBottom Then
                    CenterLadderOnCurrPrice
                End If
            End If
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAutoExitCaptions
'' Description: Set text for auto exit check box & label depending on where order bar is
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAutoExitCaptions(ByVal strCaption$)
On Error GoTo ErrSection:

    chkAutoExit.Caption = "&Auto Exit:"
    lblAutoExit.Caption = strCaption
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.SetAutoExitCaptions"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAutoExit
'' Description: Attempt to select the auto exit in the combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAutoExit()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAutoExitName As String       ' Auto exit name to select
    
    If m.frmBroker Is Nothing Then
        m.bSettingAutoExit = True
        
        strAutoExitName = g.OrderStrategies.ExitForAccountAndSymbol(TradeAccountID, m.nSymID)
        If Len(strAutoExitName) > 0 Then
            SetAutoExitCaptions strAutoExitName
            chkAutoExit.Value = vbChecked
        
            ExitCtrlAppearance Me, Nothing, strAutoExitName
        Else
            SetAutoExitCaptions "None"
            chkAutoExit.Value = vbUnchecked
        
            'JM 03-30-2011: the autoexit checkbox & label are temporarily disabled when user
            '   clicks an exit favorite button, do not need to do this right now
            If chkAutoExit.Enabled Then ExitCtrlAppearance Me, Nothing, ""
        End If
        
        m.bSettingAutoExit = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.SetAutoExit"
    
End Sub

Private Function LoadOrderFromGrid(ByVal iRow&, ByVal iCol&) As cPtOrder
On Error GoTo ErrSection:

    Dim dPrice#, strType$, i&, nQty&
    
    Dim nReturn As eGDEditOrderReturn
    Dim eOrderType As eTT_OrderType
    
    Dim PtOrder As New cPtOrder     'use for loading order via orderID
    
    Dim OrderInTree As cPtOrder     'use to walk through multiple orders at same price tree
    Dim MultiOrders As cGdTree
    Dim strOrderID As String            ' Order ID from the grid
    
    Dim bBuy As Boolean
    
    If m.frmBroker Is Nothing Then
        If m.eDisplayStyle = eView_Ladder Then
            If m.nOrderColumns = 0 Then
                'only one order column shown, load order using order ID saved to hidden column
                With fgTickDistribution
                    If PtOrder.Load(ValOfText(.TextMatrix(iRow, eTDCols_HasOrder))) Then
                        If PtOrder.OrderID > 0 Then Set LoadOrderFromGrid = PtOrder
                    End If
                End With
                
                Exit Function
            End If
            
            dPrice = ValOfColPrice(iRow)
        Else
            With fgOrdersInfo
                If PtOrder.Load(ValOfText(.TextMatrix(iRow, 2))) Then
                    If PtOrder.OrderID > 0 Then Set LoadOrderFromGrid = PtOrder
                End If
            End With
            
            Exit Function
        End If
        
        Set MultiOrders = g.Broker.PrimaryOrdersForSymbol(TradeAccountID, m.nSymID, 0&, dPrice)
        
        If MultiOrders Is Nothing Then
            If PtOrder.Load(ValOfText(fgTickDistribution.TextMatrix(iRow, eTDCols_HasOrder))) Then
                If PtOrder.OrderID > 0 Then Set LoadOrderFromGrid = PtOrder
            End If
        ElseIf MultiOrders.Count > 0 Then
            With fgTickDistribution
                If iCol = eTDCols_OrderAskX Or iCol = eTDCols_OrderAsk Or iCol = eTDCols_AskSize Then
                    OrderInfoFromGridStr .TextMatrix(iRow, eTDCols_OrderAsk), bBuy, nQty, strType
                Else
                    OrderInfoFromGridStr .TextMatrix(iRow, eTDCols_OrderBid), bBuy, nQty, strType
                End If
            End With
            For i = 1 To MultiOrders.Count
                Set OrderInTree = MultiOrders(i)
                eOrderType = OrderInTree.OrderType
                If OrderInTree.OrderID > 0 Then
                    If OrderInTree.Buy = bBuy And OrderInTree.Quantity = nQty Then
                        If strType = "Limit" Then
                            If eOrderType = eTT_OrderType_Limit Or eOrderType = eTT_OrderType_LimitCloseOnly Then
                                Set LoadOrderFromGrid = OrderInTree
                                Exit For
                            End If
                        ElseIf strType = "Stop" Then
                            If eOrderType = eTT_OrderType_Stop Or _
                               eOrderType = eTT_OrderType_StopCloseOnly Or _
                               eOrderType = eTT_OrderType_StopWithLimit Or _
                               eOrderType = eTT_OrderType_StopWithLimit Or _
                               eOrderType = eTT_OrderType_StopWithLimitCloseOnly Then
                               Set LoadOrderFromGrid = OrderInTree
                               Exit For
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf PtOrder.Load(ValOfText(fgTickDistribution.TextMatrix(iRow, eTDCols_HasOrder))) Then
            If PtOrder.OrderID > 0 Then Set LoadOrderFromGrid = PtOrder
        End If
    Else
        Set MultiOrders = GetWorkingOrders
        strOrderID = fgTickDistribution.TextMatrix(iRow, eTDCols_HasOrder)
        If MultiOrders.Exists(strOrderID) Then
            Set LoadOrderFromGrid = MultiOrders(strOrderID)
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTickDistribution.LoadOrderFromGrid"

End Function

Private Sub OrderInfoFromGridStr(ByVal strText$, bBuy As Boolean, nQty&, strOrdType$)
On Error GoTo ErrSection:

    'initialize return variables
    bBuy = False
    nQty = 0
    strOrdType = ""
    
    If Parse(strText, " ", 1) = "Buy" Then bBuy = True
    nQty = Val(Parse(strText, " ", 2))
    strOrdType = Parse(strText, " ", 3)
    

    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.OrderInfoFromGridStr"
    
End Sub

Private Function ValOfColPrice(ByVal iRow&) As Double
On Error GoTo ErrSection:

    Dim iPos&
    
    With fgTickDistribution
        iPos = iRow - .FixedRows
        If iPos >= 0 And iPos < m.Data.NumRecords Then
            ValOfColPrice = m.Data(eFld_Price, iPos)
        End If
    End With
    
    Exit Function

ErrSection:
    RaiseError "frmTickDistribution.ValOfColPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInformation
'' Description: Get the contract information (if applicable) for given symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetContractInformation()
On Error GoTo ErrSection:

    If (m.nTradeAcctID > 0) And (m.nSymID > 0) Then
        g.Broker.GetContractInfo g.Broker.AccountTypeForID(m.nTradeAcctID), m.strSym, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.GetContractInformation"
    
End Sub

Public Property Get ShowOpenEntries() As Long
    ShowOpenEntries = m.nShowOpenEntries
End Property

Public Property Let ShowOpenEntries(ByVal nShow&)
    m.nShowOpenEntries = nShow
End Property

Public Property Get ShowAvgEntry() As Long
    ShowAvgEntry = m.nShowAvgEntry
End Property

Public Property Let ShowAvgEntry(ByVal nShow&)
    m.nShowAvgEntry = nShow
End Property

Public Property Get AvgEntryColor() As Long
    AvgEntryColor = m.nAvgEntryColor
End Property

Public Property Let AvgEntryColor(ByVal nColor&)
    m.nAvgEntryColor = nColor
End Property

Private Function AvgEntryGridRow() As Long

' TLB 10/18/2014: due to "overflow" errors being raised on client machines,
' let's just ignore errors for now from this routine (until we can find and fix it).
If IsIDE Then
    On Error GoTo ErrSection:
Else
    On Error Resume Next
End If

    Dim i&, j&, d#
    Dim iNewAvgEntryRow&, strColPrice$
    
    If m.eDisplayStyle <> eView_Ladder Then Exit Function
    If m.dMinMove = 0 Then Exit Function

    iNewAvgEntryRow = 0
    
    If m.dAvgEntry > 0 And m.dLastPrice <> kNullData Then ' TLB: need to check m.dLastPrice on a Saturday
        d = (m.dLastPrice - m.dAvgEntry) / m.dMinMove
        i = m.nLastPriceRow + d
        With fgTickDistribution
            If i >= .FixedRows And i < .Rows Then
                strColPrice = m.TickBars.PriceDisplay(ValOfColPrice(i))     '4458
                If strColPrice = m.strAvgEntry Then iNewAvgEntryRow = i
                'check either side in case off by one due to rounding/truncation (causes Avg grid row to bounce between 2 prices)
                If iNewAvgEntryRow = 0 Then
                    If i - 1 >= .FixedRows Then
                        If strColPrice = m.strAvgEntry Then iNewAvgEntryRow = i - 1
                    End If
                    If iNewAvgEntryRow = 0 Then
                        If i + 1 < .Rows Then
                            If strColPrice = m.strAvgEntry Then iNewAvgEntryRow = i + 1
                        End If
                    End If
                End If
            End If
        End With
    End If
    
    If m.nPrevAvgEntryRow > 0 And m.nPrevAvgEntryRow <> i Then
        'clear out prev colored row
        With fgTickDistribution
            j = .FixedRows + 1
            If m.nPrevAvgEntryRow > .FixedRows And m.nPrevAvgEntryRow < .Rows And j < .Rows Then
                For i = 0 To .Cols - 1
                    If i <> eTDCols_Price Then
                        .Cell(flexcpBackColor, m.nPrevAvgEntryRow, i) = .Cell(flexcpBackColor, j, i)
                    End If
                Next
            End If
        End With
    End If
        
    If m.dAvgEntry <> 0 And iNewAvgEntryRow = 0 Then
        'do nothing (this was bug that caused multiple average entry highlight rows)
        'bug should be fixed but check is here just in case
        With fgTickDistribution
            i = m.nLastPriceRow + d
            If i >= .FixedRows And i < .Rows Then
                DebugLog "Multi AvgEntry Bug: actual(" & Str(m.dAvgEntry) & ") String(" & m.strAvgEntry & ") Actual Grid(" & _
                    fgTickDistribution.TextMatrix(i, eTDCols_Price) & ") Val of Grid(" & Str(ValOfColPrice(i)) & ")"
            End If
        End With
    Else
        m.nPrevAvgEntryRow = iNewAvgEntryRow
        AvgEntryGridRow = iNewAvgEntryRow
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.AvgEntryGridRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancellingAll
'' Description: Are we currently cancelling all orders?
'' Inputs:      None
'' Returns:     True if cancelling all orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancellingAll() As Boolean
On Error GoTo ErrSection:

    CancellingAll = g.FlattenQueue.IsGettingFlattened(AccountNumber, m.strSym, 0&, eGDFlattenQueueOperation_CancelAll)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.CancellingAll"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Flattening
'' Description: Are we currently flattening the position?
'' Inputs:      None
'' Returns:     True if flattening position, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Flattening() As Boolean
On Error GoTo ErrSection:

    Flattening = g.FlattenQueue.IsGettingFlattened(AccountNumber, m.strSym, 0&, eGDFlattenQueueOperation_Flatten)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.Flattening"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reversing
'' Description: Are we currently reversing the position?
'' Inputs:      None
'' Returns:     True if reversing position, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Reversing() As Boolean
On Error GoTo ErrSection:

    Reversing = g.FlattenQueue.IsGettingFlattened(AccountNumber, m.strSym, 0&, eGDFlattenQueueOperation_Reverse)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.Reversing"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelAll
'' Description: Cancel all working orders for this symbol and account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelAll()
On Error GoTo ErrSection:

    cmdCancelAll.Enabled = False

    g.Broker.BrokerDebug Broker, "Cancelling All Orders for " & m.strSym & " in account " & AccountNumber & " from the Price Ladder", True
    CancelAllForSymbol TradeAccountID, SymbolID, 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.CancelAll"
    
End Sub

Public Property Get OrdBarCtrls() As String
On Error GoTo ErrSection:

    OrdBarCtrls = m.strOrdBarCtrls

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.OrdBarCtrls.Get"

End Property

Public Property Let OrdBarCtrls(ByVal strSettings$)
On Error GoTo ErrSection:

    m.strOrdBarCtrls = strSettings
    FixOrderBarCtrlString

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmTickDistribution.OrdBarCtrls.Let"

End Property

Private Function NextOrdBarCtrl(ByVal iIndex As Long) As Control
On Error GoTo ErrSection:

    Dim aControls As New cGdArray
    
    aControls.SplitFields m.strOrdBarCtrls, "|"
    
    Set NextOrdBarCtrl = OrdBarCtrlFromCode(aControls, Me, iIndex)
    

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTickDistribution.NextOrdBarCtrl"

End Function

Private Sub PositionOrdBarCtrl(PrevCtrl As Control, CurrCtrl As Control)
On Error GoTo ErrSection:

    Dim nTop&
    
    If PrevCtrl Is Nothing Then
        nTop = cboAccounts.Top + cboAccounts.Height + 20
    ElseIf PrevCtrl.Name = "lblOrderType" Then
        nTop = PrevCtrl.Top + PrevCtrl.Height + 50
    Else
        nTop = PrevCtrl.Top + PrevCtrl.Height + 20
    End If
    
    If Not CurrCtrl Is Nothing Then
        CurrCtrl.Top = nTop
        If CurrCtrl.Name = "cmdClearQty" Then
            txtTradeQty.Top = nTop
            vscrQty.Top = nTop
            cmdClearQty.Left = lblAccounts.Left
            txtTradeQty.Left = cmdClearQty.Left + cmdClearQty.Width
            vscrQty.Left = txtTradeQty.Left + txtTradeQty.Width
        ElseIf CurrCtrl.Name = "cmdQty1" Then
            cmdQty2.Top = nTop
            cmdQty3.Top = nTop
            cmdQty1.Left = lblAccounts.Left
            cmdQty2.Left = cmdQty1.Left + cmdQty1.Width
            cmdQty3.Left = cmdQty2.Left + cmdQty2.Width
        ElseIf CurrCtrl.Name = "lblOrderType" Then
            lblOrderType.Top = CurrCtrl.Top + 30
            lblOrderType.Left = lblAccounts.Left
            cboOrderType.Move txtTradeQty.Left + 100, nTop, cmdQty2.Width + cmdQty3.Width
        ElseIf CurrCtrl.Name = "chkAutoExit" And fraExitFavorites.Visible Then
            fraExitFavorites.Top = nTop
            fraExitFavorites.Left = lblAccounts.Left
            chkAutoExit.Left = lblAccounts.Left
            chkAutoExit.Top = nTop + fraExitFavorites.Height
            ExitCtrlAppearance Me, Nothing, "", True
        Else
            CurrCtrl.Left = lblAccounts.Left
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.PositionOrdBarCtrl"
    
End Sub

Private Sub OrdBarOnRight(ByVal bShowExchanges As Boolean)
On Error Resume Next            'called from resize

    Dim nSummBarHt&, nLeft&, nTop&, nY&, iCtlIndex&
    Dim Ctrl1 As Control, Ctrl2 As Control
    Dim bConnected As Boolean
    
    fraOrderBtns.Visible = True
    
    cmdBuyMarket.Width = 1335
    cmdSellMarket.Width = 1335
    lblAutoExit.Width = 1335
    lblAutoExit.Height = kAutoExitOnRight

    nLeft = 60
    fraOrderBtns.Width = cmdBuyMarket.Width + 120
        
    If m.eDisplayStyle = eView_Detail And m.nShowSummaryBar = 1 Then
        nSummBarHt = pbBid.Height
    End If
        
    If m.nShowQuoteBar = 1 And m.nShowAccountBar = 1 Then
        nY = Me.ScaleHeight - fgQuoteBar.Height - fgAccountBar.Height - nSummBarHt
        nTop = fgQuoteBar.Top + fgQuoteBar.Height
    ElseIf m.nShowQuoteBar = 1 And m.nShowAccountBar = 0 Then
        nY = Me.ScaleHeight - fgQuoteBar.Height - nSummBarHt
        nTop = fgQuoteBar.Top + fgQuoteBar.Height
    ElseIf m.nShowQuoteBar = 0 And m.nShowAccountBar = 1 Then
        nY = Me.ScaleHeight - fgAccountBar.Height - nSummBarHt
        nTop = 0
    Else
        nY = Me.ScaleHeight - nSummBarHt
        nTop = 0
    End If
               
    With fgTickDistribution
        .Move 0, nTop, Me.ScaleWidth - fraOrderBtns.Width, nY
        vseOrderBar.Move .Left + .Width, nTop, fraOrderBtns.Width, nY
        fraOrderBtns.Move 0, 0, fraOrderBtns.Width, nY
    End With
        
    If bShowExchanges Then
        lblExchange.Move nLeft, 60, cmdBuyMarket.Width
        cboExchanges.Move nLeft, lblExchange.Top + lblExchange.Height, cmdBuyMarket.Width
        
        lblAccounts.Move nLeft, cboExchanges.Top + cboExchanges.Height, cmdBuyMarket.Width
        cboAccounts.Move nLeft, lblAccounts.Top + lblAccounts.Height, cmdBuyMarket.Width
    Else
        lblAccounts.Move nLeft, 60, cmdBuyMarket.Width
        cboAccounts.Move nLeft, lblAccounts.Top + lblAccounts.Height, cmdBuyMarket.Width
    End If
    
    If ConnectionStatus = eGDConnectionStatus_Connected Then
        bConnected = True
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JM 10-04-2011: With 6.1 all accounts are broker-style. Must be connected in order to trade.
'   Position controls as usual if connected else position controls only once to put the broker
'   disconnect label in the right place.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If bConnected Or lblBrokerDisconnect.Visible = False Then
        Set Ctrl2 = NextOrdBarCtrl(iCtlIndex)
        While Not Ctrl2 Is Nothing
            If m.eDisplayStyle = eView_Detail Then
                If Ctrl2.Name = "vseBracketOrder" Then Ctrl2.Visible = False
            End If
            
            If Ctrl2.Name = "chkConfirmOrder" Then Ctrl2.Visible = g.RealTime.Active
            
            If Ctrl2.Visible Then
                If Ctrl2.Name = "lblOrderType" And m.eDisplayStyle = eView_Detail Then
                    Ctrl2.Visible = False               'order type does not apply to market depth view
                    cboOrderType.Visible = False
                Else
                    PositionOrdBarCtrl Ctrl1, Ctrl2
                    Set Ctrl1 = Ctrl2
                    nTop = Ctrl2.Top + Ctrl2.Height + 20
                End If
            End If
            
            If Not Ctrl1 Is Nothing Then
                If Ctrl1.Name = "chkAutoExit" Then
                    Set Ctrl2 = lblAutoExit
                Else
                    iCtlIndex = iCtlIndex + 1
                    Set Ctrl2 = NextOrdBarCtrl(iCtlIndex)
                End If
            Else
                iCtlIndex = iCtlIndex + 1
                Set Ctrl2 = NextOrdBarCtrl(iCtlIndex)
            End If
        Wend
    End If
    
    If g.RealTime.Active Then
    'JM 09-21-2010 !IMPORTANT!
    'need to revisit this code since confirm order check box is no longer at the bottom
        nTop = nTop + chkConfirmOrder.Height
    End If
    
    If ShowRithmic(TradeAccountID) Then
        fraRithmicSmall.Move cmdCancelAll.Left, nTop + 105
        fraRithmic.Visible = False
        fraRithmicSmall.Visible = True
    Else
        fraRithmic.Visible = False
        fraRithmicSmall.Visible = False
    End If
    
    If m.eDisplayStyle = eView_Ladder Then
    With fgAccountBar
        If m.nShowAccountBar = 1 Then
            .Redraw = flexRDNone
            .Move 0, fgTickDistribution.Top + fgTickDistribution.Height, Me.ScaleWidth, .Height
            .Redraw = flexRDBuffered
        End If
    End With
    End If

End Sub

Private Function OrdBarVisible() As Boolean
On Error Resume Next            'to be called from form resize so do not raise error

    Dim bOrdBarVisible As Boolean

    If Left(m.strSym, 1) = "$" And Not IsForex(m.strSym) Then
        bOrdBarVisible = False
    ElseIf g.nReplaySession > 0 Or frmReplay.Visible Then
        If m.eOrderBarMode > eGDOrderBarMode_NotShown And SecurityType(m.strSym) <> "S" Then
            bOrdBarVisible = True
        Else
            bOrdBarVisible = False
        End If
    ElseIf m.bSessionCurrent And m.eOrderBarMode > eGDOrderBarMode_NotShown Then
        bOrdBarVisible = True
    Else
        bOrdBarVisible = False
    End If
            
    If bOrdBarVisible Then
        If m.nShowAccountBar = 1 Then
            fgAccountBar.Visible = True
        Else
            fgAccountBar.Visible = False
        End If
        
        If m.eDisplayStyle = eView_Detail Then
            fgOrdersInfo.Visible = True
        Else
            fgOrdersInfo.Visible = False
        End If
    Else
        fgAccountBar.Visible = False
        fgOrdersInfo.Visible = False
    End If
    
    vseOrderBar.Visible = bOrdBarVisible
    OrdBarVisible = bOrdBarVisible

End Function

Private Sub SaveSettings()
On Error GoTo ErrSection

    Dim strText$, i&
    
    SetIniFileProperty "GridFloodColor", m.nFloodColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "GridBarColor", m.nBarColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "GridUpColor", m.nUpColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "GridDownColor", m.nDownColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "TickLineColor", m.nTickLineColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "BidColor", m.nBidColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "AskColor", m.nAskColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "BidTextColor", m.nBidTextColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "AskTextColor", m.nAskTextColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "FixedPriceColor", m.nFixedPriceColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "GridFontSize", m.nFontSize, "Price Ladder", g.strIniFile
    SetIniFileProperty "GridFontName", m.strFont, "Price Ladder", g.strIniFile
    SetIniFileProperty "ShowVolumeBar", m.nShowVolBar, "Price Ladder", g.strIniFile
    SetIniFileProperty "ShowVolumeText", m.nShowVolText, "Price Ladder", g.strIniFile
    SetIniFileProperty "ShowTickLine", m.nShowTickLine, "Price Ladder", g.strIniFile
    SetIniFileProperty "ShowProfitLoss", m.nShowProfitLoss, "Price Ladder", g.strIniFile
    SetIniFileProperty "TickLineRightToLeft", m.nTickLineRL, "Price Ladder", g.strIniFile
    SetIniFileProperty "OrderColumns", m.nOrderColumns, "Price Ladder", g.strIniFile
    SetIniFileProperty "BlankRows", m.nBlankRows, "Price Ladder", g.strIniFile
    SetIniFileProperty "LadderVolStyle", m.eVolumeStyle, "Price Ladder", g.strIniFile
    SetIniFileProperty "OutlineColor", m.nOutlineColor, "Price Ladder", g.strIniFile
    
    If lblBrokerDisconnect.Visible Then Exit Sub        'don't want to save any order bar info
    
    SetIniFileProperty "OrdBarCtrls", m.strOrdBarCtrls, "Price Ladder", g.strIniFile
    If m.eDisplayStyle = eView_Ladder Then
        SetIniFileProperty "QuoteBar", m.nShowQuoteBar, "Price Ladder", g.strIniFile
        If BrokerViewMode Then
            SetIniFileProperty "AccountBar", m.nShowAcctBarSave, "Price Ladder", g.strIniFile
        Else
            SetIniFileProperty "AccountBar", m.nShowAccountBar, "Price Ladder", g.strIniFile
        End If
        SetIniFileProperty "OrderBar", m.eOrderBarMode, "Price Ladder", g.strIniFile
    Else
        SetIniFileProperty "QuoteBar", m.nShowQuoteBar, "Market Depth", g.strIniFile
        SetIniFileProperty "AccountBar", m.nShowAccountBar, "Market Depth", g.strIniFile
        SetIniFileProperty "OrderBar", m.eOrderBarMode, "Market Depth", g.strIniFile
    End If
    
    'depth of market bid/ask colors
    SetIniFileProperty "FirstColor", m.nFirstColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "SecondColor", m.nSecondColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "ThirdColor", m.nThirdColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "FourthColor", m.nFourthColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "FifthColor", m.nFifthColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "OtherColor", m.nOtherColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "InactiveColor", m.nInactiveColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "LargestSizeColor", m.nLargestSizeColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "DrawTriangle", m.nDrawTriangle, "Price Ladder", g.strIniFile
    SetIniFileProperty "BidAskUpColor", m.nBidAskUpColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "BidAskDownColor", m.nBidAskDownColor, "Price Ladder", g.strIniFile
    SetIniFileProperty "BidAskSummaryBar", m.nShowSummaryBar, "Price Ladder", g.strIniFile
    SetIniFileProperty "BidAskSummaryVert", m.nVertSummaryBar, "Price Ladder", g.strIniFile
    SetIniFileProperty "FloodMktDepth", m.eFloodMktDepth, "Price Ladder", g.strIniFile
    'volume min/max
    SetIniFileProperty "ShowVolMin", m.nShowVolMin, "Price Ladder", g.strIniFile
    SetIniFileProperty "ShowVolMax", m.nShowVolMax, "Price Ladder", g.strIniFile
    'average entry on ladder
    SetIniFileProperty "ShowOpenEntries", m.nShowOpenEntries, "Price Ladder", g.strIniFile
    SetIniFileProperty "ShowAvgEntry", m.nShowAvgEntry, "Price Ladder", g.strIniFile
    SetIniFileProperty "AvgEntryColor", m.nAvgEntryColor, "Price Ladder", g.strIniFile
    'highlight color
    SetIniFileProperty "HighlightPos", m.nHighlightPos, "Price Ladder", g.strIniFile
    SetIniFileProperty "HighlightEquity", m.nHighlightEquity, "Price Ladder", g.strIniFile
    
    SetIniFileProperty "BidAskSummaryHeight", pbBid.Height, "Price Ladder", g.strIniFile

    'bid/ask column order
    strText = m.aBidColHeader(0) & "," & m.aBidColHeader(1) & "," & m.aBidColHeader(2) & "," & m.aBidColHeader(3)
    SetIniFileProperty "BidColumns", strText, "Price Ladder", g.strIniFile
    strText = m.aAskColHeader(0) & "," & m.aAskColHeader(1) & "," & m.aAskColHeader(2) & "," & m.aAskColHeader(3)
    SetIniFileProperty "AskColumns", strText, "Price Ladder", g.strIniFile
    'quote bar column header
    strText = ""
    For i = 0 To m.aQBarColHeader.Size - 1
        If Len(strText) > 0 Then
            strText = strText & "|" & m.aQBarColHeader(i)
        Else
            strText = m.aQBarColHeader(i)
        End If
    Next
    SetIniFileProperty "QuoteBarColumns", strText, "Price Ladder", g.strIniFile
    'account bar column header
    strText = ""
    For i = 0 To m.aABarColHeader.Size - 1
        If Len(strText) > 0 Then
            strText = strText & "|" & m.aABarColHeader(i)
        Else
            strText = m.aABarColHeader(i)
        End If
    Next
    SetIniFileProperty "AccountBarColumns", strText, "Price Ladder", g.strIniFile
    
    SetIniFileProperty "AccountID", m.nTradeAcctID, "Price Ladder", g.strIniFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.SaveSettings"

End Sub

Public Property Get BracketOrderOne() As cPtOrder
    Set BracketOrderOne = m.oBracketOrdOne
End Property

Public Property Get BracketOrderTwo() As cPtOrder
    Set BracketOrderTwo = m.oBracketOrdTwo
End Property

Private Sub CreateThisOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    If m.frmBroker Is Nothing Then
        If ConfirmOrder Then
            CreateOrder , , , Order
        Else
            Order.GenesisOrderID = NextGenesisOrderID(AccountNumber)
            SubmitOrder Order
            
            If m.eDisplayStyle = eView_Detail Then
                With fgOrdersInfo
                    .Redraw = flexRDNone
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "Order generated: " & Order.OrderText
                    .Redraw = flexRDBuffered
                End With
            End If
        End If
    Else
        m.frmBroker.Ladder_CreateOrder Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.CreateThisOrder"
    
End Sub

Private Sub CancelThisOrder(Order As cPtOrder, ByVal bConfirm As Boolean)
On Error GoTo ErrSection:
    
    If m.frmBroker Is Nothing Then
        CancelOrder Order, bConfirm, , True
    
        If Not m.oBracketOrdOne Is Nothing Then
            If Order.OrderID = m.oBracketOrdOne.OrderID Then Set m.oBracketOrdOne = Nothing
        End If
    Else
        m.frmBroker.Ladder_CancelOrder Order
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.CancelThisOrder"

End Sub

Private Sub ModifyThisOrder(Order As cPtOrder, Optional ByVal dPrice# = 0#, _
    Optional ByVal iQty& = 0, Optional ByVal bConfirm As Boolean = True)
On Error GoTo ErrSection:
    
    Dim dPriceSave#
    
    If Not m.frmBroker Is Nothing Then
        m.frmBroker.Ladder_ModifyOrder Order, dPrice, iQty
    ElseIf m.oBracketOrdOne Is Nothing Then
        ModifyOrder Order, dPrice, iQty, bConfirm
    ElseIf Order.OrderID = m.oBracketOrdOne.OrderID Then
        If dPrice <> m.oBracketOrdOne.OrderPrice(False) Then
            dPriceSave = m.oBracketOrdOne.OrderPrice(False)
            m.oBracketOrdOne.OrderPrice(False) = dPrice
            If OkayToExecute(m.oBracketOrdOne, RoundToMinMove(m.TickBars(eBARS_Close, m.TickBars.Size - 1), m.dMinMove), True, Me) Then
                ParkOrder m.oBracketOrdOne
            Else
                m.oBracketOrdOne.OrderPrice(False) = dPriceSave
            End If
        End If
    Else
        ModifyOrder Order, dPrice, iQty, bConfirm
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.ModifyThisOrder"
    
End Sub

Public Sub ClearBuySellButtons(Optional ByVal bClearNow As Boolean = True)
On Error Resume Next

    If vseBracketOrder.Appearance = apInset Then vseBracketOrder.Appearance = ap3D
    
    If Not m.oBracketOrdOne Is Nothing Then
        m.oBracketOrdOne.Reload
        If m.oBracketOrdOne.Status = eTT_OrderStatus_Parked Then CancelOrder m.oBracketOrdOne, False
        Set m.oBracketOrdOne = Nothing
    End If
    
    If Not m.oBracketOrdTwo Is Nothing Then
        m.oBracketOrdTwo.Reload
        If m.oBracketOrdTwo.Status = eTT_OrderStatus_Parked Then CancelOrder m.oBracketOrdTwo, False
        Set m.oBracketOrdTwo = Nothing
    End If

End Sub

Private Sub SetCboToSimAccount()
On Error Resume Next            'call from resize event, do not use raiseerror
    
    Dim i&, nID&
    
    If cboAccounts.ListCount = 0 Then InitCboAccount

    If m.frmBroker Is Nothing Then
        For i = 0 To cboAccounts.ListCount - 1
            nID = cboAccounts.ItemData(i)
            If mTradeTracker.TypeOfAccount(nID) = eGDTypeOfAccount_Simulated Then
                TradeAccountID = nID
                cboAccounts.ListIndex = i
                Exit For
            End If
        Next
    End If

End Sub

Private Sub UpdateOpenEntries()
On Error GoTo ErrSection:

    Dim Entries As cGdTree
    Dim Fill As cPtFill
    
    Dim dPrice As Double
    
    Dim lQuantity As Long
    Dim lQtyInGrid As Long
    Dim lIndex As Long
    Dim lRec As Long
    
    Dim tbEntries As cGdTable
    Dim aEntriesIdx As cGdArray
    
    Dim i&, strFile$
    
    If m.nShowOpenEntries = 0 Then Exit Sub
    
    Set Entries = g.Broker.EntriesForSymbol(TradeAccountID, m.nSymID, 0&)
    If Entries Is Nothing Then Exit Sub
    If Entries.Count = 0 Then Exit Sub
    
    If IconPic.Tag <> m.strPos Then
        strFile = g.strAppPath & "\LadderIcon.bmp"
        If m.strPos = "Long" Then
            geFootprintIcon IconPic.hDC, 6, vbWhite, ChartGlbClrForCtl(lblTradePos, g.ChartGlobals.nLongColor, "LongColor"), strFile
        ElseIf m.strPos = "Short" Then
            geFootprintIcon IconPic.hDC, 6, vbWhite, g.ChartGlobals.nShortColor, strFile
        Else
            Set IconPic.Picture = Nothing
        End If
        If FileExist(strFile) Then
            IconPic.Picture = LoadPicture(strFile)
            KillFile strFile, True
        End If
        IconPic.Tag = m.strPos
    
        'for icon with same background color as average entry line
        Set IconPicAvgEntry.Picture = Nothing
        If m.nShowAvgEntry Then
            If m.strPos = "Long" Then
                geFootprintIcon IconPicAvgEntry.hDC, 6, m.nAvgEntryColor, ChartGlbClrForCtl(lblTradePos, g.ChartGlobals.nLongColor, "LongColor"), strFile
            ElseIf m.strPos = "Short" Then
                geFootprintIcon IconPicAvgEntry.hDC, 6, m.nAvgEntryColor, g.ChartGlobals.nShortColor, strFile
            Else
                Set IconPicAvgEntry.Picture = Nothing
            End If
            If FileExist(strFile) Then
                IconPicAvgEntry.Picture = LoadPicture(strFile)
                KillFile strFile, True
            End If
        End If
    End If
    
    Set tbEntries = New cGdTable
    tbEntries.CreateField eGDARRAY_Doubles, 0, , 0
    tbEntries.CreateField eGDARRAY_Longs, 1, , 0

    For lIndex = 1 To Entries.Count
        Set Fill = Entries(lIndex)
        dPrice = Fill.Price
        lQuantity = Fill.NumberOpen(0&)
        
        tbEntries.AddRecord ""
        tbEntries(0, lIndex - 1) = dPrice
        tbEntries(1, lIndex - 1) = lQuantity
    Next
        
    Set aEntriesIdx = tbEntries.CreateSortedIndex(0, eGdSort_Descending)
    If aEntriesIdx Is Nothing Then GoTo ErrExit
    
    dPrice = tbEntries(0, aEntriesIdx(0))
    lQuantity = tbEntries(1, aEntriesIdx(0))
    
    lIndex = 0
    lRec = tbEntries.NumRecords
    
    With fgTickDistribution
        For i = .FixedRows To .Rows - 1
            If lIndex >= lRec Then
                Exit For
            ElseIf i > .FixedRows And i < .Rows Then
                While dPrice = ValOfColPrice(i)
                    If dPrice = m.dAvgEntry And Not IconPicAvgEntry.Picture Is Nothing Then
                        .Cell(flexcpPicture, i, eTDCols_Entries) = IconPicAvgEntry.Picture
                        .Cell(flexcpPictureAlignment, i, eTDCols_Entries) = flexPicAlignLeftCenter
                    ElseIf Not IconPic.Picture Is Nothing Then
                        .Cell(flexcpPicture, i, eTDCols_Entries) = IconPic.Picture
                        .Cell(flexcpPictureAlignment, i, eTDCols_Entries) = flexPicAlignLeftCenter
                    End If
                    'add to entries already in grid
                    lQtyInGrid = ValOfText(.TextMatrix(i, eTDCols_Entries))
                    .TextMatrix(i, eTDCols_Entries) = Str(lQtyInGrid + lQuantity)
                    
                    lIndex = lIndex + 1
                    If lIndex < lRec Then
                        dPrice = tbEntries(0, aEntriesIdx(lIndex))
                        lQuantity = tbEntries(1, aEntriesIdx(lIndex))
                    Else
                        dPrice = -1#
                    End If
                Wend
            
            End If
        Next
    End With

ErrExit:
    Set Entries = Nothing
    Set tbEntries = Nothing
    Set aEntriesIdx = Nothing
    Exit Sub

ErrSection:
    Set Entries = Nothing
    Set tbEntries = Nothing
    Set aEntriesIdx = Nothing
    RaiseError "frmTickDistribution.UpdateOpenEntries"

End Sub

Public Property Get HilitePosColor() As Long
    HilitePosColor = m.nHighlightPos
End Property

Public Property Let HilitePosColor(ByVal nColor&)
    m.nHighlightPos = nColor
End Property

Public Property Get HiliteEquityColor() As Long
    HiliteEquityColor = m.nHighlightEquity
End Property

Public Property Let HiliteEquityColor(ByVal nColor&)
    m.nHighlightEquity = nColor
End Property

Public Property Get LadderVolumeStyle() As eLadderVolStyle
    LadderVolumeStyle = m.eVolumeStyle
End Property

Public Property Let LadderVolumeStyle(ByVal eStyle As eLadderVolStyle)
    If m.eVolumeStyle <> eStyle Then
        m.eVolumeStyle = eStyle
        m.bAutosizePrice = True
    End If
End Property

Public Sub OutlineCell(ByVal nRow&, ByVal nColor&, ByVal bClear As Boolean)
On Error GoTo ErrSection:

    Dim dPrice#, hArray&, i&
    Dim bFound As Boolean

    With fgTickDistribution
        If bClear And nRow = -1 Then
            .Select .FixedRows, eTDCols_Price, .Rows - 1, eTDCols_Price
            .CellBorder 0, 0, 0, 0, 0, 0, 0
            m.tbOutlineCells.NumRecords = 0
        ElseIf nRow >= .FixedRows And nRow < .Rows Then
            If bClear Then
                .Select nRow, eTDCols_Price
                .CellBorder 0, 0, 0, 0, 0, 0, 0
            Else
                .Select nRow, eTDCols_Price
                .CellBorder nColor, 2, 2, 2, 2, 0, 0
                m.nOutlineColor = nColor
            End If
            
            hArray = m.tbOutlineCells.FieldArrayHandle(0)
            dPrice = m.Data(eFld_Price, .RowData(nRow))
            
            For i = 0 To m.tbOutlineCells.NumRecords - 1
                If gdGetNum(hArray, i) = dPrice Then
                    bFound = True
                    Exit For
                End If
            Next
            
            If bFound Then
                If bClear Then
                    m.tbOutlineCells.RemoveRecords i
                Else
                    m.tbOutlineCells(1, i) = nColor
                End If
            ElseIf Not bClear Then
                m.tbOutlineCells.AddRecord ""
                i = m.tbOutlineCells.NumRecords - 1
                m.tbOutlineCells(0, i) = dPrice
                m.tbOutlineCells(1, i) = nColor
            End If

        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.OutlineCell"

End Sub

Private Sub vseOrderBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = vbRightButton Then
        frmTickDistributionCfg.ShowMe Me, , , , eLadderTab_OrderBar
    End If

End Sub

Public Sub GridTextIncrease()
On Error GoTo ErrSection:
    
    GridFontSize = GridFontSize + 1
    RefreshGrid

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.GridTextIncrease"

End Sub

Public Sub GridTextDecrease()
On Error GoTo ErrSection:
    
    GridFontSize = GridFontSize - 1
    RefreshGrid

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistribution.GridTextDecrease"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionString
'' Description: Get the position string from the appropriate location
'' Inputs:      None
'' Returns:     Position String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetPositionString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the string
    
    If m.frmBroker Is Nothing Then
        strReturn = g.Broker.PositionString(TradeAccountID, m.nSymID, 0&)
    Else
        strReturn = m.frmBroker.PositionStringForSymbol(m.strSym)
    End If
    
    GetPositionString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.GetPositionString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetWorkingOrders
'' Description: Get a collection of working orders
'' Inputs:      None
'' Returns:     Working Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetWorkingOrders() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnOrders As cGdTree         ' Collection of working orders to return

    If m.frmBroker Is Nothing Then
        Set ReturnOrders = g.Broker.PrimaryOrdersForSymbol(TradeAccountID, m.nSymID, 0&)
    Else
        Set ReturnOrders = m.frmBroker.WorkingOrdersForSymbol(m.strSym)
    End If
    
    Set GetWorkingOrders = ReturnOrders
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistribution.GetWorkingOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersToTableNormal
'' Description: Load up the orders table from a collection of Order objects
'' Inputs:      Orders, Max Price, Min Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrdersToTableNormal(Orders As cGdTree, dMaxPrice As Double, dMinPrice As Double)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cPtOrder               ' Order from the collection
    Dim dOrderPrice As Double           ' Order price

    For lIndex = 1 To Orders.Count
        Set Order = Orders(lIndex)
        If Not Order Is Nothing Then
            If Order.OrderType = eTT_OrderType_StopWithLimit Then
                dOrderPrice = RoundToMinMove(Order.OrderPrice(True), m.dMinMove)
            Else
                dOrderPrice = RoundToMinMove(Order.OrderPrice(False), m.dMinMove)
            End If
            If Order.OrderType = eTT_OrderType_Market Then
                dOrderPrice = m.dLastPrice
            End If
            If dOrderPrice <> kNullData Then
                m.tbOrders.AddRecord " "
                m.tbOrders(0, m.tbOrders.NumRecords - 1) = Order.OrderID
                m.tbOrders(1, m.tbOrders.NumRecords - 1) = Abs(Order.Buy)
                m.tbOrders(2, m.tbOrders.NumRecords - 1) = Order.RemainingQuantity
                m.tbOrders(3, m.tbOrders.NumRecords - 1) = Order.OrderType
                m.tbOrders(4, m.tbOrders.NumRecords - 1) = dOrderPrice
                m.tbOrders(5, m.tbOrders.NumRecords - 1) = Abs(OrderIsPending(Order))
                m.tbOrders(6, m.tbOrders.NumRecords - 1) = Order.Status
                
                If dOrderPrice > dMaxPrice Then
                    dMaxPrice = dOrderPrice
                ElseIf dOrderPrice < dMinPrice Then
                    dMinPrice = dOrderPrice
                End If
            End If
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.OrdersToTableNormal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrdersToTableBroker
'' Description: Load up the orders table from a collection of broker order objects
'' Inputs:      Orders, Max Price, Min Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrdersToTableBroker(Orders As cGdTree, dMaxPrice As Double, dMinPrice As Double)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cBrokerMessage         ' Order from the collection
    Dim dOrderPrice As Double           ' Order price
    Dim nOrderType As eTT_OrderType     ' Order type

    For lIndex = 1 To Orders.Count
        Set Order = Orders(lIndex)
        If Not Order Is Nothing Then
            Select Case UCase(Order("Type"))
                Case "STOP"
                    nOrderType = eTT_OrderType_Stop
                    dOrderPrice = Val(Order("StopPrice"))
                Case "LIMIT"
                    nOrderType = eTT_OrderType_Limit
                    dOrderPrice = Val(Order("LimitPrice"))
                Case "STOPWITHLIMIT"
                    nOrderType = eTT_OrderType_StopWithLimit
                    dOrderPrice = Val(Order("StopPrice"))
                Case "MARKET"
                    nOrderType = eTT_OrderType_Market
                    dOrderPrice = m.dLastPrice
            End Select
            
            If dOrderPrice > 0 Then
                m.tbOrders.AddRecord " "
                m.tbOrders(0, m.tbOrders.NumRecords - 1) = 0& ' Order.OrderID
                m.tbOrders(1, m.tbOrders.NumRecords - 1) = Abs(Left(UCase(Order("Side")), 3) = "BUY")
                m.tbOrders(2, m.tbOrders.NumRecords - 1) = Order("Quantity")
                m.tbOrders(3, m.tbOrders.NumRecords - 1) = nOrderType
                m.tbOrders(4, m.tbOrders.NumRecords - 1) = dOrderPrice
                m.tbOrders(5, m.tbOrders.NumRecords - 1) = Abs(False)
                m.tbOrders(6, m.tbOrders.NumRecords - 1) = 0& ' Order("Status")
                
                If dOrderPrice > dMaxPrice Then
                    dMaxPrice = dOrderPrice
                ElseIf dOrderPrice < dMinPrice Then
                    dMinPrice = dOrderPrice
                End If
            End If
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.OrdersToTableBroker"
    
End Sub

Public Property Get BrokerViewMode() As Boolean
On Error GoTo ErrSection:

    If Not m.frmBroker Is Nothing Then BrokerViewMode = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.BrokerViewMode.Get"
    
End Property

Private Property Get Broker() As eTT_AccountType
On Error GoTo ErrSection:

    Dim nReturn As eTT_AccountType      ' Return value for the function
    
    If m.frmBroker Is Nothing Then
        nReturn = g.Broker.AccountTypeForID(TradeAccountID)
    Else
        nReturn = m.frmBroker.Broker
    End If
    
    Broker = nReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.Broker.Get"
    
End Property

Private Property Get AccountNumber() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If m.frmBroker Is Nothing Then
        strReturn = g.Broker.AccountNumberForID(TradeAccountID)
    Else
        strReturn = m.frmBroker.Account
    End If
    
    AccountNumber = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.AccountNumber.Get"
    
End Property

Private Property Get ConnectionStatus() As eGDConnectionStatus
On Error GoTo ErrSection:

    Dim nReturn As eGDConnectionStatus  ' Return value for the function
    
    If m.frmBroker Is Nothing Then
        nReturn = g.Broker.ConnectionStatusForAccount(TradeAccountID)
    Else
        nReturn = m.frmBroker.ConnectionStatus
    End If
    
    ConnectionStatus = nReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.ConnectionStatus.Get"
    
End Property

Private Property Get ConfirmOrder() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If m.frmBroker Is Nothing Then
        bReturn = g.Broker.ConfirmOrder(TradeAccountID, SymbolOrSymbolID)
    Else
        bReturn = False
    End If
    
    ConfirmOrder = bReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.ConfirmOrder.Get"
    
End Property

Private Property Get TypeOfAccount() As eGDTypeOfAccount
On Error GoTo ErrSection:

    Dim nReturn As eGDTypeOfAccount     ' Return value for the function
    
    If m.frmBroker Is Nothing Then
        nReturn = mTradeTracker.TypeOfAccount(TradeAccountID)
    Else
        nReturn = eGDTypeOfAccount_BrokerLive
    End If
    
    TypeOfAccount = nReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmTickDistribution.TypeOfAccount.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitQuantityEditor
'' Description: Initialize the quantity editor according to the selected
''              account and symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitQuantityEditor()
On Error GoTo ErrSection:
    
    g.Broker.InitQuantityEditor m.Quantity, vscrQty, txtTradeQty, TradeAccountID, m.nSymID
    SetQuantityPresetButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.InitQuantityEditor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetQuantityPresetButtons
'' Description: Set the quantity preset buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetQuantityPresetButtons()
On Error GoTo ErrSection:

    Dim strSecType As String            ' Security type for the chart symbol

    strSecType = g.Broker.TradeSecType(m.nSymID)
    g.Broker.GetQuantityPresets TradeAccountID, m.nSymID, m.lPreset1, m.lPreset2, m.lPreset3
    cmdQty1.Caption = ShortDisplayNumber(m.lPreset1)
    cmdQty2.Caption = ShortDisplayNumber(m.lPreset2)
    cmdQty3.Caption = ShortDisplayNumber(m.lPreset3)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.SetQuantityPresetButtons"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixOrderBarCtrlString
'' Description: Remove the BC, SC, PR
'' Inputs:      None
'' Returns:     None
''
'' BC, SC, PR buttons are valid only for frmChart, frmChart2
''
'' JM - 08-23-2013 Fix for customer's Vit Funda 55896 issue with error
''      mChartLadderCtrls.OrdbarCtrlFromCode: object doesn't support this property or method
''      The error occurs with regional settings in English, but system never shows dialog
''      Changing regional settings to German is how to duplicate proble
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixOrderBarCtrlString()
On Error GoTo ErrSection

    Dim strText As String

    strText = Replace(m.strOrdBarCtrls, "BC;0|", "")
    m.strOrdBarCtrls = strText
    strText = Replace(m.strOrdBarCtrls, "SC;0|", "")
    m.strOrdBarCtrls = strText
    strText = Replace(m.strOrdBarCtrls, "PR;0|", "")
    m.strOrdBarCtrls = strText
    
    strText = Replace(m.strOrdBarCtrls, "BC;1|", "")
    m.strOrdBarCtrls = strText
    strText = Replace(m.strOrdBarCtrls, "SC;1|", "")
    m.strOrdBarCtrls = strText
    strText = Replace(m.strOrdBarCtrls, "PR;2|", "")
    m.strOrdBarCtrls = strText
    
    strText = Replace(m.strOrdBarCtrls, "BC;2|", "")
    m.strOrdBarCtrls = strText
    strText = Replace(m.strOrdBarCtrls, "SC;2|", "")
    m.strOrdBarCtrls = strText
    strText = Replace(m.strOrdBarCtrls, "PR;2|", "")
    m.strOrdBarCtrls = strText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistribution.FixOrderBarCtrlString"
    
End Sub

