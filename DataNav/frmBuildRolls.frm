VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmBuildRolls 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgDiffs 
      Height          =   915
      Left            =   180
      TabIndex        =   1
      Top             =   5700
      Width           =   4815
      _cx             =   8493
      _cy             =   1614
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
   Begin HexUniControls.ctlUniFrameWL fraDataCompare 
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   3675
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
      Caption         =   "frmBuildRolls.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBuildRolls.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBuildRolls.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraBelieve 
         Height          =   255
         Left            =   780
         TabIndex        =   17
         Top             =   2220
         Width           =   2715
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
         Caption         =   "frmBuildRolls.frx":005C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmBuildRolls.frx":007C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":009C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optDM 
            Height          =   220
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmBuildRolls.frx":00B8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmBuildRolls.frx":00F2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":0112
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optExt 
            Height          =   220
            Left            =   1380
            TabIndex        =   21
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmBuildRolls.frx":012E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmBuildRolls.frx":0160
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":0180
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraOutput 
         Height          =   255
         Left            =   780
         TabIndex        =   23
         Top             =   1920
         Width           =   2775
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
         Caption         =   "frmBuildRolls.frx":019C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmBuildRolls.frx":01BC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":01DC
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optDist 
            Height          =   220
            Left            =   1380
            TabIndex        =   24
            Top             =   0
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmBuildRolls.frx":01F8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmBuildRolls.frx":022A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":024A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optDiffs 
            Height          =   220
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmBuildRolls.frx":0266
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmBuildRolls.frx":029E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":02BE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniTextBoxXP txtOutput 
         Height          =   285
         Left            =   660
         TabIndex        =   25
         Top             =   1560
         Width           =   1755
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmBuildRolls.frx":02DA
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
         Tip             =   "frmBuildRolls.frx":02FA
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":031A
      End
      Begin HexUniControls.ctlUniComboImageXP cboFormat 
         Height          =   315
         Left            =   660
         TabIndex        =   22
         Top             =   870
         Width           =   1215
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
         Tip             =   "frmBuildRolls.frx":0336
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0356
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
         Height          =   300
         Left            =   2640
         TabIndex        =   20
         Top             =   390
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
         Caption         =   "frmBuildRolls.frx":0372
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":03A0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":03C0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPath 
         Height          =   285
         Left            =   660
         TabIndex        =   19
         Top             =   405
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmBuildRolls.frx":03DC
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
         Tip             =   "frmBuildRolls.frx":03FC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":041C
      End
      Begin HexUniControls.ctlUniLabelXP lblPathInst 
         Height          =   255
         Left            =   60
         Top             =   120
         Width           =   3495
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
         Caption         =   "frmBuildRolls.frx":0438
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":04BA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":04DA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBelieve 
         Height          =   195
         Left            =   60
         Top             =   2220
         Width           =   555
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
         Caption         =   "frmBuildRolls.frx":04F6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":0526
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0546
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblType 
         Height          =   195
         Left            =   60
         Top             =   1920
         Width           =   555
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
         Caption         =   "frmBuildRolls.frx":0562
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":058C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":05AC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOutputDesc 
         Height          =   255
         Left            =   60
         Top             =   1320
         Width           =   3435
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
         Caption         =   "frmBuildRolls.frx":05C8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":064C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":066C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOutput 
         Height          =   255
         Left            =   60
         Top             =   1575
         Width           =   615
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
         Caption         =   "frmBuildRolls.frx":0688
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":06B8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":06D8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFormat 
         Height          =   255
         Left            =   60
         Top             =   900
         Width           =   555
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
         Caption         =   "frmBuildRolls.frx":06F4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":0724
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0744
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPath 
         Height          =   255
         Left            =   60
         Top             =   420
         Width           =   495
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
         Caption         =   "frmBuildRolls.frx":0760
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":078C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":07AC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMain 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
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
      Caption         =   "frmBuildRolls.frx":07C8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBuildRolls.frx":07E8
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBuildRolls.frx":0808
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkOutputRule 
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1980
         Width           =   1995
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
         Caption         =   "frmBuildRolls.frx":0824
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":085C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":087C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fra57 
         Height          =   555
         Left            =   420
         TabIndex        =   15
         Top             =   1440
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
         Caption         =   "frmBuildRolls.frx":0898
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmBuildRolls.frx":08C4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":08E4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtRollPath 
            Height          =   285
            Left            =   1500
            TabIndex        =   7
            Top             =   15
            Width           =   1455
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmBuildRolls.frx":0900
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
            Tip             =   "frmBuildRolls.frx":0920
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":0940
         End
         Begin HexUniControls.ctlUniRadioXP optKeepDates 
            Height          =   220
            Left            =   60
            TabIndex        =   6
            Top             =   0
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmBuildRolls.frx":095C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmBuildRolls.frx":0996
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":09B6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optUseRules 
            Height          =   220
            Left            =   60
            TabIndex        =   8
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmBuildRolls.frx":09D2
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmBuildRolls.frx":0A10
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmBuildRolls.frx":0A30
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chk56 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   870
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
         Caption         =   "frmBuildRolls.frx":0A4C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0A82
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0AA2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chk57 
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1230
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
         Caption         =   "frmBuildRolls.frx":0ABE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0AF4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0B14
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chk55 
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   510
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
         Caption         =   "frmBuildRolls.frx":0B30
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0B66
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0B86
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboSymbol 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   0
         Width           =   2715
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
         Tip             =   "frmBuildRolls.frx":0BA2
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0BC2
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   255
         Left            =   0
         Top             =   30
         Width           =   675
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
         Caption         =   "frmBuildRolls.frx":0BDE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBuildRolls.frx":0C0E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0C2E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1575
      Left            =   3780
      TabIndex        =   10
      Top             =   120
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
      Caption         =   "frmBuildRolls.frx":0C4A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBuildRolls.frx":0C6A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBuildRolls.frx":0C8A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkBackupFiles 
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1320
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
         Caption         =   "frmBuildRolls.frx":0CA6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0CDC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0CFC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkBuildAll 
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   975
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
         Caption         =   "frmBuildRolls.frx":0D18
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0D4C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0D6C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdExit 
         Height          =   435
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   1155
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
         Caption         =   "frmBuildRolls.frx":0D88
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0DB2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0DD2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBuild 
         Height          =   435
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1155
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
         Caption         =   "frmBuildRolls.frx":0DEE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBuildRolls.frx":0E1A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBuildRolls.frx":0E3A
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmBuildRolls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBuildRolls.frm
'' Description: Allow for rebuilding continuous contract rolls for a symbol
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/29/2004   DAJ         Created
'' 03/07/2011   DAJ         Don't roll 55 or 56 to a contract with no data
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eFormMode
    eFormMode_BuildRolls
    eFormMode_DataComp
End Enum

Private Type mPrivate
    astrMarkets As cGdArray             ' Array of markets from the symbol universe
    Mode As eFormMode                   ' Mode we are running the forms in
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form for rebuilding rolls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    m.Mode = eFormMode_BuildRolls
    Height = 2820
    fraMain.Visible = True
    fraDataCompare.Visible = False
    fgDiffs.Visible = False
    
    Set m.astrMarkets = New cGdArray
    m.astrMarkets.Create eGDARRAY_Strings
                
    txtRollPath.Text = GetIniFileProperty("RollPath", "N:\MC1\Rolls", "BuildRolls", g.strIniFile)
    LoadCombo
    cboSymbol.Text = GetIniFileProperty("Symbol", "SP", "BuildRolls", g.strIniFile)
    chk55.Value = GetIniFileProperty("Do55", vbUnchecked, "BuildRolls", g.strIniFile)
    chk56.Value = GetIniFileProperty("Do56", vbUnchecked, "BuildRolls", g.strIniFile)
    chk57.Value = GetIniFileProperty("Do57", vbUnchecked, "BuildRolls", g.strIniFile)
    optKeepDates.Value = GetIniFileProperty("Retain", True, "BuildRolls", g.strIniFile)
    optUseRules.Value = Not optKeepDates
    chkOutputRule.Value = GetIniFileProperty("OutputRule", vbUnchecked, "BuildRolls", g.strIniFile)
    chkBackupFiles.Value = GetIniFileProperty("BackupFiles", vbUnchecked, "BuildRolls", g.strIniFile)
    chkBuildAll.Value = GetIniFileProperty("BuildAll", vbUnchecked, "BuildRolls", g.strIniFile)
    EnableControls
    
    ShowForm Me, True
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmBuildRolls.ShowMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbol_Click
'' Description: Handle when the user changes the selection in the combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbol_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.cboSymbol.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chk55_Click
'' Description: When the 55 check box is changed, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chk55_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.chk55.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chk56_Click
'' Description: When the 56 check box is changed, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chk56_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.chk56.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chk57_Click
'' Description: When the 57 check box is changed, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chk57_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.chk57.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to browse for a path
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    'txtPath.Text = frmBrowseFolders.ShowMe(txtPath.Text)
    txtPath.Text = BrowseForFolder(txtPath.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.cmdBrowse.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBuild_Click
'' Description: Build the appropriate continuous contracts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBuild_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOutput As New cGdArray      ' Array of output information
    Dim astrFile As New cGdArray        ' Input file of paths
    Dim astrReturn As New cGdArray      ' Array returned from Compare

    If m.Mode = eFormMode_DataComp Then
        Screen.MousePointer = vbHourglass
        With fgDiffs
            .Redraw = flexRDNone
            .Rows = 0
            astrOutput.Create eGDARRAY_Strings
            If UCase(Right(txtPath.Text, 4)) = ".LST" Then
                If astrFile.FromFile(txtPath.Text) Then
                    For lIndex = 0 To astrFile.Size - 1
                        StatusMsg "Processing: " & astrFile(lIndex)
                        Set astrReturn = CompareDataDirectory(astrFile(lIndex), cboFormat.Text, optDiffs, optDM)
                        astrOutput.AppendFromArray astrReturn
                    Next lIndex
                    StatusMsg ""
                End If
            Else
                Set astrOutput = CompareDataDirectory(txtPath.Text, cboFormat.Text, optDiffs, optDM)
            End If
            For lIndex = 0 To astrOutput.Size - 1
                .AddItem astrOutput(lIndex)
            Next lIndex
            .Redraw = flexRDBuffered
        End With
        If Len(Trim(txtOutput.Text)) > 0 Then astrOutput.ToFile txtOutput.Text
        Screen.MousePointer = vbDefault
    Else
        If chkBuildAll.Value = vbChecked Then
            For lIndex = 0 To cboSymbol.ListCount - 1
                cboSymbol.ListIndex = lIndex
                BuildContinuous
            Next lIndex
        Else
            BuildContinuous
        End If
    End If

ErrExit:
    Set astrFile = Nothing
    Set astrOutput = Nothing
    Set astrReturn = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "frmBuildRolls.cmdBuild.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExit_Click
'' Description: Exit the form when the user clicks on the Exit button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
On Error GoTo ErrSection:

    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.cmdExit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form loads, do some initialization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Build Continuous Contract Rolls"
    Icon = Picture16(ToolbarIcon("kBlank"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user clicks on the 'X', let ShowMe unload the form
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load the symbols combo box from the symbol universe
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol to add to the combo box
    Dim strDesc As String               ' Description of the symbol
    
    If SU_GetMarkets(m.astrMarkets) Then
        m.astrMarkets.Sort
        For lIndex = 0 To m.astrMarkets.Size - 1
            strSymbol = Parse(Parse(m.astrMarkets(lIndex), ";", 1), "-", 1)
            strDesc = Replace(Parse(m.astrMarkets(lIndex), ";", 3), " - Market", "")
            
            cboSymbol.AddItem strSymbol & " (" & strDesc & ")", lIndex
        Next lIndex
        
        cboSymbol.ListIndex = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.LoadCombo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
    
    Dim lMinWidth As Long               ' Minimum width for the form
    Dim lMinHeight As Long              ' Minimum height for the form

    If m.Mode = eFormMode_DataComp Then
        lMinWidth = fraButtons.Width + fraDataCompare.Width + (fraDataCompare.Left * 3)
        lMinHeight = fraButtons.Height * 2
        If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
        
        With fraButtons
            .Move ScaleWidth - .Width - fraDataCompare.Left
        End With
        
        With fgDiffs
            .Move .Left, .Top, ScaleWidth - (.Left * 2), ScaleHeight - .Top - fraButtons.Top
        End With
    Else
        Me.Height = 2820
        Me.Width = 5205
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Do some cleanup when the form is unloaded
'' Inputs:      Whether to Cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.astrMarkets = Nothing
    
    Select Case m.Mode
        Case eFormMode_BuildRolls
            SetIniFileProperty "Symbol", cboSymbol.Text, "BuildRolls", g.strIniFile
            SetIniFileProperty "Do55", chk55.Value, "BuildRolls", g.strIniFile
            SetIniFileProperty "Do56", chk56.Value, "BuildRolls", g.strIniFile
            SetIniFileProperty "Do57", chk57.Value, "BuildRolls", g.strIniFile
            SetIniFileProperty "Retain", optKeepDates.Value, "BuildRolls", g.strIniFile
            SetIniFileProperty "RollPath", txtRollPath.Text, "BuildRolls", g.strIniFile
            SetIniFileProperty "OutputRule", chkOutputRule.Value, "BuildRolls", g.strIniFile
            SetIniFileProperty "BackupFiles", chkBackupFiles.Value, "BuildRolls", g.strIniFile
            SetIniFileProperty "BuildAll", chkBuildAll.Value, "BuildRolls", g.strIniFile
            
        Case eFormMode_DataComp
            SetIniFileProperty "Path", txtPath.Text, "DataCompare", g.strIniFile
            SetIniFileProperty "Format", cboFormat.Text, "DataCompare", g.strIniFile
            SetIniFileProperty "Output", txtOutput.Text, "DataCompare", g.strIniFile
            SetIniFileProperty "Diff", optDiffs.Value, "DataCompare", g.strIniFile
            SetIniFileProperty "DM", optDM.Value, "DataCompare", g.strIniFile
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    If m.Mode = eFormMode_BuildRolls Then
        Enable cmdBuild, (chk55 = vbChecked) Or (chk56 = vbChecked) Or (chk57 = vbChecked)
        Enable optKeepDates, (chk57 = vbChecked)
        Enable optUseRules, (chk57 = vbChecked)
        Enable txtRollPath, optKeepDates
        Enable chkOutputRule, optUseRules And (chk57 = vbChecked)
        chkBuildAll.Visible = True
    Else
        Enable lblType, Len(Trim(txtOutput.Text)) > 0
        Enable optDiffs, Len(Trim(txtOutput.Text)) > 0
        Enable optDist, Len(Trim(txtOutput.Text)) > 0
        Enable lblBelieve, optDist And optDist.Enabled
        Enable optDM, optDist And optDist.Enabled
        Enable optExt, optDist And optDist.Enabled
        chkBuildAll.Visible = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Build55
'' Description: Build a 55 continous contract roll file for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Build55()
On Error GoTo ErrSection:

    Dim astrContracts As New cGdArray   ' List of contracts for the selected symbol
    Dim astrRollFile As New cGdArray    ' Roll file for the selected symbol
    Dim OldContract As New cGdBars      ' Bars for the old contract
    Dim NewContract As New cGdBars      ' Bars for the new contract
    Dim lSymbolID As Long               ' Symbol ID for the selected symbol
    Dim lNextID As Long                 ' Symbol ID for the next contract
    Dim lIndex As Long                  ' Index into a for loop
    Dim lDate As Long                   ' Date to look up in both contracts
    Dim lContract As Long               ' Current contract
    Dim dDiff As Double                 ' Difference in close of contracts
    Dim lLastDate As Long               ' Date of the last daily download
    Dim lPos1 As Long                   ' Position we want in the first contract
    Dim lPos2 As Long                   ' Position we want in the next contract
    
    ' Do some initialization...
    Screen.MousePointer = vbHourglass
    astrContracts.Create eGDARRAY_Strings
    astrRollFile.Create eGDARRAY_Strings
    lLastDate = LastDailyDownload
    
    ' Get the ID for the selected symbol...
    lSymbolID = CLng(ValOfText(Parse(m.astrMarkets(cboSymbol.ListIndex), ";", 2)))
    
    ' Get all of the contracts for the selected symbol...
    If SU_GetContracts(lSymbolID, astrContracts) Then
        ' Walk through all of the contracts for the selected symbol...
        For lIndex = 0 To astrContracts.Size - 2
            ' If this is the first contract, we need to load it...
            If OldContract.Size = 0 Then
                lSymbolID = CLng(ValOfText(Parse(astrContracts(lIndex), ";", 2)))
                DM_GetBars OldContract, lSymbolID
                If OldContract.Size > 0 Then
                    lDate = OldContract(eBARS_DateTime, 0)
                    astrRollFile.Add Format(lDate, "YYMMDD") & " " & Format(OldContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(0#, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                End If
            
            ' otherwise it has already been loaded...
            Else
                Set OldContract = NewContract.MakeCopy
            End If
            
            If OldContract.Size > 0 Then
                lContract = CLng(OldContract.Prop(eBARS_Contract))
                
                ' Load the next contract...
                lNextID = CLng(ValOfText(Parse(astrContracts(lIndex + 1), ";", 2)))
                DM_GetBars NewContract, lNextID
                
                If NewContract.Size = 0 Then
                    Set NewContract = OldContract.MakeCopy
                Else
                    ' Find the correct position in both contracts...
                    If OldContract.Prop(eBARS_ExpiresPriorMonth) = 0 Then
                        lDate = JulFromLong((lContract * 100) + 1)
                    Else
                        If lContract Mod 100 = 1 Then
                            lDate = JulFromLong((((lContract \ 100) - 1) * 10000) + 1201)
                        Else
                            lDate = JulFromLong(((lContract - 1) * 100) + 1)
                        End If
                    End If
                        
                    If lDate > lLastDate Then Exit For
                    lPos1 = OldContract.FindDateTime(lDate)
                    lPos2 = NewContract.FindDateTime(lDate)
                    
                    If lPos1 >= OldContract.Size Then
                        lPos1 = lPos1 - 1
                        lPos2 = NewContract.FindDateTime(OldContract(eBARS_DateTime, lPos1))
                    End If
                    
                    ' Figure out the difference in closes on the day prior and output...
                    If NewContract(eBARS_DateTime, lPos2) <> lDate Then
                        lPos1 = OldContract.FindDateTime(NewContract(eBARS_DateTime, lPos2))
                        If OldContract(eBARS_DateTime, lPos1) <> kNullData Then
                            dDiff = NewContract(eBARS_Close, lPos2) - OldContract(eBARS_Close, lPos1)
                            astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2 + 1), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                        Else
                            dDiff = 0#
                            astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2 + 1), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                        End If
                    ElseIf OldContract(eBARS_DateTime, lPos1 - 1) = NewContract(eBARS_DateTime, lPos2 - 1) Then
                        If OldContract(eBARS_DateTime, lPos1) <> kNullData Then
                            dDiff = NewContract(eBARS_Close, lPos2 - 1) - OldContract(eBARS_Close, lPos1 - 1)
                            astrRollFile.Add Format(OldContract(eBARS_DateTime, lPos1), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                        Else
                            dDiff = 0#
                            astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                        End If
                    Else
                        dDiff = -99999#
                        astrRollFile.Add Format(OldContract(eBARS_DateTime, lPos1), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                    End If
                End If
            End If
        Next lIndex
    ElseIf chkBuildAll.Value = vbUnchecked Then
        Err.Raise vbObjectError + 1000, , "No contracts could be loaded for " & cboSymbol.Text
    End If
    
    ' Output the roll file to the application directory...
    If astrRollFile.Size > 0 Then
        astrRollFile.ToFile AddSlash(App.Path) & Parse(cboSymbol.Text, "(", 1) & "-9955.ROL"
        If chkBackupFiles.Value = vbChecked Then
            astrRollFile.Remove astrRollFile.Size - 1
            astrRollFile.ToFile AddSlash(App.Path) & Parse(cboSymbol.Text, "(", 1) & "-9955.BAK"
        End If
    End If

ErrExit:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set OldContract = Nothing
    Set NewContract = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set OldContract = Nothing
    Set NewContract = Nothing
    Screen.MousePointer = vbDefault
    RaiseError "frmBuildRolls.Build55", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Build56
'' Description: Build a 56 continous contract roll file for the selected symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Build56()
On Error GoTo ErrSection:

    Dim astrContracts As New cGdArray   ' List of contracts for the selected symbol
    Dim astrRollFile As New cGdArray    ' Roll file for the selected symbol
    Dim OldContract As New cGdBars      ' Bars for the old contract
    Dim NewContract As New cGdBars      ' Bars for the new contract
    Dim lSymbolID As Long               ' Symbol ID for the selected symbol
    Dim lNextID As Long                 ' Symbol ID for the next contract
    Dim lIndex As Long                  ' Index into a for loop
    Dim lDate As Long                   ' Date to look up in both contracts
    Dim lContract As Long               ' Current contract
    Dim dDiff As Double                 ' Difference in close of contracts
    Dim lLastDate As Long               ' Date of the last daily download
    Dim lPos1 As Long                   ' Position we want in the first contract
    Dim lPos2 As Long                   ' Position we want in the next contract
    
    ' Do some initialization...
    Screen.MousePointer = vbHourglass
    astrContracts.Create eGDARRAY_Strings
    astrRollFile.Create eGDARRAY_Strings
    lLastDate = LastDailyDownload
    
    ' Get the ID for the selected symbol...
    lSymbolID = CLng(ValOfText(Parse(m.astrMarkets(cboSymbol.ListIndex), ";", 2)))
    
    ' Get all of the contracts for the selected symbol...
    If SU_GetContracts(lSymbolID, astrContracts) Then
        ' Walk through all of the contracts for the selected symbol...
        For lIndex = 0 To astrContracts.Size - 2
            ' If this is the first contract, we need to load it...
            If OldContract.Size = 0 Then
                lSymbolID = CLng(ValOfText(Parse(astrContracts(lIndex), ";", 2)))
                DM_GetBars OldContract, lSymbolID
                If OldContract.Size > 0 Then
                    lDate = OldContract(eBARS_DateTime, 0)
                    astrRollFile.Add Format(lDate, "YYMMDD") & " " & Format(OldContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(0#, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                End If
            
            ' otherwise it has already been loaded...
            Else
                Set OldContract = NewContract.MakeCopy
            End If
            
            If OldContract.Size > 0 Then
                lContract = CLng(OldContract.Prop(eBARS_Contract))
                
                ' Load the next contract...
                lNextID = CLng(ValOfText(Parse(astrContracts(lIndex + 1), ";", 2)))
                DM_GetBars NewContract, lNextID
                
                If NewContract.Size = 0 Then
                    Set NewContract = OldContract.MakeCopy
                Else
                    ' Find the correct position in both contracts...
                    lPos1 = OldContract.Size - 1
                    lDate = OldContract(eBARS_DateTime, lPos1)
                    If lDate >= lLastDate Then Exit For
                    lPos2 = NewContract.FindDateTime(lDate)
                    If NewContract(eBARS_DateTime, lPos2 + 1) < 0 Then Exit For
                    
                    ' Figure out the difference in closes on the day prior and output...
                    If OldContract(eBARS_DateTime, lPos1) = NewContract(eBARS_DateTime, lPos2) Then
                        dDiff = NewContract(eBARS_Close, lPos2) - OldContract(eBARS_Close, lPos1)
                        astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2 + 1), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                    Else
                        dDiff = -99999#
                        astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2 + 1), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
                    End If
                End If
            End If
        Next lIndex
    ElseIf chkBuildAll.Value = vbUnchecked Then
        Err.Raise vbObjectError + 1000, , "No contracts could be loaded for " & cboSymbol.Text
    End If
    
    ' Output the roll file to the application directory...
    astrRollFile.ToFile AddSlash(App.Path) & Parse(cboSymbol.Text, "(", 1) & "-9956.ROL"
    If chkBackupFiles.Value = vbChecked Then
        astrRollFile.Remove astrRollFile.Size - 1
        astrRollFile.ToFile AddSlash(App.Path) & Parse(cboSymbol.Text, "(", 1) & "-9956.BAK"
    End If

ErrExit:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set OldContract = Nothing
    Set NewContract = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set OldContract = Nothing
    Set NewContract = Nothing
    Screen.MousePointer = vbDefault
    RaiseError "frmBuildRolls.Build56", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Build57ByRule
'' Description: Build a 57 continous contract roll file using rolling rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Build57ByRule()
On Error GoTo ErrSection:

    Dim astrContracts As New cGdArray   ' List of contracts for the selected symbol
    Dim astrRollFile As New cGdArray    ' Roll file for the selected symbol
    Dim aBars As New cGdArray           ' Array of Bars for all of the contracts
    Dim lSymbolID As Long               ' Symbol ID for the selected symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Date of the last daily download
    Dim strRollByRule As String         ' Roll by rule from the futures table
    Dim lDate As Long                   ' Index into a for loop
    Dim lFront As Long                  ' Index for the front contract
    Dim lCurrent As Long                ' Index for the current contract
    Dim lPos1 As Long                   ' Position of date in bars
    Dim lPos2 As Long                   ' Position of date in anohter contract
    Dim dDiff As Double                 ' Difference in closes between contracts
    Dim bRoll As Boolean                ' Do we need to roll the commodity?
    Dim bExpPrev As Boolean             ' Does the commodity expire the previous month?
    Dim strReason As String             ' Reason for the roll
    Dim lLast As Long                   ' Last contract on the date
    
    ' Do some initialization...
    Screen.MousePointer = vbHourglass
    astrContracts.Create eGDARRAY_Strings
    astrRollFile.Create eGDARRAY_Strings
    lLastDate = LastDailyDownload
    
    ' Get the ID for the selected symbol...
    lSymbolID = CLng(ValOfText(Parse(m.astrMarkets(cboSymbol.ListIndex), ";", 2)))
    bExpPrev = (Parse(m.astrMarkets(cboSymbol.ListIndex), ";", 4) = "*")
    
    ' Get the roll by rule for the selected symbol...
    strRollByRule = GetRollByRule(Parse(cboSymbol.Text, "(", 1))
    
    ' Get all of the contracts for the selected symbol...
    If SU_GetContracts(lSymbolID, astrContracts) Then
        ' Load all of the data for all of the contracts...
        For lIndex = 0 To astrContracts.Size - 1
            lSymbolID = CLng(ValOfText(Parse(astrContracts(lIndex), ";", 2)))
            Set aBars(lIndex) = New cGdBars
            DM_GetBars aBars(lIndex), lSymbolID
        Next lIndex
        
        ' Remove any empty sets of bars...
        For lIndex = aBars.Size - 1 To 0 Step -1
            If aBars(lIndex).Size = 0 Then aBars.Remove lIndex
        Next lIndex
        
        If aBars.Size > 0 Then
            ' Add the first entry into the roll file...
            astrRollFile.Add Format(aBars(0).Item(eBARS_DateTime, 0), "YYMMDD") & " " & Format(aBars(0).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(0).PriceDisplay(0#, False) & " " & Str(aBars(0).Prop(eBARS_ConvFactor))
            
            ' Initialize the index variables...
            lFront = 0&
            lCurrent = 0&
    
            ' Go from the first date of the first contract to the last downloaded date...
            For lDate = aBars(0).Item(eBARS_DateTime, 0) To lLastDate
                If lDate >= EarliestCheck(aBars(lCurrent).Prop(eBARS_Contract)) Or IgnoreCheck(aBars(lCurrent).Prop(eBARS_BaseSymbol)) Then
                    If lCurrent + 1 = aBars.Size Then Exit For
                    If aBars(lCurrent).FindDateTime(lDate, True) >= 0 Or aBars(lCurrent + 1).FindDateTime(lDate, True) >= 0 Then
                        If aBars(lFront).FindDateTime(lDate, True) = -1 Then lFront = lFront + 1
                        lPos1 = aBars(lCurrent).FindDateTime(lDate, True)
                        
                        ' Current contract expired, so roll to the next contract...
                        If lPos1 = -1 Then
                            lPos1 = aBars(lCurrent).FindDateTime(lDate) - 1
                            lPos2 = aBars(lCurrent + 1).FindDateTime(lDate) - 1
                        
                            dDiff = aBars(lCurrent + 1).Item(eBARS_Close, lPos2) - aBars(lCurrent).Item(eBARS_Close, lPos1)
                            If chkOutputRule.Value = vbChecked Then
                                astrRollFile.Add Format(aBars(lCurrent).Item(eBARS_DateTime, lPos1), "YYMMDD") & " " & Format(aBars(lCurrent + 1).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor)) & " Expiration"
                            Else
                                astrRollFile.Add Format(aBars(lCurrent).Item(eBARS_DateTime, lPos1), "YYMMDD") & " " & Format(aBars(lCurrent + 1).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor))
                            End If
                            
                            lCurrent = lCurrent + 1
                        Else
                            bRoll = False
                            
                            If IgnoreCheck(aBars(lCurrent).Prop(eBARS_BaseSymbol)) Then
                                lLast = aBars.Size - 1
                            Else
                                lLast = lFront + 4
                            End If
                            
                            'For lIndex = lCurrent + 1 To lFront + 4
                            For lIndex = lCurrent + 1 To lLast
                                If CanRoll(aBars, lCurrent, lIndex, lDate) Or IgnoreCheck(aBars(lCurrent).Prop(eBARS_BaseSymbol)) Then
                                    If GeneralRule(aBars(lCurrent), aBars(lIndex), lDate) Then
                                        strReason = "General Rule"
                                        bRoll = True
                                    ElseIf TwoMajor(aBars(lCurrent), aBars(lIndex), lDate) Then
                                        strReason = "Two Major"
                                        bRoll = True
                                    ElseIf RollByRule(strRollByRule, lDate, aBars(lCurrent).Prop(eBARS_Contract), bExpPrev) Then
                                        strReason = "Roll By Rule"
                                        bRoll = True
                                    ElseIf lIndex = lCurrent + 1 Then
                                        If CloseToExpiration(aBars(lCurrent), aBars(lIndex), lDate, bExpPrev) Then
                                            strReason = "Close to Expiration"
                                            bRoll = True
                                        End If
                                    End If
                                
                                    If bRoll Then
                                        lPos1 = aBars(lCurrent).FindDateTime(lDate)
                                        lPos2 = aBars(lIndex).FindDateTime(lDate)
                                    
                                        'dDiff = aBars(lIndex).Item(eBARS_Close, lPos2 - 1) - aBars(lCurrent).Item(eBARS_Close, lPos1 - 1)
                                        'If chkOutputRule.Value = vbChecked Then
                                        '    astrRollFile.Add Format(aBars(lCurrent).Item(eBARS_DateTime, lPos1), "YYMMDD") & " " & Format(aBars(lIndex).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor)) & " " & strReason
                                        'Else
                                        '    astrRollFile.Add Format(aBars(lCurrent).Item(eBARS_DateTime, lPos1), "YYMMDD") & " " & Format(aBars(lIndex).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor))
                                        'End If
                                        
                                        dDiff = aBars(lIndex).Item(eBARS_Close, lPos2) - aBars(lCurrent).Item(eBARS_Close, lPos1)
                                        If aBars(lIndex).Item(eBARS_DateTime, lPos2) = lDate Then
                                            If chkOutputRule.Value = vbChecked Then
                                                astrRollFile.Add Format(lDate + 1, "YYMMDD") & " " & Format(aBars(lIndex).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor)) & " " & strReason
                                            Else
                                                astrRollFile.Add Format(lDate + 1, "YYMMDD") & " " & Format(aBars(lIndex).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor))
                                            End If
                                        Else
                                            If chkOutputRule.Value = vbChecked Then
                                                astrRollFile.Add Format(aBars(lIndex).Item(eBARS_DateTime, lPos2 + 1), "YYMMDD") & " " & Format(aBars(lIndex).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor)) & " " & strReason
                                            Else
                                                astrRollFile.Add Format(aBars(lIndex).Item(eBARS_DateTime, lPos2 + 1), "YYMMDD") & " " & Format(aBars(lIndex).Prop(eBARS_Contract) Mod 10000, "0000") & " " & aBars(lCurrent).PriceDisplay(dDiff, False) & " " & Str(aBars(lCurrent).Prop(eBARS_ConvFactor))
                                            End If
                                        End If
                                        
                                        lCurrent = lIndex
                                        
                                        Exit For
                                    End If
                                End If
                            Next lIndex
                        End If
                    End If
                End If
            Next lDate
        End If
    ElseIf chkBuildAll.Value = vbUnchecked Then
        Err.Raise vbObjectError + 1000, , "No contracts could be loaded for " & cboSymbol.Text
    End If
    
    ' Output the roll file to the application directory...
    astrRollFile.ToFile AddSlash(App.Path) & Parse(cboSymbol.Text, "(", 1) & "-9957.ROL"
    If chkBackupFiles.Value = vbChecked Then
        astrRollFile.Remove astrRollFile.Size - 1
        astrRollFile.ToFile AddSlash(App.Path) & Parse(cboSymbol.Text, "(", 1) & "-9957.BAK"
    End If

ErrExit:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set aBars = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set aBars = Nothing
    Screen.MousePointer = vbDefault
    RaiseError "frmBuildRolls.Build57ByRule", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Build57ByDates
'' Description: Build a 57 continous contract roll file using the dates in a file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Build57ByDates()
On Error GoTo ErrSection:

    Dim astrContracts As New cGdArray   ' List of contracts for the selected symbol
    Dim astrRollFile As New cGdArray    ' Roll file for the selected symbol
    Dim astrOldRoll As New cGdArray     ' Roll file to retain the dates from
    Dim OldContract As New cGdBars      ' Bars for the old contract
    Dim NewContract As New cGdBars      ' Bars for the new contract
    Dim lSymbolID As Long               ' Symbol ID for the selected symbol
    Dim lNextID As Long                 ' Symbol ID for the next contract
    Dim lIndex As Long                  ' Index into a for loop
    Dim lDate As Long                   ' Date to look up in both contracts
    Dim lContract As Long               ' Current contract
    Dim dDiff As Double                 ' Difference in close of contracts
    Dim lLastDate As Long               ' Date of the last daily download
    Dim lPos1 As Long                   ' Position we want in the first contract
    Dim lPos2 As Long                   ' Position we want in the next contract
    Dim strFileName As String           ' Name of the roll file without the path
    Dim strSymbol As String             ' Symbol to get bars for
    Dim strBaseSym As String            ' Base symbol
    
    ' Do some initialization...
    Screen.MousePointer = vbHourglass
    astrContracts.Create eGDARRAY_Strings
    astrRollFile.Create eGDARRAY_Strings
    astrOldRoll.Create eGDARRAY_Strings
    lLastDate = LastDailyDownload
    strBaseSym = Parse(cboSymbol.Text, "(", 1)
    strFileName = strBaseSym & "-9957.ROL"
    
    ' Try to open the existing roll file...
    If astrOldRoll.FromFile(AddSlash(txtRollPath.Text) & strFileName) = False Then
        Err.Raise vbObjectError + 1000, , AddSlash(txtRollPath.Text) & strFileName & " could not be opened for reading"
    End If
    
    ' Walk through the existing roll file keeping the dates and recalculating
    ' the roll amounts based on the previous day's closes...
    For lIndex = 0 To astrOldRoll.Size - 1
        strSymbol = strBaseSym & "-" & Parse(astrOldRoll(lIndex), " ", 2)
        If lIndex = 0 Then
            astrRollFile.Add astrOldRoll(lIndex)
            NewContract.Prop(eBARS_Symbol) = strSymbol
            DM_GetBars NewContract, NewContract.Prop(eBARS_Symbol)
        Else
            Set OldContract = NewContract.MakeCopy
            NewContract.Prop(eBARS_Symbol) = strSymbol
            DM_GetBars NewContract, NewContract.Prop(eBARS_Symbol)
            
            lDate = AddCentury(CLng(ValOfText(Parse(astrOldRoll(lIndex), " ", 1))))
            lPos1 = OldContract.FindDateTime(lDate)
            lPos2 = NewContract.FindDateTime(lDate)
        
            ' Figure out the difference in closes on the day prior and output...
            If OldContract(eBARS_DateTime, lPos1 - 1) = NewContract(eBARS_DateTime, lPos2 - 1) Then
                dDiff = NewContract(eBARS_Close, lPos2 - 1) - OldContract(eBARS_Close, lPos1 - 1)
                astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
            Else
                dDiff = -99999#
                astrRollFile.Add Format(NewContract(eBARS_DateTime, lPos2), "YYMMDD") & " " & Format(NewContract.Prop(eBARS_Contract) Mod 10000, "0000") & " " & OldContract.PriceDisplay(dDiff, False) & " " & Str(OldContract.Prop(eBARS_ConvFactor))
            End If
        End If
    Next lIndex
    
    ' Output the roll file to the application directory...
    astrRollFile.ToFile AddSlash(App.Path) & strFileName
    If chkBackupFiles.Value = vbChecked Then
        astrRollFile.Remove astrRollFile.Size - 1
        astrRollFile.ToFile AddSlash(App.Path) & Replace(strFileName, ".ROL", ".BAK")
    End If

ErrExit:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set astrOldRoll = Nothing
    Set OldContract = Nothing
    Set NewContract = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    Set astrContracts = Nothing
    Set astrRollFile = Nothing
    Set astrOldRoll = Nothing
    Set OldContract = Nothing
    Set NewContract = Nothing
    Screen.MousePointer = vbDefault
    RaiseError "frmBuildRolls.Build57ByDates", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDiffs_Change
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDiffs_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.optDiffs.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDist_Change
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDist_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.optDist.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDM_Change
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDM_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.optDM.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optExt_Change
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optExt_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.optExt.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optKeepDates_Click
'' Description: When the Keep Dates option is changed, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optKeepDates_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.optKeepDates.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optUseRules_Click
'' Description: When the Use Rules option is changed, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optUseRules_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.optUseRules.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CanRoll
'' Description: Determine if we can roll to the given new contract on a date
'' Inputs:      Array of Contracts, Index of old and new contracts, Date
'' Returns:     True if Can Roll To, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CanRoll(aBars As cGdArray, ByVal lOld As Long, ByVal lNew As Long, ByVal dDate As Double) As Boolean
On Error GoTo ErrSection:

    Dim lPos1 As Long                   ' Position of date in old contract
    Dim lPos2 As Long                   ' Position of date in new contract
    Dim OldBars As New cGdBars          ' Old Contract's data
    Dim NewBars As New cGdBars          ' New Contract's data

    ' Initialize the function to return false...
    CanRoll = False
    
    If lOld >= aBars.Size Or lNew >= aBars.Size Then Exit Function

    ' Get the old and new contracts out of the array...
    Set OldBars = aBars(lOld).MakeCopy
    Set NewBars = aBars(lNew).MakeCopy

    ' Find the position of the date in the two contracts...
    lPos1 = OldBars.FindDateTime(dDate, True)
    lPos2 = NewBars.FindDateTime(dDate, True)
    
    ' If the date exists in the new contract...
    If lPos2 > 1 Then
        ' and the new contract is the next contract...
        If lNew = lOld + 1 Then
            CanRoll = True
        Else
            ' or the OI of the new contract is within 27% of the total OI...
            If (NewBars(eBARS_ContOI, lPos2 - 1) > NewBars(eBARS_OI, lPos2 - 1) * 0.27) Then
                ' and the volume of the new contract for the date or the previous date
                ' is within 33% of the total volume...
                If (NewBars(eBARS_ContVol, lPos2 - 1) > NewBars(eBARS_Vol, lPos2 - 1) * 0.33) Or _
                   (NewBars(eBARS_ContVol, lPos2 - 2) > NewBars(eBARS_Vol, lPos2 - 2) * 0.33) Then
                    CanRoll = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.CanRoll", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GeneralRule
'' Description: Determine if the General Rule for rolling has been hit
'' Inputs:      Old Contract, New Contract, Date to Check
'' Returns:     True if General Rule hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GeneralRule(OldBars As cGdBars, NewBars As cGdBars, ByVal dDate As Double) As Boolean
On Error GoTo ErrSection:

    Dim lPos1 As Long                   ' Position of date in the old contract
    Dim lPos2 As Long                   ' Position of date in the new contract
    
    ' Initialize the function to return false...
    GeneralRule = False
    
    ' Find the date in the two contracts...
    lPos1 = OldBars.FindDateTime(dDate, True)
    lPos2 = NewBars.FindDateTime(dDate, True)
    
    ' Make sure that both contracts have data for previous two days...
    If lPos1 > 1 And lPos2 > 1 Then
        ' If the new contract has more open interest and at least 90% of the volume
        ' of the old contract for two consecutive days...
        If (NewBars(eBARS_ContOI, lPos2 - 1) > OldBars(eBARS_ContOI, lPos1 - 1)) And _
           (NewBars(eBARS_ContOI, lPos2 - 2) > OldBars(eBARS_ContOI, lPos1 - 2)) And _
           (NewBars(eBARS_ContVol, lPos2 - 1) > (OldBars(eBARS_ContVol, lPos1 - 1) * 0.9)) And _
           (NewBars(eBARS_ContVol, lPos2 - 2) > (OldBars(eBARS_ContVol, lPos1 - 2) * 0.9)) Then
            GeneralRule = True
        
        ' or the open interest for the current contract is 2.2 times the old contract
        ' for the date and is 1.8 times the old contract for the previous date and at
        ' least 60% of the volume of the old contract for both days...
        ElseIf (NewBars(eBARS_ContOI, lPos2 - 1) > (OldBars(eBARS_ContOI, lPos1 - 1) * 2.2)) And _
               (NewBars(eBARS_ContOI, lPos2 - 2) > (OldBars(eBARS_ContOI, lPos1 - 2) * 1.8)) And _
               (NewBars(eBARS_ContVol, lPos2 - 1) > (OldBars(eBARS_ContVol, lPos1 - 1) * 0.6)) And _
               (NewBars(eBARS_ContVol, lPos2 - 2) > (OldBars(eBARS_ContVol, lPos1 - 2) * 0.6)) Then
            GeneralRule = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.GeneralRule", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TwoMajor
'' Description: Determine if the Two Major Contracts Rule for rolling has been hit
'' Inputs:      Old Contract, New Contract, Date to Check
'' Returns:     True if Two Major Contracts Rule hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TwoMajor(OldBars As cGdBars, NewBars As cGdBars, ByVal dDate As Double) As Boolean
On Error GoTo ErrSection:

    Dim lPos1 As Long                   ' Position of date in the old contract
    Dim lPos2 As Long                   ' Position of date in the new contract
    
    ' Initialize the function to return false...
    TwoMajor = False
    
    ' Find the date in the two contracts...
    lPos1 = OldBars.FindDateTime(dDate, True)
    lPos2 = NewBars.FindDateTime(dDate, True)
    
    ' Make sure that both contracts have data for the previous day...
    If lPos1 > 0 And lPos2 > 0 Then
        ' If the two contracts make up 90% of the total open interest...
        If (NewBars(eBARS_ContOI, lPos2 - 1) + OldBars(eBARS_ContOI, lPos1 - 1)) > (NewBars(eBARS_OI, lPos2 - 1) * 0.9) And _
           (NewBars(eBARS_ContVol, lPos2 - 1) > (OldBars(eBARS_ContVol, lPos1 - 1) * 0.1)) Then
            ' If the new contracts OI is greater than the old contracts OI + 3%...
            If NewBars(eBARS_ContOI, lPos2 - 1) > (OldBars(eBARS_ContOI, lPos1 - 1) * 1.03) Then
                TwoMajor = True
                
            ' or the new contracts OI is at least 80% of the old contracts OI and the
            ' new contracts volume is 140% of the old contracts volume...
            ElseIf (NewBars(eBARS_ContOI, lPos2 - 1) > (OldBars(eBARS_ContOI, lPos1 - 1) * 0.8)) And _
                   (NewBars(eBARS_ContVol, lPos2 - 1) > (OldBars(eBARS_ContVol, lPos1 - 1) * 1.4)) Then
                TwoMajor = True
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.TwoMajor", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CloseToExpiration
'' Description: Determine if the Close to Expiration Rule for rolling has been hit
'' Inputs:      Old Contract, New Contract, Date to Check
'' Returns:     True if Close to Expiration Rule hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CloseToExpiration(OldBars As cGdBars, NewBars As cGdBars, ByVal dDate As Double, ByVal bExpPrev As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim lPos1 As Long                   ' Position of date in the old contract
    Dim lPos2 As Long                   ' Position of date in the new contract
    Dim lExpMonth As Long               ' Expiration month for the old contract
    Dim lFifteenth As Long              ' 15th of the expiration month
    Dim lNextToLast As Long             ' Next to last business day of the month
    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumDays As Long                ' Days until next to last business day
    Dim dRatio As Double                ' Ratio of open interest and volume
    
    ' Initialize the function to return false...
    CloseToExpiration = False
    
    ' Find the date in the two contracts...
    lPos1 = OldBars.FindDateTime(dDate, True)
    lPos2 = NewBars.FindDateTime(dDate, True)
    
    ' Figure out the expiration month for the contract...
    lExpMonth = OldBars.Prop(eBARS_Contract)
    If bExpPrev Then
        lExpMonth = lExpMonth - 1
        If lExpMonth Mod 100 = 0 Then lExpMonth = (lExpMonth + 12) - 100
    End If
    
    ' Calculate the 15th of the month and the next to last business day...
    lFifteenth = (lExpMonth * 100) + 15
    lNextToLast = GetDateFromRule(lExpMonth / 100&, lExpMonth Mod 100, "LB-1B")
    
    ' Make sure that both contracts have data for the previous day...
    If lPos1 > 0 And lPos2 > 0 Then
        ' If the date is past the 15th of the expiration month...
        If JulToLong(CLng(dDate), 1) >= lFifteenth Then
            ' Figure out the number of days until last business day...
            lNumDays = 0&
            For lIndex = lNextToLast - 1 To CLng(dDate) Step -1
                If IsWeekday(lIndex) Then lNumDays = lNumDays + 1
            Next lIndex
            
            ' As long as the volume and OI are non-zero for the old contract...
            If OldBars(eBARS_ContOI, lPos1 - 1) <> 0 And OldBars(eBARS_ContVol, lPos1 - 1) <> 0 Then
                ' Figure the ratio of volumes and open interests (with open interest
                ' counting double)...
                dRatio = (NewBars(eBARS_ContOI, lPos2 - 1) / OldBars(eBARS_ContOI, lPos1 - 1)) * 2#
                dRatio = (dRatio + (NewBars(eBARS_ContVol, lPos2 - 1) / OldBars(eBARS_ContVol, lPos1 - 1))) / 3#
                
                ' If the ratio is greater than the number of days until the next to
                ' last business day times 0.2...
                If dRatio > lNumDays * 0.2 Then
                    CloseToExpiration = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.CloseToExpiration", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RollByRule
'' Description: Determines if we have a roll based on the date
'' Inputs:      Rule, Date, Current Contract, Expire Previous
'' Returns:     True if rolled by rule, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RollByRule(ByVal strRule As String, ByVal dDate As Double, ByVal lContract As Long, ByVal bExpPrev As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim lRollDate As Long               ' Roll date based on the rule
    Dim lDateToCheck As Long            ' Date to check

    ' Initialize the function to return false...
    RollByRule = False

    ' If the rule is not blank...
    If Len(strRule) > 0 Then
        ' Figure out the contract based on the expires previous flag...
        If bExpPrev Then
            lContract = lContract - 1
            If lContract Mod 100 = 0 Then lContract = lContract - 100 + 12
        End If
    
        ' Get the date that corresponds to the rule...
        lRollDate = GetDateFromRule(lContract / 100&, lContract Mod 100, strRule)
        
        ' Figure out the date to check (next business day after date passed in)...
        lDateToCheck = dDate + 1
        Do While Not IsWeekday(lDateToCheck)
            lDateToCheck = lDateToCheck + 1
        Loop
        
        ' If we are on or past that date...
        If lDateToCheck >= lRollDate Then
            RollByRule = True
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.RollByRule", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetRollByRule
'' Description: Get the roll by rule for a symbol (if there is one)
'' Inputs:      Base Symbol
'' Returns:     Roll By Rule (or blank if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRollByRule(ByVal strBaseSymbol As String) As String
On Error GoTo ErrSection:

    Dim astrFutTbl As New cGdArray      ' Futures table
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Load the futures table...
    If astrFutTbl.FromFile("K:\Common\Futures.TBL") = False Then
        Err.Raise vbObjectError + 1000, , "Could not open K:\Common\Futures.TBL for reading"
    End If
    
    ' Find the base symbol in the futures table and return the rule...
    For lIndex = 0 To astrFutTbl.Size - 1
        If Parse(astrFutTbl(lIndex), " ", 1) = strBaseSymbol Then
            GetRollByRule = Trim(Mid(astrFutTbl(lIndex), 47, 7))
        End If
    Next lIndex

ErrExit:
    Set astrFutTbl = Nothing
    Exit Function
    
ErrSection:
    Set astrFutTbl = Nothing
    RaiseError "frmBuildRolls.GetRollByRule", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EarliestCheck
'' Description: Earliest date to check (1st of the month prior to contract)
'' Inputs:      Contract
'' Returns:     Earliest check date (in Julian)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EarliestCheck(ByVal lContract As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Value to return
    
    lReturn = lContract - 1
    If lReturn Mod 100 = 0 Then lReturn = lReturn - 100 + 12
    lReturn = lReturn * 100 + 1
    lReturn = JulFromLong(lReturn)
    
    EarliestCheck = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.EarliestCheck", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowDiffs
'' Description: Set up and show the form for data comparison
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowDiffs()
On Error GoTo ErrSection:

    m.Mode = eFormMode_DataComp
    fraMain.Visible = False
    fraDataCompare.Visible = True
    fgDiffs.Visible = True
    fraDataCompare.Top = fraMain.Top
    fgDiffs.Top = (fraDataCompare.Top * 2) + fraDataCompare.Height
    
    SetupGrid fgDiffs, eGridMode_List
    fgDiffs.FixedCols = 0
    fgDiffs.Cols = 1
    fgDiffs.FixedRows = 0
    fgDiffs.Rows = 0
    
    txtPath.Text = GetIniFileProperty("Path", "", "DataCompare", g.strIniFile)
    txtOutput.Text = GetIniFileProperty("Output", "", "DataCompare", g.strIniFile)
    With cboFormat
        .AddItem "CSI"
        .AddItem "MS7"
    End With
    cboFormat.Text = GetIniFileProperty("Format", "CSI", "DataCompare", g.strIniFile)
    optDiffs.Value = GetIniFileProperty("Diff", True, "DataCompare", g.strIniFile)
    optDist.Value = Not optDiffs
    optDM.Value = GetIniFileProperty("DM", True, "DataCompare", g.strIniFile)
    optExt.Value = Not optDM
    
    EnableControls
    
    ShowForm Me, True

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmBuildRolls.ShowDiffs", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOutput_Change
'' Description: Enable/Disable controls appropriately as the text changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOutput_Change()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.txtOutput.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildContinuous
'' Description: Build the continuous contracts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildContinuous()
On Error GoTo ErrSection:

    If chk55 = vbChecked Then Build55
    If chk56 = vbChecked Then Build56
    If chk57 = vbChecked Then
        If optKeepDates Then
            Build57ByDates
        Else
            Build57ByRule
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBuildRolls.BuildContinuous", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IgnoreCheck
'' Description: Ignore the check for certain base symbols
'' Inputs:      Base symbol
'' Returns:     True if Ignore, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IgnoreCheck(ByVal strBaseSym As String) As Boolean
On Error GoTo ErrSection:

    If InStr(",JGL,JRU,JKE,", "," & strBaseSym & ",") = 0 Then
        IgnoreCheck = False
    Else
        IgnoreCheck = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBuildRolls.IgnoreCheck"
    
End Function

