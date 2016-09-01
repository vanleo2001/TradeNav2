VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMmCalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Money Management Calculator"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraDrawdown 
      Height          =   3735
      Left            =   5400
      TabIndex        =   0
      Top             =   60
      Width           =   5355
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
      Caption         =   "frmMmCalc.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMmCalc.frx":0044
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMmCalc.frx":0064
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL Frame1 
         Height          =   555
         Left            =   300
         TabIndex        =   2
         Top             =   3000
         Width           =   4815
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
         Caption         =   "frmMmCalc.frx":0080
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMmCalc.frx":00DA
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":00FA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtMargin 
            Height          =   285
            Left            =   180
            TabIndex        =   4
            Top             =   240
            Width           =   915
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmMmCalc.frx":0116
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmMmCalc.frx":0144
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":0164
         End
         Begin HexUniControls.ctlUniLabelXP lblContracts 
            Height          =   255
            Left            =   1440
            Top             =   270
            Width           =   735
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
            Caption         =   "frmMmCalc.frx":0180
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":01A2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":01C2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBigPerc 
            Height          =   255
            Left            =   3900
            Top             =   270
            Width           =   795
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
            Caption         =   "frmMmCalc.frx":01DE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":0206
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":0226
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBigLoss 
            Height          =   255
            Left            =   2340
            Top             =   270
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
            Caption         =   "frmMmCalc.frx":0242
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":0270
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":0290
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   3600
            Top             =   300
            Width           =   195
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
            Caption         =   "frmMmCalc.frx":02AC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":02D0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":02F0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label17 
            Height          =   255
            Left            =   2100
            Top             =   300
            Width           =   195
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
            Caption         =   "frmMmCalc.frx":030C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":032E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":034E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label18 
            Height          =   255
            Left            =   1140
            Top             =   300
            Width           =   195
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
            Caption         =   "frmMmCalc.frx":036A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":038C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":03AC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label22 
            Height          =   255
            Left            =   3840
            Top             =   30
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
            Caption         =   "frmMmCalc.frx":03C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":03F8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":0418
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label24 
            Height          =   255
            Left            =   120
            Top             =   30
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
            Caption         =   "frmMmCalc.frx":0434
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":0470
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":0490
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label25 
            Height          =   255
            Left            =   1440
            Top             =   30
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
            Caption         =   "frmMmCalc.frx":04AC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":04DE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":04FE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label26 
            Height          =   255
            Left            =   2520
            Top             =   30
            Width           =   1095
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
            Caption         =   "frmMmCalc.frx":051A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMmCalc.frx":0556
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMmCalc.frx":0576
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid fgConsec 
         Height          =   1515
         Left            =   240
         TabIndex        =   11
         Top             =   780
         Width           =   4875
         _cx             =   8599
         _cy             =   2672
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
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   315
         Left            =   120
         Top             =   270
         Width           =   5055
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
         Caption         =   "frmMmCalc.frx":0592
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":0622
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0642
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   360
         Top             =   2730
         Width           =   4815
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
         Caption         =   "frmMmCalc.frx":065E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":06FE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":071E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label19 
         Height          =   255
         Left            =   120
         Top             =   2520
         Width           =   5115
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
         Caption         =   "frmMmCalc.frx":073A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":07EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":080C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblConsec 
         Height          =   315
         Left            =   120
         Top             =   540
         Width           =   5055
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
         Caption         =   "frmMmCalc.frx":0828
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":08D2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":08F2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdDrawdown 
      Height          =   435
      Left            =   3420
      TabIndex        =   9
      Top             =   3360
      Width           =   1815
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
      Caption         =   "frmMmCalc.frx":090E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMmCalc.frx":0952
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMmCalc.frx":0972
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdUse 
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
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
      Caption         =   "frmMmCalc.frx":098E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMmCalc.frx":09D8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMmCalc.frx":09F8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   2160
      TabIndex        =   8
      Top             =   3360
      Width           =   975
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
      Caption         =   "frmMmCalc.frx":0A14
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMmCalc.frx":0A40
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMmCalc.frx":0A60
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraCalc 
      Height          =   3195
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   5175
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
      Caption         =   "frmMmCalc.frx":0A7C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMmCalc.frx":0AD4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMmCalc.frx":0AF4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCalculate 
         Default         =   -1  'True
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   2670
         Width           =   3195
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
         Caption         =   "frmMmCalc.frx":0B10
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMmCalc.frx":0B76
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0B96
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtContracts 
         Height          =   315
         Left            =   540
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2700
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   10551295
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmMmCalc.frx":0BB2
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmMmCalc.frx":0BD4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0BF4
      End
      Begin HexUniControls.ctlUniTextBoxXP txtTypicalLoss 
         Height          =   315
         Left            =   540
         TabIndex        =   5
         Top             =   2100
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmMmCalc.frx":0C10
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmMmCalc.frx":0C3C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0C5C
      End
      Begin HexUniControls.ctlUniTextBoxXP txtAcctBal 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1275
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmMmCalc.frx":0C78
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmMmCalc.frx":0CB0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0CD0
      End
      Begin HexUniControls.ctlUniTextBoxXP txtRiskPerc 
         Height          =   315
         Left            =   540
         TabIndex        =   3
         Top             =   1260
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmMmCalc.frx":0CEC
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmMmCalc.frx":0D12
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0D32
      End
      Begin HexUniControls.ctlUniLabelXP lblRiskAmt 
         Height          =   255
         Left            =   360
         Top             =   1740
         Width           =   1095
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
         Caption         =   "frmMmCalc.frx":0D4E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":0D7C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0D9C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   255
         Left            =   180
         Top             =   2730
         Width           =   195
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
         Caption         =   "frmMmCalc.frx":0DB8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":0DDA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0DFA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   180
         Top             =   2160
         Width           =   195
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
         Caption         =   "frmMmCalc.frx":0E16
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":0E38
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0E58
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   180
         Top             =   1740
         Width           =   195
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
         Caption         =   "frmMmCalc.frx":0E74
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":0E96
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0EB6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   180
         Top             =   1260
         Width           =   195
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
         Caption         =   "frmMmCalc.frx":0ED2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":0EF4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":0F14
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   435
         Left            =   180
         Top             =   300
         Width           =   4695
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
         Caption         =   "frmMmCalc.frx":0F30
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":1044
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":1064
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label20 
         Height          =   255
         Left            =   1680
         Top             =   2280
         Width           =   3555
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
         Caption         =   "frmMmCalc.frx":1080
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":1102
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":1122
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label10 
         Height          =   255
         Left            =   1680
         Top             =   2100
         Width           =   3075
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
         Caption         =   "frmMmCalc.frx":113E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":11AA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":11CA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label14 
         Height          =   255
         Left            =   1680
         Top             =   900
         Width           =   3135
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
         Caption         =   "frmMmCalc.frx":11E6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":1226
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":1246
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label13 
         Height          =   255
         Left            =   1680
         Top             =   1320
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
         Caption         =   "frmMmCalc.frx":1262
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":12DC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":12FC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label12 
         Height          =   255
         Left            =   1680
         Top             =   1740
         Width           =   3315
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
         Caption         =   "frmMmCalc.frx":1318
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMmCalc.frx":138A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMmCalc.frx":13AA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   1620
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   1620
         Y1              =   1680
         Y2              =   1680
      End
   End
End
Attribute VB_Name = "frmMmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    nContracts As Long
    dAcctBal As Double
    dRiskPerc As Double
    dTypicalLoss As Double
    dMargin As Double
    bShowDrawdown As Boolean
    bInitialShow As Boolean
    nReturn As Long
    strSymbol As String
End Type
Dim m As mPrivate

Private Sub cmdCalculate_Click()

    Calculate
    SelectAll ActiveControl

End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDrawdown_Click()
    m.bShowDrawdown = Not m.bShowDrawdown
    ShowHideDrawdown
End Sub

Private Sub cmdUse_Click()
    
    SetIniFileProperty "Sym:" & Parse(m.strSymbol, "-", 1), m.dTypicalLoss, "TypicalLoss", g.strIniFile
    
    m.nReturn = m.nContracts
    Me.Hide

End Sub

Private Sub Form_Activate()

    If m.bInitialShow Then
        m.bInitialShow = False
        MoveFocus txtRiskPerc
        SelectAll txtRiskPerc
    End If

End Sub

Private Sub Form_Load()

    Me.Icon = Picture16(ToolbarIcon("ID_Criteria"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    'txtContracts.BackColor = ALT_GRID_ROW_COLOR

End Sub

Public Function ShowMe(Optional ByVal vSymbol As Variant = 0, Optional ByVal lAccountID As Long = 0) As Long

    Dim Bars As New cGdBars

    If lAccountID <> 0 Then
        ' get account balance for this account
        m.dAcctBal = 0#
        If Not g.Broker.Account(lAccountID) Is Nothing Then
            m.dAcctBal = g.Broker.Account(lAccountID).CurrentBalance
        End If
        cmdUse.Caption = "&Trade"
    Else
        cmdUse.Caption = "&Save"
    End If
    
    m.dMargin = 0
    m.strSymbol = GetSymbol(vSymbol)
    If Len(m.strSymbol) > 0 Then
        ' get margin for this symbol
        SetBarProperties Bars, m.strSymbol
        m.dMargin = Bars.Prop(eBARS_Margin)
    End If
    
    ' get typical loss
    m.dTypicalLoss = GetIniFileProperty("Sym:" & Parse(m.strSymbol, "-", 1), 0, "TypicalLoss", g.strIniFile)
    
    ' defaults
    If m.dAcctBal <= 0 Then m.dAcctBal = 100000
    If m.dRiskPerc <= 0 Then m.dRiskPerc = 10
    If m.dMargin <= 0 Then m.dMargin = 20000
    If m.dTypicalLoss <= 0 Then m.dTypicalLoss = Int(m.dMargin / 3)

    m.bInitialShow = True
    ShowHideDrawdown
    Calculate True
    m.nReturn = 0
    ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
    
    ShowMe = m.nReturn
    Unload Me

End Function

Private Sub LoadGrid()

    Dim i&, dAcctBal#, nContracts&, dLoss#

    SetupGrid fgConsec, eGridMode_Grid
    With fgConsec
        .Redraw = flexRDNone
        .Editable = flexEDNone
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .HighLight = flexHighlightNever
        .ScrollBars = flexScrollBarVertical
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        ' loss#, contracts, loss amt, acct bal, drawdown
        .Cols = 5
        .FixedCols = 0
        .FixedRows = 1
        .TextMatrix(0, 0) = "Loss#"
        .TextMatrix(0, 1) = "Contracts"
        .TextMatrix(0, 2) = "Loss Amount"
        .TextMatrix(0, 3) = "Acct Balance"
        .TextMatrix(0, 4) = "Drawdown"
        .ColFormat(2) = "$#,###" '"Currency"
        .ColFormat(3) = .ColFormat(2)
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        
        .Rows = 100
        dAcctBal = m.dAcctBal
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
            nContracts = CalcContracts(dAcctBal)
            .TextMatrix(i, 1) = nContracts
            If nContracts = 0 Then
                .TextMatrix(i, 2) = ""
                .TextMatrix(i, 3) = ""
                .TextMatrix(i, 4) = ""
                .Rows = i + 1
                Exit For
            End If
            dLoss = m.dTypicalLoss * nContracts
            .TextMatrix(i, 2) = dLoss
            dAcctBal = dAcctBal - dLoss
            .TextMatrix(i, 3) = dAcctBal
            .TextMatrix(i, 4) = Str(Round(100 * dAcctBal / m.dAcctBal - 100)) & "%"
        Next
        
        .Cell(flexcpFontBold, 1, 4, .Rows - 1, 4) = True
        .AutoSize 0, .Cols - 1
        .Row = -1
        .Redraw = flexRDBuffered
    End With

End Sub

Private Function CalcContracts(Optional ByVal dAcctBal# = -1) As Long

    If dAcctBal = -1 Then
        dAcctBal = m.dAcctBal
    End If
    If m.dTypicalLoss > 0 And dAcctBal > 0 And m.dRiskPerc > 0 Then
        CalcContracts = Int(dAcctBal * m.dRiskPerc / m.dTypicalLoss / 100)
    Else
        CalcContracts = 0
    End If

End Function

Private Sub Calculate(Optional ByVal bInit As Boolean = False)

    If Not bInit Then m.dAcctBal = ValOfText(txtAcctBal)
    If m.dAcctBal <= 0 Or m.dAcctBal > 999999999# Then
        m.dAcctBal = 100000
    End If
    txtAcctBal = FormatCurr(m.dAcctBal)
    
    If Not bInit Then m.dRiskPerc = Round(ValOfText(StripStr(txtRiskPerc, "%")), 2)
    If m.dRiskPerc > 100 Then
        m.dRiskPerc = 100
    ElseIf m.dRiskPerc <= 0 Then
        m.dRiskPerc = 10
    End If
    txtRiskPerc = FormatNum(m.dRiskPerc) & "%"
    lblRiskAmt = FormatCurr(m.dAcctBal * m.dRiskPerc / 100)
    lblConsec = "1)  Result of consecutive typical losses when trading " & txtRiskPerc & " of account:"
    
    If Not bInit Then m.dMargin = ValOfText(txtMargin)
    If m.dMargin <= 0 Then
        m.dMargin = 10000
    End If
    txtMargin = FormatCurr(m.dMargin)
    
    If Not bInit Then m.dTypicalLoss = ValOfText(txtTypicalLoss)
    If m.dTypicalLoss <= 0 Then
        m.dTypicalLoss = m.dMargin / 4
    End If
    txtTypicalLoss = FormatCurr(m.dTypicalLoss)
    
    m.nContracts = CalcContracts
    txtContracts = Str(m.nContracts)
    lblContracts = Str(m.nContracts)
    
    If Left(cmdUse.Caption, 2) = "&T" Then
        If m.nContracts = 1 Then
            cmdUse.Caption = "&Trade 1 contract"
        ElseIf m.nContracts > 9999 Then
            cmdUse.Caption = "&Trade " & Format(m.nContracts, "#,###")
        Else
            cmdUse.Caption = "&Trade " & Str(m.nContracts) & " contracts"
        End If
    End If
    
    lblBigLoss = FormatCurr(m.dMargin * m.nContracts)
    If m.dAcctBal > 0 Then
        lblBigPerc = Str(Round(-100 * m.dMargin * m.nContracts / m.dAcctBal)) & "%"
    Else
        lblBigPerc = "0%"
    End If
    
    LoadGrid

End Sub

Private Sub ShowHideDrawdown()

    If m.bShowDrawdown Then
        cmdDrawdown.Caption = "Hide &Drawdown <<<"
        fraDrawdown.Visible = True
        Me.Width = fraDrawdown.Left + fraDrawdown.Width + fraCalc.Left * 2 + Me.Width - Me.ScaleWidth
    Else
        cmdDrawdown.Caption = "Show &Drawdown >>>"
        fraDrawdown.Visible = False
        Me.Width = fraCalc.Width + fraCalc.Left * 2 + Me.Width - Me.ScaleWidth
    End If

End Sub

Private Function FormatCurr(ByVal dAmount As Double) As String

    If dAmount < 10000 Then
        FormatCurr = Format(dAmount, "$#0")
    Else
        FormatCurr = Format(dAmount, "$#,##0")
    End If

End Function

Private Sub txtAcctBal_GotFocus()
    SelectAll txtAcctBal
End Sub

Private Sub txtMargin_GotFocus()
    SelectAll txtMargin
End Sub

Private Sub txtRiskPerc_GotFocus()
    SelectAll txtRiskPerc
End Sub

Private Sub txtTypicalLoss_GotFocus()
    SelectAll txtTypicalLoss
End Sub

