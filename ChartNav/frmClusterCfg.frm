VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmClusterCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fibonacci Clusters"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   6285
      Left            =   135
      TabIndex        =   3
      Top             =   30
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   11086
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
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "Swing Points|Price Cluster|Time Cluster"
      Align           =   0
      Appearance      =   1
      CurrTab         =   2
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin HexUniControls.ctlUniFrameWL fraRatiosTime 
         Height          =   5910
         Left            =   45
         TabIndex        =   14
         Top             =   330
         Width           =   4800
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
         Caption         =   "frmClusterCfg.frx":0000
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmClusterCfg.frx":0020
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmClusterCfg.frx":0040
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtMaxBars 
            Height          =   285
            Left            =   3060
            TabIndex        =   26
            Top             =   1375
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frmClusterCfg.frx":005C
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
            Tip             =   "frmClusterCfg.frx":0082
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":00A2
         End
         Begin HexUniControls.ctlUniCheckXP chkL2H 
            Height          =   225
            Left            =   255
            TabIndex        =   25
            Top             =   3100
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":00BE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":00F4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0114
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkL2L 
            Height          =   225
            Left            =   255
            TabIndex        =   24
            Top             =   3480
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":0130
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":0164
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0184
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtProxTime 
            Height          =   285
            Left            =   3060
            TabIndex        =   23
            Top             =   1800
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmClusterCfg.frx":01A0
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
            Tip             =   "frmClusterCfg.frx":01C2
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":01E2
         End
         Begin HexUniControls.ctlUniFrameWL Frame2 
            Height          =   1020
            Left            =   255
            TabIndex        =   19
            Top             =   4170
            Width           =   1665
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
            Caption         =   "frmClusterCfg.frx":01FE
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmClusterCfg.frx":021E
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":023E
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdRestoreTime 
               Height          =   315
               Left            =   0
               TabIndex        =   22
               Top             =   690
               Width           =   1440
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
               Caption         =   "frmClusterCfg.frx":025A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":029A
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":02BA
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdRemoveTime 
               Height          =   315
               Left            =   0
               TabIndex        =   21
               Top             =   345
               Width           =   1440
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
               Caption         =   "frmClusterCfg.frx":02D6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":030E
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":032E
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdAddTime 
               Height          =   315
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1440
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
               Caption         =   "frmClusterCfg.frx":034A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":037C
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":039C
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkH2H 
            Height          =   225
            Left            =   255
            TabIndex        =   18
            Top             =   2340
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":03B8
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":03F0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0410
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkH2L 
            Height          =   225
            Left            =   255
            TabIndex        =   17
            Top             =   2720
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":042C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":0462
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0482
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboLabelTime 
            Height          =   315
            Left            =   2010
            TabIndex        =   16
            Top             =   920
            Width           =   1635
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
            Tip             =   "frmClusterCfg.frx":049E
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":04BE
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkShowTime 
            Height          =   300
            Left            =   240
            TabIndex        =   15
            Top             =   120
            Width           =   1965
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
            Caption         =   "frmClusterCfg.frx":04DA
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":0514
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0534
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgRatiosTime 
            Height          =   2850
            Left            =   2055
            TabIndex        =   27
            Top             =   2340
            Width           =   2460
            _cx             =   4339
            _cy             =   5027
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
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
         Begin gdOCX.gdSelectColor gdColorTime 
            Height          =   315
            Left            =   2010
            TabIndex        =   28
            Top             =   465
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   285
            Left            =   605
            Top             =   1375
            Visible         =   0   'False
            Width           =   2385
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
            Caption         =   "frmClusterCfg.frx":0550
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":05AC
            Style           =   0
            Enabled         =   0   'False
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":05CC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label8 
            Height          =   285
            Left            =   605
            Top             =   1800
            Width           =   2385
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
            Caption         =   "frmClusterCfg.frx":05E8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":062C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":064C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label6 
            Height          =   255
            Left            =   1365
            Top             =   950
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
            Caption         =   "frmClusterCfg.frx":0668
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":0692
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":06B2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   1365
            Top             =   495
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
            Caption         =   "frmClusterCfg.frx":06CE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":06FA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":071A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraRatiosPrice 
         Height          =   5910
         Left            =   -5445
         TabIndex        =   13
         Top             =   330
         Width           =   4800
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
         Caption         =   "frmClusterCfg.frx":0736
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmClusterCfg.frx":0756
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmClusterCfg.frx":0776
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkShowPrice 
            Height          =   300
            Left            =   255
            TabIndex        =   8
            Top             =   235
            Width           =   2310
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
            Caption         =   "frmClusterCfg.frx":0792
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":07CC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":07EC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkABCs 
            Height          =   225
            Left            =   255
            TabIndex        =   29
            Top             =   3240
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":0808
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":0830
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0850
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtProxPrice 
            Height          =   285
            Left            =   2055
            TabIndex        =   30
            Top             =   1920
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmClusterCfg.frx":086C
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
            Tip             =   "frmClusterCfg.frx":088E
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":08AE
         End
         Begin HexUniControls.ctlUniFrameWL frame1 
            Height          =   1020
            Left            =   240
            TabIndex        =   31
            Top             =   4290
            Width           =   1665
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
            Caption         =   "frmClusterCfg.frx":08CA
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmClusterCfg.frx":08EA
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":090A
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdRestorePrice 
               Height          =   315
               Left            =   0
               TabIndex        =   32
               Top             =   690
               Width           =   1440
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
               Caption         =   "frmClusterCfg.frx":0926
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":0966
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":0986
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdRemovePrice 
               Height          =   315
               Left            =   0
               TabIndex        =   35
               Top             =   345
               Width           =   1440
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
               Caption         =   "frmClusterCfg.frx":09A2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":09DA
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":09FA
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdAddPrice 
               Height          =   315
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   1440
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
               Caption         =   "frmClusterCfg.frx":0A16
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":0A48
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":0A68
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkResistance 
            Height          =   225
            Left            =   255
            TabIndex        =   42
            Top             =   2460
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":0A84
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":0AB8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0AD8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSupport 
            Height          =   225
            Left            =   255
            TabIndex        =   43
            Top             =   2850
            Width           =   1500
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
            Caption         =   "frmClusterCfg.frx":0AF4
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmClusterCfg.frx":0B22
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0B42
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboLabelPrice 
            Height          =   315
            Left            =   2055
            TabIndex        =   52
            Top             =   1320
            Width           =   1635
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
            Tip             =   "frmClusterCfg.frx":0B5E
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0B7E
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgRatiosPrice 
            Height          =   2850
            Left            =   2055
            TabIndex        =   45
            Top             =   2460
            Width           =   2460
            _cx             =   4339
            _cy             =   5027
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
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
         Begin gdOCX.gdSelectColor gdColorPrice 
            Height          =   315
            Left            =   2055
            TabIndex        =   47
            Top             =   840
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   600
            Top             =   1935
            Width           =   1305
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
            Caption         =   "frmClusterCfg.frx":0B9A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":0BDC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0BFC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   600
            Top             =   1350
            Width           =   1305
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
            Caption         =   "frmClusterCfg.frx":0C18
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":0C42
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0C62
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   255
            Left            =   600
            Top             =   870
            Width           =   1305
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
            Caption         =   "frmClusterCfg.frx":0C7E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":0CAA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0CCA
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSwing 
         Height          =   5910
         Left            =   -5745
         TabIndex        =   4
         Top             =   330
         Width           =   4800
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
         Caption         =   "frmClusterCfg.frx":0CE6
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmClusterCfg.frx":0D06
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmClusterCfg.frx":0D26
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraSwingInd 
            Height          =   3450
            Left            =   135
            TabIndex        =   12
            Top             =   2400
            Width           =   4485
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
            Caption         =   "frmClusterCfg.frx":0D42
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmClusterCfg.frx":0D62
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":0D82
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL Frame3 
               Height          =   660
               Left            =   120
               TabIndex        =   49
               Top             =   1380
               Width           =   4245
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
               Caption         =   "frmClusterCfg.frx":0D9E
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmClusterCfg.frx":0DFC
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":0E1C
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniCheckXP chk2nd 
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   51
                  Top             =   270
                  Width           =   1575
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
                  Caption         =   "frmClusterCfg.frx":0E38
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmClusterCfg.frx":0E70
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmClusterCfg.frx":0E90
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniCheckXP chk1st 
                  Height          =   255
                  Left            =   360
                  TabIndex        =   50
                  Top             =   270
                  Width           =   1575
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
                  Caption         =   "frmClusterCfg.frx":0EAC
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmClusterCfg.frx":0EE6
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmClusterCfg.frx":0F06
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin gdOCX.gdSelectColor gdZoneColor 
               Height          =   315
               Left            =   1845
               TabIndex        =   44
               Top             =   935
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniFrameWL Frame4 
               Height          =   1245
               Left            =   120
               TabIndex        =   36
               Top             =   2140
               Width           =   4245
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
               Caption         =   "frmClusterCfg.frx":0F22
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmClusterCfg.frx":0F68
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":0F88
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniButtonImageXP cmdFont 
                  Height          =   345
                  Left            =   3030
                  TabIndex        =   37
                  Top             =   555
                  Width           =   960
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
                  Caption         =   "frmClusterCfg.frx":0FA4
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  ShowFocus       =   -1  'True
                  Tristate        =   0   'False
                  Pressed         =   0   'False
                  Tip             =   "frmClusterCfg.frx":0FCC
                  Style           =   -1
                  RoundedBorders  =   -1  'True
                  xTranspColor    =   0
                  yTranspColor    =   0
                  MousePointer    =   0
                  MouseIcon       =   "frmClusterCfg.frx":0FEC
                  RightToLeft     =   0   'False
               End
               Begin gdOCX.gdSelectColor gdColorLong 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   38
                  Top             =   180
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin gdOCX.gdSelectColor gdColorMedium 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   39
                  Top             =   525
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin gdOCX.gdSelectColor gdColorShort 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   40
                  Top             =   870
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin HexUniControls.ctlUniLabelXP lblSwingShort 
                  Height          =   240
                  Left            =   300
                  Top             =   900
                  Width           =   1245
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
                  Caption         =   "frmClusterCfg.frx":1008
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   1
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmClusterCfg.frx":103A
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmClusterCfg.frx":105A
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblSwingMedium 
                  Height          =   240
                  Left            =   300
                  Top             =   570
                  Width           =   1245
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
                  Caption         =   "frmClusterCfg.frx":1076
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   1
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmClusterCfg.frx":10B6
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmClusterCfg.frx":10D6
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblSwingLong 
                  Height          =   240
                  Left            =   300
                  Top             =   210
                  Width           =   1245
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
                  Caption         =   "frmClusterCfg.frx":10F2
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   1
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmClusterCfg.frx":1122
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmClusterCfg.frx":1142
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniCheckXP chkSwingShow 
               Height          =   220
               Left            =   210
               TabIndex        =   34
               Top             =   0
               Width           =   2265
               _ExtentX        =   3995
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
               Caption         =   "frmClusterCfg.frx":115E
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":11B2
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":11D2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniComboImageXP cboSwingStyle 
               Height          =   315
               Left            =   1845
               TabIndex        =   33
               Top             =   210
               Width           =   1635
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
               Tip             =   "frmClusterCfg.frx":11EE
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":120E
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectColor gdColor 
               Height          =   315
               Left            =   1845
               TabIndex        =   48
               Top             =   572
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniLabelXP Label11 
               Height          =   210
               Left            =   810
               Top             =   624
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
               Caption         =   "frmClusterCfg.frx":122A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmClusterCfg.frx":1260
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":1280
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label10 
               Height          =   210
               Left            =   810
               Top             =   987
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
               Caption         =   "frmClusterCfg.frx":129C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmClusterCfg.frx":12D2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":12F2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label9 
               Height          =   210
               Left            =   810
               Top             =   255
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
               Caption         =   "frmClusterCfg.frx":130E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmClusterCfg.frx":1344
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":1364
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraSwingLevels 
            Height          =   1470
            Left            =   135
            TabIndex        =   9
            Top             =   870
            Width           =   4485
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
            Caption         =   "frmClusterCfg.frx":1380
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmClusterCfg.frx":13C6
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":13E6
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkSwingDummy 
               Height          =   225
               Left            =   4200
               TabIndex        =   10
               Top             =   60
               Visible         =   0   'False
               Width           =   180
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
               Caption         =   "frmClusterCfg.frx":1402
               Enabled         =   0   'False
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmClusterCfg.frx":143C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmClusterCfg.frx":145C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin VSFlex7LCtl.VSFlexGrid fgSwing 
               Height          =   1035
               Left            =   90
               TabIndex        =   11
               Top             =   315
               Width           =   4275
               _cx             =   7541
               _cy             =   1826
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
               ExtendLastCol   =   -1  'True
               FormatString    =   ""
               ScrollTrack     =   -1  'True
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
         Begin HexUniControls.ctlUniTextBoxXP txtBarsBack 
            Height          =   345
            Left            =   255
            TabIndex        =   7
            Top             =   180
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmClusterCfg.frx":1478
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
            Tip             =   "frmClusterCfg.frx":149E
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":14BE
         End
         Begin gdOCX.gdSelectDate gdEndDate 
            Height          =   345
            Left            =   2280
            TabIndex        =   6
            Top             =   180
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   609
         End
         Begin gdOCX.gdSelectDate gdBarTime 
            Height          =   345
            Left            =   2280
            TabIndex        =   46
            Top             =   555
            Visible         =   0   'False
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   609
            Enabled         =   0   'False
            ShowDayOfWeek   =   0   'False
            ShowCalendar    =   0   'False
            ShowPM          =   2
            ShowDate        =   0
            ShowTime        =   2
            MinDate         =   0
            MaxDate         =   0.99999
            Value           =   0
         End
         Begin HexUniControls.ctlUniLabelXP Label7 
            Height          =   285
            Left            =   840
            Top             =   210
            Width           =   1455
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
            Caption         =   "frmClusterCfg.frx":14DA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmClusterCfg.frx":1516
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmClusterCfg.frx":1536
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   600
      Left            =   825
      TabIndex        =   0
      Top             =   6230
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
      Caption         =   "frmClusterCfg.frx":1552
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmClusterCfg.frx":1572
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmClusterCfg.frx":1592
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   330
         Left            =   2565
         TabIndex        =   2
         Top             =   180
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
         Caption         =   "frmClusterCfg.frx":15AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmClusterCfg.frx":15DC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmClusterCfg.frx":15FC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   60
         TabIndex        =   5
         Top             =   180
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
         Caption         =   "frmClusterCfg.frx":1618
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmClusterCfg.frx":163E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmClusterCfg.frx":165E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveDefaults 
         Height          =   330
         Left            =   960
         TabIndex        =   1
         Top             =   180
         Width           =   1455
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
         Caption         =   "frmClusterCfg.frx":167A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmClusterCfg.frx":16BC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmClusterCfg.frx":16DC
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmClusterCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kFraLevelsTop = 870
Private Const kInfSwingTitle = "Swing Point Options"
Private Const kInfRatioTitle = "Ratio Options"

Private Enum eTabIndex
    eTab_ClusterPrice = 0
    eTab_ClusterTime
    eTab_Ratios
End Enum

Private Enum eRatioCol
    eRatioCol_Use = 0
    eRatioCol_Ratio
    eRatioCol_Weight
End Enum

Private Enum eSwingCol
    eSwingCol_Use = 0
    eSwingCol_Name
    eSwingCol_Strength
    eSwingCol_Weight
End Enum

Private Enum eSwingRow
    eSwingRow_Long = 1
    eSwingRow_Intermediate
    eSwingRow_Short
End Enum

Private Type mPrivate
    tbSwing As cGdTable
    tbRatioPrice As cGdTable
    tbRatioTime As cGdTable
    
    Chart As cChart
    IndPrice As cIndicator
    IndTime As cIndicator
    
    bUnload As Boolean
    
    lMouseDownRow As Long
    lMouseDownCol As Long
End Type
Private m As mPrivate


Public Sub ShowMe(Chart As cChart, Optional Ind As cIndicator = Nothing)
On Error GoTo ErrSection:

    If Chart Is Nothing Then Exit Sub
    If Chart.Bars Is Nothing Then Exit Sub
    If Chart.Tree Is Nothing Then Exit Sub
    
    Set m.IndPrice = Chart.Tree(kClusterPriceKey)
    If m.IndPrice Is Nothing Then Exit Sub
    
    Set m.IndTime = Chart.Tree(kClusterTimeKeyInd)
    If m.IndTime Is Nothing Then Exit Sub
    
    Set m.Chart = Chart
    m.bUnload = False
    
    If InitControls() Then
        vsTab.BoldCurrent = True
        If Not Ind Is Nothing Then
            If Ind.DisplayType = eINDIC_ClusterPrice Then
                vsTab.CurrTab = 1
            Else
                vsTab.CurrTab = 2
            End If
        End If
        CenterFormOnChart Me, Chart
        ShowForm Me, eForm_Nonmodal
    Else
        Unload Me
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.ShowMe"

End Sub

Private Sub cboLabelPrice_Change()
    If Me.Visible Then Repaint
End Sub

Private Sub cboLabelTime_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub cboSwingStyle_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chk1st_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chk2nd_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkABCs_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkH2H_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkH2L_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkL2H_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkL2L_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkResistance_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If chkResistance.Value = vbUnchecked Then
            If chkSupport.Value = vbUnchecked Then
                chkSupport.Value = vbChecked
                GoTo ErrExit
            End If
        End If
        Repaint
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.chkResistance_Click"

End Sub

Private Sub chkShowPrice_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then
        Repaint
        If m.bUnload Then
            m.Chart.GenerateChart eRedo5_RecalcInd
            Unload Me
        ElseIf Not m.IndPrice Is Nothing Then
            chkShowPrice.Value = m.IndPrice.ClusterPriceShow
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.chkShowPrice_Click"
    Unload Me

End Sub

Private Sub chkShowTime_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then
        Repaint
        If m.bUnload Then
            m.Chart.GenerateChart eRedo5_RecalcInd
            Unload Me
        ElseIf Not m.IndTime Is Nothing Then
            chkShowTime.Value = Abs(m.IndTime.Display)
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.chkShowTime_Click"
    Unload Me

End Sub

Private Sub chkSupport_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If chkSupport.Value = vbUnchecked Then
            If chkResistance.Value = vbUnchecked Then
                chkResistance.Value = vbChecked
                GoTo ErrExit
            End If
        End If
        Repaint
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.chkSupport_Click"

End Sub

Private Sub chkSwingShow_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        Repaint
        If Not m.bUnload Then
            If Not m.IndPrice Is Nothing Then
                chkSwingShow.Value = m.IndPrice.ClusterSwingShow
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.chkSwingShow_Click"

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:
    
    Me.Hide
    
    If Not m.Chart Is Nothing Then
        m.Chart.RemoveFibClusters
        m.Chart.GenerateChart eRedo5_RecalcInd
    End If
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.cmdDelete_Click"
    Unload Me

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    Dim nStyle&
    
    If m.IndPrice Is Nothing Then Exit Sub
    
    'set font currently in use
    With m.IndPrice
        Me.Font.Name = .FontName
        Me.Font.Size = .FontSize
        Me.Font.Underline = False
        Me.Font.Bold = .FontBold
        Me.Font.Italic = .FontItalic
    End With
    
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        With m.IndPrice
            .FontName = Me.Font.Name
            .FontSize = Me.Font.Size
            .FontBold = Me.Font.Bold
            .FontItalic = Me.FontItalic
        End With
        
        If Not m.Chart Is Nothing Then m.Chart.GenerateChart eRedo3_Settings
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmClusterCfg.cmdFont.Click", eGDRaiseError_Show
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrExit

    m.bUnload = True
    Repaint True
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.cmdOK_Click"

End Sub

Private Sub cmdAddTime_Click()
    RatioAdd fgRatiosTime, m.IndTime
End Sub

Private Sub cmdRemoveTime_Click()
    RatioRemove fgRatiosTime, m.IndTime
End Sub

Private Sub cmdRestoreTime_Click()
    RatioRestore fgRatiosTime, m.IndTime
End Sub

Private Sub cmdAddPrice_Click()
    RatioAdd fgRatiosPrice, m.IndPrice
End Sub

Private Sub cmdRemovePrice_Click()
    RatioRemove fgRatiosPrice, m.IndPrice
End Sub

Private Sub cmdRestorePrice_Click()
    RatioRestore fgRatiosPrice, m.IndPrice
End Sub

Private Sub cmdSaveDefaults_Click()
    m.bUnload = True
    Repaint False, True
    Unload Me
End Sub

Private Sub fgRatiosPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim i&, j&, d1#, d2#
    Dim bOK As Boolean

    If Row < fgRatiosPrice.FixedRows Or Row >= fgRatiosPrice.Rows Then Exit Sub
    
    If Col = eRatioCol_Use Then
        With fgRatiosPrice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, Col) = flexChecked Then
                    bOK = True
                    Exit For
                End If
            Next
        End With
        
        If Not bOK Then
            InfBox "At least one ratio must be selected.", "I", "Ok", kInfRatioTitle
        End If

    ElseIf Col = eRatioCol_Ratio Or Col = eRatioCol_Weight Then
        d1 = ValOfText(fgRatiosPrice.TextMatrix(Row, Col))
        
        Select Case Col
            Case eRatioCol_Ratio
                If d1 <> 0# Then
                    bOK = True
                Else
                    InfBox "Ratio for price cluster cannot be zero."
                End If
            Case eRatioCol_Weight
                If d1 > 0# Then
                    bOK = True
                Else
                    InfBox "Weight value must be greater than zero."
                End If
        End Select
    End If
    
    If bOK Then
        UpdateRatioInfo fgRatiosPrice, m.IndPrice
    Else
        SetTabRatios fgRatiosPrice, m.IndPrice
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.fgRatiosPrice_AfterEdit"

End Sub

Private Sub fgRatiosPrice_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    m.lMouseDownRow = -1
    m.lMouseDownCol = -1
End Sub

Private Sub fgRatiosPrice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleGridMouseDown fgRatiosPrice
End Sub

Private Sub fgRatiosTime_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim i&, j&, d1#, d2#
    Dim bOK As Boolean

    If Row < fgRatiosPrice.FixedRows Or Row >= fgRatiosPrice.Rows Then Exit Sub
    
    If Col = eRatioCol_Use Then
        With fgRatiosTime
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, Col) = flexChecked Then
                    bOK = True
                    Exit For
                End If
            Next
        End With
        
        If Not bOK Then
            InfBox "At least one ratio must be selected.", "I", "Ok", kInfRatioTitle
        End If

    ElseIf Col = eRatioCol_Ratio Or Col = eRatioCol_Weight Then
        d1 = ValOfText(fgRatiosTime.TextMatrix(Row, Col))
        If d1 > 0# Then
            bOK = True
        ElseIf Col = eRatioCol_Ratio Then
            InfBox "Ratio for time cluster must be greater than zero."
        Else
            InfBox "Weight value must be greater than zero."
        End If
    End If
    
    If bOK Then
        UpdateRatioInfo fgRatiosTime, m.IndTime
    Else
        SetTabRatios fgRatiosTime, m.IndTime
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.fgRatiosTime_AfterEdit"

End Sub

Private Sub fgRatiosTime_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    m.lMouseDownCol = -1
    m.lMouseDownRow = -1
End Sub

Private Sub fgRatiosTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleGridMouseDown fgRatiosTime
End Sub

Private Sub fgSwing_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    If ValidateSwingEdit(fgSwing, Row, Col) Then
        UpdateSwingInfo fgSwing, Row, Col
    Else
        SetGridSwing fgSwing
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.fgSwing_AfterEdit"

End Sub

Private Sub fgSwing_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    If Row >= fgSwing.FixedRows And Row < fgSwing.Rows Then
        If Col = eSwingCol_Name Then
            Cancel = True
        ElseIf Col = eSwingCol_Use And Row = eSwingRow_Long Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.fgSwing_BeforeEdit"

End Sub

Private Sub fgSwing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleGridMouseDown fgSwing
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    If m.IndPrice Is Nothing Then Exit Sub
    
    Me.Caption = "Fib Clusters"
    Me.Icon = Picture16(ToolbarIcon("ID_FibClusters"), , True)
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.Form_Load"

End Sub

Private Sub InitCboLabels(cbo As ctlUniComboImageXP)
    
    With cbo
        .Clear
        .AddItem "(default)"
        .AddItem "Value in axis"
        .AddItem "Value in label"
        .AddItem "Value in label/axis"
        .AddItem "No values"
        .AddItem "No labels"
        .AddItem "No labels or values"
        .AddItem "Only values"
    End With

End Sub

Private Sub InitRatiosGrid(fg As VSFlexGrid)
On Error GoTo ErrSection:
    
    With fg
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ScrollTrack = True
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionFree
        .AllowSelection = True
        .HighLight = flexHighlightNever             '6056
        .SheetBorder = RGB(128, 128, 128)
        .Editable = flexEDKbdMouse
        
        .Rows = .FixedRows
        .Cols = 3
        .TextMatrix(0, 0) = "Use"
        .TextMatrix(0, 1) = "Ratio"
        .TextMatrix(0, 2) = "Weight"
        .FillStyle = flexFillRepeat
        .CellFontBold = True
        .AutoSize 1
        .ColWidth(0) = 500
        .ColWidth(1) = 900
        .ExtendLastCol = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        .FillStyle = flexFillSingle
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.InitRatiosGrid"

End Sub

Private Sub InitSwingGrid(fg As VSFlexGrid)
On Error GoTo ErrSection:

    With fg
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ScrollTrack = True
        .ScrollBars = flexScrollBarNone
        .SelectionMode = flexSelectionFree
        .AllowSelection = True
        .HighLight = flexHighlightNever         '6056
        .SheetBorder = RGB(128, 128, 128)
        .Editable = flexEDKbdMouse
        
        .Rows = .FixedRows
        .Cols = 4
        .TextMatrix(0, 0) = "Use"
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Strength"
        .TextMatrix(0, 3) = "Weight"
        .FillStyle = flexFillRepeat
        .CellFontBold = True
        .AutoSize 1
        .ColWidth(0) = 500
        .ColWidth(1) = .Width - (500 + .ColWidth(2) + .ColWidth(3)) - 4 * Screen.TwipsPerPixelX
        .ExtendLastCol = True
        .ColAlignment(1) = flexAlignCenterCenter
        
        .FillStyle = flexFillSingle
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.InitSwingGrid"

End Sub

Private Function InitControls() As Boolean
On Error GoTo ErrSection:

    Set m.tbSwing = m.IndPrice.ClusterSwingInfoGet(0)
    If m.tbSwing Is Nothing Then Exit Function
    
    Set m.tbRatioPrice = m.IndPrice.ClusterRatioInfoGet()
    If m.tbRatioPrice Is Nothing Then Exit Function
    
    gdEndDate.MinDate = Int(m.Chart.Bars(eBARS_DateTime, m.IndPrice.ClusterBarsBack))
    gdEndDate.MaxDate = Int(m.Chart.Bars(eBARS_DateTime, m.Chart.Bars.Size - 1))
    
    If m.Chart.Bars.IsIntraday Then
        gdBarTime.Visible = True
        gdBarTime.Enabled = True
        
        fraSwingLevels.Top = kFraLevelsTop
    Else
        gdBarTime.Visible = False
        gdBarTime.Enabled = False
        
        fraSwingLevels.Top = gdBarTime.Top + 135
    End If
    fraSwingInd.Top = fraSwingLevels.Top + fraSwingLevels.Height + 135
    
    
    If Not m.IndTime Is Nothing Then
        Set m.tbRatioTime = m.IndTime.ClusterRatioInfoGet()
        If m.tbRatioTime Is Nothing Then Exit Function
    End If
    
    With cboSwingStyle
        .Clear
        .AddItem "(default)"
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
    End With
    
    InitSwingGrid fgSwing
    InitRatiosGrid fgRatiosPrice
    InitRatiosGrid fgRatiosTime
    
    InitCboLabels cboLabelPrice
    InitCboLabels cboLabelTime
    
    SetTabSwing
    SetTabRatios fgRatiosPrice, m.IndPrice
    SetTabRatios fgRatiosTime, m.IndTime
    
    If m.IndTime Is Nothing Then vsTab.TabVisible(2) = False
    
    vsTab.CurrTab = 0

    InitControls = True
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmClusterCfg.InitControls"

End Function

Private Sub SetGridSwing(fg As VSFlexGrid)
On Error GoTo ErrSection:

    Dim i&, j&, nUse&
    
    If m.tbSwing Is Nothing Then Exit Sub
    If m.tbSwing.NumRecords <= 0 Then Exit Sub

    With fg
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
'tb fields: [0]=name, [1]=strength, [2]=weight
'grid cols: ]0]=use, [1]=name, [2]=strength, [3]=weight

        For i = 0 To m.tbSwing.NumRecords - 1
            .Rows = .Rows + 1
            j = .Rows - 1
            
            .TextMatrix(j, 1) = m.tbSwing(0, i)     'name
            
            nUse = m.tbSwing(1, i)                  'swing strength must be an integer
            If nUse < 0 Then nUse = 0
            
            If j = .FixedRows Then
                'long term swing must always be on
                .Cell(flexcpChecked, j, 0) = flexNoCheckbox
                .TextMatrix(j, 2) = Str(nUse)
            Else
                If nUse = 0 Then
                    .Cell(flexcpChecked, j, 0) = flexUnchecked
                    .TextMatrix(j, 2) = ""      'strength
                Else
                    .Cell(flexcpChecked, j, 0) = flexChecked
                    .TextMatrix(j, 2) = Str(nUse)
                End If
            End If
            
            If m.tbSwing(2, i) = kNullData Or nUse = 0 Then
                .TextMatrix(j, 3) = ""      'weight
            Else
                .TextMatrix(j, 3) = Str(m.tbSwing(2, i))
            End If
        Next
        
        If .Rows >= .FixedRows Then
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignCenterCenter
            
            'show checkbox in first row grid as disabled (flexgrid does not do this)
            chkSwingDummy.Move .Left + .ColWidth(0) / 2 - chkSwingDummy.Width / 2 + 15, .Top + .RowHeight(0) + 45
            chkSwingDummy.Visible = True
            chkSwingDummy.Value = vbChecked
            chkSwingDummy.ZOrder
        Else
            chkSwingDummy.Visible = False
            chkSwingDummy.Visible = False
        End If
        
        .Redraw = flexRDBuffered
    End With
    

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.SetGridSwing"

End Sub

Private Sub SetTabSwing()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation

    SetGridSwing fgSwing

    txtBarsBack.Text = Str(m.IndPrice.ClusterBarsBack)
    chkSwingShow.Value = m.IndPrice.ClusterSwingShow
    DateTimeCtrlsSet m.IndPrice.ClusterEndDate
    
    With cboSwingStyle
        If m.IndPrice.Style >= 0 And m.IndPrice.Style < .ListCount Then
            .ListIndex = m.IndPrice.Style
        Else
            .ListIndex = 0
        End If
    End With
    
    gdColor.Color = m.IndPrice.Color
    gdColorLong.Color = m.IndPrice.ClusterSwingColor("L")
    gdColorMedium.Color = m.IndPrice.ClusterSwingColor("M")
    gdColorShort.Color = m.IndPrice.ClusterSwingColor("S")
    
    If m.IndPrice.trueRangeColor = 1 Then chk1st.Value = vbChecked
    If m.IndTime.trueRangeColor = 1 Then chk2nd.Value = vbChecked
    
    Set Annot = m.Chart.Annots(kClusterZoneRect)
    If Not Annot Is Nothing Then gdZoneColor.Color = Abs(Val(Annot.Prop("FillColor")))
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.SetTabSwing"

End Sub

Private Sub SetTabRatios(fg As VSFlexGrid, Ind As cIndicator)
On Error GoTo ErrSection:

    Dim i&, j&, d#, strWeight$
    
    Dim iTopRow As Long
    Dim bShow As Boolean
    
    Dim Pane As cPane
    Dim tb As cGdTable
    
    If Ind Is Nothing Then Exit Sub
    
    If Ind.DisplayType = eINDIC_ClusterTime Then
        Set tb = m.tbRatioTime
        Set Pane = m.Chart.Tree(kClusterTimeKeyPane)
        
        If Not Pane Is Nothing Then
            bShow = Pane.Display And Ind.Display
        Else
            bShow = Ind.Display
        End If
        
        gdColorTime.Color = Ind.Color
        cboLabelTime.ListIndex = Ind.IndLabelMode
        
        txtProxTime.Text = Str(Ind.ClusterProximity)
        chkShowTime.Value = Abs(bShow)
        
        'JM 02-07-2011 - per Tim not using this right now because not doing time clusters over entire history range
        'txtMaxBars.Text = Str(Ind.ClusterBarsMax)
        
        chkH2H.Value = Abs(Ind.ClusterH2H)
        chkH2L.Value = Abs(Ind.ClusterH2L)
        chkL2H.Value = Abs(Ind.ClusterL2H)
        chkL2L.Value = Abs(Ind.ClusterL2L)
    Else
        Set tb = m.tbRatioPrice
        
        gdColorPrice.Color = Ind.ClusterSwingColor("P")
        cboLabelPrice.ListIndex = Ind.IndLabelMode
        
        txtProxPrice.Text = Str(Ind.ClusterProximity)
        chkShowPrice.Value = Ind.ClusterPriceShow
        
        chkResistance.Value = Abs(Ind.ClusterResistance)
        chkSupport.Value = Abs(Ind.ClusterSupport)
        
        chkABCs.Value = Abs(Ind.ClusterABCs)
    End If

    If tb Is Nothing Then Exit Sub
    If tb.NumRecords <= 0 Then Exit Sub

    With fg
        iTopRow = .TopRow
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
'tb fields: [0]=ratio, [1]=weight
'grid cols: ]0]=use, [1]=ratio, [2]=weight
        
        For i = 0 To tb.NumRecords - 1
            .Rows = .Rows + 1
            j = .Rows - 1
            
            .TextMatrix(j, 1) = Str(tb(0, i))   'ratio
            
            d = tb(1, i)
            strWeight = Str(d)
            
            If Left(strWeight, 1) = "-" Then
                .Cell(flexcpChecked, j, 0) = flexUnchecked
            Else
                .Cell(flexcpChecked, j, 0) = flexChecked
            End If
            .TextMatrix(j, 2) = Str(Abs(d))
        Next
        
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignCenterCenter
        
        If iTopRow >= .FixedRows And iTopRow < .Rows Then .TopRow = iTopRow
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.SetTabRatios"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.Chart = Nothing
    
    Set m.IndPrice = Nothing
    Set m.IndTime = Nothing
    
    Set m.tbSwing = Nothing
    Set m.tbRatioPrice = Nothing
    Set m.tbRatioTime = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.Form_Unload"

End Sub

Private Sub Repaint(Optional ByVal bRepaintNow As Boolean = False, _
    Optional ByVal bSaveDefaults As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, iBar&, dDate#, strMsg$
    Dim bRepaint As Boolean
    Dim Annot As cAnnotation
    Dim Pane As cPane
    
    bRepaint = bRepaintNow
    
    i = Int(ValOfText(txtProxPrice.Text))
    If i < 0 Or i > 99 Then
        InfBox "Cluster proximity must be between 0 and 99", "I", "Ok"
        If i > 0 Then
            txtProxPrice.Text = "99"
        Else
            txtProxPrice.Text = "0"
        End If
        Exit Sub
    End If
    
    i = Int(ValOfText(txtProxTime.Text))
    If i < 0 Or i > 99 Then
        InfBox "Cluster proximity must be between 0 and 99", "I", "Ok"
        If i > 0 Then
            txtProxTime.Text = "99"
        Else
            txtProxTime.Text = "0"
        End If
        Exit Sub
    End If
    
    If m.IndPrice Is Nothing Or m.IndTime Is Nothing Then GoTo ErrExit
    
    If chkShowPrice.Value = vbUnchecked And Me.chkShowTime.Value = vbUnchecked Then
        If chkSwingShow = vbUnchecked Then
            strMsg = "You have selected not to show any indicator. Would you like to delete the fib clusters?"
            If InfBox(strMsg, "?", "+Yes|-No", "Fib Clusters") = "Y" Then
                Me.Hide
                If Not m.Chart Is Nothing Then m.Chart.RemoveFibClusters
                m.bUnload = True
                GoTo ErrExit
            Else
                GoTo ErrExit
            End If
        End If
    End If
    
    iBar = DateTimeBarNum()         '6123
'Price Cluster Properties
    dDate = m.Chart.Bars(eBARS_DateTime, iBar)
    m.IndPrice.ClusterEndDate = dDate
    m.IndPrice.ClusterBarsBack = Abs(Int(ValOfText(txtBarsBack.Text)))
    m.IndPrice.ClusterProximity = ValOfText(txtProxPrice.Text)
    m.IndPrice.ClusterSwingShow = chkSwingShow.Value
    m.IndPrice.Style = cboSwingStyle.ListIndex
    m.IndPrice.IndLabelMode = cboLabelPrice.ListIndex
    
    m.IndPrice.Color = gdColor.Color
    m.IndPrice.ClusterSwingColor("L") = gdColorLong.Color
    m.IndPrice.ClusterSwingColor("M") = gdColorMedium.Color
    m.IndPrice.ClusterSwingColor("S") = gdColorShort.Color
    m.IndPrice.ClusterSwingColor("P") = gdColorPrice.Color
    'JM 11-10-2010: not letting user set retracement & expansion flags for now
    m.IndPrice.ClusterABCs = (-1) * chkABCs.Value
    If chkSupport.Value = vbUnchecked And chkResistance.Value = vbUnchecked Then
        'do nothing - disallow unchecking both of these
    Else
        m.IndPrice.ClusterSupport = (-1) * chkSupport.Value
        m.IndPrice.ClusterResistance = (-1) * chkResistance.Value
    End If
    
    m.IndPrice.trueRangeColor = chk1st.Value    'heat map flag
    m.IndPrice.ClusterPriceShow = chkShowPrice.Value

'Time Cluster Properties
    m.IndTime.IndLabelMode = cboLabelTime.ListIndex
    m.IndTime.ClusterProximity = ValOfText(txtProxTime.Text)
    m.IndTime.Color = gdColorTime.Color
    'JM 02-07-2011 - per Tim not using this right now because not doing time clusters over entire history range
    'Ind.ClusterBarsMax = Abs(ValOfText(txtMaxBars.Text))
    m.IndTime.ClusterH2H = (-1) * chkH2H.Value
    m.IndTime.ClusterL2L = (-1) * chkL2L.Value
    m.IndTime.ClusterH2L = (-1) * chkH2L.Value
    m.IndTime.ClusterL2H = (-1) * chkL2H.Value
    
    m.IndTime.trueRangeColor = chk2nd.Value
    If chkShowTime.Value = vbChecked And Not m.IndTime.Display Then
        'need to make sure Pane is also turned on
        Set Pane = m.Chart.Tree(kClusterTimeKeyPane)
        If Not Pane Is Nothing Then Pane.Display = True
    End If
    m.IndTime.TrueRangeFlag = chkShowTime.Value     'set this to stay in sync with display flag (theoretically will always be 1 in grapheng.dll)
    m.IndTime.Display = (-1) * chkShowTime.Value    'set this last else pane check for time cluster will fail

'Cluster rect properties
    Set Annot = m.Chart.Annots(kClusterZoneRect)
    If Not Annot Is Nothing Then
        If Annot.Color <> m.IndPrice.Color Then Annot.Color = m.IndPrice.Color
        i = gdZoneColor.Color
        If m.IndPrice.trueRangeColor = 1 Or m.IndTime.trueRangeColor = 1 Then i = (-1) * i
        Annot.Prop("FillColor") = i
        If Annot.dDate(2) <> dDate Then Annot.dDate(2) = dDate
        dDate = m.Chart.Bars(eBARS_DateTime, iBar - m.IndPrice.ClusterBarsBack)
        Annot.dDate(1) = dDate
    End If
    
    If bSaveDefaults Then
        If Not Annot Is Nothing Then Annot.SaveDefaults
        m.IndPrice.SaveDefaults
        m.IndTime.SaveDefaults
    End If
    
    If Not m.bUnload Then
        SetTabSwing
        SetTabRatios fgRatiosPrice, m.IndPrice
        SetTabRatios fgRatiosTime, m.IndTime
    End If
    
    If m.IndPrice.ClusterPriceShow Then
        m.Chart.GenerateChart eRedo3_Settings
    Else
        m.Chart.ResetSplitPane
        m.Chart.RestoreChartNormal vbKeyClear
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.Repaint"

End Sub

Private Sub gdBarTime_Changed()
    If Me.Visible And gdBarTime.Visible And vsTab.CurrTab = 0 Then
        Repaint
    End If
End Sub

Private Sub gdColor_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub gdColorLong_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub gdColorMedium_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub gdColorPrice_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub gdColorShort_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub gdColorTime_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub gdEndDate_LostFocus()
    If Me.Visible And vsTab.CurrTab = 0 Then Repaint
End Sub

Private Sub gdZoneColor_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub RatioRestore(fg As VSFlexGrid, Ind As cIndicator)
On Error GoTo ErrSection:

    Dim cbo As ctlUniComboImageXP

    If Not Ind Is Nothing Then
        Ind.ClusterIndDefaults m.Chart, Ind.DisplayType, True
        
        If Ind.DisplayType = eINDIC_ClusterTime Then
            Set m.tbRatioTime = Ind.ClusterRatioInfoGet()
            Set cbo = cboLabelTime
        Else
            Set m.tbRatioPrice = Ind.ClusterRatioInfoGet()
            Set cbo = cboLabelPrice
        End If
        
        If m.tbRatioTime Is Nothing Or m.tbRatioPrice Is Nothing Then Exit Sub
        If m.tbRatioTime.NumRecords <= 0 Then Exit Sub
        If m.tbRatioPrice.NumRecords <= 0 Then Exit Sub

        SetTabRatios fg, Ind
        DoEvents
    
        If Not m.Chart Is Nothing Then m.Chart.GenerateChart
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.RatioRestore"

End Sub

Private Sub RatioRemove(fg As VSFlexGrid, Ind As cIndicator)
On Error GoTo ErrSection:

    With fg
        If m.lMouseDownRow >= .FixedRows And m.lMouseDownRow < .Rows Then
            .RemoveItem m.lMouseDownRow
            UpdateRatioInfo fg, Ind
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.RatioRemove"

End Sub

Private Sub RatioAdd(fg As VSFlexGrid, Ind As cIndicator)
On Error GoTo ErrSection:

    Dim dRatio#
    Dim bOK As Boolean
    
    'new ratio for cluster
    dRatio = ValOfText(InfBox("Enter new ratio:", "", "", "", False, 0, -1, 0, "", "NewRatio"))
    
    bOK = True
    If Ind.DisplayType = eINDIC_ClusterPrice Then
        If Abs(dRatio) = 0# Then
            InfBox "Ratio for price cluster cannot be 0.", , "Ok", "Fib Clusters"
            bOK = False
        End If
    ElseIf dRatio <= 0 Then
        InfBox "Ratio for time cluster must be greater than 0.", , "Ok", "Fib Clusters"
        bOK = False
    End If
    
    If bOK Then
        With fg
            .Rows = .Rows + 1
            'use
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignCenterCenter
            'ratio
            .TextMatrix(.Rows - 1, 1) = dRatio
            'weight
            .TextMatrix(.Rows - 1, 2) = 1
        End With
        UpdateRatioInfo fg, Ind
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.RatioAdd"

End Sub

Private Sub UpdateRatioInfo(fg As VSFlexGrid, Ind As cIndicator)
On Error GoTo ErrSection:

    Dim i&, j&, d1#, d2#
    Dim tb As cGdTable

    If Ind.DisplayType = eINDIC_ClusterTime Then
        Set tb = m.tbRatioTime
    Else
        Set tb = m.tbRatioPrice
    End If
    
    If tb Is Nothing Then Exit Sub
    tb.NumRecords = 0
    
    With fg
        For i = .FixedRows To .Rows - 1
            d1 = ValOfText(.TextMatrix(i, eRatioCol_Ratio))
            d2 = ValOfText(.TextMatrix(i, eRatioCol_Weight))
            
            If .Cell(flexcpChecked, i, eRatioCol_Use) = flexUnchecked Then
                d2 = (-1) * d2
            End If
            
            tb.AddRecord ""
            j = tb.NumRecords - 1
            'tb(0, j) = Abs(d1)
            tb(0, j) = d1
            tb(1, j) = d2
        Next
    End With
    
    Ind.ClusterRatioInfoSet tb
    
    If Ind.DisplayType = eINDIC_ClusterTime Then
        Set m.tbRatioTime = Ind.ClusterRatioInfoGet
    Else
        Set m.tbRatioPrice = Ind.ClusterRatioInfoGet
    End If
    
    SetTabRatios fg, Ind
    
    DoEvents
    
    If Not m.Chart Is Nothing Then m.Chart.GenerateChart

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.UpdateRatioInfo"

End Sub

Private Sub UpdateSwingInfo(fg As VSFlexGrid, ByVal nRow As Long, ByVal nCol As Long)
On Error GoTo ErrSection:

    Dim nUse&, nStrength&, nTbRow&, dWeight#
    
    Dim Ind As cIndicator
    
    If nRow < fg.FixedRows Or nRow >= fg.Rows Then Exit Sub
    
    nTbRow = nRow - 1
    If nTbRow < 0 Or nTbRow >= m.tbSwing.NumRecords Then Exit Sub           'precautionary
    
    With fg
        nUse = .Cell(flexcpChecked, nRow, 0)
        nStrength = Abs(ValOfText(.TextMatrix(nRow, 2)))
        dWeight = Abs(ValOfText(.TextMatrix(nRow, 3)))
    End With
    
    If nRow = eSwingRow_Long Then nUse = flexChecked   'long term is always on
    
    Set Ind = m.IndPrice
    
    If nStrength = 0 Or dWeight = 0# Then
        If nUse = flexChecked Then
            m.tbSwing(1, nTbRow) = Abs(m.tbSwing(1, nTbRow))
            m.tbSwing(2, nTbRow) = Abs(m.tbSwing(2, nTbRow))
            
            If nRow = eSwingRow_Short Then
                'make sure intermediate also gets turned on
                If fg.Cell(flexcpChecked, eSwingRow_Intermediate, nCol) = flexUnchecked Then
                    nTbRow = nTbRow - 1
                    m.tbSwing(1, nTbRow) = Abs(m.tbSwing(1, nTbRow))
                    m.tbSwing(2, nTbRow) = Abs(m.tbSwing(2, nTbRow))
                End If
            End If
        
            Ind.ClusterSwignInfoSet m.tbSwing
        End If
    ElseIf nUse = flexUnchecked Then
        'save as negative to indicate last-used value
        m.tbSwing(1, nTbRow) = (-1) * nStrength
        If dWeight > 0 Then m.tbSwing(2, nTbRow) = dWeight
        
        If nRow = eSwingRow_Intermediate Then
            'make sure short also gets turned off
            If fg.Cell(flexcpChecked, eSwingRow_Short, nCol) = flexChecked Then
                nTbRow = nTbRow + 1
                m.tbSwing(1, nTbRow) = (-1) * m.tbSwing(1, nTbRow)
                m.tbSwing(2, nTbRow) = Abs(m.tbSwing(2, nTbRow))
            End If
        End If
        
        Ind.ClusterSwignInfoSet m.tbSwing
    ElseIf nUse = flexChecked Then
        m.tbSwing(1, nTbRow) = nStrength
        If dWeight > 0 Then m.tbSwing(2, nTbRow) = dWeight
        If nRow = eSwingRow_Short Then
            nTbRow = nTbRow - 1
            If nTbRow >= 0 And nTbRow < m.tbSwing.NumRecords Then
                'make sure Intermediate is also turned on
                m.tbSwing(1, nTbRow) = Abs(m.tbSwing(1, nTbRow))
            End If
        End If
        Ind.ClusterSwignInfoSet m.tbSwing
    End If
     
    SetGridSwing fg
    DoEvents
    
    If Not m.Chart Is Nothing Then m.Chart.GenerateChart
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.UpdateSwingInfo"

End Sub

Private Function ValidateSwingEdit(fg As VSFlexGrid, ByVal Row&, ByVal Col&) As Boolean
On Error GoTo ErrSection:

    Dim strMsg$, strDescending$, strZero$
    Dim bUpdate As Boolean
    
    Dim nChkMedium&, nChkShort&
    Dim d1#, d2#, d3#
    
    
'Quick Reference:
'eSwingCol_Use = 0, eSwingCol_Name, eSwingCol_Strength, eSwingCol_Weight
'eSwingRow_Long = 1, eSwingRow_Intermediate, eSwingRow_Short

    If Row >= fg.FixedRows And Row < fg.Rows Then
        Select Case Col
            Case eSwingCol_Use
                bUpdate = True
            
            Case eSwingCol_Strength
                strMsg = ""
                
                strZero = "Strength values must be greater than zero."
                
                strDescending = "Strength values should be in descending order."
                strDescending = strDescending & vbCrLf & "( long > intermediate > short )."
                
                d1 = ValOfText(fg.TextMatrix(eSwingRow_Long, Col))
                If d1 > 0# Then
                    If fg.Cell(flexcpChecked, eSwingRow_Intermediate, eSwingCol_Use) = vbChecked Then
                        d2 = ValOfText(fg.TextMatrix(eSwingRow_Intermediate, Col))
                        If d2 <= 0# Then
                            strMsg = strZero
                        ElseIf d1 > d2 Then
                            If fg.Cell(flexcpChecked, eSwingRow_Short, eSwingCol_Use) = vbChecked Then
                                d3 = ValOfText(fg.TextMatrix(eSwingRow_Short, Col))
                                If d3 <= 0# Then
                                    strMsg = strZero
                                ElseIf d3 > d2 Then
                                    strMsg = strDescending
                                End If
                            End If
                        Else
                            strMsg = strDescending
                        End If
                    End If
                Else
                    strMsg = strZero
                End If
                
                If Len(strMsg) = 0 Then
                    bUpdate = True
                Else
                    InfBox strMsg, "I", "Ok", kInfSwingTitle
                End If
            
            Case eSwingCol_Weight
                d1 = ValOfText(fg.TextMatrix(Row, Col))
                If d1 <= 0# Then
                    strMsg = "Weight values must be greater than zero."
                    InfBox strMsg, "I", "Ok", kInfSwingTitle
                Else
                    bUpdate = True
                End If
        
        End Select
    End If
    
    ValidateSwingEdit = bUpdate

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmClusterCfg.ValidateSwingEdit"

End Function

Private Sub HandleGridMouseDown(fg As VSFlexGrid)
On Error Resume Next

    If Not fg Is Nothing Then
        With fg
            If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
                If .MouseCol >= .FixedCols And .MouseCol < .Rows Then
                    .EditCell     '6056
                    m.lMouseDownRow = .MouseRow
                    m.lMouseDownCol = .MouseCol
                End If
            End If
        End With
    End If

End Sub

Private Sub DateTimeCtrlsSet(ByVal dDateTime#)
On Error GoTo ErrSection

    Dim d#

    If dDateTime > 0 Then
        d = dDateTime
        gdEndDate.Value = d
        
        If m.Chart.Bars.IsIntraday Then
            If g.bShowInLocalTimeZone Then
                d = ConvertTimeZone(dDateTime, m.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            End If
            gdBarTime.Value = d
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmClusterCfg.DateTimeCtrlsSet"

End Sub

Private Function DateTimeBarNum() As Long
On Error GoTo ErrSection

    Dim d#, j&, iBar&, iLastGoodDatabar&
    
    If m.Chart.Bars.IsIntraday Then
        d = CDbl(gdEndDate.Value) + CDbl(gdBarTime.Value)
        
        If g.bShowInLocalTimeZone Then
            d = ConvertTimeZone(d, "", m.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf))        '6281
        End If
    Else
        d = CDbl(gdEndDate.Value)
    End If
    
    If d <= 0 Then d = m.IndPrice.ClusterEndDate     'time value invalid, theoretically should never happen
    
    iLastGoodDatabar = m.Chart.LastGoodDataBar(False)
    If d > m.Chart.Bars(eBARS_DateTime, iLastGoodDatabar) Then
        d = m.Chart.Bars(eBARS_DateTime, iLastGoodDatabar)
        iBar = iLastGoodDatabar
    Else
        iBar = m.Chart.Bars.FindDateTime(d)
        If iBar >= 0 And iBar < m.Chart.Bars.Size Then
            d = m.Chart.Bars(eBARS_DateTime, iBar)
            
            If g.bShowInLocalTimeZone Then
                d = ConvertTimeZone(d, m.Chart.Bars.Prop(eBARS_ExchangeTimeZoneInf), "")    '6281
            End If
            
            gdEndDate.Value = Int(d)
            gdBarTime.Value = Int(d) - d
        Else
            iBar = m.Chart.Bars.FindDateTime(m.IndPrice.ClusterEndDate)
        End If
    End If
    
    If iBar >= 0 And iBar < m.Chart.Bars.Size Then
        j = Abs(Int(ValOfText(txtBarsBack.Text)))
        If iBar - j < 0 Then
            j = iBar
        End If
        txtBarsBack.Text = Str(j)
    Else
        iBar = iLastGoodDatabar
    End If

    DateTimeBarNum = iBar

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmClusterCfg.DateTimeCtrlsGet"

End Function

Private Sub txtBarsBack_LostFocus()
    If Me.Visible Then Repaint
End Sub

Private Sub txtProxPrice_Change()
    If Me.Visible And Len(txtProxPrice.Text) > 0 Then Repaint
End Sub

Private Sub txtProxTime_Change()
    If Me.Visible And Len(txtProxTime.Text) > 0 Then Repaint
End Sub

