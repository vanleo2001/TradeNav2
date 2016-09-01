VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAndrewFork 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdFont 
      Height          =   330
      Left            =   3454
      TabIndex        =   4
      Top             =   645
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
      Caption         =   "frmAndrewFork.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmAndrewFork.frx":002A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":004A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkMultiChart 
      Height          =   195
      Left            =   143
      TabIndex        =   2
      Top             =   1455
      Width           =   4830
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
      Caption         =   "frmAndrewFork.frx":0066
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmAndrewFork.frx":00D6
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":00F6
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   3180
      Left            =   165
      TabIndex        =   6
      Top             =   1830
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   5609
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
      Caption         =   "Fork|Parallels|Time Intervals"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
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
      Begin HexUniControls.ctlUniFrameWL fraTicks 
         Height          =   2805
         Left            =   5835
         TabIndex        =   20
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
         Caption         =   "frmAndrewFork.frx":0112
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAndrewFork.frx":0142
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAndrewFork.frx":0162
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkAllPanes 
            Height          =   240
            Left            =   405
            TabIndex        =   5
            Top             =   1245
            Width           =   2205
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
            Caption         =   "frmAndrewFork.frx":017E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":01C0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":01E0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkMarkerText 
            Height          =   240
            Left            =   405
            TabIndex        =   15
            Top             =   990
            Width           =   2205
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
            Caption         =   "frmAndrewFork.frx":01FC
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":024A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":026A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboMarkerStyle 
            Height          =   315
            Left            =   960
            TabIndex        =   27
            Top             =   225
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
            Tip             =   "frmAndrewFork.frx":0286
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":02A6
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdMarkerAdd 
            Height          =   315
            Left            =   675
            TabIndex        =   23
            Top             =   1575
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
            Caption         =   "frmAndrewFork.frx":02C2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":02F4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0314
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdMarkerRemove 
            Height          =   315
            Left            =   675
            TabIndex        =   22
            Top             =   1935
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
            Caption         =   "frmAndrewFork.frx":0330
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":0368
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0388
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdMarkerRestore 
            Height          =   315
            Left            =   675
            TabIndex        =   21
            Top             =   2310
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
            Caption         =   "frmAndrewFork.frx":03A4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":03E4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0404
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgMarkers 
            Height          =   2370
            Left            =   2820
            TabIndex        =   24
            Top             =   210
            Width           =   1740
            _cx             =   3069
            _cy             =   4180
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
         Begin gdOCX.gdSelectColor clrMarker 
            Height          =   315
            Left            =   945
            TabIndex        =   25
            Top             =   585
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   165
            Top             =   255
            Width           =   840
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
            Caption         =   "frmAndrewFork.frx":0420
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAndrewFork.frx":044A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":046A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   255
            Left            =   150
            Top             =   615
            Width           =   840
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
            Caption         =   "frmAndrewFork.frx":0486
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAndrewFork.frx":04B2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":04D2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraParallel 
         Height          =   2805
         Left            =   5535
         TabIndex        =   8
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
         Caption         =   "frmAndrewFork.frx":04EE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAndrewFork.frx":0524
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAndrewFork.frx":0544
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboLocation 
            Height          =   315
            Left            =   870
            TabIndex        =   26
            Top             =   585
            Width           =   1110
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
            Tip             =   "frmAndrewFork.frx":0560
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0580
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkParallelText 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   990
            Width           =   1635
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
            Caption         =   "frmAndrewFork.frx":059C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":05E0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0600
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkExtendBase 
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1290
            Width           =   1890
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
            Caption         =   "frmAndrewFork.frx":061C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":0660
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0680
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor clrParallel 
            Height          =   315
            Left            =   3570
            TabIndex        =   13
            Top             =   2160
            Visible         =   0   'False
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniComboImageXP cboParallelStyle 
            Height          =   315
            Left            =   540
            TabIndex        =   12
            Top             =   180
            Width           =   1440
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
            Tip             =   "frmAndrewFork.frx":069C
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":06BC
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdRestore 
            Height          =   315
            Left            =   435
            TabIndex        =   11
            Top             =   2370
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
            Caption         =   "frmAndrewFork.frx":06D8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":0718
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0738
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
            Height          =   315
            Left            =   435
            TabIndex        =   10
            Top             =   2010
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
            Caption         =   "frmAndrewFork.frx":0754
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":078C
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":07AC
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
            Height          =   315
            Left            =   435
            TabIndex        =   9
            Top             =   1650
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
            Caption         =   "frmAndrewFork.frx":07C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":07FA
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":081A
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgParallel 
            Height          =   2415
            Left            =   2190
            TabIndex        =   14
            Top             =   180
            Width           =   2460
            _cx             =   4339
            _cy             =   4260
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
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   120
            Top             =   615
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
            Caption         =   "frmAndrewFork.frx":0836
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAndrewFork.frx":0868
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0888
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblFibStyle 
            Height          =   255
            Left            =   120
            Top             =   210
            Width           =   375
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
            Caption         =   "frmAndrewFork.frx":08A4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAndrewFork.frx":08D0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":08F0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraFork 
         Height          =   2805
         Left            =   45
         TabIndex        =   7
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
         Caption         =   "frmAndrewFork.frx":090C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAndrewFork.frx":093A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAndrewFork.frx":095A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkHideABLine 
            Height          =   210
            Left            =   600
            TabIndex        =   32
            Top             =   990
            Width           =   3870
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
            Caption         =   "frmAndrewFork.frx":0976
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":09F2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0A12
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSchiffMod 
            Height          =   255
            Left            =   2730
            TabIndex        =   34
            Top             =   600
            Width           =   1605
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
            Caption         =   "frmAndrewFork.frx":0A2E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":0A6C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0A8C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSchiff 
            Height          =   255
            Left            =   1665
            TabIndex        =   31
            Top             =   600
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
            Caption         =   "frmAndrewFork.frx":0AA8
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":0AD4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0AF4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkNormal 
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   600
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
            Caption         =   "frmAndrewFork.frx":0B10
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAndrewFork.frx":0B3C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0B5C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgTimeLines 
            Height          =   1290
            Left            =   600
            TabIndex        =   33
            Top             =   1290
            Width           =   3420
            _cx             =   6032
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
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   600
            Top             =   300
            Width           =   1050
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
            Caption         =   "frmAndrewFork.frx":0B78
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAndrewFork.frx":0BB0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAndrewFork.frx":0BD0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkPreIndicator 
      Height          =   195
      Left            =   143
      TabIndex        =   3
      Top             =   1215
      Width           =   2055
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
      Caption         =   "frmAndrewFork.frx":0BEC
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmAndrewFork.frx":0C38
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":0C58
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboStyle 
      Height          =   315
      Left            =   1647
      TabIndex        =   0
      Top             =   645
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
      Tip             =   "frmAndrewFork.frx":0C74
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":0C94
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin gdOCX.gdSelectColor clrColor 
      Height          =   315
      Left            =   1647
      TabIndex        =   1
      Top             =   225
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   615
      Left            =   908
      TabIndex        =   16
      Top             =   4980
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
      Caption         =   "frmAndrewFork.frx":0CB0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAndrewFork.frx":0CD0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":0CF0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveDefaults 
         Height          =   330
         Left            =   965
         TabIndex        =   19
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
         Caption         =   "frmAndrewFork.frx":0D0C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAndrewFork.frx":0D4E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAndrewFork.frx":0D6E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   60
         TabIndex        =   18
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
         Caption         =   "frmAndrewFork.frx":0D8A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAndrewFork.frx":0DB0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAndrewFork.frx":0DD0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   330
         Left            =   2575
         TabIndex        =   17
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
         Caption         =   "frmAndrewFork.frx":0DEC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAndrewFork.frx":0E1A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAndrewFork.frx":0E3A
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboSchiffStyle 
      Height          =   315
      Left            =   3420
      TabIndex        =   29
      Top             =   5055
      Visible         =   0   'False
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
      Tip             =   "frmAndrewFork.frx":0E56
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":0E76
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   255
      Left            =   987
      Top             =   255
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
      Caption         =   "frmAndrewFork.frx":0E92
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAndrewFork.frx":0EBE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":0EDE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblStyle 
      Height          =   255
      Left            =   987
      Top             =   675
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
      Caption         =   "frmAndrewFork.frx":0EFA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAndrewFork.frx":0F26
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAndrewFork.frx":0F46
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmAndrewFork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Chart As cChart
    Annot As cAnnotation
    nAnnotIdx As Long
    
    bWasMultiChart As Boolean
    bMultiChartOption As Boolean
    bChanged As Boolean
    bDescending_P As Boolean     'sort flag for parallel grid
    bDescending_M As Boolean     'sort flag for markers grid
    
    nFibRowDown As Long         'to keep track of color changes in grid
    nFibColDown As Long
End Type
Private m As mPrivate


Public Sub ShowMe(Annot As cAnnotation)
On Error GoTo ErrSection:
    
    If Annot Is Nothing Then Exit Sub
    
    Set m.Annot = Annot
    Set m.Chart = Annot.AnnotChart
    If m.Chart Is Nothing Then Exit Sub
    
    m.bChanged = False
    m.bDescending_P = False
    m.bDescending_M = False
    
    InitControls True
    
    CenterFormOnChart Me, m.Chart        '6434
    ShowForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.ShowMe"
    
End Sub

Private Sub cboLocation_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub cboMarkerStyle_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub cboParallelStyle_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub cboStyle_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkAllPanes_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkExtendBase_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkHideABLine_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkMarkerText_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkMultiChart_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkNormal_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkParallelText_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkPreIndicator_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkSchiff_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkSchiffMod_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub clrColor_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub clrMarker_Changed()
    If Me.Visible Then Repaint
End Sub

Private Sub clrParallel_Changed()
On Error GoTo ErrSection:
    
    Dim nColor As Long
    
    clrParallel.Visible = False
    nColor = clrParallel.Color
    If nColor = 0 Then nColor = -1  '0 is reserved color in flex grid control
    
    With fgParallel
        If m.nFibColDown > 1 And m.nFibColDown < .Cols Then
            If .TextMatrix(0, m.nFibColDown) = "Color" Then
                If m.nFibRowDown >= .FixedRows And m.nFibRowDown < .Rows Then
                    .Cell(flexcpBackColor, m.nFibRowDown, m.nFibColDown) = nColor
                    .Select 0, 0
                    m.bChanged = True
                    Repaint
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.clrParallel_Changed"

End Sub

Private Sub clrParallel_ColorClicked()
On Error Resume Next

    clrParallel.Visible = False     '6062

End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    Dim dRatio#, strAnswer$
    Dim iDontCare As Integer
    
    strAnswer = InfBox("Enter new ratio:", "", "", "", False, 0, -1, 0, "", "NewRatio")
    If Len(strAnswer) = 0 Then GoTo ErrExit         '6322
    
    'new ratio for parallel lines
    dRatio = ValOfText(strAnswer)
    
    If dRatio > 0# And dRatio <> 1# Then
        With fgParallel
            .Rows = .Rows + 1
            'use
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignCenterCenter
            'ratio
            .TextMatrix(.Rows - 1, 1) = dRatio
            'color
            .Cell(flexcpBackColor, .Rows - 1, 2) = vbRed
        End With
        
        m.bChanged = True
        Repaint
        HandleGridBeforeSort fgParallel, 1, iDontCare, Not m.bDescending_P      '6036
    Else
        InfBox "Ratio must be positive and cannot be 0 or 1.", , "Ok", "Andrews Fork"
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmAdd_Click"

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim i&
                    
    If Not m.Chart Is Nothing And Not m.Annot Is Nothing Then
        i = m.Annot.geAnnId
        m.Annot.geRemoveAnnotation (m.Chart.geChartObj)
        m.Chart.Annots.Remove i
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
    End If
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdDelete_Click"

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    Dim nStyle&
    
    'set font currently in use
    Me.Font.Name = m.Annot.Prop("FontName")
    Me.Font.Size = Val(m.Annot.Prop("FontSize"))
    Me.Font.Underline = Val(m.Annot.Prop("FontUnderline"))
    nStyle = Val(m.Annot.Prop("FontStyle"))
    Select Case nStyle
        Case 0:
            Me.Font.Italic = False
            Me.Font.Bold = False
        Case 1:
            Me.Font.Italic = False
            Me.Font.Bold = True
        Case 2:
            Me.Font.Italic = True
            Me.Font.Bold = False
        Case 3:
            Me.Font.Italic = True
            Me.Font.Bold = True
    End Select
    
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.Annot.Prop("FontName") = Me.Font.Name
        m.Annot.Prop("FontSize") = Me.Font.Size
        m.Annot.Prop("FontUnderline") = Me.Font.Underline
        
        'style - 0=reg,1=bold,2=italic,3=bold italic
        nStyle = 0
        If Me.Font.Bold = True Then
            If Me.Font.Italic = True Then
                nStyle = 3
            Else
                nStyle = 1
            End If
        ElseIf Me.Font.Italic = True Then
            nStyle = 2
        End If
        
        m.Annot.Prop("FontStyle") = nStyle
        
        m.bChanged = True
        Repaint
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAndrewFork.cmdFont.Click", eGDRaiseError_Show
    
End Sub

Private Sub cmdOK_Click()
On Error Resume Next

    If Not m.Annot Is Nothing And Not m.Chart Is Nothing Then
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
    End If
    
    Unload Me

End Sub

Private Sub cmdMarkerAdd_Click()
On Error GoTo ErrSection:
    
    Dim dRatio#, strAnswer$
    Dim iDontCare As Integer
    
    'new ratio for time intervals
    strAnswer = InfBox("Enter new ratio:", "", "", "", False, 0, -1, 0, "", "NewRatio")
    If Len(strAnswer) = 0 Then GoTo ErrExit         '6322
    
    dRatio = ValOfText(strAnswer)
    
    If dRatio > 0# Then
        With fgMarkers
            .Rows = .Rows + 1
            'use
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = flexPicAlignCenterCenter
            'ratio
            .TextMatrix(.Rows - 1, 1) = dRatio
        End With
        
        m.bChanged = True
        Repaint
        HandleGridBeforeSort fgMarkers, 1, iDontCare, Not m.bDescending_M
    Else
        InfBox "Ratio must be positive and cannot be 0.", , "Ok", "Andrews Fork"
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdMarkerAdd_Click"
    
End Sub

Private Sub cmdMarkerRemove_Click()
On Error GoTo ErrSection:

    With fgMarkers
        If .Row >= .FixedRows And .Row < .Rows Then
            .RemoveItem .Row
            m.bChanged = True
            Repaint
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdMarkerRemove_Click"
    
End Sub

Private Sub cmdMarkerRestore_Click()
On Error GoTo ErrSection:

    If Not m.Annot Is Nothing Then
        m.Annot.AndrewsDefaults "Markers"
        m.Annot.RatiosToFibGrid fgMarkers, m.bDescending_M
        If Not m.Chart Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdMarkerRestore_Click"
    
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    With fgParallel
        If .Row >= .FixedRows And .Row < .Rows Then
            .RemoveItem .Row
            m.bChanged = True
            Repaint
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdRemove_Click"
    
End Sub

Private Sub cmdRestore_Click()
On Error GoTo ErrSection:

    If Not m.Annot Is Nothing Then
        m.Annot.AndrewsDefaults "Parallel"
        m.Annot.RatiosToFibGrid fgParallel, m.bDescending_P
        If Not m.Chart Is Nothing Then m.Chart.GenerateChart eRedo1_Scrolled
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdRestore_Click"
    
End Sub

Private Sub cmdSaveDefaults_Click()
On Error GoTo ErrSection:

    If Not m.Annot Is Nothing Then
        m.Annot.SaveDefaults
        If Not m.Chart Is Nothing Then
            m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart          '6334
        End If
    End If
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.cmdSaveDefaults_Click"
End Sub

Private Sub fgMarkers_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    HandleGridBeforeMouseDown fgMarkers
End Sub

Private Sub fgMarkers_BeforeSort(ByVal Col As Long, Order As Integer)
    HandleGridBeforeSort fgMarkers, Col, Order, m.bDescending_M
End Sub

Private Sub fgMarkers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Repaint
End Sub

Private Sub fgParallel_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    HandleGridBeforeMouseDown fgParallel
End Sub

Private Sub fgParallel_BeforeSort(ByVal Col As Long, Order As Integer)
    HandleGridBeforeSort fgParallel, Col, Order, m.bDescending_P
End Sub

Private Sub fgParallel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim i&, j&
    Dim bReset As Boolean
    
    bReset = True
    With fgParallel
        If .Row >= .FixedRows And .Row < .Rows And .Col >= .FixedCols And .Col < .Cols Then
            m.bChanged = True
            .Select .Row, .Col
            m.nFibColDown = .Col
            m.nFibRowDown = .Row
            If .TextMatrix(0, .Col) = "Color" Then
                If .TopRow > .FixedRows Then
                    i = .Row - .TopRow + .FixedRows
                Else
                    i = .Row
                End If
                clrParallel.Move .Left + .ClientWidth - clrParallel.Width - 10, .Top + .RowHeight(.Row) * i
                clrParallel.Color = .Cell(flexcpBackColor, .Row, .Col)
                clrParallel.Visible = True
                
                bReset = False
            End If
        End If
    End With
    
    If bReset Then
        m.nFibColDown = -1
        m.nFibRowDown = -1
        clrParallel.Visible = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.fgParallel_MouseDown"

End Sub

Private Sub fgParallel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If fgParallel.MouseCol <> 2 Then clrParallel.Visible = False

End Sub

Private Sub fgParallel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If clrParallel.Visible And m.nFibColDown > 0 And m.nFibRowDown > 0 And Not clrParallel.DropDownVisible Then
        clrParallel.UserControl_Click
    Else
        m.bChanged = True
        Repaint
    End If
    
End Sub

Private Sub fgTimeLines_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Col = 1 Then
        m.bChanged = True
        Repaint
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.fgTimeLines_AfterEdit"

End Sub

Private Sub fgTimeLines_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    
    If Col = 0 Then Cancel = True

End Sub


Private Sub Form_Load()

    g.Styler.StyleForm Me
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    With fraFork
        .Left = 0
        .Width = Me.Width
    End With
    
End Sub

Private Sub InitControls(ByVal bSetTab As Boolean)
On Error GoTo ErrSection:

    Dim i&, strValue$
    Dim aDates As New cGdArray
    Dim aYs As New cGdArray
    Dim bIntraday As Boolean
    
    Dim eMode As eForkMode
    
    Me.Icon = Picture16(ToolbarIcon("ID_AndrewFork"), , True)
    Me.Caption = "Andrews Pitchfork"
    
    LoadAnnotPenstyle cboStyle
    LoadAnnotPenstyle cboParallelStyle
'    LoadAnnotPenstyle cboSchiffStyle
    LoadAnnotPenstyle cboMarkerStyle
    
    cboLocation.Clear
    cboLocation.AddItem "Both"
    cboLocation.AddItem "Above"
    cboLocation.AddItem "Below"
    
    If Not m.Chart.Bars Is Nothing Then bIntraday = m.Chart.Bars.IsIntraday
    
    If m.Chart.SymbolID > 0 Then
        chkMultiChart.Caption = "Show for " & m.Chart.Symbol & " in all chart windows"
    ElseIf Len(m.Chart.SpreadSymbols) > 0 Then
        chkMultiChart.Caption = "Show in all chart windows for this spread"
    Else
        chkMultiChart.Caption = "Show in all chart windows for this symbol"
    End If
    
    With m.Annot
        m.bWasMultiChart = .MultiChartFlag
        clrColor.Color = .Color
        chkPreIndicator.Value = .PreIndicator
        SetAnnotPenstyleCombo cboStyle, .Style
        
'        i = Val(.Prop("SchiffStyle"))
'        SetAnnotPenstyleCombo cboSchiffStyle, i
        
        'parallel lines properties
        i = Val(.Prop("ParallelStyle"))
        SetAnnotPenstyleCombo cboParallelStyle, i
        If Val(.Prop("ShowHandle")) = 1 Then
            chkExtendBase.Value = vbChecked
        Else
            chkExtendBase.Value = vbUnchecked
        End If
        
        If Val(.Prop("ParallelText")) = 1 Then
            chkParallelText.Value = vbChecked
        Else
            chkParallelText.Value = vbUnchecked
        End If
        
        'time intervals properties
        i = Val(.Prop("MarkerStyle"))
        SetAnnotPenstyleCombo cboMarkerStyle, i
        
        i = Val(.Prop("MarkerColor"))
        clrMarker.Color = i
        
        If Val(.Prop("MarkerText")) = 1 Then
            chkMarkerText.Value = vbChecked
        Else
            chkMarkerText.Value = vbUnchecked
        End If
        
        'show multichart option only if annotation is in price pane AND does not have alert
        If m.Chart.Tree.Key(m.Annot.gePaneId) = "PRICE PANE" And .AlertObject Is Nothing Then
            chkMultiChart.Value = Abs(.MultiChartFlag)
            m.bMultiChartOption = True
        Else
            m.bMultiChartOption = False
        End If
        chkMultiChart.Visible = m.bMultiChartOption
        
        'fork mode
        eMode = Val(.Prop("ForkMode"))
        'AB line
        chkHideABLine.Value = Val(.Prop("HideABLine"))
        'extend time intervals markers through all panes
        chkAllPanes.Value = Val(.Prop("ShowInAllPanes"))
        'paralells location
        i = Val(.Prop("ParallelLocation"))
        If i >= 0 And i <= 3 Then
            cboLocation.ListIndex = i
        Else
            cboLocation.ListIndex = 0
        End If
        
        Select Case eMode
            Case eANNOT_ForkNormal
                chkNormal.Value = vbChecked
            
            Case eANNOT_Schiff
                chkSchiff.Value = vbChecked
            
            Case eANNOT_SchiffMod
                chkSchiffMod.Value = vbChecked

            Case eANNOT_SchiffModSchiff
                chkSchiff.Value = vbChecked
                chkSchiffMod.Value = vbChecked
            
            Case eANNOT_NormalSchiff
                chkNormal.Value = vbChecked
                chkSchiff.Value = vbChecked
            
            Case eANNOT_NormalSchiffMod
                chkNormal.Value = vbChecked
                chkSchiffMod.Value = vbChecked
            
            Case eANNOT_ForkAll
                chkNormal.Value = vbChecked
                chkSchiff.Value = vbChecked
                chkSchiffMod.Value = vbChecked
        End Select
        
        aDates(0) = .dDate(1)
        aDates(1) = .dDate(2)
        aDates(2) = .DateFromArray(0)
        aYs(0) = .Y(1)
        aYs(1) = .Y(2)
        aYs(2) = .YFromArray(0)
    End With

    With fgTimeLines
        SetupGrid Me.fgTimeLines, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .ColAlignment(0) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarNone
        .Editable = flexEDKbdMouse
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 2
    
        If aDates.Size > 0 Then
            .Rows = aDates.Size + .FixedRows
            .Height = .Height - .RowHeight(1)
            .ColWidth(0) = 2200
            .ColAlignment(1) = flexAlignRightCenter
            .TextMatrix(0, 0) = "Date"
            .TextMatrix(0, 1) = "Value"
            For i = 0 To aDates.Size - 1
                strValue = ShowValue(aYs(i))
                If bIntraday Then
                    .TextMatrix(i + 1, 0) = DateFormat(aDates(i), MM_DD_YYYY, HH_MM_SS)
                Else
                    .TextMatrix(i + 1, 0) = DateFormat(aDates(i))
                End If
                .TextMatrix(i + 1, 1) = strValue
                .RowData(i + 1) = aDates(i)
            Next
            .Sort = flexSortGenericAscending
            .Select .FixedRows, .FixedCols, .Rows - 1, .Cols - 1    'forces a sort
            .Select 1, 1        'reset selection
        End If
        
        .Height = .RowHeight(0) * .Rows
    End With
    
    With fgParallel
        SetupGrid Me.fgParallel, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarVertical
        .ExtendLastCol = True
        .FixedCols = 0
        .FixedRows = 1
        
        .Rows = .FixedRows
        
        .Cols = 3
        .ColWidth(0) = 600
        .ColWidth(1) = 900
        
        .TextMatrix(0, 0) = "Use"
        .TextMatrix(0, 1) = "Ratio"
        .TextMatrix(0, 2) = "Color"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColDataType(0) = flexDTBoolean
        
        m.Annot.RatiosToFibGrid fgParallel
        .Redraw = flexRDBuffered
        
        clrParallel.Width = .ColWidth(2) - 50
    End With
    
    With fgMarkers
        .Redraw = flexRDNone
        SetupGrid Me.fgMarkers, eGridMode_Grid
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarVertical
        .ExtendLastCol = True
        .FixedCols = 0
        .FixedRows = 1
        
        .Rows = .FixedRows
        
        .Cols = 2
        .ColWidth(0) = 600
        
        .TextMatrix(0, 0) = "Use"
        .TextMatrix(0, 1) = "Ratio"
        .ColDataType(0) = flexDTBoolean
        
        m.Annot.RatiosToFibGrid fgMarkers
        
        .Redraw = flexRDBuffered
    End With
    
    If bSetTab Then vsTab.CurrTab = 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAndrewFork.InitControls"

End Sub

Private Function ShowValue(ByVal vValue As Variant) As String
On Error GoTo ErrSection:

    Dim iPane&, strValue$
    
    strValue = CStr(vValue)
    
    With m.Annot
        If Not .AnnotChart Is Nothing Then
            If .Pane = "PRICE PANE" Then
                iPane = .AnnotChart.Tree.Index(.Pane)
                If iPane > 0 Then
                    strValue = .AnnotChart.PriceDisplay(iPane, vValue)
                End If
            End If
        End If
    End With
    
    If Len(strValue) = 0 Then strValue = Format(vValue, "0.00#")

    ShowValue = strValue

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAndrewFork.ShowValue", eGDRaiseError_Raise
    
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set m.Chart = Nothing
    Set m.Annot = Nothing

End Sub

Private Sub HandleGridBeforeSort(fg As VSFlexGrid, ByVal Col As Long, ByRef Order As Integer, _
    ByRef bDescending As Boolean)
On Error GoTo ErrSection:

    'let grid do sorting for column 0 (Use column)
    'let annotation object do sorting for column 1 (Ratios column)
    'disallow sorting for all other columns
    If Col <> 0 Then
        Order = flexSortNone
        If Col = 1 And Not m.Annot Is Nothing Then
            bDescending = Not bDescending
            m.Annot.RatiosToFibGrid fg, bDescending
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAndrewFork.HandleGridBeforeSort", eGDRaiseError_Raise
    
End Sub

Private Sub HandleGridBeforeMouseDown(fg As VSFlexGrid)
On Error GoTo ErrSection:

    Dim nRow&, nCol&
    
    With fg
        nRow = .MouseRow
        nCol = .MouseCol
        
        If nCol = 0 Then
            If nRow >= .FixedRows And nRow < .Rows Then
                CheckedCell(fg, nRow, nCol) = Not CheckedCell(fg, nRow, nCol)
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAndrewFork.HandleGridBeforeMouseDown", eGDRaiseError_Raise
    
End Sub

Private Sub Repaint()
On Error GoTo ErrSection:

    Static bInProgress As Boolean
    
    Dim i&
    Dim strRatios$, strShow$, strColor$, strText$
    Dim bRepaint As Boolean
    
    Dim eMode As eForkMode
    Dim aDates As New cGdArray
    Dim aValues As New cGdArray
    
    If bInProgress Then GoTo ErrExit
    
    bInProgress = True
    bRepaint = m.bChanged
    
    If Not m.Annot Is Nothing Then
        With m.Annot
            'fork mode
            If chkNormal.Value = vbChecked Then
                eMode = eANNOT_ForkNormal
                If chkSchiff.Value = vbChecked Then eMode = eANNOT_NormalSchiff
                If chkSchiffMod.Value = vbChecked Then
                    If chkSchiff.Value = vbChecked Then
                        eMode = eANNOT_ForkAll
                    Else
                        eMode = eANNOT_NormalSchiffMod
                    End If
                End If
            ElseIf chkSchiff.Value = vbChecked Then
                eMode = eANNOT_Schiff
                If chkSchiffMod.Value = vbChecked Then eMode = eANNOT_SchiffModSchiff
            ElseIf chkSchiffMod.Value = vbChecked Then
                eMode = eANNOT_SchiffMod
            Else
                eMode = eANNOT_ForkNormal
                chkNormal.Value = vbChecked         'disallow unchecking all checkboxes
            End If
            If .Prop("ForkMode") <> eMode Then bRepaint = True
            .Prop("ForkMode") = eMode
            
            'AB line
            i = Val(.Prop("HideABLine"))
            If bRepaint Then
                .Prop("HideABLine") = chkHideABLine.Value
            ElseIf i <> chkHideABLine.Value Then
                .Prop("HideABLine") = chkHideABLine.Value
                bRepaint = True
            End If
            
            'show in all panes
            i = Val(.Prop("ShowInAllPanes"))
            If bRepaint Then
                .Prop("ShowInAllPanes") = chkAllPanes.Value
            ElseIf i <> chkAllPanes.Value Then
                .Prop("ShowInAllPanes") = chkAllPanes.Value
                bRepaint = True
            End If
            
            'parallel location
            i = Val(.Prop("ParallelLocation"))
            If bRepaint Then
                .Prop("ParallelLocation") = cboLocation.ListIndex
            ElseIf i <> cboLocation.ListIndex Then
                .Prop("ParallelLocation") = cboLocation.ListIndex
                bRepaint = True
            End If
            
            'multi chart (do not redraw until form is unloaded when user clicks Ok)
            If chkMultiChart.Value = 1 And Not .MultiChartFlag Then
                .MultiChartFlag = True
            ElseIf chkMultiChart.Value = 0 And .MultiChartFlag Then
                .MultiChartFlag = False
            End If
            
            'preindicator
            If bRepaint Then
                m.Annot.PreIndicator = chkPreIndicator.Value
            ElseIf .PreIndicator <> chkPreIndicator.Value Then
                m.Annot.PreIndicator = chkPreIndicator.Value
                bRepaint = True
            End If
            
            'line style for fork
            If bRepaint Then
                .Style = cboStyle.ItemData(cboStyle.ListIndex)
            ElseIf .Style <> cboStyle.ItemData(cboStyle.ListIndex) Then
                .Style = cboStyle.ItemData(cboStyle.ListIndex)
                bRepaint = True
            End If
            
            'color for fork
            If bRepaint Then
                .Color = clrColor.Color
            ElseIf .Color <> clrColor.Color Then
                .Color = clrColor.Color
                bRepaint = True
            End If
            
            'color for markers
            i = clrMarker.Color
            If bRepaint Then
                .Prop("MarkerColor") = i
            ElseIf i <> .Prop("MarkerColor") Then
                .Prop("MarkerColor") = i
                bRepaint = True
            End If
            
            'extend fork's base
            i = 0           'reset
            If chkExtendBase.Value = vbChecked Then i = 1
            If bRepaint Then
                .Prop("ShowHandle") = i
            ElseIf i <> Val(.Prop("ShowHandle")) Then
                .Prop("ShowHandle") = i
                bRepaint = True
            End If
            
            
            'line style for schiff fork
'JM 09-24-2010: use same pen style for Schiff Fork (project meeting mod)
'            i = cboSchiffStyle.ItemData(cboSchiffStyle.ListIndex)
'            If bRepaint Then
'                .Prop("SchiffStyle") = i
'            ElseIf i <> Val(.Prop("SchiffStyle")) Then
'                .Prop("SchiffStyle") = i
'                bRepaint = True
'            End If
            
            'line style for parallels
            i = cboParallelStyle.ItemData(cboParallelStyle.ListIndex)
            If bRepaint Then
                .Prop("ParallelStyle") = i
            ElseIf i <> Val(.Prop("ParallelStyle")) Then
                .Prop("ParallelStyle") = i
                bRepaint = True
            End If
            
            'line style for markers
            i = cboMarkerStyle.ItemData(cboMarkerStyle.ListIndex)
            If bRepaint Then
                .Prop("MarkerStyle") = i
            ElseIf i <> Val(.Prop("MarkerStyle")) Then
                .Prop("MarkerStyle") = i
                bRepaint = True
            End If
            
            'ratios, show, color, text for parallels
            strRatios = .Prop("ParallelRatios")
            strShow = .Prop("ParallelShow")
            strColor = .Prop("ParallelColor")
            
            .FibGridToRatios fgParallel
            
            If Not bRepaint Then
                If .Prop("ParallelRatios") <> strRatios Then
                    bRepaint = True
                ElseIf .Prop("ParallelShow") <> strShow Then
                    bRepaint = True
                ElseIf .Prop("ParallelColor") <> strColor Then
                    bRepaint = True
                End If
            End If

            i = 0           'reset
            If chkParallelText.Value = vbChecked Then i = 1
            If bRepaint Then
                .Prop("ParallelText") = i
            ElseIf i <> Val(.Prop("ParallelText")) Then
                .Prop("ParallelText") = i
                bRepaint = True
            End If
            
            'ratios, show, text for markers
            strRatios = .Prop("MarkerRatios")
            strShow = .Prop("MarkerShow")
            
            .FibGridToRatios fgMarkers
            
            If Not bRepaint Then
                If .Prop("MarkerRatios") <> strRatios Then
                    bRepaint = True
                ElseIf .Prop("MarkerShow") <> strShow Then
                    bRepaint = True
                End If
            End If
        
            i = 0           'reset
            If chkMarkerText.Value = vbChecked Then i = 1
            If bRepaint Then
                .Prop("MarkerText") = i
            ElseIf i <> Val(.Prop("MarkerText")) Then
                .Prop("MarkerText") = i
                bRepaint = True
            End If
            
            'y-values for A,B,C point
            With fgTimeLines
                For i = 1 To .Rows - .FixedRows
                    strText = .TextMatrix(i, 1)
                    If InStr(strText, "^") > 0 Then
                        aValues(i - 1) = m.Chart.Bars.PriceFromString(strText)
                    Else
                        aValues(i - 1) = ValOfText(.TextMatrix(i, 1))   'y-values
                    End If
                    aDates(i - 1) = .RowData(i)
                Next
            End With
            .geSetAndrewForkChange aValues, aDates
        
        End With
    
    End If
    
    If bRepaint Then
        If Not m.Chart Is Nothing Then
            m.Chart.GenerateChart eRedo1_Scrolled
        End If
        m.bChanged = False
    End If

ErrExit:
    bInProgress = False
    Exit Sub

ErrSection:
    bInProgress = False
    RaiseError "frmAndrewFork.Repaint"

End Sub

