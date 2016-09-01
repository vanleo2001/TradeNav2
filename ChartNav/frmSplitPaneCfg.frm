VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSplitPaneCfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Woodies Label Options"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraPriceTimer 
      Height          =   3030
      Left            =   93
      TabIndex        =   12
      Top             =   150
      Width           =   5635
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
      Caption         =   "frmSplitPaneCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSplitPaneCfg.frx":003A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSplitPaneCfg.frx":005A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraMisc 
         Height          =   765
         Left            =   135
         TabIndex        =   13
         Top             =   1110
         Width           =   5340
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
         Caption         =   "frmSplitPaneCfg.frx":0076
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSplitPaneCfg.frx":00A4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":00C4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtTimerHighlight 
            Height          =   285
            Left            =   2085
            TabIndex        =   14
            Top             =   45
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":00E0
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
            Tip             =   "frmSplitPaneCfg.frx":010A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":012A
         End
         Begin gdOCX.gdSelectColor gdTimerHighlight 
            Height          =   375
            Left            =   2625
            TabIndex        =   15
            Top             =   360
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   2760
            Top             =   75
            Width           =   1695
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
            Caption         =   "frmSplitPaneCfg.frx":0146
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0186
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":01A6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   1455
            Top             =   90
            Width           =   825
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
            Caption         =   "frmSplitPaneCfg.frx":01C2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":01EA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":020A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   255
            Left            =   1365
            Top             =   420
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
            Caption         =   "frmSplitPaneCfg.frx":0226
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0264
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0284
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSound 
         Height          =   1110
         Left            =   135
         TabIndex        =   16
         Top             =   1860
         Width           =   5340
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
         Caption         =   "frmSplitPaneCfg.frx":02A0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSplitPaneCfg.frx":02D0
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":02F0
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optOncePerBar 
            Height          =   240
            Left            =   2640
            TabIndex        =   17
            Top             =   300
            Width           =   2175
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
            Caption         =   "frmSplitPaneCfg.frx":030C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0356
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0376
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optEndOfBar 
            Height          =   240
            Left            =   555
            TabIndex        =   18
            Top             =   300
            Width           =   2175
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
            Caption         =   "frmSplitPaneCfg.frx":0392
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":03DC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":03FC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
            Height          =   345
            Left            =   4335
            TabIndex        =   19
            Top             =   645
            Width           =   885
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
            Caption         =   "frmSplitPaneCfg.frx":0418
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0444
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0464
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkPlaySound 
            Height          =   210
            Left            =   100
            TabIndex        =   20
            Top             =   0
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
            Caption         =   "frmSplitPaneCfg.frx":0480
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":04BE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":04DE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtSoundFile 
            Height          =   330
            Left            =   100
            TabIndex        =   21
            Top             =   645
            Width           =   4200
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":04FA
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
            Tip             =   "frmSplitPaneCfg.frx":051A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":053A
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid fgTimer 
         Height          =   810
         Left            =   210
         TabIndex        =   22
         Top             =   240
         Width           =   5220
         _cx             =   9208
         _cy             =   1429
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
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   3060
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5398
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
      Caption         =   "Highlights|Sidewinder|Text Options"
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
      Begin HexUniControls.ctlUniFrameWL fraText 
         Height          =   2685
         Left            =   6480
         TabIndex        =   23
         Top             =   330
         Width           =   5445
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
         Caption         =   "frmSplitPaneCfg.frx":0556
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSplitPaneCfg.frx":0584
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":05A4
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgText 
            Height          =   1215
            Left            =   600
            TabIndex        =   26
            Top             =   720
            Width           =   3495
            _cx             =   6165
            _cy             =   2143
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
      Begin HexUniControls.ctlUniFrameWL fraSidewinder 
         Height          =   2685
         Left            =   6180
         TabIndex        =   5
         Top             =   330
         Width           =   5445
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
         Caption         =   "frmSplitPaneCfg.frx":05C0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSplitPaneCfg.frx":05FA
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":061A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtTrending 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   25
            Top             =   1425
            Width           =   855
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":0636
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
            Tip             =   "frmSplitPaneCfg.frx":0660
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0680
         End
         Begin HexUniControls.ctlUniTextBoxXP txtFlat 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   24
            Top             =   960
            Width           =   855
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":069C
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
            Tip             =   "frmSplitPaneCfg.frx":06C4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":06E4
         End
         Begin HexUniControls.ctlUniLabelXP lblFlat 
            Height          =   255
            Left            =   480
            Top             =   975
            Width           =   4335
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
            Caption         =   "frmSplitPaneCfg.frx":0700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0776
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0796
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblNormal 
            Height          =   255
            Left            =   480
            Top             =   1920
            Width           =   4095
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
            Caption         =   "frmSplitPaneCfg.frx":07B2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0832
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0852
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTrending 
            Height          =   255
            Left            =   480
            Top             =   1447
            Width           =   4095
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
            Caption         =   "frmSplitPaneCfg.frx":086E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":08F4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0914
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblLimitsInfo 
            Height          =   495
            Left            =   750
            Top             =   240
            Width           =   3855
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
            Caption         =   "frmSplitPaneCfg.frx":0930
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0A02
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0A22
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraHighlights 
         Height          =   2685
         Left            =   45
         TabIndex        =   4
         Top             =   330
         Width           =   5445
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
         Caption         =   "frmSplitPaneCfg.frx":0A3E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSplitPaneCfg.frx":0A78
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":0A98
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtPivotUpper 
            Height          =   285
            Left            =   3000
            TabIndex        =   10
            Top             =   1424
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":0AB4
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
            Tip             =   "frmSplitPaneCfg.frx":0ADE
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0AFE
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPivotLower 
            Height          =   285
            Left            =   3000
            TabIndex        =   9
            Top             =   1806
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":0B1A
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
            Tip             =   "frmSplitPaneCfg.frx":0B44
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0B64
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDayHigh 
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   660
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":0B80
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
            Tip             =   "frmSplitPaneCfg.frx":0BAA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0BCA
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDayLow 
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Top             =   1042
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":0BE6
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
            Tip             =   "frmSplitPaneCfg.frx":0C10
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0C30
         End
         Begin HexUniControls.ctlUniTextBoxXP txtTimerSeconds 
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Top             =   2190
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSplitPaneCfg.frx":0C4C
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
            Tip             =   "frmSplitPaneCfg.frx":0C76
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0C96
         End
         Begin gdOCX.gdSelectColor gdHighlightColor 
            Height          =   375
            Left            =   2550
            TabIndex        =   11
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   1350
            Top             =   240
            Width           =   1815
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
            Caption         =   "frmSplitPaneCfg.frx":0CB2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0CF2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0D12
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPivotUpper 
            Height          =   255
            Left            =   120
            Top             =   1439
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
            Caption         =   "frmSplitPaneCfg.frx":0D2E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0D9C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0DBC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPivotUpper2 
            Height          =   255
            Left            =   3720
            Top             =   1439
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
            Caption         =   "frmSplitPaneCfg.frx":0DD8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0E24
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0E44
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPivotLower 
            Height          =   255
            Left            =   120
            Top             =   1821
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
            Caption         =   "frmSplitPaneCfg.frx":0E60
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0ECE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0EEE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPivotLower2 
            Height          =   255
            Left            =   3720
            Top             =   1821
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
            Caption         =   "frmSplitPaneCfg.frx":0F0A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0F56
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":0F76
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDayHigh 
            Height          =   255
            Left            =   120
            Top             =   675
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
            Caption         =   "frmSplitPaneCfg.frx":0F92
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":0FF4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":1014
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDayHigh2 
            Height          =   255
            Left            =   3240
            Top             =   675
            Width           =   1695
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
            Caption         =   "frmSplitPaneCfg.frx":1030
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":107C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":109C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDayLow 
            Height          =   255
            Left            =   120
            Top             =   1057
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
            Caption         =   "frmSplitPaneCfg.frx":10B8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":1118
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":1138
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDayLow2 
            Height          =   255
            Left            =   3240
            Top             =   1057
            Width           =   1695
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
            Caption         =   "frmSplitPaneCfg.frx":1154
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":11A0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":11C0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label10 
            Height          =   255
            Left            =   120
            Top             =   2205
            Width           =   2415
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
            Caption         =   "frmSplitPaneCfg.frx":11DC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":1234
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":1254
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label11 
            Height          =   255
            Left            =   3120
            Top             =   2205
            Width           =   1695
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
            Caption         =   "frmSplitPaneCfg.frx":1270
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSplitPaneCfg.frx":12B0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSplitPaneCfg.frx":12D0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1816
      TabIndex        =   0
      Top             =   3300
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
      Caption         =   "frmSplitPaneCfg.frx":12EC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSplitPaneCfg.frx":1318
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSplitPaneCfg.frx":1338
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   60
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
         Caption         =   "frmSplitPaneCfg.frx":1354
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSplitPaneCfg.frx":1380
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":13A0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   60
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
         Caption         =   "frmSplitPaneCfg.frx":13BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSplitPaneCfg.frx":13E0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSplitPaneCfg.frx":1400
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmSplitPaneCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kCols = 4
Private Const kRows = 11

Private Enum eGDCols            'text option grid columns
    eGDCol_Show = 0
    eGDCol_Value
    eGDCol_Font
    eGDCol_Placement
End Enum

Private Type mPrivate
    Pane As cPane
    PaneWood As cPaneWood
    Chart As cChart
    'alert levels
    nColor As Long
    nDayHigh As Long
    nDayLow As Long
    nUpperPivot As Long
    nLowerPivot As Long
    dTimerS As Double
End Type

Private m As mPrivate

Public Sub ShowMe(Chart As cChart, ByVal nPaneID&, ByVal nLabelIdx&)
On Error GoTo ErrSection:
    
    Dim dTimerT#, i&
    Dim fg As VSFlexGrid
    
    If Chart Is Nothing Then Exit Sub
     
    Set m.Chart = Chart
    
    Set m.Pane = m.Chart.Tree(nPaneID)
    If m.Pane Is Nothing Then Exit Sub
    
    Set m.PaneWood = m.Pane.WoodPane
    If m.PaneWood Is Nothing Then Exit Sub
    
    If m.Pane.PricePaneFlag = 1 Then
        Me.Caption = "Price Pane Timer Options"
        vsIndexTab1.Visible = False
        fraPriceTimer.Visible = True
        'countdown timer highlight
        m.PaneWood.GetHighlightLevels m.nColor, m.nDayHigh, m.nDayLow, m.nUpperPivot, m.nLowerPivot, m.dTimerS, dTimerT
        gdTimerHighlight.Color = m.nColor
        txtTimerHighlight.Text = Str(Int(m.dTimerS * 100 + 0.5))
        'countdown timer sound options
        InitSoundCtrls
        EnableSoundCtrls
        'text grid
        Set fg = fgTimer
        InitGrid fg
    Else
        Me.Caption = "Woodies Labels Options"
        vsIndexTab1.Visible = True
        fraPriceTimer.Visible = False
        'highlights tab
        m.PaneWood.GetHighlightLevels m.nColor, m.nDayHigh, m.nDayLow, m.nUpperPivot, m.nLowerPivot, m.dTimerS, dTimerT
        gdHighlightColor.Color = m.nColor
        txtDayHigh.Text = m.nDayHigh
        txtDayLow.Text = m.nDayLow
        txtPivotUpper.Text = m.nUpperPivot
        txtPivotLower.Text = m.nLowerPivot
        txtTimerSeconds.Text = m.dTimerS * 100
        EnableHighlightCtls True
        'sidewinder tab
        txtFlat.Text = Str(m.PaneWood.SideWinderFlat)
        txtTrending.Text = Str(m.PaneWood.SideWinderTrending)
        'text tab
        Set fg = fgText
        InitGrid fg
        i = m.Pane.SplitPaneLabelIdxCCI(nLabelIdx)
        If i >= 0 And i < fg.Rows Then
            i = i + fg.FixedRows
            vsIndexTab1.CurrTab = 2
            fg.Select i, 0, i, kCols - 1
        Else
            vsIndexTab1.CurrTab = 0
        End If
    End If
            
    CenterFormOnChart Me, Chart                 '6499
    
    ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
    
    Exit Sub
    
ErrSection:
    RaiseError "frmSplitPaneCfg.ShowMe"
    
End Sub

Private Sub EnableHighlightCtls(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    txtDayHigh.Enabled = bEnable
    txtDayLow.Enabled = bEnable
    txtPivotUpper.Enabled = bEnable
    txtPivotLower.Enabled = bEnable
    
    lblDayHigh.Enabled = bEnable
    lblDayLow.Enabled = bEnable
    lblPivotUpper.Enabled = bEnable
    lblPivotLower.Enabled = bEnable
    
    lblDayHigh2.Enabled = bEnable
    lblDayLow2.Enabled = bEnable
    lblPivotUpper2.Enabled = bEnable
    lblPivotLower2.Enabled = bEnable
    
    Exit Sub
    
ErrSection:
    RaiseError "frmSplitPaneCfg.EnableHighlightsCtl"
    
End Sub

Private Sub chkPlaySound_Click()
    EnableSoundCtrls
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
                
                
    Dim i&, iTrend&, iFlat&
    Dim Tree As cGdTree
    Dim Ind As cIndicator
    
    Dim eRedoMode As eChartRedoMode
                
    If m.Pane.PricePaneFlag = 0 Then
        eRedoMode = eRedo1_Scrolled
        'highlights
        m.PaneWood.HighlightColor = gdHighlightColor.Color
        m.PaneWood.HighlightDayHigh = m.nDayHigh
        m.PaneWood.HighlightDayLow = m.nDayLow
        m.PaneWood.HighlightPivotUpper = m.nUpperPivot
        m.PaneWood.HighlightPivotLower = m.nLowerPivot
        m.PaneWood.HighlightTimerS = m.dTimerS
        'sidewinder
        iTrend = RoundNum(ValOfText(txtTrending.Text), 0)
        iFlat = RoundNum(ValOfText(txtFlat.Text), 0)
        
        If iTrend <> m.PaneWood.SideWinderTrending Or iFlat <> m.PaneWood.SideWinderFlat Then
            m.PaneWood.SideWinderTrending = iTrend
            m.PaneWood.SideWinderFlat = iFlat
            eRedoMode = eRedo3_Settings
            'change parameters indicators (i.e. this is behaving as if user changed input values in chart editor)
            Set Tree = m.Chart.Tree
            If Not Tree Is Nothing Then
                For i = 1 To Tree.Count
                    If Tree.NodeLevel(i) > 0 Then
                        Set Ind = Tree(i)
                        If Not Ind Is Nothing Then
                            If UCase(Ind.CodedName) = "SIDEWINDERTRENDING" Then
                                If Ind.ParmCount > 1 Then
                                    Ind.Parm(2) = Str(iTrend)
                                End If
                            ElseIf UCase(Ind.CodedName) = "SIDEWINDERNORMAL" Then
                                If Ind.ParmCount > 1 Then
                                    Ind.Parm(2) = Str(iTrend)
                                End If
                                If Ind.ParmCount > 2 Then
                                    Ind.Parm(3) = Str(iFlat)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        'redraw
        m.Chart.GenerateChart eRedoMode
    ElseIf m.Pane.PricePaneFlag = 1 Then
        If vsIndexTab1.Visible Then
            m.PaneWood.HighlightColor = gdHighlightColor.Color
        Else
            m.PaneWood.HighlightColor = gdTimerHighlight.Color
        End If
        If chkPlaySound.Value = vbChecked Then
            If Len(txtSoundFile.Text) = 0 Then
                InfBox "Please specify a sound file.", "I"
                Exit Sub
            End If
            SetIniFileProperty "WavFileLastUsed", txtSoundFile.Text, "QuoteBoard", g.strIniFile
            If optOncePerBar.Value = True Then
                m.PaneWood.PlaySound = 1
            Else
                m.PaneWood.PlaySound = -1
            End If
        Else
            m.PaneWood.PlaySound = 0
        End If
        
        m.PaneWood.HighlightTimerS = m.dTimerS
        m.PaneWood.SoundFile = txtSoundFile.Text
        m.Chart.GenerateChart eRedo1_Scrolled
    End If
    
    Unload Me
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.cmdOK_Click"

End Sub

Private Function RowToIdxCCI(ByVal nRow&) As Long
On Error GoTo ErrSection:

    Dim i&, nIdx&
    
    i = fgText.FixedRows
    
    If m.Pane.PricePaneFlag = 1 Then
        RowToIdxCCI = -1
    Else
        Select Case nRow
            Case eCCI_DayHigh + i
                nIdx = eCCI_DayHigh
            Case eCCI_PivotAbove + i
                nIdx = eCCI_PivotAbove
            Case eCCI_BarHigh + i
                nIdx = eCCI_BarHigh
            Case eCCI_Price + i
                nIdx = eCCI_Price
            Case eCCI_BarLow + i
                nIdx = eCCI_BarLow
            Case eCCI_EMA + i
                nIdx = eCCI_EMA
            Case eCCI_Timer + i
                nIdx = eCCI_Timer
            Case eCCI_PivotBelow + i
                nIdx = eCCI_PivotBelow
            Case eCCI_DayLow + i
                nIdx = eCCI_DayLow
            Case eCCI_SideWinder + i
                nIdx = eCCI_SideWinder
            Case Else
                nIdx = -1
        End Select
    End If
    
    RowToIdxCCI = nIdx
    
    Exit Function

ErrSection:
    RaiseError "frmSplitPaneCfg.RowToIdxCCI"
    
End Function

Private Sub cmdBrowse_Click()
On Error Resume Next

    Dim strFile$

    If Len(m.PaneWood.SoundFile) > 0 Then
        strFile = m.PaneWood.SoundFile
    Else
        strFile = GetIniFileProperty("WavFileLastUsed", "", "QuoteBoard", g.strIniFile)
    End If
    
    txtSoundFile.Text = CommonDialogFile(frmMain.CommonDialog1, False, "Wave Files (*.wav)|*.wav", strFile)

End Sub

Private Sub fgText_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    HandleAfterEdit fgText, Row, Col
    
End Sub

Private Sub fgText_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = eGDCol_Value Then Cancel = True
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.fgText_BeforeEdit"

End Sub

Private Sub fgText_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    
    HandleCellBtnClick fgText, Row, Col
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.fgText_CellButtonClick"

End Sub

Private Sub fgTimer_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    HandleAfterEdit fgTimer, Row, Col
    
End Sub

Private Sub fgTimer_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = eGDCol_Value Then Cancel = True
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.fgTimer_BeforeEdit"

End Sub

Private Sub fgTimer_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    HandleCellBtnClick fgTimer, Row, Col

    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.fgTimer_CellButtonClick"

End Sub

Private Sub Form_Load()
    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
End Sub

Private Sub Form_Resize()
On Error Resume Next

    With fgText
        .Move 0, 0, fraText.Width, fraText.Height
    End With
    
    If Not m.Pane Is Nothing Then
        If m.Pane.SplitPaneType = ePANE_SplitPaneTimer Then
            With fraPriceTimer
                .Move .Left, 0, .Width, 3650            '5119
                fraMisc.Top = fgTimer.Top + fgTimer.Height + 100
                fraSound.Top = fraMisc.Top + fraMisc.Height + 50
                fraButtons.Top = .Height + 30
            End With
            Me.Height = 4700
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set m.Chart = Nothing
    Set m.Pane = Nothing
    Set m.PaneWood = Nothing

End Sub

Private Sub gdHighlightColor_LostFocus()
On Error GoTo ErrSection:

    m.nColor = gdHighlightColor.Color
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.gdHighlightColor_LostFocus"

End Sub

Private Sub txtDayHigh_Change()
On Error GoTo ErrSection:

    m.nDayHigh = Int(ValOfText(txtDayHigh.Text))
    
    Exit Sub
    
ErrSection:
    RaiseError "frmSplitPaneCfg.txtDayHigh_Change"

End Sub

Private Sub txtDayLow_Change()
On Error GoTo ErrSection:

    m.nDayLow = Int(ValOfText(txtDayLow.Text))
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.txtDayLow_Change"

End Sub

Private Sub txtPivotLower_Change()
On Error GoTo ErrSection:

    m.nLowerPivot = Int(ValOfText(txtPivotLower.Text))
    
    Exit Sub
    
ErrSection:
    RaiseError "frmSplitPaneCfg.txtPivotLower_Change"

End Sub

Private Sub txtPivotUpper_Change()
On Error GoTo ErrSection:

    m.nUpperPivot = Int(ValOfText(txtPivotUpper.Text))
    
    Exit Sub
    
ErrSection:
    RaiseError "frmSplitPaneCfg.txtPivotUpper_Change"

End Sub

Private Sub txtTimerHighlight_Change()
On Error GoTo ErrSection:

    m.dTimerS = Int(ValOfText(txtTimerHighlight.Text)) / 100
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.txtTimerHighlight_Change"


End Sub

Private Sub txtTimerSeconds_Change()
On Error GoTo ErrSection:

    m.dTimerS = Int(ValOfText(txtTimerSeconds.Text)) / 100
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.txtTimerSeconds_Change"

End Sub

Private Sub PaneDataToGrid(ByVal nRow&, ByVal eIdx As eCCIText_Index, _
    fg As VSFlexGrid, Optional ByVal strType As String)
On Error GoTo ErrSection:

    Dim strText$, strPlacement$, i&
    
    strText = m.PaneWood.GetCCITextInfo(eIdx, strType)
    
'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    i = ValOfText(Parse(strText, "|", 9))
    fg.TextMatrix(nRow, eGDCol_Font) = Parse(strText, "|", 2)
    strPlacement = Parse(strText, "|", 8)
    If strPlacement = "T" Then
        strPlacement = "T (top)"
    ElseIf strPlacement = "B" Then
        strPlacement = "B (bottom)"
    ElseIf strPlacement = "N" Then
        strPlacement = "N (next)"
    End If
    fg.TextMatrix(nRow, eGDCol_Placement) = strPlacement
    If i = 1 Then
        fg.Cell(flexcpChecked, nRow, eGDCol_Show) = 1
    Else
        fg.Cell(flexcpChecked, nRow, eGDCol_Show) = 2
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmSplitPaneCfg.PaneDataToGrid"

End Sub

Private Sub InitGrid(fg As VSFlexGrid)
On Error GoTo ErrSection:
            
    With fg
        .Redraw = flexRDNone
                
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
    
        .FixedRows = 1
        .Cols = kCols
        .FixedCols = 0
        .ColAlignment(eGDCol_Placement) = flexAlignRightCenter
        .Cell(flexcpPictureAlignment, 0, 0, kRows - 1, 0) = flexPicAlignCenterCenter
        
        'column headers
        .TextMatrix(0, eGDCol_Show) = "Show"
        .TextMatrix(0, eGDCol_Value) = "Value"
        .TextMatrix(0, eGDCol_Font) = "Font"
        .ColComboList(eGDCol_Font) = "..."
        .TextMatrix(0, eGDCol_Placement) = "Placement"
        
        If m.Pane.PricePaneFlag = 1 Then
            Select Case m.Pane.SplitPaneType
                Case ePANE_SplitPaneWood
                    .Rows = 3
                    .TextMatrix(1, eGDCol_Value) = "Price"
                    PaneDataToGrid 1, 0, fg, "PP_Price"
                    .TextMatrix(2, eGDCol_Value) = "Timer"
                    PaneDataToGrid 2, 0, fg, "PP_Timer"
                Case ePANE_SplitPaneTimer
                    .Rows = 5
                    .TextMatrix(1, eGDCol_Value) = "Price"
                    PaneDataToGrid 1, 0, fg, "PP_Price"
                    .TextMatrix(2, eGDCol_Value) = "Price Change"
                    PaneDataToGrid 2, 0, fg, "PP_PriceDelta"
                    .TextMatrix(3, eGDCol_Value) = "% Change"
                    PaneDataToGrid 3, 0, fg, "PP_PricePercent"
                    .TextMatrix(4, eGDCol_Value) = "Timer"
                    PaneDataToGrid 4, 0, fg, "PP_Timer"
                Case ePANE_SplitPaneTimer
            End Select
            .Height = .RowHeight(1) * .Rows + 30
            .ScrollBars = flexScrollBarNone
        Else
            .Rows = kRows
            'CCI pane
            .TextMatrix(eCCI_DayHigh + .FixedRows, eGDCol_Value) = "Day High"
            PaneDataToGrid eCCI_DayHigh + .FixedRows, eCCI_DayHigh, fg
    
            .TextMatrix(eCCI_PivotAbove + .FixedRows, eGDCol_Value) = "Pivot (above price)"
            PaneDataToGrid eCCI_PivotAbove + .FixedRows, eCCI_PivotAbove, fg
    
            .TextMatrix(eCCI_BarHigh + .FixedRows, eGDCol_Value) = "Prev/Curr Bar High"
            PaneDataToGrid eCCI_BarHigh + .FixedRows, eCCI_BarHigh, fg
    
            .TextMatrix(eCCI_Price + .FixedRows, eGDCol_Value) = "Price"
            PaneDataToGrid eCCI_Price + .FixedRows, eCCI_Price, fg
    
            .TextMatrix(eCCI_BarLow + .FixedRows, eGDCol_Value) = "Prev/Curr Bar Low"
            PaneDataToGrid eCCI_BarLow + .FixedRows, eCCI_BarLow, fg
    
            .TextMatrix(eCCI_EMA + .FixedRows, eGDCol_Value) = "EMA"
            PaneDataToGrid eCCI_EMA + .FixedRows, eCCI_EMA, fg
    
            .TextMatrix(eCCI_Timer + .FixedRows, eGDCol_Value) = "Timer"
            PaneDataToGrid eCCI_Timer + .FixedRows, eCCI_Timer, fg
    
            .TextMatrix(eCCI_PivotBelow + .FixedRows, eGDCol_Value) = "Pivot (below price)"
            PaneDataToGrid eCCI_PivotBelow + .FixedRows, eCCI_PivotBelow, fg
    
            .TextMatrix(eCCI_DayLow + .FixedRows, eGDCol_Value) = "Day Low"
            PaneDataToGrid eCCI_DayLow + .FixedRows, eCCI_DayLow, fg
    
            .TextMatrix(eCCI_SideWinder + .FixedRows, eGDCol_Value) = "Sidewinder"
            PaneDataToGrid eCCI_SideWinder + .FixedRows, eCCI_SideWinder, fg
            .TextMatrix(eCCI_SideWinder + .FixedRows, eGDCol_Placement) = "N/A"
        End If

        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, kCols - 1
        
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.InitGrid"
    
End Sub

Private Sub HandleAfterEdit(fg As VSFlexGrid, ByVal Row&, ByVal Col&)
On Error GoTo ErrSection:

    Dim strType$, strInfo$, strText$, strPlacement$
    Dim i&, iShow&, dPlacement#
    Dim eIdx As eCCIText_Index
    
    If m.Pane.PricePaneFlag = 1 Then
        With fg
            If Row >= .FixedRows And Row < .Rows Then
                Select Case .TextMatrix(Row, eGDCol_Value)
                    Case "Price"
                        strType = "PP_Price"
                    Case "Price Change"
                        strType = "PP_PriceDelta"
                    Case "% Change"
                        strType = "PP_PricePercent"
                    Case "Timer"
                        strType = "PP_Timer"
                End Select
            End If
        End With
    Else
        i = RowToIdxCCI(Row)
    End If
    
    If i < 0 Then
        Exit Sub
    Else
        eIdx = i
    End If
    
'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    strInfo = m.PaneWood.GetCCITextInfo(eIdx, strType)
    
    If Col = eGDCol_Show Or Col = eGDCol_Placement Then
        If fg.Cell(flexcpChecked, Row, eGDCol_Show) = 1 Then iShow = 1
        
        If Col = eGDCol_Placement And eIdx <> eCCI_SideWinder Then
            strText = fg.TextMatrix(Row, eGDCol_Placement)
            If InStr(UCase(strText), "T") Then
                strPlacement = "T"
            ElseIf InStr(UCase(strText), "B") Then
                strPlacement = "B"
            ElseIf InStr(UCase(strText), "N") Then
                strPlacement = "N"
            Else
                dPlacement = ValOfText(strText)
                strPlacement = CStr(dPlacement)
            End If
        Else
            strPlacement = Parse(strInfo, "|", 8)
        End If
        
        strText = Parse(strInfo, "|", 1)
        strText = strText & "|" & Parse(strInfo, "|", 2)
        strText = strText & "|" & Parse(strInfo, "|", 3)
        strText = strText & "|" & Parse(strInfo, "|", 4)
        strText = strText & "|" & Parse(strInfo, "|", 5)
        strText = strText & "|" & Parse(strInfo, "|", 6)
        strText = strText & "|" & Parse(strInfo, "|", 7)
        strText = strText & "|" & strPlacement
        strText = strText & "|" & Str(iShow)
        m.PaneWood.SetCCITextInfo eIdx, strText, strType
        m.Chart.GenerateChart eRedo1_Scrolled
    End If

    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.HandleAfterEdit"

End Sub

Private Sub HandleCellBtnClick(fg As VSFlexGrid, ByVal Row&, ByVal Col&)
On Error GoTo ErrSection:

    Dim strType$, strInfo$, strNewInfo$, i&
    Dim eIdx As eCCIText_Index

    If Col <> eGDCol_Font Then Exit Sub
        
    If m.Pane.PricePaneFlag = 1 Then
        With fg
            If Row >= .FixedRows And Row < .Rows Then
                Select Case .TextMatrix(Row, eGDCol_Value)
                    Case "Price"
                        strType = "PP_Price"
                    Case "Price Change"
                        strType = "PP_PriceDelta"
                    Case "% Change"
                        strType = "PP_PricePercent"
                    Case "Timer"
                        strType = "PP_Timer"
                End Select
            End If
        End With
    Else
        i = RowToIdxCCI(Row)
    End If
    
    If i < 0 Then
        Exit Sub
    Else
        eIdx = i
    End If
    
'   textType|font|fontSize|boldFlag|italicFlag|textColor|textBkColor|location|show
    strInfo = m.PaneWood.GetCCITextInfo(eIdx, strType)
    
    Me.Font.Name = Parse(strInfo, "|", 2)
    Me.Font.Size = ValOfText(Parse(strInfo, "|", 3))
    Me.Font.Bold = ValOfText(Parse(strInfo, "|", 4)) * -1
    Me.Font.Italic = ValOfText(Parse(strInfo, "|", 5)) * -1
    
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        strNewInfo = Parse(strInfo, "|", 1)
        strNewInfo = strNewInfo & "|" & Me.Font.Name
        strNewInfo = strNewInfo & "|" & CStr(Me.Font.Size)
        strNewInfo = strNewInfo & "|" & CStr(Abs(Me.Font.Bold))
        strNewInfo = strNewInfo & "|" & CStr(Abs(Me.Font.Italic))
        strNewInfo = strNewInfo & "|" & Parse(strInfo, "|", 6)
        strNewInfo = strNewInfo & "|" & Parse(strInfo, "|", 7)
        strNewInfo = strNewInfo & "|" & Parse(strInfo, "|", 8)
        strNewInfo = strNewInfo & "|" & Parse(strInfo, "|", 9)
        m.PaneWood.SetCCITextInfo eIdx, strNewInfo, strType
        fg.TextMatrix(Row, eGDCol_Font) = Me.Font.Name
        m.Chart.ResetSplitPane
        m.Chart.geForceRecalc
        m.Chart.GenerateChart eRedo3_Settings
    End If

    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.HandleCellBtnClick"

End Sub

Private Sub InitSoundCtrls()
On Error GoTo ErrSection:

    Dim i&
    
    i = m.PaneWood.PlaySound
    If i = 1 Or i = -1 Then
        chkPlaySound.Value = vbChecked
        If i = 1 Then
            optEndOfBar.Value = False
            optOncePerBar.Value = True
        Else
            optEndOfBar.Value = True
            optOncePerBar.Value = False
        End If
    Else
        chkPlaySound.Value = vbUnchecked
        optEndOfBar.Value = False
        optOncePerBar.Value = False
    End If
    
    txtSoundFile.Text = m.PaneWood.SoundFile
    
    Exit Sub

ErrSection:
    RaiseError "frmSplitPaneCfg.InitSoundCtrls"
    
End Sub

Private Sub EnableSoundCtrls()
On Error Resume Next

    Dim bEnable As Boolean
    
    If chkPlaySound.Value = vbChecked Then
        bEnable = True
        If optEndOfBar.Value = False And optOncePerBar.Value = False Then
            optEndOfBar.Value = True
        End If
        If Len(m.PaneWood.SoundFile) > 0 Then
            txtSoundFile.Text = m.PaneWood.SoundFile
        Else
            txtSoundFile.Text = GetIniFileProperty("WavFileLastUsed", "", "QuoteBoard", g.strIniFile)
        End If
    End If
    
    optEndOfBar.Enabled = bEnable
    optOncePerBar.Enabled = bEnable
    cmdBrowse.Enabled = bEnable
    txtSoundFile.Enabled = bEnable

End Sub

