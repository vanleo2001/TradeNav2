VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMarketDepthCfg 
   Caption         =   "Market Depth"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   4380
      Left            =   150
      TabIndex        =   8
      Top             =   1114
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7726
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
      Caption         =   "Bid/Ask|Quote/Order Bar"
      Align           =   0
      Appearance      =   1
      CurrTab         =   -1
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
      Begin HexUniControls.ctlUniFrameWL fraTabQuoteBar 
         Height          =   4035
         Left            =   3750
         TabIndex        =   18
         Top             =   315
         Width           =   1845
         _ExtentX        =   3254
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
         Caption         =   "frmMarketDepthCfg.frx":0000
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketDepthCfg.frx":003C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":005C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraTabOrderBar 
         Height          =   4035
         Left            =   1890
         TabIndex        =   32
         Top             =   315
         Width           =   1845
         _ExtentX        =   3254
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
         Caption         =   "frmMarketDepthCfg.frx":0078
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketDepthCfg.frx":00B4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":00D4
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgQBarCols 
            Height          =   2955
            Left            =   2880
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   2205
            _cx             =   3889
            _cy             =   5212
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
         Begin HexUniControls.ctlUniButtonImageXP cmdConfigQuote 
            Height          =   255
            Left            =   1905
            TabIndex        =   20
            Top             =   300
            Width           =   915
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
            Caption         =   "frmMarketDepthCfg.frx":00F0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":0122
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":0142
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkQuoteBar 
            Height          =   220
            Left            =   195
            TabIndex        =   21
            Top             =   345
            Width           =   1905
            _ExtentX        =   3360
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
            Caption         =   "frmMarketDepthCfg.frx":015E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":019C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":01BC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkOrderBar 
            Height          =   220
            Left            =   195
            TabIndex        =   40
            Top             =   1505
            Width           =   1755
            _ExtentX        =   3096
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
            Caption         =   "frmMarketDepthCfg.frx":01D8
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":0214
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":0234
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAccountBar 
            Height          =   220
            Left            =   195
            TabIndex        =   39
            Top             =   925
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "frmMarketDepthCfg.frx":0250
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":0290
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":02B0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQty3 
            Height          =   315
            Left            =   1665
            TabIndex        =   38
            Top             =   2835
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmMarketDepthCfg.frx":02CC
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
            Tip             =   "frmMarketDepthCfg.frx":02EE
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":030E
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQty2 
            Height          =   315
            Left            =   1065
            TabIndex        =   37
            Top             =   2835
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmMarketDepthCfg.frx":032A
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
            Tip             =   "frmMarketDepthCfg.frx":034C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":036C
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQty1 
            Height          =   315
            Left            =   465
            TabIndex        =   36
            Top             =   2835
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmMarketDepthCfg.frx":0388
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
            Tip             =   "frmMarketDepthCfg.frx":03AA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":03CA
         End
         Begin HexUniControls.ctlUniCheckXP chkOrdBarOnRight 
            Height          =   220
            Left            =   195
            TabIndex        =   35
            Top             =   2085
            Width           =   2385
            _ExtentX        =   4207
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
            Caption         =   "frmMarketDepthCfg.frx":03E6
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":0436
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":0456
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdConfigAcct 
            Height          =   255
            Left            =   1905
            TabIndex        =   34
            Top             =   895
            Width           =   915
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
            Caption         =   "frmMarketDepthCfg.frx":0472
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":04A4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":04C4
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdConfigOrder 
            Height          =   255
            Left            =   1905
            TabIndex        =   33
            Top             =   1475
            Width           =   915
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
            Caption         =   "frmMarketDepthCfg.frx":04E0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":0512
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":0532
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgAcctCols 
            Height          =   1785
            Left            =   2940
            TabIndex        =   41
            Top             =   300
            Visible         =   0   'False
            Width           =   2355
            _cx             =   4154
            _cy             =   3149
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
         Begin VSFlex7LCtl.VSFlexGrid fgOrderButtons 
            Height          =   3685
            Left            =   3150
            TabIndex        =   42
            Top             =   195
            Visible         =   0   'False
            Width           =   2355
            _cx             =   4154
            _cy             =   6500
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
            Rows            =   16
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
         Begin HexUniControls.ctlUniLabelXP Label13 
            Height          =   255
            Left            =   525
            Top             =   2580
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
            Caption         =   "frmMarketDepthCfg.frx":054E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":059C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":05BC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraTabBidAsk 
         Height          =   4035
         Left            =   30
         TabIndex        =   9
         Top             =   315
         Width           =   1845
         _ExtentX        =   3254
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
         Caption         =   "frmMarketDepthCfg.frx":05D8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketDepthCfg.frx":0610
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":0630
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkSummaryBar 
            Height          =   300
            Left            =   2955
            TabIndex        =   22
            Top             =   2160
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
            Caption         =   "frmMarketDepthCfg.frx":064C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketDepthCfg.frx":068C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":06AC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraBidAskSizeOptions 
            Height          =   1635
            Left            =   90
            TabIndex        =   25
            Top             =   2175
            Width           =   2565
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
            Caption         =   "frmMarketDepthCfg.frx":06C8
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketDepthCfg.frx":0712
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":0732
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optShowTriangles 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   600
               Width           =   1875
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
               Caption         =   "frmMarketDepthCfg.frx":074E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":078C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":07AC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optShowTriangles 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   270
               Width           =   1875
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
               Caption         =   "frmMarketDepthCfg.frx":07C8
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":080A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":082A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectColor gdBidAskUpColor 
               Height          =   315
               Left            =   1650
               TabIndex        =   28
               Top             =   870
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin gdOCX.gdSelectColor gdBidAskDownColor 
               Height          =   315
               Left            =   1650
               TabIndex        =   29
               Top             =   1230
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniLabelXP lblBidAskDownColor 
               Height          =   255
               Left            =   150
               Top             =   1260
               Width           =   1875
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
               Caption         =   "frmMarketDepthCfg.frx":0846
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":088C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":08AC
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblBidAskUpColor 
               Height          =   255
               Left            =   150
               Top             =   930
               Width           =   1875
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
               Caption         =   "frmMarketDepthCfg.frx":08C8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":090E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":092E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraBidAskColors 
            Height          =   1995
            Left            =   90
            TabIndex        =   10
            Top             =   90
            Width           =   5400
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
            Caption         =   "frmMarketDepthCfg.frx":094A
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketDepthCfg.frx":0986
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":09A6
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   1
               Left            =   1515
               TabIndex        =   11
               Top             =   705
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   2
               Left            =   1515
               TabIndex        =   12
               Top             =   1095
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   3
               Left            =   3930
               TabIndex        =   13
               Top             =   315
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   4
               Left            =   3930
               TabIndex        =   14
               Top             =   705
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   5
               Left            =   3930
               TabIndex        =   15
               Top             =   1095
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   6
               Left            =   2640
               TabIndex        =   16
               Top             =   1575
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   0
               Left            =   1515
               TabIndex        =   17
               Top             =   315
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin HexUniControls.ctlUniLabelXP Label5 
               Height          =   255
               Left            =   375
               Top             =   375
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
               Caption         =   "frmMarketDepthCfg.frx":09C2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":09FA
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0A1A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label6 
               Height          =   255
               Left            =   375
               Top             =   765
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
               Caption         =   "frmMarketDepthCfg.frx":0A36
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0A66
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0A86
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label9 
               Height          =   255
               Left            =   375
               Top             =   1155
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
               Caption         =   "frmMarketDepthCfg.frx":0AA2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0AD2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0AF2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label10 
               Height          =   255
               Left            =   2850
               Top             =   375
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
               Caption         =   "frmMarketDepthCfg.frx":0B0E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0B3E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0B5E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label11 
               Height          =   255
               Left            =   2850
               Top             =   765
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
               Caption         =   "frmMarketDepthCfg.frx":0B7A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0BAA
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0BCA
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label4 
               Height          =   255
               Left            =   2850
               Top             =   1155
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
               Caption         =   "frmMarketDepthCfg.frx":0BE6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0C20
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0C40
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label12 
               Height          =   255
               Left            =   1620
               Top             =   1635
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
               Caption         =   "frmMarketDepthCfg.frx":0C5C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0C8E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0CAE
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraSummaryBar 
            Height          =   1635
            Left            =   2805
            TabIndex        =   23
            Top             =   2175
            Width           =   2685
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
            Caption         =   "frmMarketDepthCfg.frx":0CCA
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketDepthCfg.frx":0D04
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketDepthCfg.frx":0D24
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optSummaryBar 
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   24
               Top             =   390
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
               Caption         =   "frmMarketDepthCfg.frx":0D40
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0D8E
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0DAE
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSummaryBar 
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   30
               Top             =   787
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
               Caption         =   "frmMarketDepthCfg.frx":0DCA
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0E12
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0E32
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtSummaryBarHeight 
               Height          =   315
               Left            =   1860
               TabIndex        =   31
               Top             =   1125
               Width           =   630
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketDepthCfg.frx":0E4E
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
               Tip             =   "frmMarketDepthCfg.frx":0E6E
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0E8E
            End
            Begin HexUniControls.ctlUniLabelXP lblSummaryBarHeight 
               Height          =   255
               Left            =   120
               Top             =   1185
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
               Caption         =   "frmMarketDepthCfg.frx":0EAA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketDepthCfg.frx":0EF2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketDepthCfg.frx":0F12
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1860
      TabIndex        =   5
      Top             =   5584
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
      Caption         =   "frmMarketDepthCfg.frx":0F2E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMarketDepthCfg.frx":0F5A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMarketDepthCfg.frx":0F7A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
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
         Caption         =   "frmMarketDepthCfg.frx":0F96
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarketDepthCfg.frx":0FC2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":0FE2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   60
         TabIndex        =   6
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
         Caption         =   "frmMarketDepthCfg.frx":0FFE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMarketDepthCfg.frx":1022
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":1042
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdFont 
      Height          =   315
      Left            =   4770
      TabIndex        =   4
      Top             =   611
      Width           =   1035
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
      Caption         =   "frmMarketDepthCfg.frx":105E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMarketDepthCfg.frx":1088
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMarketDepthCfg.frx":10A8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraSession 
      Height          =   975
      Left            =   150
      TabIndex        =   0
      Top             =   11
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
      Caption         =   "frmMarketDepthCfg.frx":10C4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMarketDepthCfg.frx":10EC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMarketDepthCfg.frx":110C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optCurrentSession 
         Height          =   220
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
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
         Caption         =   "frmMarketDepthCfg.frx":1128
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMarketDepthCfg.frx":116E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":118E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdSessionDate 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   540
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Enabled         =   0   'False
         AllowWeekends   =   0   'False
         MaxDate         =   42611
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
      Begin HexUniControls.ctlUniRadioXP optDate 
         Height          =   220
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "frmMarketDepthCfg.frx":11AA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMarketDepthCfg.frx":11E0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMarketDepthCfg.frx":1200
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmMarketDepthCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmMarketDepthCfg.frm
'' Description: Form to allow the user to change settings for the market depth form
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    frmTDGrid As frmTickDistribution
    bEditing As Boolean
End Type

Private m As mPrivate

Private Sub chkOrderBar_Click()
On Error Resume Next:

    If chkOrderBar.Value = vbUnchecked Then
        chkOrdBarOnRight.Enabled = False
    Else
        chkOrdBarOnRight.Enabled = True
    End If

End Sub

Private Sub chkSummaryBar_Click()
    EnableSummaryOptions
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConfigAcct_Click()
    fgQBarCols.Visible = False
    fgAcctCols.Visible = True
    fgOrderButtons.Visible = False
End Sub

Private Sub cmdConfigOrder_Click()
    fgQBarCols.Visible = False
    fgAcctCols.Visible = False
    fgOrderButtons.Visible = True
End Sub

Private Sub cmdConfigQuote_Click()
    fgQBarCols.Visible = True
    fgAcctCols.Visible = False
    fgOrderButtons.Visible = False
End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    'set font currently in use
    Me.Font.Name = m.frmTDGrid.GridFontName
    Me.Font.Size = m.frmTDGrid.GridFontSize
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.frmTDGrid.GridFontName = Me.Font.Name
        m.frmTDGrid.GridFontSize = Me.Font.Size
        m.frmTDGrid.RefreshGrid
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.cmdFont.Click", eGDRaiseError_Show
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim bDateChanged As Boolean
    Dim bColorChanged As Boolean
    Dim i&, strText$
    
    If VerifyQuantityPresets Then
        If optCurrentSession.Value <> m.frmTDGrid.IsCurrentSession Or _
            m.frmTDGrid.SessionDate <> gdSessionDate.Value Then
            bDateChanged = True
            Me.Hide
        End If
        m.frmTDGrid.SessionDate = gdSessionDate.Value
        m.frmTDGrid.IsCurrentSession = optCurrentSession.Value
        
        bColorChanged = ColorChanged()
        'market depth bid/ask colors
        m.frmTDGrid.FirstColor = gdPriceLevelColor(0).Color
        m.frmTDGrid.SecondColor = gdPriceLevelColor(1).Color
        m.frmTDGrid.ThirdColor = gdPriceLevelColor(2).Color
        m.frmTDGrid.FourthColor = gdPriceLevelColor(3).Color
        m.frmTDGrid.FifthColor = gdPriceLevelColor(4).Color
        m.frmTDGrid.OtherColor = gdPriceLevelColor(5).Color
        m.frmTDGrid.InactiveColor = gdPriceLevelColor(6).Color
        m.frmTDGrid.BidAskUpColor = gdBidAskUpColor.Color
        m.frmTDGrid.BidAskDownColor = gdBidAskDownColor.Color
        
        m.frmTDGrid.ShowSummaryBar = chkSummaryBar.Value
        If optSummaryBar(0).Value = True Then
            m.frmTDGrid.VerticalSummaryBar = 1
        Else
            m.frmTDGrid.VerticalSummaryBar = 0
        End If
        If optShowTriangles(0).Value = True Then
            m.frmTDGrid.DrawTriangles = 1
        Else
            m.frmTDGrid.DrawTriangles = 0
        End If
               
        'quote bar
        strText = ""
        m.frmTDGrid.ShowQuoteBar = chkQuoteBar.Value
        strText = ParseGridCtrl(fgQBarCols)
        m.frmTDGrid.QuoteBarHeader strText
            
        'account bar
        strText = ""
        m.frmTDGrid.ShowAccountBar = chkAccountBar.Value
        strText = ParseGridCtrl(fgAcctCols)
        m.frmTDGrid.AccountBarHeader strText
                   
        'order bar
        If chkOrderBar.Value = vbUnchecked Then
            If chkOrdBarOnRight.Value = vbUnchecked Then
                m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_NotShown
            Else
                m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_LastShownOnRight
            End If
        ElseIf chkOrdBarOnRight.Value = vbChecked Then
            m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_Right
        Else
            m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_BottomWide
        End If
        
        g.Broker.SetQuantityPresets m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, Int(ValOfText(txtQty1.Text)), Int(ValOfText(txtQty2.Text)), Int(ValOfText(txtQty3.Text))
        
        strText = ""
        m.frmTDGrid.OrderColumns = chkOrderBar.Value
        strText = ParseOrderButtonsGrid(fgOrderButtons)
        m.frmTDGrid.OrdBarCtrls = strText
        
        m.frmTDGrid.RefreshGrid bDateChanged, bColorChanged, ValOfText(txtSummaryBarHeight.Text)
        
        Unload Me
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.cmdOK.Click", eGDRaiseError_Show
    
End Sub

Private Sub fgQBarCols_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    m.bEditing = True
End Sub

Private Sub fgQBarCols_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub fgQBarCols_Click()
    ToggleShow
End Sub

Private Sub Form_Load()
    Me.Icon = Picture16(ToolbarIcon("ID_MarketDepth"), , True)
    
    g.Styler.StyleForm Me
    
    With fgAcctCols
        fgOrderButtons.Move .Left, .Top - 50
        fgQBarCols.Move .Left, .Top
    End With
    
    CenterTheForm Me
End Sub

Private Sub ToggleShow()
On Error Resume Next:

    If m.bEditing Then
        m.bEditing = False
        Exit Sub
    End If

    With fgQBarCols
        If .Col = 0 Then
            If .Cell(flexcpChecked, .Row, .Col) = 1 Then
                .Cell(flexcpChecked, .Row, .Col) = 2
            Else
                .Cell(flexcpChecked, .Row, .Col) = 1
            End If
        End If
    End With

End Sub

Public Sub ShowMe(frmCaller As Form)
On Error GoTo ErrSection:

    Dim lPreset1 As Long                ' First order quantity preset
    Dim lPreset2 As Long                ' Second order quantity preset
    Dim lPreset3 As Long                ' Third order quantity preset

    Set m.frmTDGrid = frmCaller
        
    'market depth colors & options
    gdPriceLevelColor(0).Color = m.frmTDGrid.FirstColor
    gdPriceLevelColor(1).Color = m.frmTDGrid.SecondColor
    gdPriceLevelColor(2).Color = m.frmTDGrid.ThirdColor
    gdPriceLevelColor(3).Color = m.frmTDGrid.FourthColor
    gdPriceLevelColor(4).Color = m.frmTDGrid.FifthColor
    gdPriceLevelColor(5).Color = m.frmTDGrid.OtherColor
    gdPriceLevelColor(6).Color = m.frmTDGrid.InactiveColor
    gdBidAskUpColor.Color = m.frmTDGrid.BidAskUpColor
    gdBidAskDownColor.Color = m.frmTDGrid.BidAskDownColor
            
    If m.frmTDGrid.IsCurrentSession Then
        optCurrentSession.Value = True
        optDate.Value = False
        gdSessionDate.Value = Date
    Else
        optCurrentSession.Value = False
        optDate.Value = True
        gdSessionDate.Value = m.frmTDGrid.SessionDate
    End If
    
    'quote bar options
    chkQuoteBar.Value = m.frmTDGrid.ShowQuoteBar
    'summary bar options
    chkSummaryBar.Value = m.frmTDGrid.ShowSummaryBar
    EnableSummaryOptions
    'bid/ask sizes options
    EnableBidAskSizeOptions
            
    'order bar options
    Select Case m.frmTDGrid.ShowOrderBar
        Case eGDOrderBarMode_LastShownOnRight
            chkOrderBar.Value = vbUnchecked
            chkOrdBarOnRight.Value = vbChecked
            chkOrdBarOnRight.Enabled = False
        Case eGDOrderBarMode_NotShown, eGDOrderBarMode_LastShownBottom
            chkOrderBar.Value = vbUnchecked
            chkOrdBarOnRight.Value = vbUnchecked
            chkOrdBarOnRight.Enabled = False
        Case eGDOrderBarMode_BottomWide, eGDOrderBarMode_BottomNarrow, eGDOrderBarMode_BottomContinuous
            chkOrderBar.Value = vbChecked
            chkOrdBarOnRight.Value = vbUnchecked
            chkOrdBarOnRight.Enabled = True
        Case Else
            chkOrderBar.Value = vbChecked
            chkOrdBarOnRight.Value = vbChecked
            chkOrdBarOnRight.Enabled = True
    End Select
    
    chkAccountBar.Value = m.frmTDGrid.ShowAccountBar
    
    g.Broker.GetQuantityPresets m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset1, lPreset2, lPreset3
    txtQty1.Text = Str(lPreset1)
    txtQty2.Text = Str(lPreset2)
    txtQty3.Text = Str(lPreset3)
        
    InitQBarGrid fgQBarCols, m.frmTDGrid.QBarColArray
    InitAccountGrid fgAcctCols, m.frmTDGrid.ABarColArray, m.frmTDGrid.SecType
    InitOrderButtonsGrid fgOrderButtons, m.frmTDGrid.OrdBarCtrls
                    
    ShowForm Me, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.ShowMe", eGDRaiseError_Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m.frmTDGrid = Nothing
End Sub

Private Sub optCurrentSession_Click()
    gdSessionDate.Enabled = False
End Sub

Private Sub optDate_Click()
    gdSessionDate.Enabled = True
End Sub

Private Function ColorChanged() As Boolean
On Error GoTo ErrSection:

    Dim bChanged As Boolean
    
    If m.frmTDGrid.FirstColor <> gdPriceLevelColor(0).Color Or _
      m.frmTDGrid.SecondColor <> gdPriceLevelColor(1).Color Or _
      m.frmTDGrid.ThirdColor <> gdPriceLevelColor(2).Color Or _
      m.frmTDGrid.FourthColor <> gdPriceLevelColor(3).Color Or _
      m.frmTDGrid.FifthColor <> gdPriceLevelColor(4).Color Or _
      m.frmTDGrid.OtherColor <> gdPriceLevelColor(5).Color Or _
      m.frmTDGrid.InactiveColor <> gdPriceLevelColor(6).Color Or _
      m.frmTDGrid.BidAskUpColor <> gdBidAskUpColor.Color Or _
      m.frmTDGrid.BidAskDownColor <> gdBidAskDownColor.Color Or _
      m.frmTDGrid.ShowSummaryBar <> chkSummaryBar.Value Then

            bChanged = True

    End If
    
    ColorChanged = bChanged
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmMarketDepthCfg.ColorChanged", eGDRaiseError_Raise

End Function

Private Sub optShowTriangles_Click(Index As Integer)
On Error Resume Next:

    If Index = 0 Then
        lblBidAskUpColor = "Triangle up color"
        lblBidAskDownColor = "Triangle down color"
    Else
        lblBidAskUpColor = "Bid histogram color"
        lblBidAskDownColor = "Ask histogram color"
    End If

End Sub

Private Sub txtQty1_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.txtQty1_GotFocus"
    
End Sub

Private Sub txtQty2_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.txtQty2_GotFocus"
    
End Sub

Private Sub txtQty3_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty3

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.txtQty3_GotFocus"
    
End Sub

Private Sub txtSummaryBarHeight_LostFocus()
On Error Resume Next:
    
    If ValOfText(txtSummaryBarHeight.Text) < 200 Then
        txtSummaryBarHeight.Text = Str(m.frmTDGrid.SummaryBarHeight)
    End If
    
End Sub

Private Sub EnableSummaryOptions()
On Error GoTo ErrSection:
    
    txtSummaryBarHeight.Text = Str(m.frmTDGrid.SummaryBarHeight)
    If m.frmTDGrid.VerticalSummaryBar = 1 Then
        optSummaryBar(0).Value = True
        optSummaryBar(1).Value = False
    Else
        optSummaryBar(0).Value = False
        optSummaryBar(1).Value = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.EnableSummaryOptions", eGDRaiseError_Raise

End Sub

Private Sub EnableBidAskSizeOptions()
On Error GoTo ErrSection:

    Dim bEnable As Boolean

    If m.frmTDGrid.DisplayStyle = 1 Then
        bEnable = True
    End If

    If m.frmTDGrid.DrawTriangles = 1 Then
        optShowTriangles(0).Value = True
        optShowTriangles(1).Value = False
        lblBidAskUpColor = "Triangle up color"
        lblBidAskDownColor = "Triangle down color"
    Else
        optShowTriangles(0).Value = False
        optShowTriangles(1).Value = True
        lblBidAskUpColor = "Bid histogram color"
        lblBidAskDownColor = "Ask histogram color"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMarketDepthCfg.EnableBidAskSizeOptions", eGDRaiseError_Raise

End Sub

Private Function VerifyQuantityPresets() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lPreset As Long                 ' Preset value
    
    bReturn = True
    
    lPreset = Int(Val(txtQty1.Text))
    If g.Broker.ValidQuantity(m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset) = False Then
        MoveFocus txtQty1
        InfBox "The first quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
        bReturn = False
    End If
    
    If bReturn = True Then
        lPreset = Int(Val(txtQty2.Text))
        If g.Broker.ValidQuantity(m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset) = False Then
            MoveFocus txtQty2
            InfBox "The second quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
            bReturn = False
        End If
    End If
    
    If bReturn = True Then
        lPreset = Int(Val(txtQty3.Text))
        If g.Broker.ValidQuantity(m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset) = False Then
            MoveFocus txtQty3
            InfBox "The third quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
            bReturn = False
        End If
    End If
    
    VerifyQuantityPresets = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmMarketDepthCfg.VerifyQuantityPresets"
    
End Function

