VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMkBreadth 
   Caption         =   "Market Breadth"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraSummaryView 
      Height          =   2475
      Left            =   3180
      TabIndex        =   30
      Top             =   3180
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
      Caption         =   "frmMkBreadth.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMkBreadth.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMkBreadth.frx":0040
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgSummary 
         Height          =   1995
         Left            =   480
         TabIndex        =   31
         Top             =   300
         Width           =   3015
         _cx             =   5318
         _cy             =   3519
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
   Begin HexUniControls.ctlUniFrameWL fraExchangeView 
      Height          =   2775
      Left            =   1620
      TabIndex        =   14
      Top             =   60
      Width           =   6075
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
      Caption         =   "frmMkBreadth.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMkBreadth.frx":007C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMkBreadth.frx":009C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraBRight 
         Height          =   975
         Left            =   3840
         TabIndex        =   25
         Top             =   180
         Width           =   1935
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
         Caption         =   "frmMkBreadth.frx":00B8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMkBreadth.frx":00D8
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":00F8
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgBRight 
            Height          =   615
            Left            =   120
            TabIndex        =   26
            Top             =   180
            Width           =   1575
            _cx             =   2778
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
      End
      Begin HexUniControls.ctlUniFrameWL fraTLeft 
         Height          =   1035
         Left            =   120
         TabIndex        =   23
         Top             =   180
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
         Caption         =   "frmMkBreadth.frx":0114
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMkBreadth.frx":0134
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0154
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgTLeft 
            Height          =   720
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   1365
            _cx             =   2408
            _cy             =   1270
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
            Rows            =   4
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
            AutoResize      =   0   'False
            AutoSizeMode    =   1
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
      Begin HexUniControls.ctlUniFrameWL fraTRight 
         Height          =   975
         Left            =   1080
         TabIndex        =   21
         Top             =   240
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
         Caption         =   "frmMkBreadth.frx":0170
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMkBreadth.frx":0190
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":01B0
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgTRight 
            Height          =   600
            Left            =   120
            TabIndex        =   22
            Top             =   180
            Width           =   1425
            _cx             =   2514
            _cy             =   1058
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
      Begin HexUniControls.ctlUniFrameWL fraBLeft 
         Height          =   855
         Left            =   2940
         TabIndex        =   19
         Top             =   180
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
         Caption         =   "frmMkBreadth.frx":01CC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMkBreadth.frx":01EC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":020C
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgBLeft 
            Height          =   615
            Left            =   180
            TabIndex        =   20
            Top             =   180
            Width           =   1395
            _cx             =   2461
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
      End
      Begin vsOcx6LibCtl.vsIndexTab vsTabLists 
         Height          =   1335
         Left            =   300
         TabIndex        =   15
         Top             =   1320
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2355
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
         Caption         =   "Tab0|Tab1|Tab2|Tab3|Tab4|Tab5"
         Align           =   0
         Appearance      =   1
         CurrTab         =   5
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
         Begin VSFlex7LCtl.VSFlexGrid fgTab5 
            Height          =   960
            Left            =   45
            TabIndex        =   29
            Top             =   330
            Width           =   5325
            _cx             =   9393
            _cy             =   1693
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
            AllowUserResizing=   1
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
         Begin VSFlex7LCtl.VSFlexGrid fgTab4 
            Height          =   960
            Left            =   -5970
            TabIndex        =   28
            Top             =   330
            Width           =   5325
            _cx             =   9393
            _cy             =   1693
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
            AllowUserResizing=   1
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
         Begin VSFlex7LCtl.VSFlexGrid fgTab3 
            Height          =   960
            Left            =   -6270
            TabIndex        =   27
            Top             =   330
            Width           =   5325
            _cx             =   9393
            _cy             =   1693
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
            AllowUserResizing=   1
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
         Begin VSFlex7LCtl.VSFlexGrid fgTab2 
            Height          =   960
            Left            =   -6570
            TabIndex        =   16
            Top             =   330
            Width           =   5325
            _cx             =   9393
            _cy             =   1693
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
            AllowUserResizing=   1
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
         Begin VSFlex7LCtl.VSFlexGrid fgTab1 
            Height          =   960
            Left            =   -6870
            TabIndex        =   17
            Top             =   330
            Width           =   5325
            _cx             =   9393
            _cy             =   1693
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   11
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
         Begin VSFlex7LCtl.VSFlexGrid fgTab0 
            Height          =   960
            Left            =   -7170
            TabIndex        =   18
            Top             =   330
            Width           =   5325
            _cx             =   9393
            _cy             =   1693
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
            AllowUserResizing=   1
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
   End
   Begin HexUniControls.ctlUniFrameWL fraExchange 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   60
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
      Caption         =   "frmMkBreadth.frx":0228
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmMkBreadth.frx":0258
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMkBreadth.frx":0278
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   4050
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":0294
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":02BA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":02DA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   3660
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":02F6
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":031C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":033C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSummary 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4440
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
         Caption         =   "frmMkBreadth.frx":0358
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":0392
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":03B2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   3270
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":03CE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":03F4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0414
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":0430
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":0456
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0476
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2490
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":0492
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":04B8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":04D8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   2100
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":04F4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":051A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":053A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   5340
         Width           =   1095
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
         Caption         =   "frmMkBreadth.frx":0556
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":0586
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":05A6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRefresh 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   4920
         Width           =   1095
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
         Caption         =   "frmMkBreadth.frx":05C2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":05F0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0610
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":062C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmMkBreadth.frx":0652
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0672
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   930
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":068E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":06B4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":06D4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":06F0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":0716
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0736
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExchange 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   1710
         Width           =   1035
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
         Caption         =   "frmMkBreadth.frx":0752
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmMkBreadth.frx":0778
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmMkBreadth.frx":0798
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmMkBreadth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kColOneWidth = 1000
Private Const kExchButtonMax = 10
Private Const kTabsMax = 6
Private Const kDataFile = "MkBreadth.dat"
Private Const kTabsCaption = "Price % Gainers|Price % Losers|Volume Leaders"

Private Type mPrivate
    oData As New cMkBreadth
    strCurrExchange As String
    bExchangeChanged As Boolean
    bLoaded As Boolean
End Type
Private m As mPrivate

Public Sub ShowMe()
    
    CenterTheForm Me
    ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR

End Sub

Private Sub LoadSummary()
    m.oData.DataToGrid fgSummary, eSummary
End Sub

Private Sub LoadTabs()

    Dim i&, k&
    Dim aData As cGdArray
    Dim strCaption$
       
    Set aData = m.oData.TabTittles
    k = aData.Size - 1
    
    If k < 0 Then
        vsTabLists.Caption = kTabsCaption       'use generic default
        fgTab0.Rows = 0
        fgTab1.Rows = 0
        fgTab2.Rows = 0
        Exit Sub
    End If
    
    For i = 0 To k
        If Len(aData(i)) > 0 Then
            If Len(strCaption) = 0 Then
                strCaption = aData(i)
            Else
                strCaption = strCaption & "|" & aData(i)
            End If
        End If
    Next
    
    vsTabLists.Caption = strCaption
    m.oData.DataToGrid fgTab0, eTab, 0
    
    Form_Resize

End Sub

Private Sub LoadGrids()
    
    m.oData.DataToGrid fgTLeft, eTopLeft
    
    With fgTLeft
        If .Rows > 0 Then
            .BackColorAlternate = ALT_GRID_ROW_COLOR
            .ColWidth(1) = kColOneWidth
        End If
    End With
    
    m.oData.DataToGrid fgTRight, eTopRight
    With fgTRight
        If .Rows > 0 Then
            .BackColorAlternate = ALT_GRID_ROW_COLOR
            .ColWidth(1) = kColOneWidth
        End If
    End With
    
    m.oData.DataToGrid fgBLeft, eBottomLeft
    With fgBLeft
        If .Rows > 0 Then
            .BackColorAlternate = ALT_GRID_ROW_COLOR
            .ColWidth(1) = kColOneWidth
        End If
    End With
    
    m.oData.DataToGrid fgBRight, eBottomRight
    With fgBRight
        If .Rows > 0 Then
            .BackColorAlternate = ALT_GRID_ROW_COLOR
            .ColWidth(1) = kColOneWidth
        End If
    End With
    
    Form_Resize
    
End Sub

Private Sub cmdSettings_Click()
    
    frmMkBreadthCfg.ShowMe m.oData
    
    If fraSummaryView.Visible Then
        LoadSummary
        SizeSummaryGrid
    End If
    
End Sub

Private Sub Form_Load()

    g.Styler.StyleForm Me
    
    m.oData.DataFileName = g.strAppPath & "\" & kDataFile
    m.bExchangeChanged = True
    
    vsTabLists.CurrTab = 0
    InitExchangeControls
    If m.oData.ExchangeCount = 0 Then
        MsgBox "Data not found."
        Exit Sub
    End If
    
    m.oData.ExchangeData = optExchange(0).Caption
    LoadTabs
    LoadGrids
    
    'RH commented out fraTLeft.BorderStyle = 0
    'RH commented out fraBLeft.BorderStyle = 0
    'RH commented out fraTRight.BorderStyle = 0
    'RH commented out fraBRight.BorderStyle = 0
    'RH commented out fraExchangeView.BorderStyle = 0
    fraExchangeView.Caption = ""
    
    'RH commented out fraSummaryView.BorderStyle = 0
    fraSummaryView.Caption = ""
    fraSummaryView.Visible = False
    
    m.bLoaded = True
    
End Sub

Private Sub Form_Resize()

    SizeFrames
    
    If fraExchangeView.Visible Then
        SizeTopFrames
        SizeBottomFrames
        SizeTabControls
    ElseIf fraSummaryView.Visible Then
        SizeSummaryGrid
    End If

        
End Sub

Private Sub InitExchangeControls()

    Dim i&, j&, nCount&
    Dim aExchanges As cGdArray
    
    Set aExchanges = m.oData.ExchangeNames
    nCount = aExchanges.Size

    If nCount = 0 Then
        Exit Sub
    ElseIf nCount > kExchButtonMax Then
        j = kExchButtonMax - 1
    Else
        j = nCount - 1
    End If

    For i = 0 To j
        optExchange(i).Caption = aExchanges(i)
        optExchange(i).Visible = True
    Next
    
    For i = j + 1 To kExchButtonMax - 1
        optExchange(i).Visible = False
    Next
    
    If j < kExchButtonMax - 1 Then
        optSummary.Move optSummary.Left, optExchange(j + 1).Top
    End If
    
    optExchange(0).Value = True
    m.strCurrExchange = aExchanges(0)
    
End Sub

Private Sub SizeFrames()
On Error Resume Next

    Dim nLeft&, nTop&, nHeight&, nWidth&
            
    With fraExchange
        nHeight = Me.ScaleHeight - 50
        .Move 50, 50, .Width, nHeight
        nLeft = .Left + .Width + 50
        nWidth = Me.ScaleWidth - .Width - 100
    End With
    
    fraExchangeView.Move nLeft, 50, nWidth, nHeight
    fraSummaryView.Move nLeft, 50, nWidth, nHeight
    
    nHeight = cmdSettings.Height + cmdRefresh.Height
    nTop = Me.ScaleHeight - nHeight - 200
    With cmdRefresh
        .Move .Left, nTop
        nTop = .Top + .Height + 50
        cmdSettings.Move .Left, nTop
    End With
        
End Sub

Private Sub SizeTopFrames()
On Error Resume Next

    Dim nWidth&, nHeight&, nLeft&

    With fgTLeft
        If .Rows > 0 Then
            fraTLeft.Visible = True
            nWidth = (Me.ScaleWidth - fraExchange.Width - 250) / 2
            nHeight = .RowHeight(0) * .Rows
            fraTLeft.Move 0, 150, nWidth, nHeight
            .Move 0, 0, fraTLeft.Width, fraTLeft.Height + .RowHeight(0)
            .ColWidth(0) = .ClientWidth - kColOneWidth
            nLeft = .Left + .Width + 100
        Else
            fraTLeft.Visible = False
        End If
    End With
    
    With fgTRight
        If .Rows > 0 Then
            fraTRight.Visible = True
            nHeight = .RowHeight(0) * .Rows
            fraTRight.Move nLeft, 150, nWidth, nHeight
            .Move 0, 0, fraTRight.Width, fraTRight.Height + .RowHeight(0)
            .ColWidth(0) = .ClientWidth - kColOneWidth
        Else
            fraTRight.Visible = False
        End If
    End With

End Sub

Private Sub SizeBottomFrames()
On Error Resume Next

    Dim nWidth&, nHeight&, nTop&
        
    nWidth = fraTLeft.Width
    nTop = fraTLeft.Top + fraTLeft.Height + 150
        
    With fgBLeft
        If .Rows > 0 Then
            fraBLeft.Visible = True
            nHeight = .RowHeight(0) * .Rows
            fraBLeft.Move fraTLeft.Left, nTop, nWidth, nHeight
            .Move 0, 0, fraBLeft.Width, fraBLeft.Height + .RowHeight(0)
            .ColWidth(0) = .ClientWidth - kColOneWidth
        Else
            fraBLeft.Visible = False
        End If
    End With
    
    With fgBRight
        If .Rows > 0 Then
            fraBRight.Visible = True
            nHeight = .RowHeight(0) * .Rows
            fraBRight.Move fraTRight.Left, nTop, nWidth, nHeight
            .Move 0, 0, fraBRight.Width, fraBRight.Height + .RowHeight(0)
            .ColWidth(0) = .ClientWidth - kColOneWidth
        Else
            fraBRight.Visible = False
        End If
    End With
        
End Sub

Private Sub SizeTabControls()
On Error Resume Next

    Dim nWidth&, nHeight&, nTop&, nLeft&, i&
    Dim fgToMove As VSFlexGrid

    If Not m.bLoaded Then Exit Sub
    
    Select Case vsTabLists.CurrTab
        Case 0:
            Set fgToMove = fgTab0
        Case 1:
            Set fgToMove = fgTab1
        Case 2:
            Set fgToMove = fgTab2
        Case 3:
            Set fgToMove = fgTab3
        Case 4:
        Case 5:
    End Select
    
    If fgToMove.Rows = 0 Then Exit Sub
    
    nHeight = fgToMove.RowHeight(0) * 12 + 50
    nWidth = fraExchangeView.Width - 50
    nLeft = fraTLeft.Left
    nTop = Me.ScaleHeight - nHeight - 100
   
    vsTabLists.Move nLeft, nTop, nWidth, nHeight
    
    With fgToMove
        nWidth = nWidth - kColOneWidth - 150
        nWidth = nWidth / (.Cols - 1)
        .ColWidth(0) = kColOneWidth
        For i = 1 To .Cols - 1
            .ColWidth(i) = nWidth
        Next
    End With
           
End Sub

Private Sub SizeSummaryGrid()

    Dim nWidth&, i&
    
    With fgSummary
        .Move 0, fraExchange.Top + 80, fraSummaryView.Width, fraSummaryView.Height
        If .Cols > 2 Then
            .ColWidth(0) = kColOneWidth * 1.5
            nWidth = (.ClientWidth - kColOneWidth * 1.5) / (.Cols - 1)
            For i = 1 To .Cols - 1
                .ColWidth(i) = nWidth
            Next
        End If
    End With

End Sub

Private Sub optExchange_Click(Index As Integer)

    fraExchangeView.Visible = True
    fraSummaryView.Visible = False
    
    If optExchange(Index).Caption <> m.strCurrExchange Then
        m.oData.ExchangeData = optExchange(Index).Caption
        m.strCurrExchange = optExchange(Index).Caption
        LoadGrids
        LoadTabs
    End If

End Sub

Private Sub optSummary_Click()

    fraExchangeView.Visible = False
    fraSummaryView.Visible = True
    LoadSummary
    SizeSummaryGrid

End Sub

Private Sub vsTabLists_Click()
    
    If Not m.bLoaded Then Exit Sub
    
    Select Case vsTabLists.CurrTab
        Case 0:
            m.oData.DataToGrid fgTab0, eTab, 0
        Case 1:
            m.oData.DataToGrid fgTab1, eTab, 1
        Case 2:
            m.oData.DataToGrid fgTab2, eTab, 2
        Case 3:
            m.oData.DataToGrid fgTab3, eTab, 3
        Case 4:
            m.oData.DataToGrid fgTab4, eTab, 4
        Case 5:
            m.oData.DataToGrid fgTab5, eTab, 5
    End Select
    
    SizeTabControls
    
End Sub

