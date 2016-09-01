VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmBrokerScreen 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   10080
      TabIndex        =   76
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdParkOrder 
         Caption         =   "&Park Order"
         Height          =   495
         Left            =   0
         TabIndex        =   82
         Top             =   2820
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refres&h"
         Height          =   495
         Left            =   0
         TabIndex        =   83
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubmitOrder 
         Caption         =   "Su&bmit Order"
         Height          =   495
         Left            =   0
         TabIndex        =   81
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Ca&ncel Order"
         Height          =   495
         Left            =   0
         TabIndex        =   80
         Top             =   1740
         Width           =   1215
      End
      Begin VB.CommandButton cmdModifyOrder 
         Caption         =   "&Modify Order"
         Height          =   495
         Left            =   0
         TabIndex        =   79
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Pr&int"
         Height          =   495
         Left            =   0
         TabIndex        =   78
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Clos&e"
         Height          =   495
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   1215
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab tabInfo 
      Height          =   3675
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   6482
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
      Caption         =   "Wor&king Orders|&Filled Orders|Parke&d Orders|Acco&unt Status|S&tatement|&Open Positions|Ne&w Order"
      Align           =   0
      Appearance      =   1
      CurrTab         =   6
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
      Begin VB.Frame fraParkedOrders 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3300
         Left            =   -10350
         TabIndex        =   17
         Top             =   330
         Width           =   9705
         Begin VSFlex7LCtl.VSFlexGrid fgPositions 
            Height          =   2415
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   8535
            _cx             =   15055
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
      Begin VB.Frame fraStatement 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3300
         Left            =   -10650
         TabIndex        =   15
         Top             =   330
         Width           =   9705
         Begin VSFlex7LCtl.VSFlexGrid fgStatement 
            Height          =   2415
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   8535
            _cx             =   15055
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
      Begin VB.Frame fraAccountStatus 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3300
         Left            =   -10950
         TabIndex        =   13
         Top             =   330
         Width           =   9705
         Begin VSFlex7LCtl.VSFlexGrid fgAccountStatus 
            Height          =   2415
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   8535
            _cx             =   15055
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
      Begin VB.Frame fraPositions 
         BorderStyle     =   0  'None
         Height          =   3300
         Left            =   -11250
         TabIndex        =   11
         Top             =   330
         Width           =   9705
         Begin VSFlex7LCtl.VSFlexGrid fgParkedOrders 
            Height          =   2415
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   8535
            _cx             =   15055
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
      Begin VB.Frame fraFilledOrders 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3300
         Left            =   -11550
         TabIndex        =   9
         Top             =   330
         Width           =   9705
         Begin VSFlex7LCtl.VSFlexGrid fgFilledOrders 
            Height          =   2415
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   8535
            _cx             =   15055
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
      Begin VB.Frame fraWorkingOrders 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3300
         Left            =   -11850
         TabIndex        =   7
         Top             =   330
         Width           =   9705
         Begin VSFlex7LCtl.VSFlexGrid fgWorkingOrders 
            Height          =   2415
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   8535
            _cx             =   15055
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
      Begin VB.Frame fraNewOrder 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3300
         Left            =   45
         TabIndex        =   19
         Top             =   330
         Width           =   9705
         Begin VB.TextBox txtOrderID 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   75
            Width           =   2175
         End
         Begin VB.CommandButton cmdResetOrder 
            Caption         =   "&Reset Order"
            Height          =   315
            Left            =   7620
            TabIndex        =   73
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdVerifyOrder 
            Caption         =   "&Verify Order"
            Height          =   315
            Left            =   6300
            TabIndex        =   72
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtCommandLine 
            Height          =   315
            Left            =   1320
            TabIndex        =   75
            Top             =   2400
            Width           =   8115
         End
         Begin VB.ComboBox cboSession 
            Height          =   315
            Left            =   4860
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   1920
            Width           =   855
         End
         Begin VB.ComboBox cboGoodThru 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   1920
            Width           =   855
         End
         Begin VB.ComboBox cboBuySell 
            Height          =   315
            Index           =   2
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1500
            Width           =   735
         End
         Begin VB.ComboBox cboCommodity 
            Height          =   315
            Index           =   2
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1500
            Width           =   1695
         End
         Begin VB.TextBox txtQuantity 
            Height          =   315
            Index           =   2
            Left            =   900
            TabIndex        =   57
            Top             =   1500
            Width           =   975
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Index           =   2
            Left            =   3660
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1500
            Width           =   615
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Index           =   2
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   1500
            Width           =   795
         End
         Begin VB.TextBox txtStrike 
            Height          =   315
            Index           =   2
            Left            =   5160
            TabIndex        =   61
            Top             =   1500
            Width           =   735
         End
         Begin VB.ComboBox cboPutCall 
            Height          =   315
            Index           =   2
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox txtPrice 
            Height          =   315
            Index           =   2
            Left            =   6480
            TabIndex        =   63
            Top             =   1500
            Width           =   735
         End
         Begin VB.ComboBox cboOpenClose 
            Height          =   315
            Index           =   2
            Left            =   7260
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1500
            Width           =   495
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Index           =   2
            Left            =   7860
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1500
            Width           =   795
         End
         Begin VB.TextBox txtLimit 
            Height          =   315
            Index           =   2
            Left            =   8700
            TabIndex        =   66
            Top             =   1500
            Width           =   735
         End
         Begin VB.ComboBox cboBuySell 
            Height          =   315
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox cboCommodity 
            Height          =   315
            Index           =   1
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtQuantity 
            Height          =   315
            Index           =   1
            Left            =   900
            TabIndex        =   46
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Index           =   1
            Left            =   3660
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Index           =   1
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1080
            Width           =   795
         End
         Begin VB.TextBox txtStrike 
            Height          =   315
            Index           =   1
            Left            =   5160
            TabIndex        =   50
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox cboPutCall 
            Height          =   315
            Index           =   1
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtPrice 
            Height          =   315
            Index           =   1
            Left            =   6480
            TabIndex        =   52
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox cboOpenClose 
            Height          =   315
            Index           =   1
            Left            =   7260
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1080
            Width           =   495
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Index           =   1
            Left            =   7860
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1080
            Width           =   795
         End
         Begin VB.TextBox txtLimit 
            Height          =   315
            Index           =   1
            Left            =   8700
            TabIndex        =   55
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtLimit 
            Height          =   315
            Index           =   0
            Left            =   8700
            TabIndex        =   44
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Index           =   0
            Left            =   7860
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   660
            Width           =   795
         End
         Begin VB.ComboBox cboOpenClose 
            Height          =   315
            Index           =   0
            Left            =   7260
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtPrice 
            Height          =   315
            Index           =   0
            Left            =   6480
            TabIndex        =   41
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cboPutCall 
            Height          =   315
            Index           =   0
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtStrike 
            Height          =   315
            Index           =   0
            Left            =   5160
            TabIndex        =   39
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Index           =   0
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   660
            Width           =   795
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Index           =   0
            Left            =   3660
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtQuantity 
            Height          =   315
            Index           =   0
            Left            =   900
            TabIndex        =   35
            Top             =   660
            Width           =   975
         End
         Begin VB.ComboBox cboCommodity 
            Height          =   315
            Index           =   0
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   660
            Width           =   1695
         End
         Begin VB.ComboBox cboBuySell 
            Height          =   315
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   660
            Width           =   735
         End
         Begin VB.CheckBox chkSpread 
            Caption         =   "&Spread"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   915
         End
         Begin gdOCX.gdSelectDate gdThruDate 
            Height          =   315
            Left            =   1980
            TabIndex        =   69
            Top             =   1920
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            ShowDayOfWeek   =   0   'False
         End
         Begin VB.Label lblOrderID 
            Caption         =   "Order ID:"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblError 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   2880
            Width           =   9315
         End
         Begin VB.Label lblCommand 
            Caption         =   "Command Line:"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   2460
            Width           =   1155
         End
         Begin VB.Label lblSession 
            Caption         =   "Session:"
            Height          =   195
            Left            =   4140
            TabIndex        =   70
            Top             =   1980
            Width           =   675
         End
         Begin VB.Label lblGoodThru 
            Caption         =   "Good Thru:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1980
            Width           =   855
         End
         Begin VB.Label lblLimit 
            Caption         =   "Limit:"
            Height          =   195
            Left            =   8700
            TabIndex        =   33
            Top             =   420
            Width           =   735
         End
         Begin VB.Label lblType 
            Caption         =   "Type:"
            Height          =   195
            Left            =   7860
            TabIndex        =   32
            Top             =   420
            Width           =   735
         End
         Begin VB.Label lblOpenClose 
            Caption         =   "O/C:"
            Height          =   195
            Left            =   7260
            TabIndex        =   31
            Top             =   420
            Width           =   435
         End
         Begin VB.Label lblPrice 
            Caption         =   "Price:"
            Height          =   195
            Left            =   6480
            TabIndex        =   30
            Top             =   420
            Width           =   675
         End
         Begin VB.Label lblPutCall 
            Caption         =   "P/C:"
            Height          =   195
            Left            =   5940
            TabIndex        =   29
            Top             =   420
            Width           =   435
         End
         Begin VB.Label lblStrike 
            Caption         =   "Strike:"
            Height          =   195
            Left            =   5160
            TabIndex        =   28
            Top             =   420
            Width           =   675
         End
         Begin VB.Label lblYear 
            Caption         =   "Year:"
            Height          =   195
            Left            =   4320
            TabIndex        =   27
            Top             =   420
            Width           =   735
         End
         Begin VB.Label lblMonth 
            Caption         =   "Month:"
            Height          =   195
            Left            =   3660
            TabIndex        =   26
            Top             =   420
            Width           =   555
         End
         Begin VB.Label lblCommodity 
            Caption         =   "Symbol:"
            Height          =   195
            Left            =   1920
            TabIndex        =   25
            Top             =   420
            Width           =   855
         End
         Begin VB.Label lblQuantity 
            Caption         =   "Quantity:"
            Height          =   195
            Left            =   900
            TabIndex        =   24
            Top             =   420
            Width           =   855
         End
         Begin VB.Label lblBuySell 
            Caption         =   "B/S:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   420
            Width           =   675
         End
      End
   End
   Begin VB.Frame fraAccountSelection 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7395
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   315
         Left            =   4380
         TabIndex        =   3
         Top             =   0
         Width           =   555
      End
      Begin VB.OptionButton optSpecificAccount 
         Caption         =   "&Account:"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optAllAccounts 
         Caption         =   "All A&ccounts"
         Height          =   195
         Left            =   6120
         TabIndex        =   5
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccountLookup 
         Caption         =   "&Lookup"
         Height          =   315
         Left            =   4980
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "Orders"
      Begin VB.Menu mnuModifyOrder 
         Caption         =   "Modify Order"
      End
      Begin VB.Menu mnuCancelOrder 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuSubmitOrder 
         Caption         =   "Submit Order"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshAccount 
         Caption         =   "Refresh Account"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "frmBrokerScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBrokerScreen.frm
'' Description: Screen to allow for broker capabilities
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 04/25/2012   DAJ         Removed broker from account lookup call
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDWorkingCols
    eGDWorkingCol_Account = 0
    eGDWorkingCol_OrderNumber
    eGDWorkingCol_ReferenceID
    eGDWorkingCol_OrderDate
    eGDWorkingCol_BuySell
    eGDWorkingCol_Quantity
    eGDWorkingCol_Symbol
    eGDWorkingCol_Strike
    eGDWorkingCol_Price
    eGDWorkingCol_OC
    eGDWorkingCol_OrderType
    eGDWorkingCol_Limit
    eGDWorkingCol_Tic
    eGDWorkingCol_GoodThru
    eGDWorkingCol_Session
    eGDWorkingCol_CreditControl
    eGDWorkingCol_ExchangeMessage
    eGDWorkingCol_SpecialInstruction
    eGDWorkingCol_EnteredBy
    eGDWorkingCol_Salesman
    eGDWorkingCol_Exchange
    eGDWorkingCol_BaseSymbol
    eGDWorkingCol_Contract
    eGDWorkingCol_NumCols
End Enum

Private Enum eGDFilledCols
    eGDFilledCol_Account = 0
    eGDFilledCol_OrderNumber
    eGDFilledCol_TradeDate
    eGDFilledCol_Position
    eGDFilledCol_Quantity
    eGDFilledCol_Symbol
    eGDFilledCol_Strike
    eGDFilledCol_TradePrice
    eGDFilledCol_Tic
    eGDFilledCol_FuturesProfit
    eGDFilledCol_OptionsValue
    eGDFilledCol_Currency
    eGDFilledCol_OrderTime
    eGDFilledCol_FillTime
    eGDFilledCol_NumCols
End Enum

Private Enum eGDPositionCols
    eGDPositionCol_Account = 0
    eGDPositionCol_TradeDate
    eGDPositionCol_Position
    eGDPositionCol_Quantity
    eGDPositionCol_Symbol
    eGDPositionCol_Strike
    eGDPositionCol_TradePrice
    eGDPositionCol_Tic
    eGDPositionCol_FuturesProfit
    eGDPositionCol_OptionValue
    eGDPositionCol_Currency
    eGDPositionCol_NumCols
End Enum

Private Enum eGDTabs
    eGDTab_WorkingOrders = 0
    eGDTab_FilledOrders
    eGDTab_ParkedOrders
    eGDTab_AccountStatus
    eGDTab_Statement
    eGDTab_OpenPositions
    eGDTab_NewOrder
End Enum

Private Type mPrivate
    nBroker As eTT_AccountType          ' Broker type
    bParsing As Boolean                 ' Parsing the command line
    bUpdating As Boolean                ' Updating the command line
    
    astrAccounts As cGdArray            ' Array of account information
    astrAcctStatus As cGdArray          ' Array of account status information
    astrOrders As cGdArray              ' Table of order information
    astrPositions As cGdArray           ' Array of open position information
    astrSecurities As cGdArray          ' Array of security information
    astrStatement As cGdArray           ' Table of account statement information
    
    bAccountsLoaded As Boolean          ' Are we done getting the accounts?
    bAccountStatusLoaded As Boolean     ' Are we done getting the account status?
    bOrdersLoaded As Boolean            ' Are we done getting the orders?
    bPositionsLoaded As Boolean         ' Are we done getting the positions?
    bSecuritiesLoaded As Boolean        ' Are we done getting the securites list?
    bStatementLoaded As Boolean         ' Are we done getting the statement?
    
    bLoadingOrders As Boolean           ' Are we currently loading orders?
End Type
Private m As mPrivate

Private Function WorkingOrdersCol(ByVal nCol As eGDWorkingCols) As Long
    WorkingOrdersCol = nCol
End Function
Private Function FilledOrdersCol(ByVal nCol As eGDFilledCols) As Long
    FilledOrdersCol = nCol
End Function
Private Function PositionsCol(ByVal nCol As eGDPositionCols) As Long
    PositionsCol = nCol
End Function
Private Function Tabs(ByVal nTab As eGDTabs) As Long
    Tabs = nTab
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Broker Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    m.nBroker = nBroker
    
    InitWorkingOrdersGrid
    InitFilledOrdersGrid
    InitPositionsGrid
    InitAccountStatusGrid
    InitStatementGrid
    InitParkedOrdersGrid
    
    ClearSide 0
    ClearSide 1
    ClearSide 2
    
    tabInfo.CurrTab = Tabs(eGDTab_WorkingOrders)
    chkSpread.Value = vbUnchecked
    cboGoodThru.Text = "Today"
    cboSession.Text = "Day"
    
    EnableControls
    FilterGrids
    
    ShowForm Me, eForm_Nonmodal, frmMain

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Show the print preview form for the currently selected tab
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    PrintMe = frmPrintPreview.ShowMe("CNV BrokerScreen", frmBrokerScreen, , , , , , True)

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmBrokerScreen.PrintMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Generate the print preview for the currently selected tab
'' Inputs:      Arguments
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim fg As VSFlexGrid                ' Flex grid object

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        Select Case tabInfo.CurrTab
            Case Tabs(eGDTab_WorkingOrders)
                Set fg = fgWorkingOrders
            Case Tabs(eGDTab_FilledOrders)
                Set fg = fgFilledOrders
            Case Tabs(eGDTab_OpenPositions)
                Set fg = fgPositions
            Case Tabs(eGDTab_AccountStatus)
                Set fg = fgAccountStatus
            Case Tabs(eGDTab_Statement)
                Set fg = fgStatement
            Case Tabs(eGDTab_ParkedOrders)
                Set fg = fgParkedOrders
        End Select
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fg
        Else
            .RenderControl = fg.hWnd
        End If

        .EndDoc
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reload
'' Description: Reload the form information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reload()
On Error GoTo ErrSection

    Dim lTimeOut As Long                ' Time-out variable

    Screen.MousePointer = vbHourglass
    
    If g.Alaron.ConnectionStatus = eGDConnectionStatus_Connecting Then
        Do While g.Alaron.ConnectionStatus <> eGDConnectionStatus_Connected And lTimeOut < 30&
            Sleep 1
            lTimeOut = lTimeOut + 1&
        Loop
    End If
    
    If g.Alaron.ConnectionStatus = eGDConnectionStatus_Connected Then
        g.Alaron.DumpDebug "Getting Account Information from Broker Screen Reload"
        
        InfBox "Retrieving Account Information...", , , "Alaron Information Retrieval", True
        GetAccounts
        LoadAccountsCombo
        
        InfBox "Retrieving Security List...", , , "Alaron Information Retrieval", True
        GetSecurities
        LoadSecuritiesCombo
        
        If (cboAccount.ListIndex = -1&) And (cboAccount.ListCount = 1) Then
            cboAccount.ListIndex = 0&
        End If
        
        InfBox ""
        
        EnableControls
    End If
    
    Screen.MousePointer = vbDefault

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.Reload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearForm
'' Description: Clear out the form information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearForm()
On Error GoTo ErrSection:

    cboAccount.Clear
    InitAccountStatusGrid
    fgFilledOrders.Rows = fgFilledOrders.FixedRows
    fgParkedOrders.Rows = fgParkedOrders.FixedRows
    fgPositions.Rows = fgPositions.FixedRows
    fgStatement.Rows = fgStatement.FixedRows
    fgWorkingOrders.Rows = fgWorkingOrders.FixedRows
    
    m.astrAccounts.Clear
    m.astrAcctStatus.Clear
    m.astrOrders.Clear
    m.astrPositions.Clear
    m.astrSecurities.Clear
    m.astrStatement.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.ClearForm"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccount_Click
'' Description: When the user chooses an account, refilter the grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccount_Click()
On Error GoTo ErrSection:
    
    If Visible Then
        DoRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccount_KeyUp
'' Description: As the user changes the text, try to select the account
'' Inputs:      Code of Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccount_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ''SelectAccount Parse(cboAccount.Text, "-", 1)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboAccount_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBuySell_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBuySell_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboBuySell_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCommodity_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCommodity_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboCommodity_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboGoodThru_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboGoodThru_Click()
On Error GoTo ErrSection:

    If Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboGoodThru_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboMonth_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboMonth_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboMonth_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOpenClose_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOpenClose_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboOpenClose_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboPutCall_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboPutCall_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboPutCall_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboType_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboType_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        txtLimit(Index).Visible = (cboType(Index).Text = "STWL")
        
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboType_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboYear_Click
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboYear_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cboYear_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkSpread_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkSpread_Click()
On Error GoTo ErrSection:

    If Visible Then
        ClearSide 1
        ClearSide 2
        EnableControls
        
        If m.bParsing = False Then
            UpdateCommandLine
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.chkSpread_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAccountLookup_Click
'' Description: Allow the user to lookup an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAccountLookup_Click()
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account returned from lookup form
    
    strAccount = frmAccountLookup.ShowMe(m.astrAccounts)
    If Len(Trim(strAccount)) > 0 Then
        SelectAccount strAccount
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.cmdAccountLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancelOrder_Click
'' Description: Allow the broker to cancel an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancelOrder_Click()
On Error GoTo ErrSection:

    CancelOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdCancelOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: Close the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.cmdClose_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdGo_Click
'' Description: Select the current item in the accounts combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdGo_Click()
On Error GoTo ErrSection:

    SelectAccount Parse(cboAccount.Text, "-", 1)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdGo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdModifyOrder_Click
'' Description: Allow the broker to modify the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdModifyOrder_Click()
On Error GoTo ErrSection:

    ModifyOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdModifyOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdParkOrder_Click
'' Description: Allow the user to park an order on the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdParkOrder_Click()
On Error GoTo ErrSection:

    ParkOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdParkOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Allow the user to print the currently selected tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRefresh_Click
'' Description: Refresh either all accounts or the selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRefresh_Click()
On Error GoTo ErrSection:

    DoRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdRefresh_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdResetOrder_Click
'' Description: Reset all of the order controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdResetOrder_Click()
On Error GoTo ErrSection:

    ResetOrder

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.cmdResetOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmit_Click
'' Description: Submit the selected order to the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmitOrder_Click()
On Error GoTo ErrSection:

    SubmitOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdSubmitOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdVerifyOrder_Click
'' Description: Verify the order information supplied with the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdVerifyOrder_Click()
On Error GoTo ErrSection:

    VerifyOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.cmdVerifyOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilledOrders_AfterSort
'' Description: After the user sorts the grid, re-color the backcolor
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilledOrders_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgFilledOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgFilledOrders_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFilledOrders_BeforeMouseDown
'' Description: Popup a menu on a right click
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFilledOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        mnuSubmitOrder.Visible = False
        mnuModifyOrder.Visible = False
        mnuCancelOrder.Visible = False
        mnuSep1.Visible = False
        
        mnuRefreshAccount.Enabled = True
        mnuPrint.Enabled = True
        
        PopupMenu mnuOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgFilledOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgParkedOrders_AfterRowColChange
'' Description: After a row/column change, enable/disable controls
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgParkedOrders_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgParkedOrders_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgParkedOrders_AfterSort
'' Description: After the user sorts the grid, re-color the backcolor
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgParkedOrders_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgParkedOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgParkedOrders_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgParkedOrders_BeforeMouseDown
'' Description: Popup a menu on a right click
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgParkedOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim bValidRow As Boolean            ' Is this a valid row?
    Dim strType As String               ' Type from the grid
    Dim bCancelled As Boolean           ' Has this order been cancelled?

    If Button = vbRightButton Then
        lMouseRow = fgParkedOrders.MouseRow
        bValidRow = (lMouseRow >= fgParkedOrders.FixedRows) And (lMouseRow < fgParkedOrders.Rows)
        If bValidRow Then
            strType = UCase(Trim(fgParkedOrders.TextMatrix(lMouseRow, WorkingOrdersCol(eGDWorkingCol_OrderType))))
            bCancelled = (strType = "CXL") Or (strType = "CXLR")
        
            fgParkedOrders.Row = lMouseRow
            fgParkedOrders.RowSel = lMouseRow
        End If
        
        mnuSubmitOrder.Visible = bValidRow And (Not bCancelled)
        mnuModifyOrder.Enabled = bValidRow And (Not bCancelled)
        mnuCancelOrder.Enabled = bValidRow And (Not bCancelled)
        
        mnuRefreshAccount.Enabled = True
        mnuPrint.Enabled = True
        
        PopupMenu mnuOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgParkedOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_AfterSort
'' Description: After the user sorts the grid, re-color the backcolor
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgPositions

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgPositions_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgWorkingOrders_AfterRowColChange
'' Description: After a row/column change, enable/disable controls
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgWorkingOrders_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgWorkingOrders_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgWorkingOrders_AfterSort
'' Description: After the user sorts the grid, re-color the backcolor
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgWorkingOrders_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgWorkingOrders

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgWorkingOrders_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgWorkingOrders_BeforeMouseDown
'' Description: Popup a menu on a right click
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgWorkingOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim bValidRow As Boolean            ' Is this a valid row?
    Dim strType As String               ' Type from the grid
    Dim bCancelled As Boolean           ' Has this order been cancelled?

    If Button = vbRightButton Then
        lMouseRow = fgWorkingOrders.MouseRow
        bValidRow = (lMouseRow >= fgWorkingOrders.FixedRows) And (lMouseRow < fgWorkingOrders.Rows)
        If bValidRow Then
            strType = UCase(Trim(fgWorkingOrders.TextMatrix(lMouseRow, WorkingOrdersCol(eGDWorkingCol_OrderType))))
            bCancelled = (strType = "CXL") Or (strType = "CXLR")
            
            fgWorkingOrders.Row = lMouseRow
            fgWorkingOrders.RowSel = lMouseRow
        End If
        
        mnuSubmitOrder.Visible = False
        mnuModifyOrder.Enabled = bValidRow And (Not bCancelled)
        mnuCancelOrder.Enabled = bValidRow And (Not bCancelled)
        
        mnuRefreshAccount.Enabled = True
        mnuPrint.Enabled = True
        
        PopupMenu mnuOrders
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.fgWorkingOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Perform some actions when the form is activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    If (m.astrAccounts.Size = 0) Or (cboAccount.ListCount = 0) Then
        Reload
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement string from the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Integer              ' Index into a for loop

    Caption = "Broker Management"
    Icon = Picture16("kBlank")

    strPlacement = GetIniFileProperty("frmBrokerScreen", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    mnuOrders.Visible = False
    
    Set m.astrAccounts = New cGdArray
    Set m.astrAcctStatus = New cGdArray
    Set m.astrOrders = New cGdArray
    Set m.astrPositions = New cGdArray
    Set m.astrSecurities = New cGdArray
    Set m.astrStatement = New cGdArray
    
    For lIndex = 0 To 2
        cboBuySell(lIndex).AddItem "Buy"
        cboBuySell(lIndex).AddItem "Sell"
    Next lIndex
    
    For lIndex = 0 To 2
        For lIndex2 = 1 To 12
            cboMonth(lIndex).AddItem MonthName(lIndex2)
        Next lIndex2
    Next lIndex
    
    For lIndex = 0 To 2
        For lIndex2 = 0 To 9
            cboYear(lIndex).AddItem Str(Year(Now) + lIndex2)
        Next lIndex2
    Next lIndex
    
    For lIndex = 0 To 2
        cboPutCall(lIndex).AddItem "Put"
        cboPutCall(lIndex).AddItem "Call"
    Next lIndex
    
    For lIndex = 0 To 2
        cboOpenClose(lIndex).AddItem "Open"
        cboOpenClose(lIndex).AddItem "Close"
    Next lIndex
    
    For lIndex = 0 To 2
        'cboType(lIndex).AddItem "DRT"
        cboType(lIndex).AddItem "FOK"
        cboType(lIndex).AddItem "MIT"
        cboType(lIndex).AddItem "MOC"
        cboType(lIndex).AddItem "OB"
        cboType(lIndex).AddItem "OBOO"
        cboType(lIndex).AddItem "OO"
        cboType(lIndex).AddItem "SCO"
        cboType(lIndex).AddItem "SLCO"
        cboType(lIndex).AddItem "SOO"
        cboType(lIndex).AddItem "STL"
        cboType(lIndex).AddItem "STOP"
        cboType(lIndex).AddItem "STWL"
    Next lIndex
    
    cboGoodThru.AddItem "Today"
    cboGoodThru.AddItem "Cancel"
    cboGoodThru.AddItem "Date"
    
    cboSession.AddItem "Day"
    cboSession.AddItem "Electronic"
    
    optSpecificAccount.Value = True
    optSpecificAccount.Visible = False
    cboAccount.Left = optSpecificAccount.Left
    cmdGo.Left = cboAccount.Width + (cboAccount.Left * 2)
    cmdAccountLookup.Left = cmdGo.Left + cmdGo.Width + cboAccount.Left
    optAllAccounts.Visible = False
    chkSpread.Visible = False
    
    lblOrderID.Visible = False
    txtOrderID.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Determine whether we want to allow the form to unload
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, 11430, 4900) Then Exit Sub
    
    With fraAccountSelection
        .Move 60, 60
    End With
    
    With fraButtons
        .Move ScaleWidth - .Width - 60, 60
    End With
    
    With tabInfo
        .Move 60, fraAccountSelection.Height + 120, ScaleWidth - fraButtons.Width - 180, ScaleHeight - fraAccountSelection.Height - 180
        .Refresh
    End With
    
    With fgWorkingOrders
        .Move 60, 60, fraWorkingOrders.Width - 120, fraWorkingOrders.Height - 120
    End With

    With fgFilledOrders
        .Move 60, 60, fraFilledOrders.Width - 120, fraFilledOrders.Height - 120
    End With

    With fgPositions
        .Move 60, 60, fraPositions.Width - 120, fraPositions.Height - 120
    End With

    With fgAccountStatus
        .Move 60, 60, fraAccountStatus.Width - 120, fraAccountStatus.Height - 120
    End With

    With fgStatement
        .Move 60, 60, fraStatement.Width - 120, fraStatement.Height - 120
    End With

    With fgParkedOrders
        .Move 60, 60, fraParkedOrders.Width - 120, fraParkedOrders.Height - 120
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up and save settings when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmBrokerScreen", GetFormPlacement(Me), "Placement", g.strIniFile
    
    Set m.astrAccounts = Nothing
    Set m.astrAcctStatus = Nothing
    Set m.astrOrders = Nothing
    Set m.astrPositions = Nothing
    Set m.astrSecurities = Nothing
    Set m.astrStatement = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCancelOrder_Click
'' Description: Cancel the order on the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCancelOrder_Click()
On Error GoTo ErrSection:

    CancelOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.mnuCancelOrder_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuModifyOrder_Click
'' Description: Cancel/Replace the order on the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuModifyOrder_Click()
On Error GoTo ErrSection:

    ModifyOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.mnuModifyOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrint_Click
'' Description: Allow the user to print the currently selected tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.mnuPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRefresh_Click
'' Description: Refresh either all accounts or the selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRefreshAccount_Click()
On Error GoTo ErrSection:

    DoRefresh

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.mnuRefreshAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmitOrder_Click
'' Description: Submit the order to the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmitOrder_Click()
On Error GoTo ErrSection:

    SubmitOrder

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.mnuSubmitOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAllAccounts_Click
'' Description: Enable/Disable/Hide/Show controls and columns appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAllAccounts_Click()
On Error GoTo ErrSection:

    If Visible Then
        DoRefresh
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.optAllAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSpecificAccount_Click
'' Description: Enable/Disable/Hide/Show controls and columns appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSpecificAccount_Click()
On Error GoTo ErrSection:

    If Visible Then
        DoRefresh
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.optSpecificAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitWorkingOrdersGrid
'' Description: Initialize the working orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitWorkingOrdersGrid()
On Error GoTo ErrSection:

    With fgWorkingOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = WorkingOrdersCol(eGDWorkingCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Account)) = "Account"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OrderNumber)) = "Order #"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_ReferenceID)) = "Ref ID"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OrderDate)) = "Order Date"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_BuySell)) = "B/S"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Quantity)) = "Quantity"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Symbol)) = "Symbol"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Strike)) = "Strike"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Price)) = "Price"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OC)) = "O/C"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OrderType)) = "Type"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Limit)) = "Limit"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Tic)) = "Tic"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_GoodThru)) = "Good Thru"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Session)) = "Session"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_CreditControl)) = "Credit Control"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_ExchangeMessage)) = "Exchange Msg"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_SpecialInstruction)) = "Special Instructions"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_EnteredBy)) = "Entered By"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Salesman)) = "Sales"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Exchange)) = "Exchange"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_BaseSymbol)) = "Base Symbol"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Contract)) = "Contract"
        
        .ColAlignment(WorkingOrdersCol(eGDWorkingCol_Account)) = flexAlignRightTop
        .ColAlignment(WorkingOrdersCol(eGDWorkingCol_ReferenceID)) = flexAlignRightTop
        
        .ColHidden(WorkingOrdersCol(eGDWorkingCol_Exchange)) = True
        .ColHidden(WorkingOrdersCol(eGDWorkingCol_BaseSymbol)) = True
        .ColHidden(WorkingOrdersCol(eGDWorkingCol_Contract)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.InitWorkingOrdersGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFilledOrdersGrid
'' Description: Initialize the filled orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFilledOrdersGrid()
On Error GoTo ErrSection:

    With fgFilledOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = FilledOrdersCol(eGDFilledCol_NumCols)
        .FixedCols = 0
                        
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Account)) = "Account"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_OrderNumber)) = "Order #"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_TradeDate)) = "Trade Date"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Position)) = "Position"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Quantity)) = "Quantity"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Symbol)) = "Symbol"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Strike)) = "Strike"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_TradePrice)) = "Trade Price"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Tic)) = "Tic"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_FuturesProfit)) = "Futures Profit"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_OptionsValue)) = "Options Value"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_Currency)) = "Currency"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_OrderTime)) = "Order Time"
        .TextMatrix(0, FilledOrdersCol(eGDFilledCol_FillTime)) = "Fill Time"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.InitFilledOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitPositionsGrid
'' Description: Initialize the positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitPositionsGrid()
On Error GoTo ErrSection:

    With fgPositions
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = PositionsCol(eGDPositionCol_NumCols)
        .FixedCols = 0
                        
        .TextMatrix(0, PositionsCol(eGDPositionCol_Account)) = "Account"
        .TextMatrix(0, PositionsCol(eGDPositionCol_TradeDate)) = "Trade Date"
        .TextMatrix(0, PositionsCol(eGDPositionCol_Position)) = "Position"
        .TextMatrix(0, PositionsCol(eGDPositionCol_Quantity)) = "Quantity"
        .TextMatrix(0, PositionsCol(eGDPositionCol_Symbol)) = "Symbol"
        .TextMatrix(0, PositionsCol(eGDPositionCol_Strike)) = "Strike"
        .TextMatrix(0, PositionsCol(eGDPositionCol_TradePrice)) = "Trade Price"
        .TextMatrix(0, PositionsCol(eGDPositionCol_Tic)) = "Tic"
        .TextMatrix(0, PositionsCol(eGDPositionCol_FuturesProfit)) = "Futures Profit"
        .TextMatrix(0, PositionsCol(eGDPositionCol_OptionValue)) = "Options Value"
        .TextMatrix(0, PositionsCol(eGDPositionCol_Currency)) = "Currency"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.InitPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitParkedOrdersGrid
'' Description: Initialize the parked orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitParkedOrdersGrid()
On Error GoTo ErrSection:

    With fgParkedOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = WorkingOrdersCol(eGDWorkingCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Account)) = "Account"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OrderNumber)) = "Order #"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_ReferenceID)) = "Ref ID"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OrderDate)) = "Order Date"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_BuySell)) = "B/S"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Quantity)) = "Quantity"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Symbol)) = "Symbol"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Strike)) = "Strike"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Price)) = "Price"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OC)) = "O/C"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_OrderType)) = "Type"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Limit)) = "Limit"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Tic)) = "Tic"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_GoodThru)) = "Good Thru"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Session)) = "Session"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_CreditControl)) = "Credit Control"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_ExchangeMessage)) = "Exchange Msg"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_SpecialInstruction)) = "Special Instructions"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_EnteredBy)) = "Entered By"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Salesman)) = "Sales"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Exchange)) = "Exchange"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_BaseSymbol)) = "Base Symbol"
        .TextMatrix(0, WorkingOrdersCol(eGDWorkingCol_Contract)) = "Contract"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .ColHidden(WorkingOrdersCol(eGDWorkingCol_Exchange)) = True
        .ColHidden(WorkingOrdersCol(eGDWorkingCol_BaseSymbol)) = True
        .ColHidden(WorkingOrdersCol(eGDWorkingCol_Contract)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.InitParkedOrdersGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitAccountStatusGrid
'' Description: Initialize the account status grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitAccountStatusGrid()
On Error GoTo ErrSection:

    With fgAccountStatus
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorBkg = .BackColor
        .Editable = flexEDNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .MergeCells = flexMergeFree
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = .BackColor
        
        .Rows = 14
        .FixedRows = 0
        .Cols = 5
        .FixedCols = 0
        
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "Margin Requirements"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .MergeRow(0) = True
        
        .Cell(flexcpText, 1, 0, 1, 1) = "Current"
        .Cell(flexcpText, 1, 3, 1, 4) = "Working Orders(estimated)"
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontUnderline, 1, 0, 1, .Cols - 1) = True
        .MergeRow(1) = True
        
        .TextMatrix(2, 0) = "Initial:"
        .TextMatrix(2, 3) = "Initial:"
        .TextMatrix(3, 0) = "Maintenance:"
        .TextMatrix(3, 3) = "Maintenance:"
        
        .Cell(flexcpText, 4, 0, 4, .Cols - 1) = "Option Valuation"
        .Cell(flexcpAlignment, 4, 0, 4, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 4, 0, 4, .Cols - 1) = True
        .MergeRow(4) = True
        
        .TextMatrix(5, 0) = "Long Option Value:"
        .TextMatrix(5, 3) = "Short Option Value:"
        
        .Cell(flexcpText, 6, 0, 6, .Cols - 1) = "Account Status"
        .Cell(flexcpAlignment, 6, 0, 6, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 6, 0, 6, .Cols - 1) = True
        .MergeRow(6) = True
        
        .Cell(flexcpText, 7, 0, 7, 1) = "Start of Day"
        .Cell(flexcpText, 7, 3, 7, 4) = "Marked to the Market"
        .Cell(flexcpAlignment, 7, 0, 7, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontUnderline, 7, 0, 7, .Cols - 1) = True
        .MergeRow(7) = True
        
        .TextMatrix(8, 0) = "Ledger Balance:"
        .TextMatrix(8, 3) = "Ledger Balance:"
        .TextMatrix(9, 0) = "Open Trade Equity:"
        .TextMatrix(9, 3) = "Open Trade Equity:"
        .TextMatrix(10, 0) = "Total Account Equity:"
        .TextMatrix(10, 3) = "Total Account Equity:"
        .TextMatrix(11, 3) = "Net Option Value:"
        .TextMatrix(12, 0) = "Securities on Deposit:"
        .TextMatrix(12, 3) = "Net Liquidation Value:"
        .TextMatrix(13, 3) = "Excess Funds:"
        
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignRightTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop
        
        .ColWidth(0) = .Width / .Cols
        .ColWidth(1) = .Width / .Cols
        .ColWidth(2) = .Width / .Cols
        .ColWidth(3) = .Width / .Cols
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.InitAccountStatusGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitStatementGrid
'' Description: Initialize the statement grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitStatementGrid()
On Error GoTo ErrSection:

    With fgStatement
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorBkg = .BackColor
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = .BackColor
        
        .FontName = "Courier"
        
        .Rows = 0
        .FixedRows = 0
        .Cols = 1
        .FixedCols = 0
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.InitStatementGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable/Hide/Show controls and columns appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls(Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim bSpecificAccount As Boolean     ' Is this a specific account?

    bSpecificAccount = optSpecificAccount.Value
    
    Enable cboAccount, ((bSpecificAccount = True) And (cboAccount.ListCount > 1))
    Enable cmdAccountLookup, ((bSpecificAccount = True) And (cboAccount.ListCount > 1))
    Enable cmdGo, ((bSpecificAccount = True) And (cboAccount.ListCount > 1))
    
    Enable cmdPrint, (lTab <> Tabs(eGDTab_NewOrder))

    fgWorkingOrders.ColHidden(WorkingOrdersCol(eGDWorkingCol_Account)) = bSpecificAccount
    fgFilledOrders.ColHidden(FilledOrdersCol(eGDFilledCol_Account)) = bSpecificAccount
    fgPositions.ColHidden(PositionsCol(eGDPositionCol_Account)) = bSpecificAccount
    fgParkedOrders.ColHidden(WorkingOrdersCol(eGDWorkingCol_Account)) = bSpecificAccount
    
    If bSpecificAccount = False And tabInfo.CurrTab >= Tabs(eGDTab_AccountStatus) Then
        tabInfo.CurrTab = Tabs(eGDTab_WorkingOrders)
    End If
    
    tabInfo.TabEnabled(Tabs(eGDTab_AccountStatus)) = bSpecificAccount
    tabInfo.TabEnabled(Tabs(eGDTab_Statement)) = bSpecificAccount
    tabInfo.TabEnabled(Tabs(eGDTab_OpenPositions)) = bSpecificAccount
    tabInfo.TabEnabled(Tabs(eGDTab_NewOrder)) = bSpecificAccount
    
    If lTab = -1& Then lTab = tabInfo.CurrTab
    Select Case lTab
        Case Tabs(eGDTab_WorkingOrders)
            With fgWorkingOrders
                Enable cmdModifyOrder, (.Row >= .FixedRows) And (.Row < .Rows)
                Enable cmdCancelOrder, (.Row >= .FixedRows) And (.Row < .Rows)
                'Enable cmdParkOrder, (.Row >= .FixedRows) And (.Row < .Rows)
            End With
            Enable cmdParkOrder, False
            Enable cmdSubmitOrder, False
            
        Case Tabs(eGDTab_FilledOrders)
            Enable cmdModifyOrder, False
            Enable cmdCancelOrder, False
            Enable cmdSubmitOrder, False
            Enable cmdParkOrder, False
            
        Case Tabs(eGDTab_ParkedOrders)
            With fgParkedOrders
                Enable cmdModifyOrder, (.Row >= .FixedRows) And (.Row < .Rows)
                Enable cmdCancelOrder, (.Row >= .FixedRows) And (.Row < .Rows)
                Enable cmdSubmitOrder, (.Row >= .FixedRows) And (.Row < .Rows)
            End With
            Enable cmdParkOrder, False
        
        Case Tabs(eGDTab_AccountStatus)
            Enable cmdModifyOrder, False
            Enable cmdCancelOrder, False
            Enable cmdSubmitOrder, False
            Enable cmdParkOrder, False
            
        Case Tabs(eGDTab_Statement)
            Enable cmdModifyOrder, False
            Enable cmdCancelOrder, False
            Enable cmdSubmitOrder, False
            Enable cmdParkOrder, False
            
        Case Tabs(eGDTab_OpenPositions)
            Enable cmdModifyOrder, False
            Enable cmdCancelOrder, False
            Enable cmdSubmitOrder, False
            Enable cmdParkOrder, False
            
        Case Tabs(eGDTab_NewOrder)
            Enable cmdModifyOrder, False
            Enable cmdCancelOrder, False
            Enable cmdSubmitOrder, True
            Enable cmdParkOrder, True
            
    End Select
    
    EnableOrderControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableOrderControls
'' Description: Enable/Disable/Hide/Show controls on the new order tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableOrderControls()
On Error GoTo ErrSection:

    Dim bNewOrder As Boolean            ' Is this a new order or a modify order?

    bNewOrder = (Len(Trim(txtOrderID.Text)) = 0)

    cboBuySell(1).Visible = (chkSpread.Value = vbChecked)
    cboBuySell(2).Visible = (chkSpread.Value = vbChecked)
    txtQuantity(1).Visible = (chkSpread.Value = vbChecked)
    txtQuantity(2).Visible = (chkSpread.Value = vbChecked)
    cboCommodity(1).Visible = (chkSpread.Value = vbChecked)
    cboCommodity(2).Visible = (chkSpread.Value = vbChecked)
    cboMonth(1).Visible = (chkSpread.Value = vbChecked)
    cboMonth(2).Visible = (chkSpread.Value = vbChecked)
    cboYear(1).Visible = (chkSpread.Value = vbChecked)
    cboYear(2).Visible = (chkSpread.Value = vbChecked)
    txtStrike(1).Visible = (chkSpread.Value = vbChecked)
    txtStrike(2).Visible = (chkSpread.Value = vbChecked)
    cboPutCall(1).Visible = (chkSpread.Value = vbChecked)
    cboPutCall(2).Visible = (chkSpread.Value = vbChecked)
    txtPrice(1).Visible = (chkSpread.Value = vbChecked)
    txtPrice(2).Visible = (chkSpread.Value = vbChecked)
    cboOpenClose(1).Visible = (chkSpread.Value = vbChecked)
    cboOpenClose(2).Visible = (chkSpread.Value = vbChecked)
    cboType(1).Visible = (chkSpread.Value = vbChecked)
    cboType(2).Visible = (chkSpread.Value = vbChecked)
    txtLimit(1).Visible = (chkSpread.Value = vbChecked)
    txtLimit(2).Visible = (chkSpread.Value = vbChecked)
    
    gdThruDate.Visible = (UCase(Trim(cboGoodThru.Text)) = "DATE")
    
    chkSpread.Enabled = bNewOrder
    
    cboBuySell(0).Enabled = bNewOrder
    cboBuySell(1).Enabled = bNewOrder
    cboBuySell(2).Enabled = bNewOrder
    
    cboCommodity(0).Enabled = bNewOrder
    cboCommodity(1).Enabled = bNewOrder
    cboCommodity(2).Enabled = bNewOrder
    
    cboMonth(0).Enabled = bNewOrder
    cboMonth(1).Enabled = bNewOrder
    cboMonth(2).Enabled = bNewOrder
    
    cboYear(0).Enabled = bNewOrder
    cboYear(1).Enabled = bNewOrder
    cboYear(2).Enabled = bNewOrder
    
    txtStrike(0).Enabled = bNewOrder
    txtStrike(1).Enabled = bNewOrder
    txtStrike(2).Enabled = bNewOrder
    
    cboPutCall(0).Enabled = bNewOrder
    cboPutCall(1).Enabled = bNewOrder
    cboPutCall(2).Enabled = bNewOrder
    
    cboOpenClose(0).Enabled = bNewOrder
    cboOpenClose(1).Enabled = bNewOrder
    cboOpenClose(2).Enabled = bNewOrder
    
    lblSession.Enabled = bNewOrder
    cboSession.Enabled = bNewOrder
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.EnableOrderControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectAccount
'' Description: Select the appropriate account in the combo box
'' Inputs:      Selection
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectAccount(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    cboAccount.ListIndex = -1&
    If Len(strAccount) > 0 Then
        For lIndex = 0 To cboAccount.ListCount - 1
            If Parse(cboAccount.List(lIndex), "-", 1) = strAccount Then
                cboAccount.ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End If
    
    If cboAccount.ListIndex = -1& And cboAccount.ListCount = 1& Then
        cboAccount.ListIndex = 0&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SelectAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrids
'' Description: Filter the grids according to the chosen account (or all)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrids()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Currently selected account
    
    strAccount = Parse(cboAccount.Text, "-", 1)
    
    With fgWorkingOrders
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If (optAllAccounts.Value = True) Or (.TextMatrix(lIndex, WorkingOrdersCol(eGDWorkingCol_Account)) = strAccount) Then
                .RowHidden(lIndex) = False
            Else
                .RowHidden(lIndex) = True
            End If
        Next lIndex
        
        SetBackColors fgWorkingOrders
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

    With fgFilledOrders
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If (optAllAccounts.Value = True) Or (.TextMatrix(lIndex, FilledOrdersCol(eGDFilledCol_Account)) = strAccount) Then
                .RowHidden(lIndex) = False
            Else
                .RowHidden(lIndex) = True
            End If
        Next lIndex
        
        SetBackColors fgFilledOrders
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

    With fgPositions
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If (optAllAccounts.Value = True) Or (.TextMatrix(lIndex, PositionsCol(eGDPositionCol_Account)) = strAccount) Then
                .RowHidden(lIndex) = False
            Else
                .RowHidden(lIndex) = True
            End If
        Next lIndex
        
        SetBackColors fgPositions
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

    With fgParkedOrders
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If (optAllAccounts.Value = True) Or (.TextMatrix(lIndex, WorkingOrdersCol(eGDWorkingCol_Account)) = strAccount) Then
                .RowHidden(lIndex) = False
            Else
                .RowHidden(lIndex) = True
            End If
        Next lIndex
        
        SetBackColors fgParkedOrders
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.FilterGrids"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountStatusCell
'' Description: Set the text and the color for the cell based on the value
'' Inputs:      Row, Column, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AccountStatusCell(ByVal lRow As Long, ByVal lCol As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    With fgAccountStatus
        .TextMatrix(lRow, lCol) = strValue
        If Len(strValue) > 0 Then
            If Val(strValue) >= 0 Then
                .Cell(flexcpForeColor, lRow, lCol) = vbGreen
            Else
                .Cell(flexcpForeColor, lRow, lCol) = vbRed
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AccountStatusCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccount
'' Description: Refresh the account information for the given account
'' Inputs:      Account Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshAccount(ByVal strAccountNumber As String)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    If Not g.Alaron.Refreshing Then
        g.Alaron.DumpDebug "Refresh Account from Broker Screen for " & strAccountNumber
        
        m.bOrdersLoaded = False
        m.bPositionsLoaded = False
        m.bAccountStatusLoaded = False
        m.bStatementLoaded = False
        
        m.astrOrders.Size = 0
        m.astrPositions.Size = 0
        m.astrAcctStatus.Size = 0
        m.astrStatement.Size = 0
        
        InfBox "Refreshing Account Information for " & strAccountNumber & "....", , , "Refreshing Alaron Account", True
        g.Alaron.RefreshAccountInfo strAccountNumber
        
        lTimeOut = 0&
        Do While (g.Alaron.Refreshing = True) And (lTimeOut < 30&)
            Sleep 1&
            lTimeOut = lTimeOut + 1&
        Loop
        
        LoadOrdersGrids
        LoadOpenPositionsGrid
        LoadAccountStatusGrid
        LoadStatementGrid
        
        InfBox ""
    Else
        g.Alaron.DumpDebug "Refresh Account called for but refresh already in progress"
    End If
    
#If 0 Then
    ' Get the orders for the given account number...
    InfBox "Refreshing Orders for " & strAccountNumber & "....", , , "Refreshing Alaron Account", True
    GetOrders strAccountNumber
    LoadOrdersGrids

    ' Get the positions for the given account number...
    InfBox "Refreshing Open Positions for " & strAccountNumber & "....", , , "Refreshing Alaron Account", True
    GetPositions strAccountNumber
    LoadOpenPositionsGrid
    
    ' Get the Account Status record for the given account number...
    InfBox "Refreshing Account Status for " & strAccountNumber & "....", , , "Refreshing Alaron Account", True
    GetAccountStatus strAccountNumber
    LoadAccountStatusGrid

    ' Get the statement for the given account number...
    InfBox "Refreshing Statement Information for " & strAccountNumber & "....", , , "Refreshing Alaron Account", True
    GetStatement strAccountNumber
    LoadStatementGrid

    InfBox ""
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.RefreshAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParseCommandLine
'' Description: Parse the command line and fill in controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParseCommandLine()
On Error GoTo ErrSection:

    Dim astrSide As New cGdArray        ' Array of spread sides
    Dim astrCommand As New cGdArray     ' Array of parsed commands
    Dim lIndex As Long                  ' Index into a for loop
    
    m.bParsing = True
        
    astrSide.SplitFields Trim(txtCommandLine.Text), "/"
    If astrSide.Size < 2 Then chkSpread.Value = vbUnchecked Else chkSpread.Value = vbChecked
    
    For lIndex = 0 To astrSide.Size - 1
        astrCommand.SplitFields astrSide(lIndex), " "
        
        SetBuySell lIndex, astrCommand(0)
        txtQuantity(lIndex).Text = astrCommand(1)
        SetCommodity lIndex, astrCommand(2)
        If Len(astrCommand(3)) > 0 Then
            If UCase(Left(astrCommand(3), 1)) = "M" Then
                SetMarketOrder lIndex, False
            ElseIf IsDigit(Left(astrCommand(3), 1)) Then
                If UCase(Right(astrCommand(3), 1)) = "C" Or UCase(Right(astrCommand(3), 1)) = "P" Then
                    SetOptionInfo lIndex, astrCommand(3)
                    If Len(astrCommand(4)) > 0 Then
                        If UCase(Left(astrCommand(4), 1)) = "M" Then
                            SetMarketOrder lIndex, True
                            SetOpenClose lIndex, astrCommand(5)
                        Else
                            SetPriceInfo lIndex, astrCommand(4), True
                            SetOpenClose lIndex, astrCommand(5)
                            SetOrderType lIndex, astrCommand(6)
                            If txtLimit(lIndex).Visible = True Then
                                If Val(astrCommand(7)) <> 0 Then
                                    txtLimit(lIndex).Text = Str(Val(astrCommand(7)))
                                Else
                                    txtLimit(lIndex).Text = ""
                                End If
                            End If
                        End If
                    Else
                        SetPriceInfo lIndex, "", True
                    End If
                Else
                    SetPriceInfo lIndex, astrCommand(3), False
                    SetOrderType lIndex, astrCommand(4)
                    If txtLimit(lIndex).Visible = True Then
                        If Val(astrCommand(5)) <> 0 Then
                            txtLimit(lIndex).Text = Str(Val(astrCommand(5)))
                        Else
                            txtLimit(lIndex).Text = ""
                        End If
                    End If
                End If
            End If
        Else
            txtPrice(lIndex).Text = ""
            txtStrike(lIndex).Visible = True
            txtStrike(lIndex).Text = ""
            cboPutCall(lIndex).Visible = True
            cboPutCall(lIndex).ListIndex = -1&
            cboOpenClose(lIndex).Visible = True
            cboOpenClose(lIndex).ListIndex = -1&
            cboType(lIndex).Visible = True
            cboType(lIndex).ListIndex = -1&
            txtLimit(lIndex).Visible = False
            txtLimit(lIndex).Text = ""
        End If
    Next lIndex
    
    For lIndex = astrSide.Size To 2
        ClearSide lIndex
    Next lIndex
    
ErrExit:
    m.bParsing = False
    Exit Sub
    
ErrSection:
    m.bParsing = False
    RaiseError "frmBrokerScreen.ParseCommandLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetBuySell
'' Description: Set the buy/sell combo appropriately
'' Inputs:      Control Index, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetBuySell(ByVal lIndex As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    If UCase(Left(strValue, 1)) = "B" Then
        cboBuySell(lIndex).Text = "Buy"
    ElseIf UCase(Left(strValue, 1)) = "S" Then
        cboBuySell(lIndex).Text = "Sell"
    Else
        cboBuySell(lIndex).ListIndex = -1&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetBuySell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCommodity
'' Description: Set the commodity combo appropriately
'' Inputs:      Control Index, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCommodity(ByVal lIndex As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim lIndex2 As Long                 ' Index into a for loop
    Dim iMonth As Integer               ' Month entered
    
    If Len(strValue) >= 3 Then
        cboCommodity(lIndex).ListIndex = -1&
        For lIndex2 = 0 To cboCommodity(lIndex).ListCount - 1
            If UCase(Parse(cboCommodity(lIndex).List(lIndex2), "-", 1)) = UCase(Left(strValue, Len(strValue) - 2)) Then
                cboCommodity(lIndex).ListIndex = lIndex2
                Exit For
            End If
        Next lIndex2
        
        iMonth = CodeToMonth(Mid(strValue, Len(strValue) - 1, 1))
        If iMonth = 0 Then
            cboMonth(lIndex).ListIndex = -1&
        Else
            cboMonth(lIndex).Text = MonthName(iMonth)
        End If
        
        cboYear(lIndex).ListIndex = -1&
        For lIndex2 = 0 To cboYear(lIndex).ListCount - 1
            If Right(cboYear(lIndex).List(lIndex2), 1) = Right(strValue, 1) Then
                cboYear(lIndex).ListIndex = lIndex2
                Exit For
            End If
        Next lIndex2
    Else
        cboCommodity(lIndex).ListIndex = -1&
        cboMonth(lIndex).ListIndex = -1&
        cboYear(lIndex).ListIndex = -1&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetCommodity"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetMarketOrder
'' Description: Set the controls appropriately for a market order
'' Inputs:      Control Index, Is Option?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMarketOrder(ByVal lIndex As Long, ByVal bOption As Boolean)
On Error GoTo ErrSection:

    txtPrice(lIndex).Text = "Market"
    If bOption = False Then
        txtStrike(lIndex).Visible = False
        txtStrike(lIndex).Text = ""
        cboPutCall(lIndex).Visible = False
        cboPutCall(lIndex).ListIndex = -1&
        cboOpenClose(lIndex).Visible = False
        cboOpenClose(lIndex).ListIndex = -1&
    End If
    cboType(lIndex).Visible = False
    cboType(lIndex).ListIndex = -1&
    txtLimit(lIndex).Visible = False
    txtLimit(lIndex).Text = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetMarketOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOptionInfo
'' Description: Set the controls appropriately for options
'' Inputs:      Control Index, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOptionInfo(ByVal lIndex As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    If Len(strValue) > 1 Then
        txtStrike(lIndex).Text = Left(strValue, Len(strValue) - 1)
        txtStrike(lIndex).Visible = True
        If UCase(Right(strValue, 1)) = "C" Then
            cboPutCall(lIndex).Text = "Call"
        Else
            cboPutCall(lIndex).Text = "Put"
        End If
        cboPutCall(lIndex).Visible = True
        cboOpenClose(lIndex).Visible = True
    Else
        If UCase(Right(strValue, 1)) = "C" Then
            cboPutCall(lIndex).Text = "Call"
        Else
            cboPutCall(lIndex).Text = "Put"
        End If
        cboPutCall(lIndex).Visible = True
        cboOpenClose(lIndex).Visible = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetOptionInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceInfo
'' Description: Set the controls appropriately for a price
'' Inputs:      Control Index, Value, Is Option?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceInfo(ByVal lIndex As Long, ByVal strValue As String, ByVal bOption As Boolean)
On Error GoTo ErrSection:

    If bOption = False Then
        txtStrike(lIndex).Visible = False
        txtStrike(lIndex).Text = ""
        cboPutCall(lIndex).Visible = False
        cboPutCall(lIndex).ListIndex = -1&
        cboOpenClose(lIndex).Visible = False
        cboOpenClose(lIndex).ListIndex = -1&
    Else
        txtStrike(lIndex).Visible = True
        cboPutCall(lIndex).Visible = True
        cboOpenClose(lIndex).Visible = True
    End If
    
    If Val(strValue) = 0 Then
        txtPrice(lIndex).Text = ""
    Else
        txtPrice(lIndex).Text = Str(Val(strValue))
    End If
    
    cboType(lIndex).Visible = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetPriceInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOpenClose
'' Description: Set the controls appropriately for open/close
'' Inputs:      Control Index, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOpenClose(ByVal lIndex As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    If Len(strValue) > 0 Then
        If UCase(Left(strValue, 1)) = "O" Then
            cboOpenClose(lIndex).Text = "Open"
        ElseIf UCase(Left(strValue, 1)) = "C" Then
            cboOpenClose(lIndex).Text = "Close"
        Else
            cboOpenClose(lIndex).ListIndex = -1&
        End If
    Else
        cboOpenClose(lIndex).ListIndex = -1&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetOpenClose"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOrderType
'' Description: Set the controls appropriately for order type
'' Inputs:      Control Index, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOrderType(ByVal lIndex As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim lIndex2 As Long                 ' Index into a for loop
    
    With cboType(lIndex)
        .ListIndex = -1&
        For lIndex2 = 0 To .ListCount - 1
            If UCase(.List(lIndex2)) = UCase(strValue) Then
                .ListIndex = lIndex2
                Exit For
            End If
        Next lIndex2
    End With
    
    If UCase(strValue) = "STWL" Then
        txtLimit(lIndex).Visible = True
    Else
        txtLimit(lIndex).Visible = False
        txtLimit(lIndex).Text = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SetOrderType"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabInfo_Switch
'' Description: Enable/Disable buttons when tab is switched
'' Inputs:      Old Tab, New Tab, Cancel Switch?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabInfo_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    If OldTab = Tabs(eGDTab_NewOrder) Then
        ResetOrder
    End If

    EnableControls NewTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.tabInfo_Switch"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtCommandLine_Change
'' Description: As the command line changes, set the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtCommandLine_Change()
On Error GoTo ErrSection:

    If m.bUpdating = False Then
        ParseCommandLine
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.txtCommandLine_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoRefresh
'' Description: Refresh either all accounts or the selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoRefresh()
On Error GoTo ErrSection:

    If Len(Trim(cboAccount.Text)) > 0 Then
        RefreshAccount Parse(cboAccount.Text, "-", 1)
    End If
    
    EnableControls
    FilterGrids

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.DoRefresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateCommandLine
'' Description: Update the command line from the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateCommandLine()
On Error GoTo ErrSection:

    Dim astrCommand As New cGdArray     ' Command line for a side of the spread
    Dim lIndex As Long                  ' Index into a for loop
    
    m.bUpdating = True
    
    txtCommandLine.Text = ""
    For lIndex = 0 To 2
        astrCommand.Clear
        astrCommand(0) = cboBuySell(lIndex).Text
        astrCommand(1) = txtQuantity(lIndex).Text
        If cboCommodity(lIndex).ListIndex >= 0 Then
            astrCommand(2) = Trim(Parse(cboCommodity(lIndex).Text, "-", 1))
            If cboMonth(lIndex).ListIndex >= 0 Then
                astrCommand(2) = astrCommand(2) & MonthToCode(cboMonth(lIndex).ListIndex + 1)
                If cboYear(lIndex).ListIndex >= 0 Then
                    astrCommand(2) = astrCommand(2) & Right(cboYear(lIndex).Text, 1)
                End If
            End If
        End If
        If txtStrike(lIndex).Visible = True Then
            If cboPutCall(lIndex).ListIndex >= 0 Then
                astrCommand(3) = txtStrike(lIndex).Text & Left(cboPutCall(lIndex).Text, 1)
            Else
                astrCommand(3) = txtStrike(lIndex).Text
            End If
            astrCommand(4) = txtPrice(lIndex).Text
            astrCommand(5) = cboOpenClose(lIndex).Text
            If cboType(lIndex).Visible = True Then
                astrCommand(6) = cboType(lIndex).Text
                If txtLimit(lIndex).Visible = True Then
                    astrCommand(7) = txtLimit(lIndex).Text
                End If
            End If
        Else
            astrCommand(3) = txtPrice(lIndex).Text
            If cboType(lIndex).Visible = True Then
                astrCommand(4) = cboType(lIndex).Text
                If txtLimit(lIndex).Visible = True Then
                    astrCommand(5) = txtLimit(lIndex).Text
                End If
            End If
        End If
        
        txtCommandLine.Text = txtCommandLine.Text & Trim(astrCommand.JoinFields(" "))
        Select Case lIndex
            Case 0:
                If cboBuySell(1).ListIndex > -1& Then txtCommandLine.Text = txtCommandLine.Text & " / "
            Case 1:
                If cboBuySell(2).ListIndex > -1& Then txtCommandLine.Text = txtCommandLine.Text & " / "
        End Select
    Next lIndex
    
ErrExit:
    m.bUpdating = False
    Exit Sub
    
ErrSection:
    m.bUpdating = False
    RaiseError "frmBrokerScreen.UpdateCommandLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearSide
'' Description: Clear the controls for a given side of the spread
'' Inputs:      Side of the Spread
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearSide(ByVal lIndex As Long)
On Error GoTo ErrSection:

    SetBuySell lIndex, ""
    txtQuantity(lIndex).Text = ""
    SetCommodity lIndex, ""
    txtPrice(lIndex).Text = ""
    txtStrike(lIndex).Text = ""
    cboPutCall(lIndex).ListIndex = -1&
    cboOpenClose(lIndex).ListIndex = -1&
    cboType(lIndex).ListIndex = -1&
    txtLimit(lIndex).Visible = False
    txtLimit(lIndex).Text = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.ClearSide"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtLimit_Change
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLimit_Change(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.txtLimit_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPrice_Change
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPrice_Change(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        If (Len(Trim(txtStrike(Index).Text)) = 0) And (Len(Trim(txtPrice(Index).Text)) > 0) Then
            txtStrike(Index).Visible = False
            cboPutCall(Index).Visible = False
            cboOpenClose(Index).Visible = False
        Else
            txtStrike(Index).Visible = True
            cboPutCall(Index).Visible = True
            cboOpenClose(Index).Visible = True
        End If
        
        If Len(Trim(txtPrice(Index).Text)) > 0 Then
            If UCase(Left(Trim(txtPrice(Index).Text), 1)) = "M" Then
                cboType(Index).Visible = False
                txtLimit(Index).Visible = False
            Else
                cboType(Index).Visible = True
                txtLimit(Index).Visible = (UCase(cboType(Index).Text) = "STWL")
            End If
        Else
            cboType(Index).Visible = True
            txtLimit(Index).Visible = (UCase(cboType(Index).Text) = "STWL")
        End If
        
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.txtPrice_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtQuantity_Change
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtQuantity_Change(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.txtQuantity_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStrike_Change
'' Description: Update the command line if the user changes the control
'' Inputs:      Control Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStrike_Change(Index As Integer)
On Error GoTo ErrSection:

    If m.bParsing = False Then
        UpdateCommandLine
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.txtStrike_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromControls
'' Description: Fill a gdTable from the order controls
'' Inputs:      None
'' Returns:     Filled in Table
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderFromControls() As cGdTable
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tblOrder As New cGdTable        ' Table to return from the function
    Dim lTo As Long                     ' To variable for a for loop

    With tblOrder
        .CreateField eGDARRAY_Strings, 0, "TraderAccount"   ' Text 20
        .CreateField eGDARRAY_Strings, 1, "OrderType"       ' Text 10
        .CreateField eGDARRAY_Strings, 2, "ExchangeName"    ' Text 10
        .CreateField eGDARRAY_Strings, 3, "SecurityName"    ' Text 10
        .CreateField eGDARRAY_Strings, 4, "ContractDate"    ' Text 50
        .CreateField eGDARRAY_Strings, 5, "Session"         ' 'R' or 'E'
        .CreateField eGDARRAY_Strings, 6, "BuyOrSell"       ' 'B' or 'S'
        .CreateField eGDARRAY_Strings, 7, "Price"           ' Text 20
        .CreateField eGDARRAY_Strings, 8, "Price2"          ' Text 20
        .CreateField eGDARRAY_Strings, 9, "Lots"            '
        .CreateField eGDARRAY_Strings, 10, "OpenOrClose"    ' 'O' or 'C'
        .CreateField eGDARRAY_Strings, 11, "XRefPersist"    '
        .CreateField eGDARRAY_Strings, 12, "GoodTillDate"   ' MMDDYYYY
        .CreateField eGDARRAY_Strings, 13, "StrikePrice"
        .CreateField eGDARRAY_Strings, 14, "PutCall"
        .CreateField eGDARRAY_Strings, 15, "NewOffset"      ' 'N'ew or 'O'ffset
        .NumRecords = 0
        
        If (chkSpread.Value = vbUnchecked) Then
            lTo = 0
        Else
            lTo = 2
        End If
        
        For lIndex = 0 To lTo
            If (lIndex = 0) Or (cboBuySell(lIndex).ListIndex >= 0) Then
                .NumRecords = .NumRecords + 1
                
                If cboAccount.ListIndex >= 0 Then
                    .Item(0, lIndex) = Parse(cboAccount.Text, "-", 1)
                Else
                    .Item(0, lIndex) = ""
                End If
                If cboCommodity(lIndex).ListIndex >= 0 Then
                    .Item(2, lIndex) = Parse(cboCommodity(lIndex).Text, "-", 3)
                    .Item(3, lIndex) = Parse(cboCommodity(lIndex).Text, "-", 1)
                Else
                    .Item(2, lIndex) = ""
                    .Item(3, lIndex) = ""
                End If
                If (cboYear(lIndex).ListIndex >= 0) And (cboMonth(lIndex).ListIndex >= 0) Then
                    .Item(4, lIndex) = cboYear(lIndex).Text & Format(cboMonth(lIndex).ListIndex + 1, "00")
                Else
                    .Item(4, lIndex) = ""
                End If
                If cboSession.ListIndex >= 0 Then
                    If cboSession.Text = "Day" Then .Item(5, lIndex) = "R" Else .Item(5, lIndex) = "E"
                Else
                    .Item(5, lIndex) = ""
                End If
                If cboBuySell(lIndex).ListIndex >= 0 Then
                    If cboBuySell(lIndex).Text = "Buy" Then .Item(6, lIndex) = "B" Else .Item(6, lIndex) = "S"
                Else
                    .Item(6, lIndex) = ""
                End If
                .Item(9, lIndex) = txtQuantity(lIndex).Text
                Select Case UCase(cboGoodThru.Text)
                    Case "TODAY"
                        .Item(12, lIndex) = "DAY"
                    Case "CANCEL"
                        .Item(12, lIndex) = "GTC"
                    Case "DATE"
                        .Item(12, lIndex) = Format(gdThruDate.Value, "MMDDYYYY")
                End Select
                .Item(13, lIndex) = txtStrike(lIndex).Text
                If cboPutCall(lIndex).ListIndex >= 0 Then
                    If cboPutCall(lIndex).Text = "Put" Then .Item(14, lIndex) = "P" Else .Item(14, lIndex) = "C"
                Else
                    .Item(14, lIndex) = ""
                End If
                If cboOpenClose(lIndex).ListIndex >= 0 Then
                    If cboOpenClose(lIndex).Text = "Open" Then .Item(15, lIndex) = "N" Else .Item(16, lIndex) = "O"
                Else
                    .Item(15, lIndex) = ""
                End If
                    
                If UCase(Trim(txtPrice(lIndex).Text)) = "MARKET" Then
                    .Item(1, lIndex) = "Market"
                    .Item(7, lIndex) = ""
                    .Item(8, lIndex) = ""
                Else
                    .Item(1, lIndex) = OrderTypeFromDisplay(cboType(lIndex).Text)
                    .Item(7, lIndex) = txtPrice(lIndex).Text
                    .Item(8, lIndex) = txtLimit(lIndex).Text
                End If
            End If
        Next lIndex
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerScreen.OrderFromControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeFromDisplay
'' Description: Convert an order type from the combo to what the DLL expects
'' Inputs:      Display Order Type
'' Returns:     DLL Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderTypeFromDisplay(ByVal strOrderType As String) As String
On Error GoTo ErrSection:

    Select Case UCase(strOrderType)
        Case ""
            OrderTypeFromDisplay = "Limit"
        Case "FOK"
            OrderTypeFromDisplay = "FillOrKill"
        Case "MIT"
            OrderTypeFromDisplay = "MarketIfTouched"
        Case "MOC"
            OrderTypeFromDisplay = "MarketOnClose"
        Case "OB"
            OrderTypeFromDisplay = "OrBetter"
        Case "OBOO"
            OrderTypeFromDisplay = "OrBetterOnOpen"
        Case "OO"
            OrderTypeFromDisplay = "MarketOnOpen"
        Case "SCO"
            OrderTypeFromDisplay = "StopCloseOnly"
        Case "SLCO"
            OrderTypeFromDisplay = "StopLimitCloseOnly"
        Case "SOO"
            OrderTypeFromDisplay = "StopOnOpen"
        Case "STL"
            OrderTypeFromDisplay = "StopAndLimit"
        Case "STOP"
            OrderTypeFromDisplay = "Stop"
        Case "STWL"
            OrderTypeFromDisplay = "Stop with Limit"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerScreen.OrderTypeFromDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeToDisplay
'' Description: Convert an order type from what the DLL expects to the display
'' Inputs:      DLL Order Type
'' Returns:     Display Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderTypeToDisplay(ByVal strOrderType As String) As String
On Error GoTo ErrSection:

    Select Case UCase(strOrderType)
        Case "LIMIT", ""
            OrderTypeToDisplay = ""
        Case "FILLORKILL"
            OrderTypeToDisplay = "FOK"
        Case "MARKETIFTOUCHED"
            OrderTypeToDisplay = "MIT"
        Case "MARKETONCLOSE"
            OrderTypeToDisplay = "MOC"
        Case "ORBETTER"
            OrderTypeToDisplay = "OB"
        Case "ORBETTERONOPEN"
            OrderTypeToDisplay = "OBOO"
        Case "MARKETONOPEN"
            OrderTypeToDisplay = "OO"
        Case "STOPCLOSEONLY"
            OrderTypeToDisplay = "SCO"
        Case "STOPLIMITCLOSEONLY"
            OrderTypeToDisplay = "SLCO"
        Case "STOPONOPEN"
            OrderTypeToDisplay = "SOO"
        Case "STOPANDLIMIT"
            OrderTypeToDisplay = "STL"
        Case "STOP"
            OrderTypeToDisplay = "STOP"
        Case "STOP WITH LIMIT"
            OrderTypeToDisplay = "STWL"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerScreen.OrderTypeToDisplay"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Ask the Alaron servers for the list of accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAccounts()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    m.bAccountsLoaded = False
    m.astrAccounts.Size = 0
    
    g.Alaron.GetAccounts False
    
    Do While (m.bAccountsLoaded = False) And (lTimeOut < 30&)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.GetAccounts"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountCallback
'' Description: Account received from the Alaron servers
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AccountCallback(ByVal strAccount As String)
On Error GoTo ErrSection:

    If UCase(strAccount) = "BEGIN" Then
        m.bAccountsLoaded = False
    ElseIf UCase(strAccount) = "END" Then
        m.bAccountsLoaded = True
    Else
        m.astrAccounts.Add strAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AccountCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsCombo
'' Description: Load the accounts combo from the accounts list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadAccountsCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSelection As String          ' Selection from the combo box
    Dim strAccount As String            ' Account number
    Dim strUserName As String           ' User name

    strSelection = Parse(cboAccount.Text, "-", 1)
    cboAccount.Clear
    For lIndex = 0 To m.astrAccounts.Size - 1
        strAccount = Parse(m.astrAccounts(lIndex), vbTab, 1)
        strUserName = Parse(m.astrAccounts(lIndex), vbTab, 3)
        
        If Len(strAccount) > 0 Then
            If Len(strUserName) > 0 Then
                cboAccount.AddItem strAccount & " - " & strUserName
            Else
                cboAccount.AddItem strAccount
            End If
        End If
    Next lIndex
    
    SelectAccount strSelection

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.LoadAccountsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountStatus
'' Description: Ask the Alaron servers for the account status for the account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAccountStatus(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    m.bAccountStatusLoaded = False
    m.astrAcctStatus.Size = 0
    
    g.Alaron.GetAccountStatus strAccount
    
    Do While (m.bAccountStatusLoaded = False) And (lTimeOut < 30&)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.GetAccountStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountStatusCallback
'' Description: AccountStatus received from the Alaron servers
'' Inputs:      AccountStatus
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AccountStatusCallback(ByVal strAccountStatus As String)
On Error GoTo ErrSection:

    m.astrAcctStatus.Add strAccountStatus
    m.bAccountStatusLoaded = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AccountStatusCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountStatusGrid
'' Description: Load the account status grid based on the current account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAccountStatusGrid()
On Error GoTo ErrSection:

    Dim strAccount As String            ' Currently selected account number
    Dim lIndex As Long                  ' Index into a for loop

    If optSpecificAccount.Value = True And Len(cboAccount.Text) > 0 Then
        strAccount = Parse(cboAccount.Text, "-", 1)
        
        For lIndex = 0 To m.astrAcctStatus.Size - 1
            If Parse(m.astrAcctStatus(lIndex), vbTab, 1) = strAccount Then
                With fgAccountStatus
                    AccountStatusCell 2, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 6)
                    AccountStatusCell 2, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 17)
                    AccountStatusCell 3, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 11)
                    AccountStatusCell 3, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 18)
                    AccountStatusCell 5, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 7)
                    AccountStatusCell 5, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 16)
                    AccountStatusCell 8, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 2)
                    AccountStatusCell 8, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 8)
                    AccountStatusCell 9, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 3)
                    AccountStatusCell 9, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 9)
                    AccountStatusCell 10, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 4)
                    AccountStatusCell 10, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 10)
                    AccountStatusCell 11, 4, ""
                    AccountStatusCell 12, 1, Parse(m.astrAcctStatus(lIndex), vbTab, 15)
                    AccountStatusCell 12, 4, Parse(m.astrAcctStatus(lIndex), vbTab, 12)
                    AccountStatusCell 13, 4, ""
                End With
                
                Exit For
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.LoadAccountStatusGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Ask the Alaron servers for the orders for the account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetOrders(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    m.bLoadingOrders = True
    m.bOrdersLoaded = False
    
    m.astrOrders.Size = 0
    
    g.Alaron.GetOrders strAccount
    
    Do While (m.bOrdersLoaded = False) And (lTimeOut < 30&)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop
    
    m.bLoadingOrders = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.GetOrders"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderCallback
'' Description: Order received from the Alaron servers
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderCallback(ByVal strOrder As String)
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Did the order already exist?

    If UCase(strOrder) = "BEGIN" Then
        m.bOrdersLoaded = False
    ElseIf UCase(strOrder) = "END" Then
        m.bOrdersLoaded = True
    Else
        astrOrder.SplitFields strOrder, vbTab
        
        If astrOrder(17) <> "V001" Then
            If optAllAccounts Or (astrOrder(4) = Parse(cboAccount.Text, "-", 1)) Then
                bFound = False
                For lIndex = 0 To m.astrOrders.Size - 1
                    If Parse(m.astrOrders(lIndex), vbTab, 2) = astrOrder(1) Then
                        bFound = True
                        m.astrOrders(lIndex) = strOrder
                    End If
                Next lIndex
            End If
            
            If bFound = False Then m.astrOrders.Add strOrder
            
            If m.bLoadingOrders = False Then LoadOrdersGrids
        Else
            If UCase(astrOrder(14)) = "REJECTED" Then
                If Len(astrOrder(39)) > 0 Then
                    lblError.Caption = astrOrder(39)
                End If
            Else
                lblError.Caption = "Order OK"
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.OrderCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrdersGrids
'' Description: Load the orders grids form the order information table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOrdersGrids()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
            
    fgWorkingOrders.Redraw = flexRDNone
    fgParkedOrders.Redraw = flexRDNone
    fgFilledOrders.Redraw = flexRDNone
    
    fgWorkingOrders.Rows = fgWorkingOrders.FixedRows
    fgParkedOrders.Rows = fgParkedOrders.FixedRows
    fgFilledOrders.Rows = fgFilledOrders.FixedRows
    
    For lIndex = 0 To m.astrOrders.Size - 1
        If UCase(Parse(m.astrOrders(lIndex), vbTab, 30)) = "P" Then
            AddParkedOrder lIndex
        ElseIf UCase(Parse(m.astrOrders(lIndex), vbTab, 37)) = "F" Then
            AddFilledOrder lIndex
        Else
            AddWorkingOrder lIndex
        End If
    Next lIndex
    
    fgWorkingOrders.AutoSize 0, fgWorkingOrders.Cols - 1, False, 75
    fgParkedOrders.AutoSize 0, fgParkedOrders.Cols - 1, False, 75
    fgFilledOrders.AutoSize 0, fgFilledOrders.Cols - 1, False, 75
    
    If fgWorkingOrders.Rows > fgWorkingOrders.FixedRows Then
        fgWorkingOrders.Col = WorkingOrdersCol(eGDWorkingCol_OrderNumber)
        fgWorkingOrders.Sort = flexSortGenericAscending
    End If
    
    If fgParkedOrders.Rows > fgParkedOrders.FixedRows Then
        fgParkedOrders.Col = WorkingOrdersCol(eGDWorkingCol_OrderNumber)
        fgParkedOrders.Sort = flexSortGenericAscending
    End If
    
    If fgFilledOrders.Rows > fgFilledOrders.FixedRows Then
        fgFilledOrders.Col = FilledOrdersCol(eGDFilledCol_OrderNumber)
        fgFilledOrders.Sort = flexSortGenericAscending
    End If
    
    fgWorkingOrders.Redraw = flexRDBuffered
    fgParkedOrders.Redraw = flexRDBuffered
    fgFilledOrders.Redraw = flexRDBuffered

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.LoadOrdersGrids"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddWorkingOrder
'' Description: Add a working order from the table to the grid
'' Inputs:      Table Index, Grid Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddWorkingOrder(ByVal lTableIndex As Long, Optional ByVal lGridRow As Long = -1&)
On Error GoTo ErrSection:

    With fgWorkingOrders
        If lGridRow = -1& Then
            .Rows = .Rows + 1
            lGridRow = .Rows - 1
        End If
        
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Account)) = Parse(m.astrOrders(lTableIndex), vbTab, 5)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OrderNumber)) = Parse(m.astrOrders(lTableIndex), vbTab, 2)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_ReferenceID)) = Parse(m.astrOrders(lTableIndex), vbTab, 3)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OrderDate)) = Parse(m.astrOrders(lTableIndex), vbTab, 16) & " " & Parse(m.astrOrders(lTableIndex), vbTab, 17)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_BuySell)) = Parse(m.astrOrders(lTableIndex), vbTab, 9)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Quantity)) = Parse(m.astrOrders(lTableIndex), vbTab, 12)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Symbol)) = Parse(m.astrOrders(lTableIndex), vbTab, 32)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Strike)) = Parse(m.astrOrders(lTableIndex), vbTab, 29)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Price)) = Parse(m.astrOrders(lTableIndex), vbTab, 10)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OC)) = ""
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OrderType)) = Parse(m.astrOrders(lTableIndex), vbTab, 25)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Limit)) = Parse(m.astrOrders(lTableIndex), vbTab, 11)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Tic)) = Parse(m.astrOrders(lTableIndex), vbTab, 31)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_GoodThru)) = Parse(m.astrOrders(lTableIndex), vbTab, 19)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Session)) = Parse(m.astrOrders(lTableIndex), vbTab, 28)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_CreditControl)) = Parse(m.astrOrders(lTableIndex), vbTab, 21)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_ExchangeMessage)) = ""
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_SpecialInstruction)) = ""
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_EnteredBy)) = Parse(m.astrOrders(lTableIndex), vbTab, 4)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Salesman)) = Parse(m.astrOrders(lTableIndex), vbTab, 27)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Exchange)) = Parse(m.astrOrders(lTableIndex), vbTab, 6)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_BaseSymbol)) = Parse(m.astrOrders(lTableIndex), vbTab, 7)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Contract)) = Parse(m.astrOrders(lTableIndex), vbTab, 8)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AddWorkingOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddParkedOrder
'' Description: Add a parked order from the table to the grid
'' Inputs:      Table Index, Grid Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddParkedOrder(ByVal lTableIndex As Long, Optional ByVal lGridRow As Long = -1&)
On Error GoTo ErrSection:

    With fgParkedOrders
        If lGridRow = -1& Then
            .Rows = .Rows + 1
            lGridRow = .Rows - 1
        End If
        
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Account)) = Parse(m.astrOrders(lTableIndex), vbTab, 5)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OrderNumber)) = Parse(m.astrOrders(lTableIndex), vbTab, 2)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_ReferenceID)) = Parse(m.astrOrders(lTableIndex), vbTab, 3)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OrderDate)) = Parse(m.astrOrders(lTableIndex), vbTab, 16) & " " & Parse(m.astrOrders(lTableIndex), vbTab, 17)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_BuySell)) = Parse(m.astrOrders(lTableIndex), vbTab, 9)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Quantity)) = Parse(m.astrOrders(lTableIndex), vbTab, 12)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Symbol)) = Parse(m.astrOrders(lTableIndex), vbTab, 32)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Strike)) = Parse(m.astrOrders(lTableIndex), vbTab, 29)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Price)) = Parse(m.astrOrders(lTableIndex), vbTab, 10)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OC)) = ""
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_OrderType)) = Parse(m.astrOrders(lTableIndex), vbTab, 25)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Limit)) = Parse(m.astrOrders(lTableIndex), vbTab, 11)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Tic)) = Parse(m.astrOrders(lTableIndex), vbTab, 31)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_GoodThru)) = Parse(m.astrOrders(lTableIndex), vbTab, 19)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Session)) = Parse(m.astrOrders(lTableIndex), vbTab, 28)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_CreditControl)) = Parse(m.astrOrders(lTableIndex), vbTab, 21)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_ExchangeMessage)) = ""
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_SpecialInstruction)) = ""
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_EnteredBy)) = Parse(m.astrOrders(lTableIndex), vbTab, 4)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Salesman)) = Parse(m.astrOrders(lTableIndex), vbTab, 27)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Exchange)) = Parse(m.astrOrders(lTableIndex), vbTab, 6)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_BaseSymbol)) = Parse(m.astrOrders(lTableIndex), vbTab, 7)
        .TextMatrix(lGridRow, WorkingOrdersCol(eGDWorkingCol_Contract)) = Parse(m.astrOrders(lTableIndex), vbTab, 8)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AddParkedOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFilledOrder
'' Description: Add a Filled order from the table to the grid
'' Inputs:      Table Index, Grid Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFilledOrder(ByVal lTableIndex As Long, Optional ByVal lGridRow As Long = -1&)
On Error GoTo ErrSection:

    With fgFilledOrders
        If lGridRow = -1& Then
            .Rows = .Rows + 1
            lGridRow = .Rows - 1
        End If
            
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Account)) = Parse(m.astrOrders(lTableIndex), vbTab, 5)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_OrderNumber)) = Parse(m.astrOrders(lTableIndex), vbTab, 2)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_TradeDate)) = Parse(m.astrOrders(lTableIndex), vbTab, 34)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Position)) = Parse(m.astrOrders(lTableIndex), vbTab, 9)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Quantity)) = Parse(m.astrOrders(lTableIndex), vbTab, 14)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Symbol)) = Parse(m.astrOrders(lTableIndex), vbTab, 32)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Strike)) = Parse(m.astrOrders(lTableIndex), vbTab, 29)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_TradePrice)) = Parse(m.astrOrders(lTableIndex), vbTab, 36)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Tic)) = Parse(m.astrOrders(lTableIndex), vbTab, 31)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_FuturesProfit)) = Parse(m.astrOrders(lTableIndex), vbTab, 33)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_OptionsValue)) = ""
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_Currency)) = Parse(m.astrOrders(lTableIndex), vbTab, 22)
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_OrderTime)) = ""
        .TextMatrix(lGridRow, FilledOrdersCol(eGDFilledCol_FillTime)) = Parse(m.astrOrders(lTableIndex), vbTab, 34) & " " & Parse(m.astrOrders(lTableIndex), vbTab, 35)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AddFilledOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Ask the Alaron servers for the positions for the account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetPositions(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    m.bPositionsLoaded = False
    m.astrPositions.Size = 0
    
    g.Alaron.GetPositions strAccount
    
    Do While (m.bPositionsLoaded = False) And (lTimeOut < 30&)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.GetPositions"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionCallback
'' Description: Position received from the Alaron servers
'' Inputs:      Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PositionCallback(ByVal strPosition As String)
On Error GoTo ErrSection:

    If UCase(strPosition) = "BEGIN" Then
        m.bPositionsLoaded = False
    ElseIf UCase(strPosition) = "END" Then
        m.bPositionsLoaded = True
    Else
        m.astrPositions.Add strPosition
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.PositionCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOpenPositionsGrid
'' Description: Load the open positions grid from the positions table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOpenPositionsGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgPositions
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        For lIndex = 0 To m.astrPositions.Size - 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Account)) = Parse(m.astrPositions(lIndex), vbTab, 1)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_TradeDate)) = Parse(m.astrPositions(lIndex), vbTab, 15)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Position)) = Parse(m.astrPositions(lIndex), vbTab, 4)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Quantity)) = Parse(m.astrPositions(lIndex), vbTab, 13)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Symbol)) = Parse(m.astrPositions(lIndex), vbTab, 11)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Strike)) = Parse(m.astrPositions(lIndex), vbTab, 9)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_TradePrice)) = Parse(m.astrPositions(lIndex), vbTab, 14)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Tic)) = Parse(m.astrPositions(lIndex), vbTab, 10)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_FuturesProfit)) = Parse(m.astrPositions(lIndex), vbTab, 12)
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_OptionValue)) = ""
            .TextMatrix(.Rows - 1, PositionsCol(eGDPositionCol_Currency)) = Parse(m.astrPositions(lIndex), vbTab, 6)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.LoadOpenPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSecurities
'' Description: Ask the Alaron servers for the security list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSecurities()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    m.bSecuritiesLoaded = False
    m.astrSecurities.Size = 0
    
    g.Alaron.GetSecurityList g.Alaron.UserName
    
    Do While (m.bSecuritiesLoaded = False) And (lTimeOut < 30&)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop
    
    If m.astrSecurities.Size = 0 Then
        m.astrSecurities.FromFile AddSlash(App.Path) & "Provided\AlrnSyms.TXT"
        m.bSecuritiesLoaded = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.GetSecutities"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SecurityCallback
'' Description: Security received from the Alaron servers
'' Inputs:      Security
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SecurityCallback(ByVal strSecurity As String)
On Error GoTo ErrSection:

    If UCase(strSecurity) = "BEGIN" Then
        m.bSecuritiesLoaded = False
    ElseIf UCase(strSecurity) = "END" Then
        m.bSecuritiesLoaded = True
    Else
        m.astrSecurities.Add strSecurity
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SecurityCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSecuritiesCombo
'' Description: Load the securities combo from the securities list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSecuritiesCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    cboCommodity(0).Clear
    cboCommodity(1).Clear
    cboCommodity(2).Clear

    For lIndex = 0 To m.astrSecurities.Size - 1
        If Parse(m.astrSecurities(lIndex), vbTab, 5) <> "O" Then
            cboCommodity(0).AddItem Parse(m.astrSecurities(lIndex), vbTab, 1) & " - " & Parse(m.astrSecurities(lIndex), vbTab, 3) & " - " & Parse(m.astrSecurities(lIndex), vbTab, 4)
            cboCommodity(1).AddItem Parse(m.astrSecurities(lIndex), vbTab, 1) & " - " & Parse(m.astrSecurities(lIndex), vbTab, 3) & " - " & Parse(m.astrSecurities(lIndex), vbTab, 4)
            cboCommodity(2).AddItem Parse(m.astrSecurities(lIndex), vbTab, 1) & " - " & Parse(m.astrSecurities(lIndex), vbTab, 3) & " - " & Parse(m.astrSecurities(lIndex), vbTab, 4)
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.LoadSecuritiesCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetStatement
'' Description: Ask the Alaron servers for the statement for the account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetStatement(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    m.bStatementLoaded = False
    m.astrStatement.Size = 0
    
    g.Alaron.GetStatement strAccount
    
    Do While (m.bStatementLoaded = False) And (lTimeOut < 30&)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerScreen.GetStatement"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StatementCallback
'' Description: Statement information received from the Alaron servers
'' Inputs:      Statement Line
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StatementCallback(ByVal astrStatement As cGdArray)
On Error GoTo ErrSection:

    Set m.astrStatement = astrStatement.MakeCopy
    m.bStatementLoaded = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.StatementCallback"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStatementGrid
'' Description: Load the statement grid from the table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadStatementGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgStatement
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        For lIndex = 0 To m.astrStatement.Size - 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = m.astrStatement(lIndex)
        Next lIndex
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.LoadStatementGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectSymbolInCombo
'' Description: Select the given symbol in the combo box (if found)
'' Inputs:      Base Symbol, Exchange, Combo Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectSymbolInCombo(ByVal strBaseSym As String, ByVal strExchange As String, ByVal lComboIndex As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To cboCommodity(lComboIndex).ListCount - 1
        If Parse(cboCommodity(lComboIndex).List(lIndex), " - ", 1) = strBaseSym Then
            If Parse(cboCommodity(lComboIndex).List(lIndex), " - ", 3) = strExchange Then
                cboCommodity(lComboIndex).ListIndex = lIndex
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SelectSymbolInCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlaronOrderFromControls
'' Description: Put together an Alaron order string from the controls
'' Inputs:      Control Index, Order Verification?
'' Returns:     Alaron Order String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AlaronOrderFromControls(ByVal lIndex As Long, ByVal bVerify As Boolean) As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order information to join and return
    
    astrOrder.Create eGDARRAY_Strings
    
    astrOrder(0) = Parse(cboAccount.Text, "-", 1)
    astrOrder(2) = Parse(cboCommodity(lIndex).Text, " - ", 3)
    astrOrder(3) = Parse(cboCommodity(lIndex).Text, " - ", 1)
    astrOrder(4) = cboYear(lIndex).Text & Format(cboMonth(lIndex).ListIndex + 1, "00")
    If cboSession.Text = "Day" Then astrOrder(5) = "P" Else astrOrder(5) = "E"
    If cboBuySell(lIndex).Text = "Buy" Then astrOrder(6) = "B" Else astrOrder(6) = "S"
    If UCase(Trim(txtPrice(lIndex).Text)) = "MARKET" Then
        astrOrder(1) = "Market"
        astrOrder(7) = ""
        astrOrder(8) = ""
    Else
        astrOrder(1) = OrderTypeFromDisplay(cboType(lIndex).Text)
        astrOrder(7) = txtPrice(lIndex).Text
        astrOrder(8) = txtLimit(lIndex).Text
    End If
    astrOrder(9) = txtQuantity(lIndex).Text
    astrOrder(10) = ""                  ' Open or Close
    If bVerify Then astrOrder(11) = "V001" Else astrOrder(11) = "NONE"
    Select Case UCase(cboGoodThru.Text)
        Case "TODAY"
            astrOrder(12) = "DAY"
        Case "CANCEL"
            astrOrder(12) = "GTC"
        Case "DATE"
            astrOrder(12) = Format(gdThruDate.Value, "MMDDYYYY")
    End Select
    astrOrder(13) = txtStrike(lIndex).Text
    If cboPutCall(lIndex).Text = "Put" Then astrOrder(14) = "P" Else astrOrder(14) = "C"
    If cboOpenClose(lIndex).Text = "Open" Then astrOrder(15) = "N" Else astrOrder(15) = "O"
    astrOrder(16) = Trim(txtOrderID.Text)
    
    AlaronOrderFromControls = astrOrder.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerScreen.AlaronOrderFromControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlaronOrderToControls
'' Description: Fill in the controls with the given Alaron order string
'' Inputs:      Alaron Order String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AlaronOrderToControls(ByVal strOrder As String)
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order information to join and return
    
    astrOrder.SplitFields strOrder, vbTab
    
    If astrOrder(6) = "B" Then
        cboBuySell(0).Text = "Buy"
    ElseIf astrOrder(6) = "S" Then
        cboBuySell(0).Text = "Sell"
    Else
        cboBuySell(0).ListIndex = -1&
    End If
    txtQuantity(0).Text = astrOrder(9)
    SelectSymbolInCombo astrOrder(3), astrOrder(2), 0
    If Len(astrOrder(4)) > 2 Then
        cboMonth(0).ListIndex = CLng(Val(Right(astrOrder(4), 2))) - 1
        cboYear(0).Text = Left(astrOrder(4), Len(astrOrder(4)) - 2)
    End If
    txtStrike(0).Text = astrOrder(13)
    If astrOrder(14) = "P" Then
        cboPutCall(0).Text = "Put"
    ElseIf astrOrder(14) = "C" Then
        cboPutCall(0).Text = "Call"
    Else
        cboPutCall(0).ListIndex = -1&
    End If
    If astrOrder(15) = "N" Then
        cboPutCall(0).Text = "Open"
    ElseIf astrOrder(15) = "O" Then
        cboPutCall(0).Text = "Close"
    Else
        cboPutCall(0).ListIndex = -1&
    End If
    If Len(Trim(astrOrder(7))) = 0 Then
        cboType(0).ListIndex = -1&
        txtPrice(0).Text = "Market"
        txtLimit(0).Text = ""
    Else
        If astrOrder(1) = "Limit" Or astrOrder(1) = "" Then
            cboType(0).ListIndex = -1&
        Else
            cboType(0).Text = OrderTypeToDisplay(astrOrder(1))
        End If
        txtPrice(0).Text = astrOrder(7)
        txtLimit(0).Text = astrOrder(8)
    End If
    If astrOrder(5) = "P" Then cboSession.Text = "Day" Else cboSession.Text = "Electronic"
    Select Case astrOrder(12)
        Case "DAY", ""
            cboGoodThru.Text = "Today"
        Case "GTC"
            cboGoodThru.Text = "Cancel"
        Case Else
            cboGoodThru.Text = "Date"
            If Len(astrOrder(12)) = 8 Then
                gdThruDate.YYYYMMDD = Right(astrOrder(12), 4) & Left(astrOrder(12), 2) & Mid(astrOrder(12), 3, 2)
            End If
    End Select
    
    lblOrderID.Visible = True
    txtOrderID.Visible = True
    txtOrderID.Text = astrOrder(16)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.AlaronOrderToControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlaronOrderFromGrid
'' Description: Put together an Alaron order string from the grid
'' Inputs:      Grid
'' Returns:     Alaron Order String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AlaronOrderFromGrid(fg As VSFlexGrid, ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Array of order information to join together
    
    astrOrder.Create eGDARRAY_Strings
    
    With fg
        astrOrder(0) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Account))
        astrOrder(1) = OrderTypeFromDisplay(.TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_OrderType)))
        astrOrder(2) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Exchange))
        astrOrder(3) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_BaseSymbol))
        astrOrder(4) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Contract))
        astrOrder(5) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Session))
        If astrOrder(5) = "R" Then astrOrder(5) = "P"
        astrOrder(6) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_BuySell))
        astrOrder(7) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Price))
        astrOrder(8) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Limit))
        astrOrder(9) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Quantity))
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_GoodThru))
        If Len(.TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Strike))) > 0 Then
            astrOrder(13) = Left(.TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Strike)), Len(.TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Strike))) - 1)
            astrOrder(14) = Right(.TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_Strike)), 1)
        Else
            astrOrder(13) = ""
            astrOrder(14) = ""
        End If
        astrOrder(15) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_OC))
        astrOrder(16) = .TextMatrix(lRow, WorkingOrdersCol(eGDWorkingCol_OrderNumber))
    End With
    
    AlaronOrderFromGrid = astrOrder.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerScreen.AlaronOrderFromGrid"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyOrder
'' Description: Verify the order information in the appropriate controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyOrder()
On Error GoTo ErrSection:

    Dim strOrder As String              ' Order string to pass to the servers
    
    strOrder = AlaronOrderFromControls(0, True)
    
    If Len(Trim(txtOrderID.Text)) = 0 Then
        g.Alaron.DumpDebug "Verifying Add Order from Broker Screen: " & strOrder
        g.Alaron.AddBrokerOrder strOrder, eGDRanOrderMode_Validate
    Else
        g.Alaron.DumpDebug "Verifying Amend Order from Broker Screen: " & strOrder
        g.Alaron.AmendBrokerOrder strOrder, eGDRanOrderMode_Validate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.VerifyOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrder
'' Description: Submit an order to the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOrder()
On Error GoTo ErrSection:

    Dim strOrder As String              ' Order string to submit to server

    Select Case tabInfo.CurrTab
        Case Tabs(eGDTab_NewOrder)
            strOrder = AlaronOrderFromControls(0, False)
            
        Case Tabs(eGDTab_ParkedOrders)
            strOrder = AlaronOrderFromGrid(fgParkedOrders, fgParkedOrders.Row)
            
    End Select
    
    If Len(strOrder) > 0 Then
        If Len(Trim(txtOrderID.Text)) = 0 Then
            g.Alaron.DumpDebug "Submit Add Order from Broker Screen: " & strOrder
            g.Alaron.AddBrokerOrder strOrder, eGDRanOrderMode_Submit
        Else
            g.Alaron.DumpDebug "Submit Amend Order from Broker Screen: " & strOrder
            g.Alaron.AmendBrokerOrder strOrder, eGDRanOrderMode_Submit
        End If
        tabInfo.CurrTab = Tabs(eGDTab_WorkingOrders)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.SubmitOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Park an order on the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkOrder()
On Error GoTo ErrSection:

    Dim strOrder As String              ' Order string to submit to server

    Select Case tabInfo.CurrTab
        Case Tabs(eGDTab_NewOrder)
            strOrder = AlaronOrderFromControls(0, False)
            
        Case Tabs(eGDTab_WorkingOrders)
            strOrder = AlaronOrderFromGrid(fgWorkingOrders, fgWorkingOrders.Row)
            
    End Select
    
    If Len(strOrder) > 0 Then
        If Len(Trim(txtOrderID.Text)) = 0 Then
            g.Alaron.DumpDebug "Park Add Order from Broker Screen: " & strOrder
            g.Alaron.AddBrokerOrder strOrder, eGDRanOrderMode_Park
        Else
            g.Alaron.DumpDebug "Park Amend Order from Broker Screen: " & strOrder
            g.Alaron.AmendBrokerOrder strOrder, eGDRanOrderMode_Park
        End If
        tabInfo.CurrTab = Tabs(eGDTab_ParkedOrders)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.ParkOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Cancel an order on the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelOrder()
On Error GoTo ErrSection:

    Select Case tabInfo.CurrTab
        Case Tabs(eGDTab_WorkingOrders)
            With fgWorkingOrders
                If (.Row >= .FixedRows) And (.Row < .Rows) Then
                    g.Alaron.DumpDebug "Cancel Working Order from Broker Screen: " & .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_Account)) & ", " & .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_OrderNumber))
                    g.Alaron.CancelBrokerOrder .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_Account)), .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_OrderNumber))
                End If
            End With
            
        Case Tabs(eGDTab_ParkedOrders)
            With fgParkedOrders
                If (.Row >= .FixedRows) And (.Row < .Rows) Then
                    g.Alaron.DumpDebug "Cancel Parked Order from Broker Screen: " & .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_Account)) & ", " & .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_OrderNumber))
                    g.Alaron.CancelBrokerOrder .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_Account)), .TextMatrix(.Row, WorkingOrdersCol(eGDWorkingCol_OrderNumber))
                End If
            End With
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.CancelOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrder
'' Description: Cancel/Replace an order on the Alaron servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ModifyOrder()
On Error GoTo ErrSection:

    Dim strOrder As String              ' Order from the grid

    Select Case tabInfo.CurrTab
        Case Tabs(eGDTab_WorkingOrders)
            strOrder = AlaronOrderFromGrid(fgWorkingOrders, fgWorkingOrders.Row)
            
        Case Tabs(eGDTab_ParkedOrders)
            strOrder = AlaronOrderFromGrid(fgParkedOrders, fgParkedOrders.Row)
            
    End Select
    
    If Len(strOrder) > 0 Then
        AlaronOrderToControls strOrder
        tabInfo.CurrTab = Tabs(eGDTab_NewOrder)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.ModifyOrder"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetOrder
'' Description: Reset the order controls on the new order tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetOrder()
On Error GoTo ErrSection:

    txtCommandLine.Text = ""
    cboGoodThru.Text = "Today"
    cboSession.Text = "Day"
    
    lblOrderID.Visible = False
    txtOrderID.Visible = False
    txtOrderID.Text = ""
    
    lblError.Caption = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerScreen.ResetOrder"
    
End Sub
