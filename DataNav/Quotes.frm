VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmQuotes 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quote Board"
   ClientHeight    =   5910
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9510
   Icon            =   "Quotes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   4800
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2145
      Left            =   8340
      TabIndex        =   10
      Top             =   120
      Width           =   1110
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
      Caption         =   "Quotes.frx":038A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "Quotes.frx":03BE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Quotes.frx":03DE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSortBySym 
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   1600
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
         Caption         =   "Quotes.frx":03FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":0422
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":0442
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAlerts 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   885
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
         Caption         =   "Quotes.frx":045E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":048C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":04AC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1243
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
         Caption         =   "Quotes.frx":04C8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":04FA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":051A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRefresh 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   120
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
         Caption         =   "Quotes.frx":0536
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":0566
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":0586
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditList 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   495
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
         Caption         =   "Quotes.frx":05A2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":05D0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":05F0
         RightToLeft     =   0   'False
      End
   End
   Begin VB.Timer tmrRealTime 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   8700
      Top             =   2520
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons3 
      Height          =   570
      Left            =   60
      TabIndex        =   7
      Top             =   5340
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
      Caption         =   "Quotes.frx":060C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "Quotes.frx":0638
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Quotes.frx":0658
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
         Height          =   360
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   120
         Width           =   900
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
         Caption         =   "Quotes.frx":0674
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":06A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":06C6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAlerts 
         Height          =   360
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   120
         Width           =   900
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
         Caption         =   "Quotes.frx":06E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":0710
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":0730
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons2 
      Height          =   570
      Left            =   60
      TabIndex        =   4
      Top             =   120
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
      Caption         =   "Quotes.frx":074C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "Quotes.frx":0782
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Quotes.frx":07A2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdEditList 
         Height          =   360
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   120
         Width           =   900
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
         Caption         =   "Quotes.frx":07BE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":07EC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":080C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRefresh 
         Height          =   360
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   900
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
         Caption         =   "Quotes.frx":0828
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Quotes.frx":0858
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Quotes.frx":0878
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   4305
      Left            =   60
      TabIndex        =   16
      Top             =   780
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   7594
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
      Caption         =   "ALL|Indices|Stocks|Futures|(custom)"
      Align           =   0
      Appearance      =   1
      CurrTab         =   4
      FirstTab        =   0
      Style           =   0
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin VSFlex7LCtl.VSFlexGrid fgQuotes 
         Height          =   3930
         Left            =   -9135
         TabIndex        =   0
         Top             =   45
         Width           =   7890
         _cx             =   13917
         _cy             =   6932
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
      End
      Begin VB.PictureBox pbQuoteBoard 
         AutoRedraw      =   -1  'True
         Height          =   3930
         Left            =   -9435
         ScaleHeight     =   3870
         ScaleWidth      =   7830
         TabIndex        =   1
         Top             =   45
         Width           =   7890
         Begin gdOCX.gdScrollBar gdsHorz 
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            Horizontal      =   -1  'True
         End
         Begin gdOCX.gdScrollBar gdsVert 
            Height          =   1335
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   2355
         End
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Begin VB.Menu mnuTest1 
         Caption         =   "Test1"
      End
      Begin VB.Menu mnuTest2 
         Caption         =   "Test2"
      End
   End
   Begin VB.Menu mnuQuotes 
      Caption         =   "Quotes"
      Begin VB.Menu mnuAddSymbol 
         Caption         =   "Add a Symbol   (hotkey: 'Insert')"
      End
      Begin VB.Menu mnuLabelRow 
         Caption         =   "Add a Label Row"
      End
      Begin VB.Menu mnuRemoveSymbol 
         Caption         =   "Remove Symbol   (hotkey: 'Delete')"
      End
      Begin VB.Menu mnuToCategory 
         Caption         =   "Put symbol into Category"
         Begin VB.Menu mnuCategory 
            Caption         =   "< Edit tabs >"
            Index           =   0
         End
      End
      Begin VB.Menu mnuChangePeriod 
         Caption         =   "Change Period for All Symbols on Tab"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuy 
         Caption         =   "Buy"
      End
      Begin VB.Menu mnuSell 
         Caption         =   "Sell"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuoteBoardStyle 
         Caption         =   "Quote Tab Style"
         Begin VB.Menu mnuConvertToGrid 
            Caption         =   "Convert to Grid Style"
         End
         Begin VB.Menu mnuConvertToBox 
            Caption         =   "Convert to Box Style"
         End
         Begin VB.Menu mnuConvertToForex 
            Caption         =   "Convert to Forex Style"
         End
      End
      Begin VB.Menu mnuCopyQuoteTab 
         Caption         =   "Copy Quote Tab"
      End
      Begin VB.Menu mnuDetachQBTab 
         Caption         =   "Detach Quote Tab"
      End
      Begin VB.Menu mnuExportQuoteTab 
         Caption         =   "Export Quote Tab"
      End
      Begin VB.Menu mnuPrintQuoteTab 
         Caption         =   "Print Quote Tab"
      End
      Begin VB.Menu mnuChangeFilter 
         Caption         =   "Change Filter"
      End
      Begin VB.Menu mnuClearFilter 
         Caption         =   "Clear Filter"
      End
      Begin VB.Menu mnuFields 
         Caption         =   "Edit Fields (columns) to display"
      End
      Begin VB.Menu mnuQBF 
         Caption         =   "Edit Quote Board Field"
      End
      Begin VB.Menu mnuRenameQBF 
         Caption         =   "Rename Quote Board Field File"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddAlert 
         Caption         =   "Add Alert"
      End
      Begin VB.Menu mnuEditAlert 
         Caption         =   "Edit Alert"
      End
      Begin VB.Menu mnuAlerts 
         Caption         =   "Alerts"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuShowButtons 
         Caption         =   "Show Buttons"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSortBySym 
         Caption         =   "Sort"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmQuotes.frm
'' Description: Quote board interface
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/26/2009   DAJ         Fixed request for % of H-L when High=Low
'' 09/03/2009   DAJ         Implemented and called ChangeSymbolOnBox function
'' 03/05/2010   DAJ         Fixes for importing quote board tabs
'' 03/05/2010   DAJ         Check required module codes on QBFs
'' 12/11/2012   DAJ         Fix for LoadQbfs after a QBT import
'' 04/30/2014   DAJ         Change period for all rows in grid
'' 05/01/2014   DAJ         Include period on caption for 'change period' menu item
'' 05/01/2014   DAJ         Fix for filter symbols not being removed when period not daily
'' 06/12/2014   DAJ         Added 240 Minute to period combo boxes on grid
'' 10/16/2014   DAJ         Removed unused reference to File System Object
'' 06/11/2015   DAJ         Allow custom indexes to be added to the quote board
'' 05/18/2016   DAJ         Added '180 Minute' and '120 Minute' bar periods to the combo dropdown lists in the grid
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const MAX_DAYS = 10

Public Enum eGDQuoteStyle
    eGDQuoteStyle_Grid = -1
    eGDQuoteStyle_Thermometer
    eGDQuoteStyle_OHLC
    eGDQuoteStyle_Candlestick
    eGDQuoteStyle_Bollinger
    eGDQuoteStyle_NoBarPicture
    eGDQuoteStyle_Forex
End Enum

Public Enum eGDTabFuncWrapper
    eGDTabFuncWrapper_AddSymbol = 0
    eGDTabFuncWrapper_AddLabel
    eGDTabFuncWrapper_ChangePeriod
    eGDTabFuncWrapper_RemoveSymbol
    eGDTabFuncWrapper_SaveTab
End Enum

Private Type mPrivate
    QBData As cGdTable                  ' unique symbol-period (combined tabs)
    hQBData As Long
    alDataIndex As cGdArray             ' sorted index for QBData (row # of grid or cell # of box)
    hDataIndex As Long
    QB As cQuoteCellBoard
    tblTabInfo As cGdTable              ' tab info (1 record per tab)
    QBFs As cGdTree                     ' collection of criteria
    BarsColl As cGdTree                 ' unique symbol-period collection of bars (same size as QBData)

    astrSymbols As cGdArray             ' unique symbol-period list (same size as QBData)
    astrFields As cGdArray              ' info for each field
    astrCriteria As cGdArray            ' Array of criteria ID's
    aDetachedTabs As cGdArray           ' array of forms so don't have to search forms collection
    
    lQbfCol As Long
    nRefreshRow As Long
    bAbortedLoad As Boolean
    lTotWidth As Long
    bUserResize As Boolean
    DefaultStyle As eGDQuoteStyle
    
    strSaveSymbol As String
    strSavePeriod As String
    
    astrRequests As cGdArray
    strDefaultFields As String
    
    lCurrentTab As Long                 ' Currently selected tab

    nMouseDownRow As Long               ' used to start dragging during MouseMove
    nMouseDownButton As Long
    
    strFilterPeriod As String           ' Filter tab period
    alFilterVols As cGdArray            ' Array of volumes for the filter tab
    alFilterIdx As cGdArray             ' Index for the filter volumes array
    
    lFilterSortCol As Long              ' Last column sorted on the filter tab
    lFilterSortDir As Long              ' Direction of sort on last sort on the filter tab
    
    frmActiveDetTab As frmDetachedQBTab ' handle to detached quote tab form that invoked popup menu
    
    bShowButtons As Boolean
    dLastTotalRefresh As Double         ' time when Total Refresh last executed
    dWhenStreamingStarted As Double
End Type
Private m As mPrivate

Private Enum eTblColumns
    eQbTbl_SearchKey = 0
    eQbTbl_SecType
    eQbTbl_Symbol
    eQbTbl_SymbolID
    eQbTbl_FeedSymbol
    eQbTbl_Delay
    eQbTbl_Rows
    eQbTbl_Criteria
    eQbTbl_Recalc   ' need to recalc criteria for this symbol
    eQbTbl_Dirty    ' displayed data is older than what's in the table
End Enum

Private Enum eGDCols
    eGDCol_SymbolID = 0
    eGDCol_SecType
    eGDCol_Symbol
    eGDCol_Period
    eGDCol_Delay
    eGDCol_NumFixed
End Enum

Public Enum eGDTabSettings
    eGDTabSettings_Name = 0
    eGDTabSettings_Style
    eGDTabSettings_Symbols
    eGDTabSettings_Fields
    eGDTabSettings_FilterID
    eGDTabSettings_Form            'hwnd to detached tab
End Enum

' Wrapper functions for enumerations...
Private Function TblField(ByVal FldNum As eTblColumns) As Long
    TblField = FldNum
End Function
Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function
Private Function TabField(ByVal Num As eGDTabSettings) As Long
    TabField = Num
End Function
Private Function QStyle(ByVal Style As eGDQuoteStyle) As Long
    QStyle = Style
End Function

Public Property Get SymbolID() As Long
    If CurrentTabStyle = eGDQuoteStyle_Grid Then
        If fgQuotes.Row >= fgQuotes.FixedRows And fgQuotes.Row < fgQuotes.Rows Then
            SymbolID = CLng(Val(fgQuotes.TextMatrix(fgQuotes.Row, GDCol(eGDCol_SymbolID))))
        Else
            SymbolID = 0&
        End If
    Else
        If Not m.QB.Cell(m.QB.Row, m.QB.Col) Is Nothing Then
            SymbolID = m.QB.Cell(m.QB.Row, m.QB.Col).SymbolID
        Else
            SymbolID = 0&
        End If
    End If
End Property

Public Property Get Symbol() As String
    If CurrentTabStyle = eGDQuoteStyle_Grid Then
        If fgQuotes.Row >= fgQuotes.FixedRows And fgQuotes.Row < fgQuotes.Rows Then
            Symbol = fgQuotes.TextMatrix(fgQuotes.Row, GDCol(eGDCol_Symbol))
        Else
            Symbol = ""
        End If
    Else
        If Not m.QB.Cell(m.QB.Row, m.QB.Col) Is Nothing Then
            Symbol = m.QB.Cell(m.QB.Row, m.QB.Col).Symbol
        Else
            Symbol = ""
        End If
    End If
End Property

Private Property Get TblStr(ByVal nField As eTblColumns, ByVal lRecord As Long) As String
    'tblStr = m.QBData.Item(nField, lRecord)
    TblStr = gdGetTableString(m.hQBData, nField, lRecord)
End Property

Private Property Let TblStr(ByVal nField As eTblColumns, ByVal lRecord As Long, ByVal strData As String)
    'm.QBData.Item(nField, lRecord) = strData
    gdSetTableStr m.hQBData, nField, lRecord, strData
End Property

Private Property Get TblNum(ByVal nField As eTblColumns, ByVal lRecord As Long) As Double
    'TblNum = m.QBData.Num(nField, lRecord)
    TblNum = gdGetTableNum(m.hQBData, nField, lRecord)
End Property

Private Property Let TblNum(ByVal nField As eTblColumns, ByVal lRecord As Long, ByVal dData As Double)
    'm.QBData.Num(nField, lRecord) = dData
    gdSetTableNum m.hQBData, nField, lRecord, dData
End Property

Public Property Get TabStr(ByVal nField As eGDTabSettings, ByVal lRecord As Long) As String
    TabStr = m.tblTabInfo.Item(nField, lRecord)
End Property

Public Property Let TabStr(ByVal nField As eGDTabSettings, ByVal lRecord As Long, ByVal strData As String)
    m.tblTabInfo.Item(nField, lRecord) = strData
End Property

Public Property Get TabNum(ByVal nField As eGDTabSettings, ByVal lRecord As Long) As Double
    TabNum = m.tblTabInfo.Num(nField, lRecord)
End Property

Public Property Let TabNum(ByVal nField As eGDTabSettings, ByVal lRecord As Long, ByVal dData As Double)
    m.tblTabInfo.Num(nField, lRecord) = dData
End Property

Public Property Get TabRecords() As Long
    TabRecords = m.tblTabInfo.NumRecords
End Property

Public Property Get AlertFields2() As String
    AlertFields2 = TabStr(eGDTabSettings_Fields, 0)
End Property
Public Property Get AlertFields() As cGdArray
    Set AlertFields = m.astrFields
End Property
Public Property Get AlertSymbols() As cGdArray
    Set AlertSymbols = m.QBData.FieldArray(TblField(eQbTbl_SearchKey), True)
End Property

Public Property Get UpColor() As Long
    Dim lColor&
    lColor = m.QB.UpColor
    If g.nColorTheme = kDarkThemeColor Then
        If IsBlueRange(lColor) Then
            lColor = vbCyan
        ElseIf IsGreenRange(lColor, True) Then
            lColor = vbGreen
        ElseIf lColor = vbBlack Then
            lColor = vbWhite
        End If
    End If
    UpColor = lColor
End Property

Public Property Get DownColor() As Long
    Dim lColor&
    lColor = m.QB.DownColor
    If g.nColorTheme = kDarkThemeColor Then
        If IsBlueRange(lColor) Then
            lColor = vbCyan
        ElseIf IsGreenRange(lColor, True) Then
            lColor = vbGreen
        ElseIf lColor = vbBlack Then
            lColor = vbWhite
        End If
    End If
    DownColor = lColor
End Property
Public Property Get UnchColor() As Long
    Dim lColor&
    lColor = m.QB.UnchColor
    If g.nColorTheme = kDarkThemeColor Then
        If IsBlueRange(lColor) Then
            lColor = vbCyan
        ElseIf IsGreenRange(lColor, True) Then
            lColor = vbGreen
        ElseIf lColor = vbBlack Then
            lColor = vbWhite
        End If
    End If
    UnchColor = lColor
End Property
Public Property Get UpdateColor() As Long
    Dim lColor&
    lColor = m.QB.UpdateColor
    If g.nColorTheme = kDarkThemeColor Then
        If IsBlueRange(lColor) Then
            lColor = vbCyan
        ElseIf IsGreenRange(lColor, True) Then
            lColor = vbGreen
        ElseIf lColor = vbBlack Then
            lColor = vbWhite
        End If
    End If
    UpdateColor = lColor
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Symbols (Get)
'' Description: Get a unique set of symbols that are on the quote board
'' Inputs:      None
'' Returns:     Array of Symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Symbols() As cGdArray
On Error GoTo ErrSection:

    Dim astrTemp As cGdArray

    Set astrTemp = m.QBData.FieldArray(TblField(eQbTbl_Symbol), True).MakeCopy
    astrTemp.Sort eGdSort_DeleteDuplicates
    
    Set Symbols = astrTemp

ErrExit:
    Set astrTemp = Nothing
    Exit Property
    
ErrSection:
    Set astrTemp = Nothing
    RaiseError "frmQuotes.Symbols.Get", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAlerts_Click
'' Description: If the user clicks on the Alerts button, show the Alerts menu
'' Inputs:      Which Alerts button the user clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAlerts_Click(Index As Integer)
On Error GoTo ErrSection:

    MoveFocusToQb
    
    frmAlertsSetup.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.cmdAlerts_Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditList_Click
'' Description: If the user clicks on the Edit List button, allow them to pick
''              their list of symbols to get quotes on
'' Inputs:      Which Edit List button the user clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditList_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim X&, Y&
    
    
    MoveFocusToQb
    EditCategories          'JM (05-07-2010): aardvark 5711

'JM (05-07-2010): original code, leave awhile then remove
'    With cmdEditList(Index)
'        If Index = 0 Then
'            X = fraButtons.Left + .Left + .Width
'            Y = fraButtons.Top + .Top + .Height
'        Else
'            X = fraButtons2.Left + .Left + .Width
'            Y = fraButtons2.Top + .Top + .Height
'        End If
'
'        mnuQBF.Visible = False
'        mnuRenameQBF.Visible = False
'
'        ShowQuotesPopup X, Y 'PopupMenu mnuQuotes, vbPopupMenuLeftAlign, X, Y
'    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.cmdEditList.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRefresh_Click
'' Description: If the user clicks on the refresh button, refresh the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshQB()
On Error GoTo ErrSection:

    Dim strMsg$, dWait#, iRefreshMode As Integer
    Static dOneTimeAgo#, dTwoTimesAgo#
    Static bMsgShown As Boolean

    If ProcessIsBusy Then Exit Sub
    
    ' if real-time, give message
    iRefreshMode = 1
    If g.RealTime.Active Then
#If 0 Then
        iRefreshMode = 2
        If Not bMsgShown Then
            bMsgShown = True
            strMsg = "When realtime is active, the 'Refresh' button on the quote board will refresh just the symbols which are not being updated in the data stream."
            InfBox strMsg, "i", , "Reminder" '"Quote Board Refresh"
        End If
        g.dLastQuoteBoardRefresh = Now
#Else
        strMsg = "Refresh which symbols?||(it is much quicker to refresh just the symbols which are not included in the realtime stream)"
        Select Case InfBox(strMsg, "?", "+Unstreamed|All Symbols|-Cancel", "Quote Board Refresh")
        Case "C"
            Exit Sub
        Case "U"
            iRefreshMode = 2
        End Select
#End If
    End If
    
    Select Case iRefreshMode
    Case 1
        ' limit refresh times (twice every 10 minutes)
        dWait = dTwoTimesAgo + 10 * 60000# - gdTickCount
        If dWait > 2000 And dTwoTimesAgo > 0 Then
            If dWait < 100000 Then
                strMsg = "There is a limit of 2 refreshes every 10 minutes|(" _
                    & Str(Round(dWait / 1000)) & " seconds until next allowed)"
            Else
                strMsg = "There is a limit of 2 refreshes every 10 minutes|(" _
                    & Str(Round(dWait / 60000)) & " minutes until next allowed)"
            End If
            InfBox strMsg, "!", , "Quote Board Refresh"
            Exit Sub
        End If
    Case 2
        ' is just unstreamed, set this flag now (even if unsuccessful)
        g.dLastQuoteBoardRefresh = Now
    End Select
               
    If Not g.RealTime.RefreshSymbolList(iRefreshMode) Then
        If iRefreshMode = 2 Then
            InfBox "All symbols are being streamed.", "i", , "Quote Board Refresh"
        End If
        Exit Sub
    End If
    
    ' if successful, store current time
    If FileExist(App.Path & "\ftp\data.dat") And (iRefreshMode <> 2) Then
        dTwoTimesAgo = dOneTimeAgo
        dOneTimeAgo = gdTickCount
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.RefreshQB"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRefresh_Click
'' Description: If the user clicks on the refresh button, refresh the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRefresh_Click(Index As Integer)
On Error GoTo ErrSection:

    MoveFocusToQb
    
    RefreshQB
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.cmdRefresh.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSettings_Click
'' Description: If the user clicks on the settings button, bring up the
''              program settings form on the Snap Quote tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSettings_Click(Index As Integer)
On Error GoTo ErrSection:

    ''tmrRealTime.Interval = Val(InfBox("Interval in milliseconds:", "?", , "Quote Board", , , , , , "s", Str(tmrRealTime.Interval)))
    
    ShowSettings

    MoveFocusToQb

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.cmdSettings.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSortBySym_Click
'' Description: Don't allow a sort during a total refresh
''              Sort box/forex quoteboard by symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSortBySym_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the infbox
    Dim strTabInfo As String
    
    Dim QB As cQuoteCellBoard
    Dim lIndex As Long
    
    If m.frmActiveDetTab Is Nothing Then
        Set QB = m.QB
        lIndex = vsTab.CurrTab
    Else
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        lIndex = m.frmActiveDetTab.MyTabIndex
    End If

    If m.nRefreshRow <> 0 Or (CurrentTabStyle = eGDQuoteStyle_Grid And m.frmActiveDetTab Is Nothing) Then
        'total refresh in progress or incorrect quoteboard style (do nothing)
    ElseIf Not QB Is Nothing Then
        strReturn = InfBox("How do you wish to sort the symbols?", "?", "Ascending|Descending|+-Cancel", "Confirmation")
        If strReturn = "A" Or strReturn = "D" Then
            strTabInfo = TabStr(eGDTabSettings_Symbols, lIndex)
            If strReturn = "A" Then
                strTabInfo = QB.SortSymbols(strTabInfo, eGdSort_Default)
            Else
                strTabInfo = QB.SortSymbols(strTabInfo, eGdSort_Descending)
            End If
            If Len(strTabInfo) > 0 Then
                TabStr(eGDTabSettings_Symbols, lIndex) = strTabInfo
                If m.frmActiveDetTab Is Nothing Then
                    If Len(strTabInfo) > 0 Then ShowCategory
                Else
                    m.frmActiveDetTab.DrawBoxQB
                End If
            End If
        End If
    End If
    
    Set m.frmActiveDetTab = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.cmdSortBySym_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterEdit
'' Description: Handle the user's change
'' Inputs:      Row and Column of the edited cell
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current row in the grid
    Dim strSymbol As String             ' Symbol from the grid
    Dim strPeriod As String             ' Period from the grid
    Dim bRemoved As Boolean             ' Was removed from the data table
    Dim bAdded As Boolean               ' Was added to the data table

    If vsTab.TabCaption(vsTab.CurrTab) = "(Filter)" Then
        If fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period)) <> m.strSavePeriod Then
            m.strFilterPeriod = fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period))
            SetIniFileProperty "FilterPeriod", m.strFilterPeriod, "QuoteBoard", g.strIniFile
            
            ChangePeriodForAllRows m.strFilterPeriod, m.strSavePeriod
        End If
    Else
        If fgQuotes.MergeRow(Row) = False Then
            Select Case Col
                Case GDCol(eGDCol_Period)
                    If fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period)) <> m.strSavePeriod Then
                        ChangePeriodForRow Row, fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period)), m.strSavePeriod
                        
'                        With fgQuotes
'                            .Redraw = flexRDNone
'                            lRow = Row
'                            strSymbol = Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)
'                            strPeriod = .TextMatrix(lRow, GDCol(eGDCol_Period))
'                            bRemoved = RemoveSymbolFromGrid(lRow)
'                            bAdded = AddSymbolToGrid(strSymbol, strPeriod, lRow)
'                            If (bRemoved Or bAdded) And g.RealTime.Active Then
'                                'g.RealTime.UpdateSymbolList
'                            End If
'                            .Redraw = flexRDBuffered
'                        End With
                        
                        ResetRows
                        TotalRefresh False
                        m.strSavePeriod = ""
                    End If
                    
                Case GDCol(eGDCol_Symbol)
                
            End Select
        ElseIf Row = fgQuotes.Rows - 1 Then
            If Col >= fgQuotes.FrozenCols And Len(Trim(fgQuotes.TextMatrix(Row, Col))) > 0 Then
                AddLabelRow "", fgQuotes.Rows
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterMoveRow
'' Description: After a user has moved a row in the quotes grid, turn the
''              sort triangle off
'' Inputs:      Row moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    ' this clears the triangle displayed for sorting
    ' (after moving a row, it's now "unsorted")
    With fgQuotes
        .ExplorerBar = flexExNone
        .ExplorerBar = flexExSortShowAndMove
    End With
    
    ' Make sure that the alternate coloring is correct...
    ColorQuoteRows
    
    ' Reset the rows field of the data table...
    ResetRows

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterMoveRow", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterRowColChange
'' Description: After the user changes the cell, make sure that the cell is
''              shown and go into edit mode if applicable
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With fgQuotes
        'if no horiz scrollbars, we need to "manually scroll"
        'by forcing the cell to show when changing columns
        If .Redraw <> flexRDNone And .ScrollBars <> flexScrollBarBoth Then
            If OldCol <> NewCol Then
                .ShowCell NewRow, NewCol
            End If
        End If
        
        If NewCol = GDCol(eGDCol_Period) Or .MergeRow(NewRow) = True Then
            If NewRow >= .FixedRows And NewRow < .Rows And m.nMouseDownButton <> vbRightButton Then
                fgQuotes.EditCell
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterScroll
'' Description: After a scroll, update the rows that haven't been visible
'' Inputs:      Old Top and Left, New Top and Left
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lFrom As Long                   ' First of the rows that just became visible
    Dim lTo As Long                     ' Last of the rows that just became visible
    
    ' Figure out the rows that we will need to update (the new visible rows)...
    If OldTopRow > NewTopRow Then
        lFrom = NewTopRow
        lTo = OldTopRow - 1
    ElseIf OldTopRow < NewTopRow Then
        lFrom = OldTopRow + (fgQuotes.BottomRow - fgQuotes.TopRow) + 1
        lTo = NewTopRow + (fgQuotes.BottomRow - fgQuotes.TopRow)
    End If
    
    ' Update the new visible rows...
    For lIndex = lFrom To lTo
        UpdateCols lIndex
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterScroll", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterSort
'' Description: After the user resorts the grid, save off what symbols are in
''              what rows
'' Inputs:      Column sorted, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    Dim i&, strText$

    With fgQuotes
        i = fgQuotes.Rows - 1
        If .MergeRow(i) = True Then         '5799
            strText = .TextMatrix(i, .Cols - 1)
            If InStr(strText, "Click here") <> 0 Then
                .RemoveItem i
                AddLabelRow strText, .FixedRows
            End If
        End If
    End With
    
    ' Add the blank row back as the last row of the grid...
    AddLabelRow "", fgQuotes.Rows

    ' Make sure that the alternate coloring is correct...
    ColorQuoteRows
    
    ' Reset the rows field of the data table...
    ResetRows
    
    If vsTab.TabCaption(vsTab.CurrTab) = "(Filter)" Then
        m.lFilterSortCol = Col
        m.lFilterSortDir = Order
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterSort", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterUserFreeze
'' Description: Make sure that the first column is frozen
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterUserFreeze()
On Error GoTo ErrSection:

    If fgQuotes.FrozenCols < 1 Then fgQuotes.FrozenCols = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterUserFreeze", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_AfterUserResize
'' Description: If the user resized a column, set the flag
'' Inputs:      Row and Column resized
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Col <> -1 Then
        m.bUserResize = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_BeforeEdit
'' Description: Only allow the user to edit the "Period" column
'' Inputs:      Row and Column being edited, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strComboList As String          ' Combo list string
    
    strComboList = "Daily|240 Minute|180 Minute|120 Minute|60 Minute|30 Minute|15 Minute|10 Minute|5 Minute"

    If vsTab.TabCaption(vsTab.CurrTab) = "(Filter)" Then
        'Cancel = True
        If fgQuotes.MergeRow(Row) = True Then
            Cancel = True
        ElseIf Col <> GDCol(eGDCol_Period) Then
            Cancel = True
        Else
            m.strSavePeriod = fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period))
            fgQuotes.ComboList = strComboList & "|1 Minute"
        End If
    Else
        If fgQuotes.MergeRow(Row) = True Then
            fgQuotes.ComboList = ""
            If Col = GDCol(eGDCol_Symbol) Then
                Cancel = True
            End If
        Else
            Select Case Col
                Case GDCol(eGDCol_Period)
                    m.strSavePeriod = fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period))
                    fgQuotes.ComboList = "|" & strComboList
                
                Case GDCol(eGDCol_Symbol)
                    m.strSaveSymbol = fgQuotes.TextMatrix(Row, GDCol(eGDCol_Symbol))
                    fgQuotes.ComboList = "..."
                
                Case Else
                    fgQuotes.ComboList = ""
                    Cancel = True
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_BeforeMouseDown
'' Description: Either bring up the pop up menu (right click) or set up the row
''              for dragging purposes
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Location of Mouse,
''              Whether to Cancel the Mouse Down
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRow As Long, nCol As Long      ' Current mouse row
    Dim strPeriod As String             ' Period for the row
    Dim nPeriod As Long                 ' Periodicity for the row

    Set m.frmActiveDetTab = Nothing

    If fgQuotes.Redraw = flexRDNone Then Exit Sub
    
    m.nMouseDownButton = Button
    
    With fgQuotes
        nRow = .MouseRow
        nCol = .MouseCol
        If nRow >= .FixedRows And nRow < .Rows Then         'aardvark 4193
            If nCol >= .FixedCols And nCol < .Cols Then
                .Row = nRow
                .Col = nCol
                .Refresh
                If Button = vbRightButton Then
                    ' show popup menu
                    Cancel = True
                    If nCol >= 0 And nCol < .Cols Then
                        m.lQbfCol = nCol
                        mnuQBF.Visible = (.ColData(m.lQbfCol) <> "")
                        mnuRenameQBF.Visible = (.ColData(m.lQbfCol) <> "")
                    Else
                        m.lQbfCol = 0
                        mnuQBF.Visible = False
                        mnuRenameQBF.Visible = False
                    End If
                    mnuFields.Visible = True
                    mnuLabelRow.Visible = True
                    
                    If .MergeRow(nRow) Then
                        strPeriod = ""
                        nPeriod = 0
                    Else
                        strPeriod = .TextMatrix(nRow, GDCol(eGDCol_Period))
                        nPeriod = GetPeriodicity(strPeriod)
                    End If
                    Enable mnuChangePeriod, (nPeriod >= ePRD_Days) Or (GetPeriodType(nPeriod) = ePRD_Minutes)
                    
                    ShowQuotesPopup X, Y ' PopupMenu mnuQuotes, , X, Y
                
                ElseIf (vsTab.CurrTab = vsTab.NumTabs - 2) And (nRow = .FixedRows) Then     '5711 - And (nCol = GDCol(eGDCol_Symbol)) Then
                    ChangeFilter
                End If
            End If
        End If
    End With
    
    m.nMouseDownButton = 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_BeforeMoveColumn
'' Description: When the user tries to move a column, just make sure that the
''              symbol column stays where it is
'' Inputs:      Column to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    ' Keep the symbol where it is...
    If Col = GDCol(eGDCol_Symbol) Then
        Position = GDCol(eGDCol_Symbol)
    ElseIf Col = GDCol(eGDCol_Period) Then
        Position = GDCol(eGDCol_Period)
    ElseIf Col = GDCol(eGDCol_Delay) Then
        Position = GDCol(eGDCol_Delay)
    ElseIf Col <> GDCol(eGDCol_Symbol) And Position = GDCol(eGDCol_Symbol) Then
        Position = Col
    ElseIf Col <> GDCol(eGDCol_Period) And Position = GDCol(eGDCol_Period) Then
        Position = Col
    ElseIf Col <> GDCol(eGDCol_Delay) And Position = GDCol(eGDCol_Delay) Then
        Position = Col
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.BeforeMoveColumn", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_BeforeMoveRow
'' Description: Don't allow a move row during a total refresh
'' Inputs:      Row to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_BeforeMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    If m.nRefreshRow <> 0 Then
        Position = Row
    ElseIf Position = fgQuotes.Rows - 1 Then
        Position = Position - 1
    ElseIf Row = fgQuotes.Rows - 1 Then
        Position = Row
    ElseIf UCase(vsTab.TabCaption(m.lCurrentTab)) = "(FILTER)" Then
        If Row = fgQuotes.FixedRows Then
            Position = Row
        ElseIf Position = fgQuotes.FixedRows Then
            Position = fgQuotes.FixedRows + 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.BeforeMoveRow", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgQuotes_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck fgQuotes, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_BeforeSort
'' Description: Don't allow a sort during a total refresh
'' Inputs:      Column to sort, Order to sort it
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the infbox

    If m.nRefreshRow <> 0 Then
        Order = flexSortNone
    Else
        strReturn = InfBox("How do you wish to sort this column?", "?", "Ascending|Descending|+-Cancel", "Confirmation")
        Select Case strReturn
            Case "A"
                Order = flexSortGenericAscending
                If fgQuotes.MergeRow(fgQuotes.Rows - 1) = True Then fgQuotes.RemoveItem fgQuotes.Rows - 1
                
            Case "D"
                Order = flexSortGenericDescending
                If fgQuotes.MergeRow(fgQuotes.Rows - 1) = True Then fgQuotes.RemoveItem fgQuotes.Rows - 1
            
            Case "C"
                Order = flexSortNone
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.BeforeSort", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_CellButtonClick
'' Description: Bring up the symbol selector for the user to choose a symbol
'' Inputs:      Row and Column of the Cell Button
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Col = GDCol(eGDCol_Symbol) And ((Row >= fgQuotes.FixedRows) And (Row < fgQuotes.Rows)) Then
        ChangeSymbolOnGrid Row
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.CellButtonClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_Click
'' Description: When the user clicks on the Quotes grid, make sure that the
''              current row gets changed to the mouse row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current grid row of the mouse

    lMouseRow = fgQuotes.MouseRow
    If lMouseRow >= fgQuotes.FixedRows And lMouseRow < fgQuotes.Rows Then
        fgQuotes.Row = lMouseRow
        fgQuotes.RowSel = lMouseRow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_ComboCloseUp
'' Description: When the user closes the combo box, tell the flex grid to finish
''              the edit so that the AfterEdit happens right away
'' Inputs:      Row and Column of the Edit, Whether to Finish the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.ComboCloseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_DblClick
'' Description: If the user double clicks on the grid, chart the symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_DblClick()
On Error GoTo ErrSection:

    Dim nRow As Long                    ' Row user double clicked on

    ' Make sure that the user double clicked on a non-fixed row
    With fgQuotes
        nRow = .MouseRow
        If nRow >= .FixedRows Then
            If .MergeRow(nRow) = False Then
                .Row = nRow
                .Col = GDCol(eGDCol_Symbol)
                SetActiveChartSymbol Parse(.TextMatrix(nRow, GDCol(eGDCol_Symbol)), "(", 1)
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_GotFocus
'' Description: When the grid gets the focus, make sure that something a row
''              is selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_GotFocus()
On Error GoTo ErrSection:

    With fgQuotes
        If (.Row < .FixedRows Or .Row >= .Rows) And .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
            .Col = GDCol(eGDCol_Symbol)
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgQuotes_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If fgQuotes.Col = GDCol(eGDCol_Symbol) Then
        Select Case Chr(KeyAscii)
            Case "A" To "Z", "a" To "z", "$" ', "#"
                m.strSaveSymbol = fgQuotes.TextMatrix(fgQuotes.Row, GDCol(eGDCol_Symbol))
                ChangeSymbolOnGrid fgQuotes.Row, Chr(KeyAscii)
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_KeyUp
'' Description: If the user presses the delete key on a symbol in the quote
''              board, remove that symbol
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If fgQuotes.EditWindow = 0 Then
        Select Case KeyCode
            Case vbKeyDelete
                DoRemoveSymbol
            
            Case vbKeyInsert
                DoAddSymbol
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes.KeyUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_MouseDown
'' Description: Save off the current MouseRow for dragging purposes
'' Inputs:      Mouse button being pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseRow As Long               ' Current mouse row in the grid

    lMouseRow = fgQuotes.MouseRow
    
    If (UCase(vsTab.TabCaption(m.lCurrentTab)) = "(FILTER)") And (lMouseRow = fgQuotes.FixedRows) Then
        m.nMouseDownRow = 0
    ElseIf m.nMouseDownRow = kNullData Then
        'this is set in changefilter routine to prevent dragging after user cancel from filter dialog
        'that is brought up by single clicking in the symbol column (does not happen from right-click menu)
        m.nMouseDownRow = 0
    Else
        ' save row when MouseDown occurred in order to start dragging in MouseMove
        m.nMouseDownRow = lMouseRow
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_MouseMove
'' Description: Show the tooltip for the grid if applicable
'' Inputs:      Mouse button being pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim nRow As Long
    Dim lMouseCol As Long
    Dim lMouseRow As Long
    Dim strCol As String

    With fgQuotes
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        If m.nRefreshRow = 0 Then
            ' if row has changed since mouse went down, then start dragging
            If m.nMouseDownRow <> .MouseRow And m.nMouseDownRow >= .FixedRows And .MouseRow >= .FixedRows Then
                nRow = m.nMouseDownRow
                m.nMouseDownRow = 0
                .DragRow nRow
            End If
        End If
        If lMouseCol >= .FixedCols And lMouseCol < .Cols Then
            If .TextMatrix(0, lMouseCol) = "T" Then
                strCol = "Tick/Settle"
            ElseIf (vsTab.CurrTab = vsTab.NumTabs - 2) And (lMouseCol = GDCol(eGDCol_Symbol)) And (lMouseRow = .FixedRows) Then
                strCol = "Click here to change or clear the filter"
            End If
        End If
        GridTooltip fgQuotes, , strCol
    End With

End Sub

Private Sub fgQuotes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m.nMouseDownRow = 0

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgQuotes_ValidateEdit
'' Description: Make sure that the period column is only Daily or X Minute Bars
'' Inputs:      Row and Column of edit, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgQuotes_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If fgQuotes.MergeRow(Row) = True Then
        If Len(fgQuotes.EditText) = 0 Then
            fgQuotes.EditText = " "
        ElseIf Not ValidLabel(fgQuotes.EditText) Then
            Cancel = True
        End If
    Else
        Select Case Col
            Case GDCol(eGDCol_Period)
                fgQuotes.EditText = GetPeriodStr(fgQuotes.EditText)
                
                ' We now want to start allowing periods greater than daily on the quote board
                ' so we no longer need this test (DAJ: 04/30/2008)...
                'If GetPeriodicity(fgQuotes.EditText) > ePRD_Days Then
                '    fgQuotes.EditText = "Daily"
                'End If
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.fgQuotes_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form gets activated, reset the toolbar and change
''              the window list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean      ' Have we been here before?
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
    End If
        
    Me.Caption = "Quote Board:  hit 'Insert' to add symbol, 'Delete' to remove"
    
    EnableButtons
    
    fgQuotes.BackColorAlternate = ALT_GRID_ROW_COLOR
    fgQuotes.Redraw = flexRDBuffered
    'MoveFocus fgQuotes  '(don't do this since it makes people hit a button twice the first time)
    
    ToolbarSync Me
    
    TextIncDecRegisterForm Me, True
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Deactivate
'' Description: When the form gets deactivated, set the previous form to me
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Deactivate()
On Error GoTo ErrSection:
    
    SetPrevActiveForm Me
    Me.Caption = "Quote Board"
    
    
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Handle various special keys at the form level
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        g.Help.ShowF1Help Me
    ElseIf KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        KeyPress KeyCode, Shift
    ElseIf KeyCode = vbKeyInsert Then
        KeyCode = 0
        If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
            DoAddSymbol fgQuotes.Row
        Else
            DoAddSymbol m.QB.Row, m.QB.Col
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyPress
'' Description: Handle various special keys at the form level
'' Inputs:      Ascii representation of the key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form gets loaded, do some initialization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Number of the data field
    Dim lRow As Long                    ' Index for a for loop
    Dim strTemp As String               ' Temporary string
    Dim strField As String              ' Field
    Dim strDefault As String            ' Default field order from ini file
    Dim iAutoQuotes As Integer          ' Auto Quote value from the ini file
    Dim strID As String                 ' ID of the QBF if one exists
    Dim QBF As New cCriteria            ' QBF to store on the quotes grid
    Dim strFont As String               ' Font from the ini file
    Dim i&
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrTemp As New cGdArray        ' Temporary array of strings
    Dim lIndex2 As Long                 ' Index into a for loop
    
    g.Styler.StyleForm Me
        
    If g.bUnloading Then
        ' if main program is unloading, then get out now
        m.bAbortedLoad = True
        DebugLog "AbortedLoad for frmQuotes"
        Unload Me
        Exit Sub
    End If
        
    m.strFilterPeriod = GetIniFileProperty("FilterPeriod", "Daily", "QuoteBoard", g.strIniFile)
    m.lFilterSortCol = GetIniFileProperty("FilterSortCol", -1&, "QuoteBoard", g.strIniFile)
    m.lFilterSortDir = GetIniFileProperty("FilterSortDir", -1&, "QuoteBoard", g.strIniFile)
    
    ' Default the box style controls to invisible...
    pbQuoteBoard.Visible = False
    gdsHorz.Visible = False
    gdsVert.Visible = False
       
    ' Default the pop up menu's to be invisible...
    mnuTest.Visible = False
    mnuQuotes.Visible = False
    
    'RH commented out fraButtons.BorderStyle = vbBSNone
    'RH commented out fraButtons2.BorderStyle = vbBSNone
    
    ' Initialize the quote cell board...
    Set m.QB = New cQuoteCellBoard
    LoadSettings
    m.QB.Init pbQuoteBoard, gdsHorz, gdsVert

    Set m.QBData = New cGdTable
    With m.QBData
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_SearchKey), "SearchKey"
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_SecType), "SecType"
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_Symbol), "Symbol"
        .CreateField eGDARRAY_Longs, TblField(eQbTbl_SymbolID), "SymbolID"
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_FeedSymbol), "FeedSymbol"
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_Delay), "Delay"
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_Rows), "Rows"
        .CreateField eGDARRAY_Strings, TblField(eQbTbl_Criteria), "Criteria"
        .CreateField eGDARRAY_TinyInts, TblField(eQbTbl_Recalc), "Recalc"
        .CreateField eGDARRAY_TinyInts, TblField(eQbTbl_Dirty), "Dirty"
    End With
    m.hQBData = m.QBData.TableHandle
    
    Set m.BarsColl = New cGdTree
    
    Set m.astrFields = New cGdArray
    m.astrFields.Create eGDARRAY_Strings
    Set astrTemp = New cGdArray
    astrTemp.Create eGDARRAY_Strings
    Set m.alDataIndex = New cGdArray
    m.alDataIndex.Create eGDARRAY_Longs
    Set m.alFilterVols = New cGdArray
    m.alFilterVols.Create eGDARRAY_Longs
    Set m.alFilterIdx = New cGdArray
    m.alFilterIdx.Create eGDARRAY_Longs
    Set m.aDetachedTabs = New cGdArray
    m.aDetachedTabs.Create eGDARRAY_Objects
    
    LoadTabInfo                         'restore saved quoteboard tabs information from file or set up default ones
    m.lCurrentTab = GetIniFileProperty("QuoteBoard", 0, "CurrTab", g.strIniFile)
    InitQuoteTabs m.lCurrentTab         'set up tab control based on info in tab talbe
    
    strDefault = "SymbolID|SecType|Symbol|Date|Open|High|Low|Last|Volume|Bid|Ask|Prev Close|Change|% Change|Last Tick|"
    strTemp = GetIniFileProperty("DisplayFields", strDefault, "QuoteList", g.strIniFile)
    If InStr(strTemp, "Date") = 0 Then
        lRow = InStr(strTemp, "Symbol|")
        If lRow > 0 Then strTemp = Left(strTemp, lRow + 6) & "Date|" & Mid(strTemp, lRow + 7)
    End If
    If InStr(strTemp, "Last Tick") = 0 Then
        lRow = InStr(strTemp, "Date|")
        If lRow > 0 Then strTemp = Left(strTemp, lRow + 4) & "Last Tick|" & Mid(strTemp, lRow + 5)
    End If
    If InStr(strTemp, "Bid;") = 0 And InStr(strTemp, "Bid|") = 0 Then strTemp = strTemp & "Bid;-1;|"
    If InStr(strTemp, "Ask;") = 0 And InStr(strTemp, "Ask|") = 0 Then strTemp = strTemp & "Ask;-1;|"
    If InStr(strTemp, "Bid Size") = 0 Then strTemp = strTemp & "Bid Size;-1;|"
    If InStr(strTemp, "Ask Size") = 0 Then strTemp = strTemp & "Ask Size;-1;|"
    If InStr(strTemp, "Prev Close") = 0 Then strTemp = strTemp & "Prev Close;-1;|"
    If InStr(strTemp, "Open Interest") = 0 Then strTemp = strTemp & "Open Interest;-1;|"
    If InStr(strTemp, "Exchange") = 0 Then strTemp = strTemp & "Exchange;-1;|"
    If InStr(strTemp, "% of H-L") = 0 Then strTemp = strTemp & "% of H-L;-1;|"
    ''If Left(strTemp, 5) <> "Dirty" Then strTemp = "Dirty;-1;|" & strTemp
    
    'field format: Name;ShowFlag(0/-1);criteria.scn file(optional)|Name;ShowFlag(0/-1);criteria.scn file(optional)|...
    Set astrTemp = New cGdArray
    astrTemp.SplitFields strTemp, "|"
    
    'make sure period, feed symbol, delay, description and T are in fields list
    If Left(astrTemp(GDCol(eGDCol_Period)), 6) <> "Period" Then
        astrTemp.Add "Period;0;", GDCol(eGDCol_Period)
        strTemp = astrTemp.JoinFields("|")
        If Right(strTemp, 1) <> "|" Then strTemp = strTemp & "|"
    End If
    If InStr(strTemp, "Feed Symbol") = 0 Then strTemp = strTemp & "Feed Symbol;-1;|"
    If InStr(strTemp, "Delay") = 0 Then strTemp = strTemp & "Delay;-1;|"
    If InStr(strTemp, "Description") = 0 Then strTemp = strTemp & "Description;-1|"
    If InStr(strTemp, "|T;") = 0 Then
        astrTemp.SplitFields strTemp, "|"
        For lIndex = 0 To astrTemp.Size - 1
            If Parse(astrTemp(lIndex), "|", 1) = "Last" Then
                astrTemp.Add "T;0;", lIndex + 1
                strTemp = astrTemp.JoinFields("|")
                If Right(strTemp, 1) <> "|" Then strTemp = strTemp & "|"
                Exit For
            End If
        Next lIndex
    End If
    
    ' Remove the dirty flag from the default list of columns...
    astrTemp.SplitFields strTemp, "|"
    For lIndex = 0 To astrTemp.Size - 1
        If Parse(astrTemp(lIndex), ";", 1) = "Dirty" Then
            astrTemp.Remove lIndex
        End If
        If Parse(astrTemp(lIndex), ";", 1) = "Date" Then
            If Len(Trim(Parse(astrTemp(lIndex), ";", 3))) = 0 Then
                astrTemp(lIndex) = "Session;" & Parse(astrTemp(lIndex), ";", 2) & ";"
            End If
        End If
    Next lIndex
    strTemp = astrTemp.JoinFields("|")
    If Right(strTemp, 1) <> "|" Then strTemp = strTemp & "|"
    
    ' Save off the default fields for use with new categories...
    m.strDefaultFields = strTemp
        
    ' If the column information for a category is not set yet, assign the default...
    For lIndex = 0 To m.tblTabInfo.NumRecords - 1
        If Len(TabStr(eGDTabSettings_Fields, lIndex)) = 0 Then
            TabStr(eGDTabSettings_Fields, lIndex) = m.strDefaultFields
        Else
            If InStr(TabStr(eGDTabSettings_Fields, lIndex), "Feed Symbol") = 0 Then TabStr(eGDTabSettings_Fields, lIndex) = TabStr(eGDTabSettings_Fields, lIndex) & "Feed Symbol;-1;|"
            If InStr(TabStr(eGDTabSettings_Fields, lIndex), "Delay") = 0 Then
                If Right(TabStr(eGDTabSettings_Fields, lIndex), 1) = "|" Then
                    TabStr(eGDTabSettings_Fields, lIndex) = TabStr(eGDTabSettings_Fields, lIndex) & "Delay;-1;|"
                Else
                    TabStr(eGDTabSettings_Fields, lIndex) = TabStr(eGDTabSettings_Fields, lIndex) & "|Delay;-1;|"
                End If
            End If
            If InStr(TabStr(eGDTabSettings_Fields, lIndex), "Description") = 0 Then
                If Right(TabStr(eGDTabSettings_Fields, lIndex), 1) = "|" Then
                    TabStr(eGDTabSettings_Fields, lIndex) = TabStr(eGDTabSettings_Fields, lIndex) & "Description;-1;|"
                Else
                    TabStr(eGDTabSettings_Fields, lIndex) = TabStr(eGDTabSettings_Fields, lIndex) & "|Description;-1;|"
                End If
            End If
            If InStr(TabStr(eGDTabSettings_Fields, lIndex), "|T;") = 0 Then
                astrTemp.SplitFields TabStr(eGDTabSettings_Fields, lIndex), "|"
                For lIndex2 = 0 To astrTemp.Size - 1
                    If Parse(astrTemp(lIndex2), ";", 1) = "Last" Then
                        astrTemp.Add "T;0;", lIndex2 + 1
                        TabStr(eGDTabSettings_Fields, lIndex) = astrTemp.JoinFields("|")
                        If Right(TabStr(eGDTabSettings_Fields, lIndex), 1) <> "|" Then
                            TabStr(eGDTabSettings_Fields, lIndex) = TabStr(eGDTabSettings_Fields, lIndex) & "|"
                        End If
                        Exit For
                    End If
                Next lIndex2
            End If
            
            ' Remove the dirty flag from the default list of columns...
            astrTemp.SplitFields TabStr(eGDTabSettings_Fields, lIndex), "|"
            For i = 0 To astrTemp.Size - 1
                If Parse(astrTemp(i), ";", 1) = "Dirty" Then
                    astrTemp.Remove i
                End If
                If Parse(astrTemp(i), ";", 1) = "Date" Then
                    If Len(Trim(Parse(astrTemp(i), ";", 3))) = 0 Then
                        astrTemp(i) = "Session;" & Parse(astrTemp(i), ";", 2) & ";;" & Parse(astrTemp(i), ";", 4)
                    End If
                End If
                If Parse(astrTemp(i), ";", 1) = "Session" Then
                    If Len(Trim(Parse(astrTemp(i), ";", 3))) > 0 Then
                        astrTemp(i) = "Session;" & Parse(astrTemp(i), ";", 2) & ";;" & Parse(astrTemp(i), ";", 4)
                    End If
                End If
            Next i
            TabStr(eGDTabSettings_Fields, lIndex) = astrTemp.JoinFields("|")
        End If
        
        ' Move the delay column to it's new location after the period...
        astrTemp.SplitFields TabStr(eGDTabSettings_Fields, lIndex), "|"
        'For i = 0 To astrTemp.Size - 1
        '    If Parse(astrTemp(i), ";", 1) = "Delay" Then
        '        astrTemp.MoveItems i, 1, (i - GDCol(eGDCol_Delay)) * -1
        '        Exit For
        '    End If
        'Next i
        FixPeriodAndDelay astrTemp
        TabStr(eGDTabSettings_Fields, lIndex) = astrTemp.JoinFields("|")
    Next lIndex
    
    Set m.astrCriteria = New cGdArray
    m.astrCriteria.Create eGDARRAY_Strings
    
    InitGrid
    LoadQbfs2

    ' Load the global alerts collection...
    If g.Alerts Is Nothing Then Set g.Alerts = New cAlerts
    g.Alerts.Load AddSlash(App.Path) & "Custom\QuoteList.ALR"
    
    If LoadTable = False Then
        Err.Raise vbObjectError + 1000, , "Could not load Quote Board List"
    End If
    
    ' Make a backup of the QuoteBoard.INF file if successfully loaded...
    If FileExist(AddSlash(App.Path) & "Custom\QuoteBoard.INF") Then
        FileCopy AddSlash(App.Path) & "Custom\QuoteBoard.INF", AddSlash(App.Path) & "Custom\QuoteBoard.BAK", True
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    InfBox "There was an error loading the quote board.  Trade Navigator is restoring the quote board information from a backup and will shut down.  Please restart the program.|", "e", , "Quote Board Error"
    If FileExist(AddSlash(App.Path) & "Custom\QuoteBoard.BAK") Then
        FileCopy AddSlash(App.Path) & "Custom\QuoteBoard.BAK", AddSlash(App.Path) & "Custom\QuoteBoard.INF", True
    Else
        KillFile AddSlash(App.Path) & "Custom\QuoteBoard.INF"
    End If
    End
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the form was unloaded from the system menu, hide it instead
''              of unloading it.
'' Inputs:      Whether or not to cancel unload, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    TextIncDecUnregisterForm Me
    
    SaveAllQbSettings
    
    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_Quote").State = ssUnchecked
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the form is resized, resize the grid accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lIndex As Long

    'If LimitFormSize(Me, fraButtons.Width * 5, 3000) Then Exit Sub
    
    If m.bShowButtons Then
        If Me.ScaleWidth >= Me.ScaleHeight Then
            fraButtons.Visible = True
            fraButtons2.Visible = False
            fraButtons3.Visible = False
            
            With fraButtons
                .Move ScaleWidth - .Width, 0, .Width, ScaleHeight
            End With
            With vsTab
                .Move 0, 0, ScaleWidth - fraButtons.Width, ScaleHeight
                .Refresh
            End With
            
            With pbQuoteBoard
                .Move 0, 0, vsTab.ClientWidth, vsTab.ClientHeight
            End With
        Else
            fraButtons.Visible = False
            fraButtons2.Visible = True
            fraButtons3.Visible = True
            
            With fraButtons2
                .Move 0, 0, ScaleWidth, .Height
            End With
            With fraButtons3
                .Move 0, ScaleHeight - .Height, ScaleWidth, .Height
            End With
            With vsTab
                .Move 0, fraButtons2.Height, ScaleWidth, ScaleHeight - fraButtons2.Height - fraButtons3.Height
                .Refresh
            End With
        End If
    Else
        fraButtons.Visible = False
        fraButtons2.Visible = False
        fraButtons3.Visible = False
        With vsTab
            .Move 0, 0, ScaleWidth, ScaleHeight
            .Refresh
        End With
    End If
    
    ' Make sure to redraw the visible area of the grid...
    If CurrentTabStyle = eGDQuoteStyle_Grid Then
        For lIndex = fgQuotes.TopRow To fgQuotes.BottomRow
            UpdateCols lIndex
        Next lIndex
    End If
    
    AutoSizeChart
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form gets unloaded, reset the window list
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrRealtime.Enabled = False
    
    ' Clean up...
    Set m.alDataIndex = Nothing
    Set m.QBFs = Nothing
    Set m.BarsColl = Nothing
    Set m.QB = Nothing
    Set m.QBData = Nothing
    Set m.tblTabInfo = Nothing
    Set m.alFilterVols = Nothing
    
    ToolbarSync Me, False
    
    frmMain.DockPro.RemoveForm Me.Name
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.Form_Unload"
    Resume ErrExit
    
End Sub

Public Sub SaveAllQbSettings()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strTemp As String               ' Temporary string
       
    If m.bAbortedLoad Then Exit Sub
     
    ' Save the tab information...
    SaveTabInfo vsTab.CurrTab
    FileFromString AddSlash(App.Path) & "Custom\QuoteBoard.INF", m.tblTabInfo.ToString(vbLf, vbTab)
    SetIniFileProperty "QuoteBoard", vsTab.CurrTab, "CurrTab", g.strIniFile
    
    ' Save the alerts...
    g.Alerts.Save
    
    ' Save the column order of the quotes grid...
    strTemp = ""
    On Error Resume Next
    For lIndex = 0 To fgQuotes.Cols - 1
        If fgQuotes.TextMatrix(0, lIndex) <> "#DELETED#" Then
            strTemp = strTemp & fgQuotes.TextMatrix(0, lIndex) & ";" & CStr(CLng(fgQuotes.ColHidden(lIndex))) & ";" & fgQuotes.ColData(lIndex) & "|"
        End If
    Next lIndex
    SetIniFileProperty "DisplayFields", strTemp, "QuoteList", g.strIniFile
    
    SetIniFileProperty "FilterSortCol", m.lFilterSortCol, "QuoteBoard", g.strIniFile
    SetIniFileProperty "FilterSortDir", m.lFilterSortDir, "QuoteBoard", g.strIniFile
    
    ' Save miscellaneous settings about the grid and box style quote boards...
    SaveSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.SaveAllQbSettings"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbol
'' Description: Add a symbol to the grid
'' Inputs:      Record number of the symbol to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddSymbol(ByVal lRecNum As Long, ByVal strPeriod As String, _
            Optional ByVal strSymbol As String = "", Optional ByVal lPos As Long = -1&)
On Error GoTo ErrSection:

    If lRecNum <> -1 Then strSymbol = g.SymbolPool.Symbol(lRecNum)
    
    If Len(strSymbol) > 0 Then
        If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
            AddSymbolToGrid strSymbol, strPeriod, lPos
        Else
            AddSymbolToBox strSymbol
        End If
        
        ' Reset the rows information in the data table...
        ResetRows
        
        TotalRefresh False
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.AddSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSecType
'' Description: Given the security type from the symbol pool, return the
''              security type for the request file
'' Inputs:      Security type from the symbol pool
'' Returns:     The security type that the request file is expecting
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSecType(ByVal pSecType As eSYM_SecType) As String
On Error GoTo ErrSection:

    Select Case pSecType
        Case eSYMType_Index
            GetSecType = "I"
        Case eSYMType_Stock
            GetSecType = "S"
        Case eSYMType_Future
            GetSecType = "F"
        Case Else
            GetSecType = "SO"
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.GetSecType", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolData
'' Description: Get the value of the cell for the given symbol and field
'' Inputs:      Symbol, Field name
'' Returns:     Value at that cell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolData(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, ByVal strField As String) As Double
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index for a for loop
    Dim Bars As cGdBars                 ' Bars structure to get price from display
    Dim dClose As Double                ' Current Close from the bars
    Dim dHigh As Double                 ' Current High from the bars
    Dim dLow As Double                  ' Current Low from the bars
    Dim dPrev As Double                 ' Previous Close from the bars
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' Criteria ID for quote board field
    Dim lSymbolID As Long               ' Symbol ID for the given symbol
    Dim strSymbol As String             ' Symbol passed in
    Dim dValue As Double
    
    If VarType(vSymbolOrSymbolID) = vbString Then
        strSymbol = vSymbolOrSymbolID
        lSymbolID = GetSymbolID(strSymbol)
    Else
        lSymbolID = vSymbolOrSymbolID
        strSymbol = GetSymbol(lSymbolID)
    End If
    SymbolData = kNullData
    
    ' Find the row in the table with the given symbol...
    If TblSearch(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod, lRow) Then
        Set Bars = GetBars(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod)
        If Not Bars Is Nothing Then
            Select Case UCase(strField)
                Case "SESSION"
                    SymbolData = Bars.SessionDate(Bars.Size - 1)
                Case "OPEN"
                    SymbolData = Bars(eBARS_Open, Bars.Size - 1)
                Case "HIGH"
                    SymbolData = Bars(eBARS_High, Bars.Size - 1)
                Case "LOW"
                    SymbolData = Bars(eBARS_Low, Bars.Size - 1)
                Case "LAST"
                    SymbolData = Bars(eBARS_Close, Bars.Size - 1)
                Case "PREV CLOSE"
                    SymbolData = GetPrevCloseForQB(Bars)
                Case "CHANGE"
                    dClose = Bars(eBARS_Close, Bars.Size - 1)
                    dPrev = GetPrevCloseForQB(Bars)
                    If dClose <> kNullData And dPrev <> kNullData Then
                        SymbolData = dClose - dPrev
                    End If
                Case "% CHANGE"
                    dClose = Bars(eBARS_Close, Bars.Size - 1)
                    dPrev = GetPrevCloseForQB(Bars)
                    If dClose <> kNullData And dPrev <> kNullData Then
                        SymbolData = (dClose - dPrev) / dPrev
                    End If
                Case "% OF H-L"
                    dClose = Bars(eBARS_Close, Bars.Size - 1)
                    dHigh = Bars(eBARS_High, Bars.Size - 1)
                    dLow = Bars(eBARS_Low, Bars.Size - 1)
                    If dClose <> kNullData And dHigh <> kNullData And dLow <> kNullData Then
                        If (dHigh <= dLow) Or (dClose > dHigh) Or (dClose < dLow) Then
                            SymbolData = kNullData
                        Else
                            SymbolData = (dClose - dLow) / (dHigh - dLow)
                        End If
                    End If
                Case "BID"
                    SymbolData = Bars(eBARS_Bid, Bars.Size - 1)
                Case "BID SIZE"
                    SymbolData = Bars(eBARS_BidSize, Bars.Size - 1)
                Case "ASK"
                    SymbolData = Bars(eBARS_Ask, Bars.Size - 1)
                Case "ASK SIZE"
                    SymbolData = Bars(eBARS_AskSize, Bars.Size - 1)
                Case "VOLUME"
                    SymbolData = Bars(eBARS_Vol, Bars.Size - 1)
                Case "OPEN INTEREST"
                    SymbolData = Bars(eBARS_OI, Bars.Size - 1)
                Case "LAST TICK"
                    dValue = 0
                    If Bars(eBARS_DateTime, Bars.Size - 1) > 0 Then
                        dValue = Bars.Prop(eBARS_LastTickTime)
                        If dValue <> 0 Then
                            dValue = Int(Bars(eBARS_DateTime, Bars.Size - 1)) + dValue / 1440#
                            If g.bShowInLocalTimeZone Then
                                dValue = ConvertTimeZone(dValue, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                            End If
                            dValue = dValue - Int(dValue)
                        End If
                    End If
                    SymbolData = dValue
                Case "EXCHANGE"
                    'SymbolData = Bars(eBARS_DateTime, Bars.Size - 1)
                Case Else
                    For lIndex = 1 To m.QBFs.Count
                        If UCase(m.QBFs(lIndex).Name) = UCase(strField) Then
                            strID = m.QBFs(lIndex).ID
                            Exit For
                        End If
                    Next lIndex
                    SymbolData = m.QBData(m.QBData.FieldNum(strID), lRow)
            End Select
        End If
    End If
            
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.SymbolData", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableButtons
'' Description: Enable/Disable certain buttons under certain circumstances
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableButtons()
On Error GoTo ErrSection:

    Dim i&, bEnable As Boolean              ' Whether or not to enable a control
    
    'bEnable = (fgQuotes.Rows > fgQuotes.FixedRows)
    If Not g.RealTime.SalmonIsRunning Then
        bEnable = (m.QBData.NumRecords > 0)
    End If
    If mnuRefresh.Enabled <> bEnable Then
        mnuRefresh.Enabled = bEnable
        For i = cmdRefresh.LBound To cmdRefresh.UBound
            cmdRefresh(i).Enabled = bEnable
        Next
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.EnableButtons", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddAlert_Click
'' Description: Allow the user to add an alert from the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddAlert_Click()
On Error GoTo ErrSection:

    Dim Alert As New cAlert             ' Temporary alert object
    Dim lIndex As Long                  ' Index where new alert was added
    
    Dim lTabIdx As Long
    Dim fgCurrGrid As VSFlexGrid
    
    Dim QB As cQuoteCellBoard
    Dim QbCell As cQuoteCell
    
    Dim bBoxQB As Boolean
    
    If m.frmActiveDetTab Is Nothing Then
        lTabIdx = vsTab.CurrTab
        Set fgCurrGrid = fgQuotes
        Set QB = m.QB
    Else
        lTabIdx = m.frmActiveDetTab.MyTabIndex
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        Set QB = m.frmActiveDetTab.QuoteCellBoard
    End If
    
    Alert.IsSymbol = True
    If TabStr(eGDTabSettings_Style, lTabIdx) = Str(QStyle(eGDQuoteStyle_Grid)) Then
    
        With fgCurrGrid
            'make sure alert object gets a valid symbol else the frmAlerts dialog will give an error
            If .Row >= .Rows - 1 Or .Row < .FixedRows Then .Row = .FixedRows
        End With
    
        Alert.SymbolID = CLng(Val(fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_SymbolID))))
        Alert.Symbol = Parse(fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_Symbol)), "(", 1)
        Alert.Period = fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_Period))
        If Len(fgCurrGrid.ColData(fgCurrGrid.Col)) = 0 Then
            Alert.field = fgCurrGrid.TextMatrix(0, fgCurrGrid.Col)
        Else
            Alert.CriteriaID = fgCurrGrid.ColData(fgCurrGrid.Col)
            Alert.field = fgCurrGrid.TextMatrix(0, fgCurrGrid.Col)
        End If
    Else
    
'JM 01-28-2011 - Tim's email 12/31/2010
'2. Go ahead and display the "Add Alert" menu item for the box-style qb (we're currently only showing that for the grid-style)
'   just have it bring up the alert editor for that symbol and choose "Last" as the default
        
        Set QbCell = QB.Cell(QB.Row, QB.Col)
        If QbCell Is Nothing Then Set QbCell = QB.FirstCellInTree
    
        If QbCell Is Nothing Then
            GoTo ErrExit            'nothing on QB just exit
        Else
            Alert.SymbolID = QbCell.SymbolID
            Alert.Symbol = Parse(QbCell.Symbol, "(", 1)
            Alert.Period = "Daily"
            Alert.field = "Last"
        End If
        bBoxQB = True

    End If
    
    If Len(Alert.Symbol) > 0 And Len(Alert.field) > 0 Then
        Alert.Operator = "="
        If Alert.SymbolID = 0 Then
            Alert.Value = SymbolData(Alert.Symbol, Alert.Period, Alert.field)
        Else
            Alert.Value = SymbolData(Alert.SymbolID, Alert.Period, Alert.field)
        End If
    End If
    
    If frmAlerts.ShowMe(Alert, eGDAlertType_QuoteBoard, bBoxQB) = True Then
        lIndex = g.Alerts.Add(Alert)
        Alert.CheckAlert
        If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid      '6085
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuAddAlert.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddSymbol_Click
'' Description: Allow the user to add a symbol to the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddSymbol_Click()
On Error GoTo ErrSection:
    
    If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        DoAddSymbol fgQuotes.Row
    Else
        DoAddSymbol m.QB.Row, m.QB.Col
    End If
    
ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuAddSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuAlerts_Click()
On Error GoTo ErrSection:

    cmdAlerts_Click 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuAlerts_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuBuy_Click
'' Description: Allow the user to buy a security
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuBuy_Click()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol to pass to create order
    
    Dim eStyle As eGDQuoteStyle
    Dim fgCurrGrid As VSFlexGrid
    Dim QB As cQuoteCellBoard
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
        Set QB = m.QB
        eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), vsTab.CurrTab)
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), m.frmActiveDetTab.MyTabIndex)
    End If
    
    If eStyle = eGDQuoteStyle_Grid And Not fgCurrGrid Is Nothing Then
        strSymbol = RollSymbolForDate(Parse(fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_Symbol)), "(", 1), Now)
    ElseIf Not QB Is Nothing Then
        If Not QB.Cell(QB.Row, QB.Col) Is Nothing Then
            strSymbol = RollSymbolForDate(Parse(QB.Cell(QB.Row, QB.Col).Symbol, "(", 1), Now)
        End If
    End If
    CreateOrder strSymbol, , 1, , , "Quote Board"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuBuy_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCategory_Click
'' Description: Move/Put a symbol into a custom category
'' Inputs:      Index of the menu item selected
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCategory_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim lFromCat As Long                ' Tab moving from
    Dim lToCat As Long                  ' Tab to move to
    Dim strSymbol As String             ' Symbol to move
    Dim strPeriod As String             ' Period of the symbol to move

    Dim eStyle As eGDQuoteStyle
    Dim iTabIdx&, hwndTo&, i&
    
    Dim fgCurrGrid As VSFlexGrid
    Dim QB As cQuoteCellBoard
    Dim frmToDetached As frmDetachedQBTab
        
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
        Set QB = m.QB
        iTabIdx = vsTab.CurrTab
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        iTabIdx = m.frmActiveDetTab.MyTabIndex
    End If
    
    eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), iTabIdx)
    
    If Index = 0 Then
        EditCategories
    Else
        ' Figure out the tab coming from and tab going to...
        lFromCat = iTabIdx
        lToCat = Index - 1
                                
        'check for going to a detached tab
        hwndTo = m.tblTabInfo(TabField(eGDTabSettings_Form), lToCat)
        If hwndTo <> 0 Then
            For i = 0 To m.aDetachedTabs.Size - 1
                If Not m.aDetachedTabs(i) Is Nothing Then
                    If hwndTo = m.aDetachedTabs(i).hWnd Then
                        Set frmToDetached = m.aDetachedTabs(i)
                        Exit For
                    End If
                End If
            Next
        End If
        
        'failed to find detached tab - something very wrong, just exit
        If hwndTo <> 0 And frmToDetached Is Nothing Then GoTo ErrExit
        
        ' Get symbol information for the symbol/period to move...
        If eStyle = eGDQuoteStyle_Grid Then
            strSymbol = Parse(fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_Symbol)), "(", 1)
            strPeriod = fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_Period))
        Else
            strSymbol = Parse(QB.Cell(QB.Row, QB.Col).Symbol, "(", 1)
            strPeriod = "Daily"
        End If
        
        ' If going to a new tab, remove from the old, switch to the new, and add...
        If lFromCat <> lToCat And Len(strSymbol) > 0 Then
            If TabStr(eGDTabSettings_Style, lFromCat) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                RemoveSymbolFromGrid fgCurrGrid.Row
            Else
                RemoveSymbolFromBox QB.Row, QB.Col
            End If
            
            If Not m.frmActiveDetTab Is Nothing Then
                TabStr(eGDTabSettings_Symbols, m.frmActiveDetTab.MyTabIndex) = m.frmActiveDetTab.MySymbols(True)
            End If
            
            If frmToDetached Is Nothing Then
                Set m.frmActiveDetTab = Nothing
                vsTab.CurrTab = lToCat
            Else
                Set m.frmActiveDetTab = frmToDetached
            End If
            
            If TabStr(eGDTabSettings_Style, lToCat) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                AddSymbolToGrid strSymbol, strPeriod
            Else
                AddSymbolToBox strSymbol
            End If
                        
            ' Reset the rows information in the data table...
            If m.frmActiveDetTab Is Nothing Then
                ResetRows
                TotalRefresh False
            Else
                TotalRefresh False
                TabStr(eGDTabSettings_Symbols, m.frmActiveDetTab.MyTabIndex) = m.frmActiveDetTab.MySymbols(True)
                UpdateCols m.frmActiveDetTab.fgQuotes.Rows - 2, , , , m.frmActiveDetTab.fgQuotes
            End If
        End If
    End If

ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuCategory.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFilter_Click
'' Description: Allow the user to change the filter on the filter tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFilter_Click()
On Error GoTo ErrSection:

    ChangeFilter

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuChangeFilter_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Change the font of the quotes grid if the user chooses to
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    If vsTab.Tag = Str(eGDQuoteStyle_Grid) Then
        ChangeGridFont fgQuotes, True
    Else
        If CommonDialogFont(frmMain.CommonDialog1, m.QB.Font) = True Then
            m.QB.Font = m.QB.Font
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangePeriod_Click
'' Description: Allow the user to change the bar period for all of the symbols
''              on the current tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangePeriod_Click()
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Period to change to
    Dim nPeriod As Long                 ' Periodicity
    Dim QbGrid As VSFlexGrid            ' Grid to modify
    
    If m.frmActiveDetTab Is Nothing Then
        Set QbGrid = fgQuotes
    Else
        Set QbGrid = m.frmActiveDetTab.fgQuotes
    End If
    
    With QbGrid
        If (ValidGridRow(QbGrid) = True) And (.MergeRow(.Row) = False) Then
            strPeriod = .TextMatrix(.Row, GDCol(eGDCol_Period))
            If Len(strPeriod) > 0 Then
                nPeriod = GetPeriodicity(strPeriod)
                If (nPeriod >= ePRD_Days) Or (GetPeriodType(nPeriod) = ePRD_Minutes) Then
                    ChangePeriodForAllRows strPeriod
                End If
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuChangePeriod_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuClearFilter_Click
'' Description: Allow the user to clear the filter that has been applied
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuClearFilter_Click()
On Error GoTo ErrSection:
    
    ClearFilterTab
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuClearFilter_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromGridToCells
'' Description: Changes the quoteboard from using grid to using cells (i.e box or forex)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FromGridToCells(ByVal lTab&) As Boolean
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols in the category
    Dim lIndex As Long                  ' Index into a for loop
    
    
    astrSymbols.Create eGDARRAY_Strings
    astrSymbols.SplitFields TabStr(eGDTabSettings_Symbols, lTab), ","
    For lIndex = astrSymbols.Size - 1 To 0 Step -1
        If Len(astrSymbols(lIndex)) > 0 Then
            If Parse(astrSymbols(lIndex), ";", 1) = "Label" Then
                astrSymbols.Remove lIndex
            Else
                astrSymbols(lIndex) = Parse(astrSymbols(lIndex), ";", 1)
            End If
        End If
    Next lIndex
    TabStr(eGDTabSettings_Symbols, lTab) = astrSymbols.JoinFields(",")

    FromGridToCells = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.FromGridToCells"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuConvertToBox_Click
'' Description: Allow user to convert the quote board to box style for this tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuConvertToBox_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    If m.frmActiveDetTab Is Nothing Then
        lIndex = vsTab.CurrTab
    Else
        lIndex = m.frmActiveDetTab.MyTabIndex
    End If

    If TabStr(eGDTabSettings_Style, lIndex) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        If Not FromGridToCells(lIndex) Then GoTo ErrExit
        TabStr(eGDTabSettings_Style, lIndex) = Str(m.DefaultStyle)
    ElseIf TabStr(eGDTabSettings_Style, lIndex) = Str(QStyle(eGDQuoteStyle_Forex)) Then
        TabStr(eGDTabSettings_Style, lIndex) = Str(m.DefaultStyle)
    Else
        'something very wrong here - just ignore
    End If
    
    If m.frmActiveDetTab Is Nothing Then
        ShowCategory vsTab.CurrTab
    Else
        m.frmActiveDetTab.UpdateStyle
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.mnuConvertToBox_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuConvertToForex_Click
'' Description: Allow user to convert the quote board to forex style for this tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuConvertToForex_Click()
    
    Dim lTab As Long
    
    If m.frmActiveDetTab Is Nothing Then
        lTab = vsTab.CurrTab
    Else
        lTab = m.frmActiveDetTab.MyTabIndex
    End If
    
    If TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        If Not FromGridToCells(lTab) Then GoTo ErrExit
        TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Forex))
    ElseIf TabStr(eGDTabSettings_Style, lTab) <> Str(QStyle(eGDQuoteStyle_Forex)) Then
        TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Forex))
    Else
        'something very wrong here - just ignore
    End If
    
    If m.frmActiveDetTab Is Nothing Then
        ShowCategory vsTab.CurrTab
    Else
        m.frmActiveDetTab.UpdateStyle
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.mnuConvertToForex_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuConvertToGrid_Click
'' Description: Allow user to convert the quote board to grid style for this tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuConvertToGrid_Click()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols in the category
    Dim lIndex As Long                  ' Index into a for loop
    Dim lTab As Long
    
    Dim lPos As Long
    Dim lSymbolID As Long
    Dim lPoolRec As Long
    Dim strSymbol As String
    Dim bAdded As Boolean
    
    If m.frmActiveDetTab Is Nothing Then
        lTab = vsTab.CurrTab
    Else
        lTab = m.frmActiveDetTab.MyTabIndex
    End If

    astrSymbols.Create eGDARRAY_Strings
    astrSymbols.SplitFields TabStr(eGDTabSettings_Symbols, lTab), ","

    If TabStr(eGDTabSettings_Style, lTab) <> Str(QStyle(eGDQuoteStyle_Grid)) Then  'precautionary check
        For lIndex = 0 To astrSymbols.Size - 1
            If Len(astrSymbols(lIndex)) > 0 Then
                lSymbolID = Val(Parse(astrSymbols(lIndex), ";", 1))
                strSymbol = ""
                
                If TblSearch(SymbolOrSymbolID(lSymbolID, strSymbol), "Daily", lPos) = False Then
                    lPoolRec = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
                    bAdded = AddSymbolToTable(lPoolRec, "Daily")
                End If
                
                strSymbol = Str(lSymbolID)
                astrSymbols(lIndex) = strSymbol & ";Daily"
            End If
        Next lIndex
        TabStr(eGDTabSettings_Symbols, lTab) = astrSymbols.JoinFields(",")
        TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Grid))
    End If

    If m.frmActiveDetTab Is Nothing Then
        ShowCategory vsTab.CurrTab
        TotalRefresh False
    Else
        If m.frmActiveDetTab.fgQuotes.FixedRows = 0 Then InitGrid m.frmActiveDetTab.fgQuotes
        PopulateGrid m.frmActiveDetTab.fgQuotes, TabStr(eGDTabSettings_Symbols, lTab)
        m.frmActiveDetTab.UpdateStyle
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.mnuConvertToGrid_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuCopyQuoteTab_Click()
    EditCategories True
End Sub

Private Sub mnuDetachQBTab_Click()
On Error GoTo ErrSection:

    Dim bAllow As Boolean

    If ExtremeCharts >= 1 Then
        bAllow = True
    ElseIf HasLevel(eTN3_Standard) Then
        bAllow = True
    End If
    
    If bAllow Then
        If InStr(mnuDetachQBTab.Caption, "Detach") = 0 Then
            If Not m.frmActiveDetTab Is Nothing Then
                'if your detached QB has 30+ symbols (like the DOW 30)
                'it takes several tries before the form closes unless the call to Statusmsg is made
                'just moving the focus to a different control does not help
                StatusMsg Space(4)
                Unload m.frmActiveDetTab
                Set m.frmActiveDetTab = Nothing
            Else
                DebugLog "No active detached QB. Attach QB failed."
            End If
        Else
            If m.aDetachedTabs.Size = m.tblTabInfo.NumRecords - 2 And Not HasFilterTab Then
                InfBox "You cannot detach all quote tabs.", "I"
            Else
                DetachTab m.lCurrentTab
            End If
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.mnuDetachQBTab_Click", eGDRaiseError_Show

End Sub

Private Sub mnuEditAlert_Click()
On Error GoTo ErrSection:

    Dim Alert As cAlert
    Dim strKey As String
        
    Dim lTabIdx As Long
    Dim fgCurrGrid As VSFlexGrid
    
    Dim bWasTabAlert As Boolean
    Dim bIsTabAlert As Boolean
    
    Dim eTabStyle As eGDQuoteStyle
    
    If m.frmActiveDetTab Is Nothing Then
        lTabIdx = vsTab.CurrTab
        Set fgCurrGrid = fgQuotes
    Else
        lTabIdx = m.frmActiveDetTab.MyTabIndex
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
    End If
    
    eTabStyle = TabStr(eGDTabSettings_Style, lTabIdx)
    
    If eTabStyle = eGDQuoteStyle_Grid Then
        With fgCurrGrid
            If m.frmActiveDetTab Is Nothing Then
                If m.lQbfCol > .FixedCols And m.lQbfCol < .Cols Then
                    .Col = m.lQbfCol
                End If
            End If
            
            If .Row >= 0 And .Row < .Rows Then
                If .Col >= .FixedCols And .Col < .Cols Then
                    If .Cell(flexcpPicture, .Row, .Col) Is Nothing Then .Row = 0    'try header row
                    If Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then
                        strKey = Parse(.Cell(flexcpData, .Row, .Col), "|", 2)
                        If Len(strKey) > 0 Then
                            Set Alert = g.Alerts(strKey)
                        End If
                    End If
                End If
            End If
        End With
        
        If Not Alert Is Nothing Then
            If Len(Alert.TabName) > 0 Then bWasTabAlert = True
            
            If frmAlerts.ShowMe(Alert, eGDAlertType_QuoteBoard) = True Then
                If Len(Alert.TabName) > 0 Then bIsTabAlert = True
                If (bIsTabAlert And Not bWasTabAlert) Or (Not bIsTabAlert And bWasTabAlert) Then
                    'a single cell alert got changed to a tab alert or vice versa
                    'remove the bell icon, remove alert key, set back color to normal if was cell alert
                    With fgCurrGrid
                        .Cell(flexcpPicture, .Row, .Col) = Nothing
                        .Cell(flexcpData, .Row, .Col) = Parse(.Cell(flexcpData, .Row, .Col), "|", 1)
                        If Not bWasTabAlert Then
                            .Cell(flexcpBackColor, .Row, .Col) = .Cell(flexcpBackColor, .Row, 0)
                        End If
                    End With
                    
                End If
                Alert.CheckAlert
                If FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid      '6085
                DisplayAlert Alert
            End If
        End If
    End If

    If Alert Is Nothing Or eTabStyle <> eGDQuoteStyle_Grid Then
        cmdAlerts_Click 0
    End If

    Set Alert = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuEditAlert_Click"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuExportQuoteTab_Click
'' Description: Allow the user to export the current quote board tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuExportQuoteTab_Click()
On Error GoTo ErrSection:

    ExportCurrentTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuExportQuoteTab_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuFields_Click
'' Description: Allow the user to add/remove/change order of the fields in the
''              quote grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuFields_Click()
On Error GoTo ErrSection:

    EditFields

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuFields.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuPrintQuoteTab_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuPrintQuoteTab_Click"
    Resume ErrExit

End Sub

Private Sub mnuRefresh_Click()
On Error GoTo ErrSection:

    cmdRefresh_Click 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuRefresh_Click"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRenameQBF_Click
'' Description: Allow the user to rename the underlying file for the criteria
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRenameQBF_Click()
On Error GoTo ErrSection:

    Dim QBF As New cCriteria            ' Temporary holder for QBF object
    
    Set QBF = m.QBFs(fgQuotes.ColData(m.lQbfCol))
    If QBF.Custom Then
        mDataNav.RenameCriteriaFile QBF.ID, QBF.Name
    Else
        InfBox "You cannot rename the file for a Provided criteria", "!", , "Rename Criteria Error"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuRenameQBF_Click"
    
End Sub

Private Sub mnuShowButtons_Click()
On Error GoTo ErrSection:

    m.bShowButtons = Not m.bShowButtons
    PromptShowButtons
    FormResize Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuShowButtons_Click"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuLabelRow_Click
'' Description: Ask the user for a label, then add the row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuLabelRow_Click()
On Error GoTo ErrSection:

    Dim strLabel As String              ' Label for the label row
    
    strLabel = AskBox("h=Label ; i=? ; b=+OK ; g=string ; Please enter a label for the row (leave blank to insert a blank row")
    If ValidLabel(strLabel) Then
        AddLabelRow strLabel
        ColorQuoteRows
        
        'Reset the rows information in the data table...
        If m.frmActiveDetTab Is Nothing Then
            ResetRows
        Else
            TabStr(eGDTabSettings_Symbols, m.frmActiveDetTab.MyTabIndex) = m.frmActiveDetTab.MySymbols(True)
        End If
    End If

ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuLabelRow_Click"
    Resume ErrExit
      
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuQBF_Click
'' Description: If the user clicks on the Edit Quote Board Field menu item,
''              bring up the criteria editor to allow the user to change the
''              QBF
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuQBF_Click()
On Error GoTo ErrSection:

    Dim QBF As New cCriteria            ' Temporary holder for QBF object
    Dim frm As New frmCriteria          ' Criteria editor
    
    Set QBF = m.QBFs(fgQuotes.ColData(m.lQbfCol))
    If frm.ShowMe(AddSlash(App.Path) & "Custom\", QBF.ID, True, eCriteria_QuoteBoardField) <> "" Then
        m.QBFs(fgQuotes.ColData(m.lQbfCol)).FromFile AddSlash(App.Path) & "Custom\", QBF.ID
        TotalRefresh True
    End If

ErrExit:
    Set QBF = Nothing
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuQBF.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemoveSymbol_Click
'' Description: Allow the user to remove a symbol from the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemoveSymbol_Click()
On Error GoTo ErrSection:
    
    DoRemoveSymbol
    
ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuRemoveSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSell_Click
'' Description: Allow the user to sell the selected security
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSell_Click()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol to pass to create order
    
    Dim eStyle As eGDQuoteStyle
    Dim fgCurrGrid As VSFlexGrid
    Dim QB As cQuoteCellBoard
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
        Set QB = m.QB
        eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), vsTab.CurrTab)
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), m.frmActiveDetTab.MyTabIndex)
    End If
    
    If eStyle = eGDQuoteStyle_Grid And Not fgCurrGrid Is Nothing Then
        strSymbol = RollSymbolForDate(Parse(fgCurrGrid.TextMatrix(fgCurrGrid.Row, GDCol(eGDCol_Symbol)), "(", 1), Now)
    ElseIf Not QB Is Nothing Then
        If Not QB.Cell(QB.Row, QB.Col) Is Nothing Then
            strSymbol = RollSymbolForDate(Parse(QB.Cell(QB.Row, QB.Col).Symbol, "(", 1), Now)
        End If
    End If
    CreateOrder strSymbol, , 0, , , "Quote Board"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuSell_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSettings_Click
'' Description: Allow the user to change settings about the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSettings_Click()
On Error GoTo ErrSection:

    ShowSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuSettings.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuSortBySym_Click()
On Error GoTo ErrSection:

    cmdSortBySym_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.mnuSortBySym_Click"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_DblClick
'' Description: When the user double clicks on a cell, chart the symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_DblClick()
On Error GoTo ErrSection:

    If Not m.QB.Cell(m.QB.Row, m.QB.Col) Is Nothing Then
        SetActiveChartSymbol m.QB.Cell(m.QB.Row, m.QB.Col).SymbolID     '5338
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub pbQuoteBoard_GotFocus()
On Error GoTo ErrSection:
    
    m.QB.GotFocus
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_KeyDown
'' Description: When the user presses a key, send it off to the control class
'' Inputs:      Code of the key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    m.QB.KeyDown KeyCode, Shift
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub pbQuoteBoard_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    Select Case Chr(KeyAscii)
        Case "A" To "Z", "a" To "z", "$"
            If m.QB.Cell(m.QB.Row, m.QB.Col) Is Nothing Then
                DoAddSymbol m.QB.Row, m.QB.Col, Chr(KeyAscii)
            Else
                m.strSaveSymbol = m.QB.Cell(m.QB.Row, m.QB.Col).Symbol
                ChangeSymbolOnBox m.QB.Row, m.QB.Col, Chr(KeyAscii)
            End If
        
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.KeyPress", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_KeyUp
'' Description: Handle key presses accordingly
'' Inputs:      Code of the key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            DoRemoveSymbol
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.KeyUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub pbQuoteBoard_LostFocus()
On Error Resume Next

    m.QB.LostFocus

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_MouseDown
'' Description: When the user clicks the mouse, highlight the "mouse cell"
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    m.QB.MouseDown Button, Shift, X, Y
        
    If Button = vbRightButton Then
        m.QB.Row = m.QB.MouseRow
        m.QB.Col = m.QB.MouseCol
    
        mnuQBF.Visible = False
        mnuRenameQBF.Visible = False
        mnuFields.Visible = False
        mnuLabelRow.Visible = False
        
        ShowQuotesPopup X, Y
    Else
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.MouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_MouseMove
'' Description: Call the class version of mouse move
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    m.QB.MouseMove Button, Shift, X, Y

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_MouseUp
'' Description: Call the class version of MouseUp
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If m.QB.MouseUp(Button, Shift, X, Y) = 1& Then
        ResetRows
        m.QB.DrawBoard
        pbQuoteBoard.Refresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.pbQuoteBoard.MouseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    pbQuoteBoard_Resize
'' Description: When the quote board gets resized, redraw the quote cell board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub pbQuoteBoard_Resize()
On Error Resume Next

    m.QB.ResizeBoard False
    m.QB.DrawBoard
    pbQuoteBoard.Refresh

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRealTime_Timer
'' Description: Handles updating the quote board from real-time data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRealTime_Timer()
On Error GoTo ErrSection:
   
    'JM: Total Refresh is called from multiple places this static only account for Total Refresh being called from here
    'Static dLastTotal As Double         ' Time of the last TotalRefresh
    
    Dim dTime#
    Static dQuoteCount As Double        ' previous quote count
    Static bInProgress As Boolean       ' Are we already doing an update?
    Static dUpdateTime#, dUpdateCount#
    Static dLastFeedTime#, bPastFirstMinute As Boolean
    
    Dim nDetachTabs As Long
        
    If g.bUnloading Then Exit Sub
    
TimerStart "frmQuotes.tmrRealTime"
    
    ' when streaming is active, display the time in the main title
    If Not g.RealTime Is Nothing Then
        If g.RealTime.Active And g.nReplaySession = 0 Then
            dTime = g.RealTime.FeedTime
            If dTime > 0 Then
                ' if LastFeedTime was too long ago (e.g. reconnected), then clear it out
                If dTime > dLastFeedTime + 15 / 86400# Then
                    dLastFeedTime = 0
                    bPastFirstMinute = False
                ElseIf Second(dTime) <> Second(dLastFeedTime) Then
                    If bPastFirstMinute Then
                        SetMainCaption Format(ConvertTimeZone(dTime, "NY", ""), "hh:mm:ss")
                    Else
                        SetMainCaption Format(ConvertTimeZone(dTime, "NY", ""), "hh:mm")
                        If Minute(dTime) <> Minute(dLastFeedTime) Then
                            bPastFirstMinute = True
                        End If
                    End If
                End If
                dLastFeedTime = dTime
            End If
        End If
    End If
       
    ' Don't allow re-entering when hasn't exited from last call...
    ' or if Total Refresh has not exited
    If bInProgress Or m.nRefreshRow > 0 Then
        Exit Sub
    End If
       
    bInProgress = True
    
    CheckStreamStart
    
    If Not m.aDetachedTabs Is Nothing Then nDetachTabs = m.aDetachedTabs.Size
   
    ' See if enough time has elapsed...
    dTime = gdTickCount(False)
    If Now > m.dLastTotalRefresh + (5# / 86400#) Then
        TotalRefresh False
        ClearUpdatedColors
        dTime = gdTickCount(False) - dTime
        If dUpdateCount > 0 Then
            If frmTest.Visible Then
                'frmTest.AddList "UpdateTable = " & Str(Int(dUpdateTime / dUpdateCount)) & " ms avg (" _
                    & Str(dUpdateCount) & " times), TotalRefresh = " & Str(Int(dTime)) & " ms"
            End If
            dUpdateCount = 0
            dUpdateTime = 0
        End If
    ElseIf Me.Visible Or nDetachTabs > 0 Then       '5652
        ' if no quotes (trade,bid,ask) received since last checked,
        ' then no need to call UpdateTable
        fgQuotes.Redraw = flexRDNone
        If g.RealTime.QuoteCount > dQuoteCount Or g.RealTime.SalmonIsRunning Then
            dQuoteCount = g.RealTime.QuoteCount
            ' If UpdateTable returns true, then we need to do a
            ' TotalRefresh since one of the symbols has a new day
            If UpdateTable Then
                If Not g.bUnloading Then
                    TotalRefresh False
                End If
            Else
                g.Alerts.CheckAlerts False
            End If
        End If
        ClearUpdatedColors
        dTime = gdTickCount(False) - dTime
        dUpdateTime = dUpdateTime + dTime
        dUpdateCount = dUpdateCount + 1
    End If
        
    fgQuotes.Redraw = flexRDDirect
    
    g.RealTime.SetRTbutton
    
    WebPageCheck
           
    TimerEnd "frmQuotes.tmrRealTime", tmrRealtime.Interval
           
ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmQuotes.tmrRealTime.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsTab_GotFocus
'' Description: If the tab gets the focus, move the focus to the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsTab_GotFocus()
On Error GoTo ErrSection:

    MoveFocusToQb

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.vsTab_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsTab_KeyUp
'' Description: Handle certain special keystrokes
'' Inputs:      Code of the key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsTab_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            DoRemoveSymbol
            
        Case vbKeyInsert
            DoAddSymbol
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.vsTab.KeyUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Static lTab As Long
    
    If vsTab.MouseOver <> lTab Then
        lTab = vsTab.MouseOver
        If lTab < 0 Or lTab >= vsTab.NumTabs - 1 Then
            vsTab.ToolTipText = ""
        Else
            vsTab.ToolTipText = Str(NumSymbolsForTab(lTab)) & " symbols"
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsTab_MouseUp
'' Description: If the user right clicks on the tab, edit the categories
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lTab As Long
    Dim bDoEdit As Boolean

    If Button = 2 Then
        lTab = vsTab.MouseOver
        If lTab >= 0 And lTab < vsTab.NumTabs Then
'            bDoEdit = vsTab.TabCaption(lTab) <> "(new)"
            bDoEdit = vsTab.TabCaption(lTab) <> "(manage)"
            vsTab.CurrTab = lTab
            If bDoEdit Then EditCategories
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.vsTab.MouseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsTab_Switch
'' Description: As the user changes tabs, save info from old and set up new
'' Inputs:      Old Tab, New Tab, Whether to Cancel the Switch
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsTab_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

'    Dim strOldFilterID As String        ' Current Filter ID of the selected filter
'    Dim strNewFilterID As String        ' New Filter ID of the selected filter
'    Dim bRefresh As Boolean             ' Do we need to do a refresh?
'    Dim bFilterLoaded As Boolean        ' Did we just load a filter?

    Static bInProgress As Boolean
    
    If bInProgress Then Exit Sub        '6772

    If vsTab.Visible Then
        If IsBusy Then
            DoEvents
            If IsBusy Then
                InfBox "Please wait until the quote board|has finished refreshing...", "t", "+-OK", "Quote Board", True, 1&
                Cancel = True
                Exit Sub
            End If
        End If
            
        bInProgress = True ' will maybe help #6534?
            
'        If vsTab.TabCaption(NewTab) = "(new)" Then
        If vsTab.TabCaption(NewTab) = "(manage)" Then
            EditCategories      '5711 - rename to manage & just bring up edit tab form
            Cancel = True
        ElseIf OldTab <> NewTab Then
            If OldTab < vsTab.NumTabs Then
                ' Save the tab information...
                SaveTabInfo OldTab
                FileFromString AddSlash(App.Path) & "Custom\QuoteBoard.INF", m.tblTabInfo.ToString(vbLf, vbTab)
            End If
            If (vsTab.TabCaption(NewTab) = "(Filter)") Then
                If Len(TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1)) = 0 Then
                    ChangeFilter
                    If Len(TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1)) = 0 Then
                        'user cancelled from selection dialog, don't show filter tab with blank row
                        NewTab = m.lCurrentTab
                    End If
                End If
            End If
                        
            m.lCurrentTab = NewTab
            ShowCategory NewTab
                        
'            If bRefresh Then TotalRefresh False
'            If bFilterLoaded Then SortFilterTab NewTab
        
        End If
    End If

ErrExit:
    bInProgress = False
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmQuotes.vsTab.Switch", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview control for the quotes grid
'' Inputs:      Args to pass to print preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lAvailWidth As Long             ' Available width on the page
    Dim lAvailHeight As Long            ' Available height on the page
    Dim lPicHeight As Long              ' Picture height
    Dim lPicWidth As Long               ' Picture width
    Dim lResize As Long                 ' Resize amount for the picture
    Dim lSaveFontSize As Long           ' Font Size before printing
    Dim strFile As String               ' Name of the file to save off
    
    Dim lIndex As Long                  'variables for handling printing of detached QB tab
    Dim fg As VSFlexGrid
    Dim QB As cQuoteCellBoard
    Dim pB As PictureBox
    
    Dim bQbBox As Boolean
    
    lIndex = CurrentTab
    
    If m.frmActiveDetTab Is Nothing Then
        Set fg = fgQuotes
        Set QB = m.QB
        Set pB = pbQuoteBoard
    Else
        Set fg = m.frmActiveDetTab.fgQuotes
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        Set pB = m.frmActiveDetTab.pbQuoteBoard
    End If
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        If TabStr(eGDTabSettings_Style, lIndex) = Str(QStyle(eGDQuoteStyle_Grid)) Then
            If frmPrintPreview.GoingToFile Then
                .Text = Right(.Header, Len(.Header) - 1)        'strip of the leading | character
                .Paragraph = ""
                .Paragraph = ""
                frmPrintPreview.GridToTable fg
            Else
                .RenderControl = fg.hWnd
            End If
        Else
            bQbBox = True
            strFile = AddSlash(App.Path) & "BoxQB.BMP"
            gePrintQuoteWin QB.QuoteBoardObj, pB.hWnd, pB.hDC, 0, 0, QB.LastDataRow, QB.LastDataCol, strFile
            Picture1.Picture = LoadPicture(strFile)
            ' Redraw the diagram on the given device context
            lAvailWidth = .PageWidth - .MarginLeft - .MarginRight
            lAvailHeight = .PageHeight - .CurrentY - .MarginBottom
            .CalcPicture = Picture1.Picture
            lPicHeight = .Y2 - .Y1
            lPicWidth = .X2 - .X1
            
            .X1 = 0
            .Y1 = 0
            
            If lPicHeight > lAvailHeight And lPicWidth > lAvailWidth Then
                .DrawPicture Picture1.Picture, _
                     .MarginLeft, .CurrentY, lAvailWidth, lAvailHeight, vppaZoom
            ElseIf lPicHeight > lAvailHeight Then
                lResize = (lPicWidth * (1 - (lAvailHeight / lPicHeight))) / 2
                .DrawPicture Picture1.Picture, _
                     .MarginLeft - lResize, .CurrentY, lPicWidth, lAvailHeight, vppaZoom
            ElseIf lPicWidth > lAvailWidth Then
                lResize = (lPicHeight * (1 - (lAvailWidth / lPicWidth))) / 2
                .DrawPicture Picture1.Picture, _
                     .MarginLeft, .CurrentY - lResize, lAvailWidth, lPicHeight, vppaZoom
            Else
                .DrawPicture Picture1.Picture, _
                     .MarginLeft, .CurrentY, lPicWidth, lPicHeight, vppaZoom
            End If
        End If
        
        .EndDoc
    End With

    If bQbBox Then
        'must do this outside the startdoc/enddoc block otherwise will get file sharing error
        With frmPrintPreview
            If .GoingToFile Then
                If .vp.ExportFile = "BOXQBTOCLIPBOARD" Then
                    geSaveQuoteWin QB.QuoteBoardObj, pB.hWnd, pB.hDC, "", Right(.vp.Header, Len(.vp.Header) - 1)
                Else
                    geSaveQuoteWin QB.QuoteBoardObj, pB.hWnd, pB.hDC, .vp.ExportFile, Right(.vp.Header, Len(.vp.Header) - 1)
                End If
            End If
        End With
    End If

ErrExit:
    If Len(strFile) > 0 Then KillFile strFile
    Exit Sub
    
ErrSection:
    If Len(strFile) > 0 Then KillFile strFile
    RaiseError "frmQuotes.GenerateReport", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Bring up the Print Preview form to allow the user to print
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection
    
    Dim lIndex As Long
    
    lIndex = CurrentTab
    
    If TabStr(eGDTabSettings_Style, lIndex) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        PrintMe = frmPrintPreview.ShowMe("CNV QuoteBoard", frmQuotes, , 0.75, 0.75, 0.75, 0.75, True)
    Else
        PrintMe = frmPrintPreview.ShowMe("CNV QuoteBoard", frmQuotes, , 0.75, 0.75, 0.75, 0.75, True, , ePrintToFile_Image)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.PrintMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TotalRefresh
'' Description: Do a total refresh of the quote board with the daily bars,
''              snapshot ticks, and real time ticks and recalculate the QBF's
'' Inputs:      Whether or not to reload daily data
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TotalRefresh(ByVal bReloadData As Boolean)
On Error GoTo ErrSection:
    
    Dim lCol As Long                    ' Index into a for loop
    Dim hArray As Long                  ' Array handle
    Dim lMaxBars As Long                ' Number of bars of data to load
    Dim lNumDays As Long                ' Number of days of data to load
    Dim lFromDate As Long               ' Date to load data from
    Dim lLastDateOfData As Long         ' Last date of data
    Dim Bars As cGdBars                 ' Data for the main market
    Dim GC As New cGdBars               ' Data for Gold 67 contract
    Dim Daily As New cGdBars            ' Data for daily bars of main market
    Dim Weekly As New cGdBars           ' Data for weekly bars of main market
    Dim Monthly As New cGdBars          ' Data for monthly bars of main market
    Dim rc As Long                      ' Return code from function calls
    Dim strSymbol As String             ' Symbol to get data for
    Dim lSymbolID As Long               ' Symbol ID to get data for
    Dim dLastDate As Double             ' Last date of data
    Dim lRow As Long                    ' Index into a for loop
    Dim lRow2 As Long                   ' Index into a for loop
    Dim i As Long
    Dim dValue As Double                ' Price from the results array
    Dim strValue As String
    Dim bReloadThisSymbol As Boolean
    Dim bNewBar As Boolean
    Dim strMonth As String
    Dim strPeriod As String
    Dim nPeriodicity As Long
    Dim lTblField As Long               ' Field number for a criteria
    Dim nSaveUseUpdateColor As Long
    
    Dim astrParms As New cGdArray       ' Paramaters array for the engine
    Dim astrBarNames As New cGdArray    ' Array of bar names
    Dim aScanExpr As New cGdArray       ' Array of coded text expressions
    Dim aArrayOfResults As New cGdArray ' Array of results
    Dim aArrayOfBars As New cGdArray    ' Array of bars structures
    Dim aScanArrays As New cGdArray     ' Array of results for a scan
    Dim aQbfNames As New cGdArray
    Dim aTemp As New cGdArray
    Dim Criteria As cCriteria
    
    Dim bDoInit As Boolean
    Dim bDoEvents As Boolean
    Dim bUsingStatusMsg As Boolean
    
    Dim SecondaryMarkets As New cGdTree ' Bars collection of secondary markets
    Dim lBars As Long                   ' Index into a for loop
    Dim strNewSymbol As String          ' Potentially new symbol
    
    Static bRestart As Boolean
    Static bDetachedDone As Boolean
                                
Dim bDebugRefreshRow As Boolean
If FileExist(App.Path & "\RefreshRow.flg") Then
    bDebugRefreshRow = frmTest.Visible
End If
If bDebugRefreshRow Then
    frmTest.AddList "1: RefreshRow = " & Str(m.nRefreshRow)
End If
                                
    ' check if this routine already in progress
    If m.nRefreshRow > 0 Then
        ' if needing to reload data, then set Restart flag, else just quit
'        If bReloadData Then bRestart = True
        Exit Sub
    End If
    
    CheckStreamStart
    
    If bReloadData Then
        If frmStatus.Status = eStatus_Running Then
            If InStr(UCase(frmStatus.Caption), "RETRIEVING DATA") <> 0 Then
                frmStatus.AddDetail "Updating Quote Board..."
            End If
        End If
    End If
#If 0 Then ' TLB 10/7/2009: don't check this anymore -- it sometimes causes the QB to stop updating
    Else
        ' check if any new ticks since last time it was checked
        For lRow = 0 To m.QBData.NumRecords - 1
            If TblNum(eQbTbl_Recalc, lRow) = 1 Then
                bDoInit = True
                Exit For
            End If
        Next lRow
    End If
    If Not bDoInit And Not bReloadData Then Exit Sub
#End If
    
    m.nRefreshRow = 1 ' (set flag: just need to set > 0 for now)
    
If bDebugRefreshRow Then
    frmTest.AddList "2: RefreshRow = " & Str(m.nRefreshRow)
End If
'    Screen.MousePointer = vbHourglass

    ' Create the arrays
    aScanExpr.Create eGDARRAY_Strings
    aScanArrays.Create eGDARRAY_Longs
    aArrayOfResults.Create eGDARRAY_Longs
    lMaxBars = 0
    
    ' Create
    With m.QBFs
        For lCol = 1 To .Count
            With .Item(lCol)
                If Len(Trim(.CodedText)) > 0 Then
                    aScanExpr.Add Trim(.CodedText)
                    aQbfNames.Add .Name
                    
                    ' get values array handle, clear array (so no longer
                    ' a const array), pre-size array, and store handle
                    hArray = .ValuesArray.ArrayHandle
                    gdClear hArray, True
                    'gdSetSize hArray, fgQuotes.Rows - fgQuotes.FixedRows - 1, False
                    gdSetSize hArray, m.QBData.NumRecords, False
                    aScanArrays.Add hArray
                    
                    ' create a temporary result array to be used
                    ' by the expression evaluator
                    hArray = gdCreateArray(eGDARRAY_Doubles, lMaxBars)
                    aArrayOfResults.Add hArray
                    
                    ' see if NumDays bigger
                    If .NumDays = 0 Then
                        .NumDays = frmCriteria.AutoDetect(.CodedText)
                        .ToFile
                    End If
                    If .NumDays > lMaxBars Then lMaxBars = .NumDays
                End If
            End With
        Next lCol
    End With

    ' calc FromDate, adjusting for weekends and holidays
    ' (need to fudge a little to the safe side)
    If lMaxBars < 2 Then lMaxBars = 2
    lLastDateOfData = LastDailyDownload
    'lFromDate = lLastDateOfData - Int(lMaxBars * 1.46 + 0.5) - 2
    lFromDate = lLastDateOfData - Int(lMaxBars * 1.6 + 0.5) - 2
    
    astrBarNames(0) = "Market1"
    astrBarNames(1) = "Daily"
    astrBarNames(2) = "Weekly"
    astrBarNames(3) = "Monthly"
    MarketsInExpressions aScanExpr, lFromDate, True, astrBarNames, SecondaryMarkets, "Daily"

    If aScanExpr.Size > 0 Then
        'DM_GetBars GC, "GC-067", , lFromDate, 0
        ' Init the expression evaluator with list of scan expressions
        'astrBarNames(0) = "Market1"
        'astrBarNames(1) = "Weekly"
        'astrBarNames(2) = "GC"
        astrParms(0) = "TotalRefresh"   ' 0) Expression Set Name
        If Not SetupExpressions(astrParms, astrBarNames, aScanExpr) Then
            ' try to find out which one's bombing
            For i = 0 To aScanExpr.Size - 1
                aTemp(0) = aScanExpr(i)
                aTemp.Size = 1
                If Not SetupExpressions(astrParms, astrBarNames, aTemp) Then
                    InfBox "i=[] ; h=Calculate QBF ; An error exists in the Quote Board Field expression for:|" _
                        & aQbfNames(i)
                    GoTo ErrExit
                End If
            Next
            InfBox "i=[] ; h=Calculate QBF ; An error exists in a Quote Board Field expression."
            GoTo ErrExit
        End If
        
        aArrayOfBars.Create eGDARRAY_Longs
    End If
    
    ' TLB 2/11/2014: init this flag now just in case no symbols on the quote board
    If frmStatus.Status <> eStatus_Running Then
        bUsingStatusMsg = True
    End If
    
    'TLB 7/8/02: just buffer the redraw (total refresh can be
    ' slow enough that the grid's refresh is too long between)
    fgQuotes.Redraw = flexRDBuffered ' = flexRDNone
    'For lRow = fgQuotes.FixedRows To fgQuotes.Rows - 1

    For lRow2 = 0 To m.QBData.NumRecords - 1
        If g.bUnloading Then GoTo ErrExit
        
        ' Get the index into the table from the sorted index...
        lRow = gdGetNum(m.hDataIndex, lRow2)
        
        ' see if Restart flag got set (if so, then start over)
        If bRestart Then
            bRestart = False
            bReloadData = True '(make sure flag is set since need to reload data)
            'lRow = fgQuotes.FixedRows
            lRow = 0&
        End If
        
        'If fgQuotes.MergeRow(lRow) = False Then
            ' set flag to current row (so UpdateTable will skip it)
            m.nRefreshRow = lRow
If bDebugRefreshRow Then
    frmTest.AddList "3: RefreshRow = " & Str(m.nRefreshRow)
End If
            'If lRow >= fgQuotes.Rows Then Exit For '(need this for some reason!)
            If lRow >= m.QBData.NumRecords Then Exit For '(need this for some reason!)
            
            'strSymbol = Parse(fgQuotes.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)
            'lSymbolID = CLng(fgQuotes.TextMatrix(lRow, GDCol(eGDCol_SymbolID)))
            'strPeriod = fgQuotes.TextMatrix(lRow, GDCol(eGDCol_Period))
            
            strSymbol = TblStr(eQbTbl_Symbol, lRow)
            strPeriod = Parse(TblStr(eQbTbl_SearchKey, lRow), vbTab, 2)
            'lSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
            lSymbolID = TblNum(eQbTbl_SymbolID, lRow)
            
            strNewSymbol = GetSymbol(lSymbolID)
            If strNewSymbol <> strSymbol And Len(strNewSymbol) > 0 Then
                TblStr(eQbTbl_Symbol, lRow) = strNewSymbol
                strSymbol = strNewSymbol
            End If
            
            'If VarType(fgQuotes.RowData(lRow)) = vbObject Then
            '    Set Bars = fgQuotes.RowData(lRow)
            'Else
            '    Set Bars = Nothing
            'End If
            Set Bars = m.BarsColl(TblStr(eQbTbl_SearchKey, lRow))
                    
            bNewBar = False
            bReloadThisSymbol = False
            If bReloadData Or (Bars Is Nothing) Then
                ' we'll need to reload data
                bReloadThisSymbol = True
            ElseIf Bars.ArrayMask = eBARS_Ask Then
                bReloadThisSymbol = True '(not loaded yet)
            Else
                ' try updating the existing bars
                If g.RealTime.UpdateBars(Bars, bNewBar) Then
                    'CheckedCell(fgQuotes, lRow, GDCol(eGDCol_Recalc)) = True
                    TblNum(eQbTbl_Recalc, lRow) = 1
                    TblNum(eQbTbl_Dirty, lRow) = 1
                End If
                ' but if updated ticks are in a new bar (next day), then
                ' we'll need to do a reload for this symbol
                If bNewBar Then
                    bReloadThisSymbol = True
                End If
            End If
            
            If bReloadThisSymbol Then
                ' reload the data for this symbol
                bDoEvents = True
                
                nPeriodicity = GetPeriodicity(strPeriod)
                
                ' When a new bar starts for an intraday bar period, all we need to do is
                ' a SpliceBars (06/11/2008 DAJ)...
                If (bNewBar = False) Or (nPeriodicity >= ePRD_Days) Then
                    Set Bars = New cGdBars
                    'Select Case fgQuotes.TextMatrix(lRow, GDCol(eGDCol_SecType))
                    Select Case TblStr(eQbTbl_SecType, lRow)
                        Case "SO", "FO"
                            Bars.ArrayMask = eBARS_Prices Or eBARS_VolOI Or eBARS_BidAsk
                        Case "F"
                            Bars.ArrayMask = eBARS_Eod Or eBARS_BidAsk
                        Case "S"
                            Bars.ArrayMask = eBARS_Prices Or eBARS_Vol Or eBARS_BidAsk
                        Case "I"
                            If Not IsForex(strSymbol) Then
                                Bars.ArrayMask = eBARS_Prices 'Or eBARS_Vol
                            ElseIf Right(strSymbol, 4) = "@PFG" Then
                                ' treat PFG Forex like a future (has bid/ask sizes, volumes, etc)
                                Bars.ArrayMask = eBARS_Eod Or eBARS_BidAsk
                            Else
                                Bars.ArrayMask = eBARS_Prices Or eBARS_Bid Or eBARS_Ask
                            End If
                        Case Else
                            Bars.ArrayMask = eBARS_Eod Or eBARS_BidAsk
                    End Select
                    'fgQuotes.RowData(lRow) = Bars
                    Set m.BarsColl(TblStr(eQbTbl_SearchKey, lRow)) = Bars
    
                    ' Instead of adding all of the fluff here, we can now pass the number of bars
                    ' that we want to DM_GetBars by passing in a negative number for the from date,
                    ' so that is what we are going to do here (it helps for allowing periods greater
                    ' than daily on the quote board) (DAJ 04/30/2008)...
                    lNumDays = lMaxBars * -1&
                    If lNumDays < 0 Then
                        lFromDate = lNumDays
                    Else
                        If lNumDays < 2 Then lNumDays = 2
                        lFromDate = lLastDateOfData - lNumDays - 2
                    End If
                    
'frmTest.AddList "Loading " & Bars.Prop(eBARS_Symbol) & "  " & CStr(lFromDate), True
                    If Not DM_GetBars(Bars, strSymbol, nPeriodicity, lFromDate, , , False) Then
                        Bars.Size = 0
                    Else
                        ' append contract month to continuous contract for display
                        strMonth = RollSymbolForDate(strSymbol, Bars(eBARS_DateTime, Bars.Size - 1))
                        If strMonth <> strSymbol And Len(strMonth) > 0 Then
                            strMonth = MonthName(Val(Right(strMonth, 2)), True)
                            If Len(strMonth) > 0 Then
                                'fgQuotes.TextMatrix(lRow, GDCol(eGDCol_Symbol)) = strSymbol & " (" & strMonth & ")"
                            End If
                        End If
                    End If
        
'frmTest.AddList "Loaded " & Bars.Prop(eBARS_Symbol) & "  " & CStr(Bars.Size), True
                    g.RealTime.AddTickBuffer Bars, False
                End If
                
                g.RealTime.SpliceBars Bars

                'CheckedCell(fgQuotes, lRow, GDCol(eGDCol_Recalc)) = True
                TblNum(eQbTbl_Recalc, lRow) = 1
                TblNum(eQbTbl_Dirty, lRow) = 1
            End If
            
            ' TLB 8/25/2009: need to check the delay more often (since the info can come seconds after it's loaded)
            i = g.RealTime.SymbolDelay(strSymbol)
            If i > 0 Then
                TblStr(eQbTbl_Delay, lRow) = Str(i)
            Else
                TblStr(eQbTbl_Delay, lRow) = ""
            End If
                            
            'If Bars.Size > 0 And aScanExpr.Size > 0 And CheckedCell(fgQuotes, lRow, GDCol(eGDCol_Recalc)) = True Then
                    'Weekly.BuildBars "Weekly", Bars.BarsHandle
                
            ' DAJ 06/11/2015: If this is a custom index, don't display the OHL because it could be
            ' invalid if it is being used as a spread symbol...
            If lSymbolID < 0 Then
                Bars(eBARS_Open, Bars.Size - 1) = kNullData
                Bars(eBARS_High, Bars.Size - 1) = kNullData
                Bars(eBARS_Low, Bars.Size - 1) = kNullData
            End If
            
            If Bars.Size > 0 And aScanExpr.Size > 0 And TblNum(eQbTbl_Recalc, lRow) = 1 And Len(TblStr(eQbTbl_Criteria, lRow)) > 0 Then
                Daily.BuildBars "Daily", Bars.BarsHandle
                Weekly.BuildBars "Weekly", Daily.BarsHandle
                Monthly.BuildBars "Monthly", Daily.BarsHandle
                
                ' make sure date of last bar is within last 5 days
                dLastDate = Bars(eBARS_DateTime, Bars.Size - 1)
                If dLastDate >= lLastDateOfData - 5 Then
                    bDoEvents = True
                    
                    ' run engine to evaluate expressions for this symbol
                    aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
                    aArrayOfBars.Num(1) = Daily.BarsHandle
                    aArrayOfBars.Num(2) = Weekly.BarsHandle
                    aArrayOfBars.Num(3) = Monthly.BarsHandle
                    'aArrayOfBars.Num(2) = GC.BarsHandle
                    For lBars = 4 To astrBarNames.Size - 1
                        aArrayOfBars.Num(lBars) = SecondaryMarkets(lBars + 1).BarsHandle
                    Next lBars
                    astrParms.Size = 1
                    rc = RunExpressions(astrParms.ArrayHandle, _
                        astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                        aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
                    If rc <> 0 Then
                        'Me.Caption = CStr(rc) & ": " & aStrParms(aStrParms.Size - 1)
                    Else
                        ' set current value for each expression
                        For i = 0 To aScanArrays.Size - 1
                            ' get most recent value
                            hArray = aArrayOfResults.Num(i)
                            dValue = gdGetNum(hArray, Bars.Size - 1)
    
                            ' store into the scan's array for this symbol
                            hArray = aScanArrays.Num(i)
                            If dValue = kNullData Then
                                dValue = gdNullValue(hArray)
                            End If
                            gdSetNum hArray, lRow, dValue
                        Next
                    End If
                End If
                TblNum(eQbTbl_Recalc, lRow) = 0
                TblNum(eQbTbl_Dirty, lRow) = 1
            End If
                
            ' Update all of the columns in the row
            ' TLB: need to only update cells if dirty or should have recalced
            If TblNum(eQbTbl_Recalc, lRow) = 1 Or TblNum(eQbTbl_Dirty, lRow) = 1 Then
                ' TLB 8/24/2009: just to avoid the "all blue" visual effect when streaming has just started,
                ' temporarily turn off the update color for a symbol with a new bar within 30 seconds of
                ' starting or when loading a new symbol and we still only have the old data
                nSaveUseUpdateColor = kNullData
                If g.RealTime.Active Then
                    If Bars.SessionDate(Bars.Size - 1) <= LastDailyDownload Then
                        nSaveUseUpdateColor = m.QB.UseUpdateColor
                        m.QB.UseUpdateColor = 0
                    ElseIf bNewBar Then
                        If gdTickCount <= m.dWhenStreamingStarted + 30000 Then
                            nSaveUseUpdateColor = m.QB.UseUpdateColor
                            m.QB.UseUpdateColor = 0
                        End If
                    End If
                End If
                
                For lCol = 1 To m.QBFs.Count
                    Set Criteria = m.QBFs(lCol)
                    lTblField = m.QBData.FieldNum(Criteria.ID)
                    TblStr(lTblField, lRow) = Criteria.ValuesArray(lRow)
                Next lCol
                Set Criteria = Nothing
                If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                    UpdateSymbols lRow, False, bReloadThisSymbol, Bars
                ElseIf Bars.Prop(eBARS_Periodicity) = ePRD_Days + 1 Then
                    m.QB.UpdateSymbol Bars, False
                    If Not g.RealTime.Active And m.aDetachedTabs.Size > 0 Then
                        UpdateSymbols lRow, False, bReloadThisSymbol, Bars      '5271
                    End If
                End If
                
                TblNum(eQbTbl_Dirty, lRow) = 0
                
                If nSaveUseUpdateColor <> kNullData Then
                    m.QB.UseUpdateColor = nSaveUseUpdateColor
                End If
                
''                    If FormIsLoaded("frmTTSummary") Then frmTTSummary.RefreshPrices Bars
            End If
        
            ' Refresh if past visible part of grid
            ' (so will seem like it shows faster)
            If bReloadData Then
                'If .Rows > 5 Then
                'If m.QBData.NumRecords > 15 And (lRow2 Mod 5 = 4 Or lRow2 >= m.QBData.NumRecords - 1) Then
                    If frmStatus.Status = eStatus_Running Then
                        frmStatus.UpdateProgress "Reloading Data", 100 * (lRow2 + 1) / m.QBData.NumRecords
                    ElseIf (lRow2 + 1) Mod 10 = 0 Or lRow2 + 1 = m.QBData.NumRecords Then
                        'strValue = "Loading " & CStr(lRow) & " of " & CStr(.Rows) & " quote symbols"
                        strValue = "Loading " & CStr(lRow2 + 1) & " of " & CStr(m.QBData.NumRecords) & " quote symbols"
                        StatusMsg strValue, vbRed
                        bUsingStatusMsg = True
                    End If
                'End If
                If lRow = fgQuotes.BottomRow + 1 And m.lTotWidth = 0 Then
                    fgQuotes.AutoSize 0, fgQuotes.Cols - 1, False, 75
                    fgQuotes.Redraw = flexRDBuffered
                    fgQuotes.Refresh
                End If
            End If
            Set Bars = Nothing
        'End If
                
        ' We'll yield to other threads only every 1/2 second
        'Sleep -0.5
        If bDoEvents Then
            'TLB: do we need this anymore?
            'DoEvents
            bDoEvents = False
        End If
        ' Turn back off, since might have been turned on by
        ' the timer's call to UpdateTable
        ''fgQuotes.Redraw = flexRDNone
    Next lRow2

    With fgQuotes
        If bReloadData Then
            'Me.Caption = "Quotes"
            'SetMainCaption
            ShowDelayColumn
            If bUsingStatusMsg Then StatusMsg
            If m.lTotWidth = 0& Then .AutoSize 0, .Cols - 1, False, 75
        End If
        .Redraw = flexRDBuffered
    End With
        
    If aScanExpr.Size > 0 Then
        ' clear the expression evaluator when done with it
        SetupExpressions astrParms '(clear expressions)
    End If
    
   ' Destroy all the temporary result arrays
    For i = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(i)
    Next
    aArrayOfResults.Size = 0
    
    If Not bDetachedDone Then
        Dim iSaveTab As Long
        Dim bRestoreTab As Boolean
        
        iSaveTab = vsTab.CurrTab
        'show detached tabs
        For i = 0 To m.tblTabInfo.NumRecords - 1
            If m.tblTabInfo(TabField(eGDTabSettings_Form), i) <> 0 Then
                DetachTab i
                bRestoreTab = True
            End If
        Next
        If bRestoreTab Then ShowCategory iSaveTab
        bDetachedDone = True
    End If
         
    'do not call alerts collection updatebars when streaming is on because the trade console (frmttsummary) does this in its timer
    If Not g.RealTime.Active And bReloadData Then g.Alerts.UpdateBars       '5200
    g.Alerts.CheckAlerts bReloadData, , , True
                
    If bReloadData Then
        ClearUpdatedColors
    End If
    
ErrExit:
    m.dLastTotalRefresh = Now
    ' clear the flags
    m.nRefreshRow = 0
If bDebugRefreshRow Then
    frmTest.AddList "end: RefreshRow = " & Str(m.nRefreshRow) & vbTab & Format(m.dLastTotalRefresh, "hh:mm:ss")
End If
    Exit Sub
    
ErrSection:
    ' clear the flags
    m.nRefreshRow = 0
    RaiseError "frmQuotes.TotalRefresh", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateCols
'' Description: Update each of the columns in the row with the data from the
''              Bars that is stored in the row data for the given row
'' Inputs:      Row to update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateCols(ByVal lRow As Long, Optional ByVal Bars As cGdBars = Nothing, Optional ByVal bOnlyIfVisible As Boolean = True, _
    Optional ByVal bRefreshSymbol As Boolean = False, Optional fg As VSFlexGrid = Nothing)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Column to update
    Dim strFormat As String             ' Format for the prices
    Dim strSecType As String            ' Security type for this row
    Dim lRedraw As Long                 ' Previous state of the grid redraw
    Dim dValue#, dPrev#, dHigh#, dLow#
    Dim strValue As String
    Dim bCurrentSession As Boolean
    Dim lTblPos As Long                 ' Position in the table
    Dim strSymbol As String             ' Symbol to update
    Dim strPeriod As String             ' Period to update
    Dim lTblField As Long               ' Field number for the criteria
    Dim strMonth As String
    Dim hBars As Long
    Dim lItem As Long
    Dim lColor As Long
    Dim lDeltaColor As Long
    Dim Criteria As New cCriteria
    Dim lDelay As Long
    Dim lSymbolID As Long               ' Symbol ID for the row
    Dim strNewSymbol As String          ' Potentially new symbol
    Dim fgCurrGrid As VSFlexGrid

    If fg Is Nothing Then
        Set fgCurrGrid = fgQuotes
    Else
        Set fgCurrGrid = fg
    End If
    
    With fgCurrGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
                
        If ((lRow >= .TopRow) And (lRow <= .BottomRow)) Or (bOnlyIfVisible = False) Then
            'verify that symbol name in symbol column is for symbolID in symbol ID column
            lSymbolID = Val(.TextMatrix(lRow, GDCol(eGDCol_SymbolID)))
            strSymbol = Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)
            strPeriod = .TextMatrix(lRow, GDCol(eGDCol_Period))
            strSecType = .TextMatrix(lRow, GDCol(eGDCol_SecType))
            strNewSymbol = GetSymbol(lSymbolID)
            If strNewSymbol <> strSymbol And Len(strNewSymbol) > 0 Then
                .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = strNewSymbol
                strSymbol = strNewSymbol
            End If
            
            'get row index into table to access fields to be displayed
            TblSearch SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod, lTblPos
            'get bars object if not passed in
            If Bars Is Nothing Then Set Bars = GetBars(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod)
            
            If Not Bars Is Nothing Then
                hBars = Bars.BarsHandle
                lItem = gdGetSize(hBars) - 1
                'dPrev = gdBarsData(hBars, eBARS_Close, lItem - 1)
                dPrev = GetPrevCloseForQB(Bars)
                
                ' Append contract month to continuous contract for display...
                'If (strSecType = "F") And (Bars.Size > 0) And ((InStr(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(") = 0) Or (bRefreshSymbol = True)) Then
                ' TLB 9/10/2009: with salmon, the 67 contract can change when data comes in (without a reload)
                If strSecType = "F" And Bars.Size > 0 Then
                    If InStr(strSymbol, "-0") > 0 Then
                        strMonth = RollSymbolForDate(strSymbol, Bars(eBARS_DateTime, Bars.Size - 1))
                        If strMonth <> strSymbol And Len(strMonth) > 0 Then
                            strMonth = MonthName(Val(Right(strMonth, 2)), True)
                            If Len(strMonth) > 0 Then
                                .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = strSymbol & " (" & strMonth & ")"
                            End If
                        End If
                    End If
                End If
                
                If strSecType = "S" Or strSecType = "I" Then
                    strFormat = "#0.00"
                Else
                    strFormat = "#0.0####"
                End If
                
                'For lCol = GDCol(eGDCol_NumFixed) To .Cols - 1
                For lCol = GDCol(eGDCol_Delay) To .Cols - 1
                    strValue = ""
                    Select Case UCase(.TextMatrix(0, lCol))
                        Case "SESSION"
                            dValue = Bars.SessionDate(lItem) ' gdBarsData(hBars, eBARS_DateTime, lItem)
                            If dValue > 0 Then
                                strValue = DateFormat(dValue, MM_DD_YY)
                                If dValue > LastDailyDownload Then
                                    bCurrentSession = True
                                End If
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "OPEN"
                            dValue = gdBarsData(hBars, eBARS_Open, lItem)
                            If dValue <> kNullData Then
                                strValue = Bars.PriceDisplay(dValue, True)
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "HIGH"
                            dValue = gdBarsData(hBars, eBARS_High, lItem)
                            If dValue <> kNullData Then
                                strValue = Bars.PriceDisplay(dValue, True)
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "LOW"
                            dValue = gdBarsData(hBars, eBARS_Low, lItem)
                            If dValue <> kNullData Then
                                strValue = Bars.PriceDisplay(dValue, True)
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "LAST"
                            dValue = gdBarsData(hBars, eBARS_Close, lItem)
                            If dValue <> kNullData Then
                                strValue = Bars.PriceDisplay(dValue, True)
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "PREV CLOSE"
                            If dPrev <> kNullData Then
                                strValue = Bars.PriceDisplay(dPrev, True)
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "CHANGE"
                            dValue = gdBarsData(hBars, eBARS_Close, lItem)
                            If dValue <> kNullData And dPrev <> kNullData Then
                                strValue = Bars.PriceDisplay(dValue - dPrev, True)
                                If Left(strValue, 1) <> "-" Then strValue = "+" & strValue
                            End If
                            ChangeCell lRow, lCol, strValue, True, fg
                            lDeltaColor = .Cell(flexcpForeColor, lRow, lCol)
                        Case "% CHANGE"
                            dValue = gdBarsData(hBars, eBARS_Close, lItem)
                            ' TLB 10/27/2008: Percent Change only makes sense for values which stay positive
                            If dValue > 0 And dPrev > 0 Then
                                strValue = Format((dValue - dPrev) / dPrev, "+#0.00%;-#0.00%")
                            End If
                            ChangeCell lRow, lCol, strValue, True, fg
                        Case "% OF H-L"
                            dValue = gdBarsData(hBars, eBARS_Close, lItem)
                            dHigh = gdBarsData(hBars, eBARS_High, lItem)
                            dLow = gdBarsData(hBars, eBARS_Low, lItem)
                            If dValue <> kNullData And dHigh > dLow And dValue >= dLow And dValue <= dHigh Then
                                strValue = Format((dValue - dLow) / (dHigh - dLow), "#0%")
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "BID"
                            dValue = gdBarsData(hBars, eBARS_Bid, lItem)
                            If dValue > kNullData Then
                                strValue = Bars.PriceDisplay(dValue, True)
                                If .ColWidth(lCol) < 600 Then
                                    .ColWidth(lCol) = 800
                                End If
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "BID SIZE"
                            dValue = gdBarsData(hBars, eBARS_BidSize, lItem)
                            If dValue > 0 Then
                                strValue = Format(dValue, "#,##0")
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "ASK"
                            dValue = gdBarsData(hBars, eBARS_Ask, lItem)
                            If dValue > kNullData Then
                                strValue = Bars.PriceDisplay(dValue, True)
                                If .ColWidth(lCol) < 600 Then
                                    .ColWidth(lCol) = 800
                                End If
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "ASK SIZE"
                            dValue = gdBarsData(hBars, eBARS_AskSize, lItem)
                            If dValue > 0 Then
                                strValue = Format(dValue, "#,##0")
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "VOLUME"
                            dValue = gdBarsData(hBars, eBARS_Vol, lItem)
                            If dValue > 0 Then
                                strValue = Format(dValue, "#,##0")
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "OPEN INTEREST"
                            dValue = gdBarsData(hBars, eBARS_OI, lItem)
                            If dValue > 0 Then
                                strValue = Format(dValue, "#,##0")
                            Else
                                dValue = gdBarsData(hBars, eBARS_ContOI, lItem)
                                If dValue > 0 Then
                                    strValue = Format(dValue, "#,##0")
                                End If
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "LAST TICK"
                            If Bars(eBARS_DateTime, Bars.Size - 1) > 0 Then
                                dValue = Bars.Prop(eBARS_LastTickTime)
                                If dValue <> 0 Then
                                    dValue = Int(Bars(eBARS_DateTime, Bars.Size - 1)) + dValue / 1440#
                                    If g.bShowInLocalTimeZone Then
                                        dValue = ConvertTimeZone(dValue, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                                    End If
                                    dValue = gdFixDateTime(dValue)
                                    'If tmrRealTime.Enabled Then
                                    If Second(dValue) <> 0 Or InStr(4, .TextMatrix(lRow, lCol), ":") > 0 Then
                                        strValue = Format(dValue, "hh:mm:ss")
                                    Else
                                        strValue = Format(dValue, "hh:mm")
                                    End If
                                End If
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        Case "EXCHANGE"
                            strValue = Bars.Prop(eBARS_Exchange)
                            ChangeCell lRow, lCol, strValue, , fg
                            
                        Case "FEED SYMBOL"
                            strValue = TblStr(eQbTbl_FeedSymbol, lTblPos)
                            ChangeCell lRow, lCol, strValue, , fg
                            
                        Case "DELAY"
                            If g.RealTime.Active Then
                                lDelay = g.RealTime.SymbolDelay(strSymbol)
                                If lDelay <= 0 Then
                                    strValue = ""
                                Else
                                    strValue = Str(lDelay) & " min"
                                    ' TLB 8/24/2009: if still hidden, need to show the delay column now
                                    If .ColHidden(GDCol(eGDCol_Delay)) Then
                                        .ColHidden(GDCol(eGDCol_Delay)) = False
                                    End If
                                End If
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                            '.Cell(flexcpForeColor, lRow, lCol) = vbBlue
                        
                        Case "DESCRIPTION"
                            strValue = Bars.Prop(eBARS_Desc)
                            ChangeCell lRow, lCol, strValue, , fg
                        
                        Case "T"
                            If Bars.Prop(eBARS_PriceHasSettled) <> 0 And Bars.SecurityType = "F" Then
                                strValue = "s"
                            ElseIf g.RealTime.Active = True Then
                                If Bars.Prop(eBARS_LastTickDown) = 0 Then
                                    strValue = "+"
                                Else
                                    strValue = "-"
                                End If
                            Else
                                strValue = ""
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                        
                        Case Else 'color for Custom Quote Board Fields
                            If Len(.ColData(lCol)) > 0 And lTblPos >= 0 Then
                                lTblField = m.QBData.FieldNum(.ColData(lCol))
                                dValue = TblNum(lTblField, lTblPos)
                                Set Criteria = m.QBFs(.ColData(lCol))
                                
                                If dValue = kNullData Then
                                    strValue = ""
                                Else
                                    If Criteria.IsBoolean Then
                                        Select Case dValue
                                            Case -128
                                                strValue = ""
                                            Case 0
                                                strValue = "False"
                                            Case Else
                                                strValue = "TRUE"
                                        End Select
                                    ElseIf InStr(UCase(Criteria.Name), "DATE") > 0 Then
                                        strValue = DateFormat(dValue) ' TLB 6/28/2011: for displaying dates
                                    Else
                                        Select Case Criteria.PriceDisplay
                                            Case eCriteria_AutoRound
                                                If Abs(dValue) >= 100000 Then
                                                    strValue = Format(dValue, "#,##0")
                                                ElseIf Abs(dValue) > 10000 Or Int(dValue) = dValue Then
                                                    strValue = Format(dValue, "0")
                                                ElseIf Abs(dValue) < 10 Then
                                                    strValue = Format(dValue, "0.####")
                                                Else
                                                    strValue = Format(dValue, "0.##")
                                                End If
                                            
                                            Case eCriteria_RoundToDecimal
                                                strValue = Trim(NumStr(dValue, 0, Criteria.DecimalPlaces))
                                                
                                            Case eCriteria_TradingUnits
                                                strValue = Bars.PriceDisplay(dValue, True)
                                                
                                        End Select
                                    End If
                                End If
                            Else
                                strValue = .TextMatrix(lRow, lCol)
                            End If
                            ChangeCell lRow, lCol, strValue, , fg
                    End Select
                Next lCol
                    
                If m.QB.ColorSymbol = 0 Or Not g.RealTime.Active Then
                    ' - Color symbol up/down only if realtime is on and "color symbol" is selected (otherwise symbol is black)
                    lColor = vbBlack
                ElseIf Bars.Prop(eBARS_LastTickTime) = 0 Then
                    ' - hasn't traded today yet, then unchanged
                    lColor = m.QB.UnchColor
                ElseIf Bars(eBARS_High, Bars.Size - 1) <> Bars(eBARS_Low, Bars.Size - 1) Then
                    '- if Realtime and have up or down trades for the day, then color as an Up or Down tick
                    If Bars.Prop(eBARS_LastTickDown) = 0 Then
                        lColor = m.QB.UpColor
                    Else
                        lColor = m.QB.DownColor
                    End If
                Else
                    '- else color same as "change" (This used to get hit on QB refresh. Don't think this will get hit any more.)
                    lColor = lDeltaColor
                End If
                If g.nColorTheme = kDarkThemeColor Then
                    If IsBlueRange(lColor) Then
                        lColor = vbCyan
                    ElseIf IsGreenRange(lColor, True) Then
                        lColor = vbGreen
                    End If
                End If
                .Cell(flexcpForeColor, lRow, GDCol(eGDCol_Symbol)) = lColor
            End If
        End If
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Set Bars = Nothing
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    Set Criteria = Nothing
    RaiseError "frmQuotes.UpdateCols", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeCell
'' Description: To change text and forecolor of grid cell (called by UpdateCols
''              and TotalRefresh)
'' Inputs:      Row and Column to change, New Value, Is this a Green/Red cell?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeCell(ByVal lRow&, ByVal lCol&, ByVal strCellText$, _
        Optional ByVal bGreenRed As Boolean = False, Optional fg As VSFlexGrid = Nothing)
On Error GoTo ErrSection:

    Dim nForeColor&, lColor&, dValue#, dTickCount#
    Dim strFlexData$        'format: tickCount|alertKey
    
    Dim fgCurrGrid As VSFlexGrid
    
    If fg Is Nothing Then
        Set fgCurrGrid = fgQuotes
    Else
        Set fgCurrGrid = fg
    End If
    
    lColor = m.QB.UpdateColor
    If g.nColorTheme = kDarkThemeColor Then
        If IsBlueRange(m.QB.UpdateColor) Then
            lColor = vbCyan
        ElseIf IsGreenRange(m.QB.UpdateColor, True) Then
            lColor = vbGreen
        End If
    End If
    
    With fgCurrGrid
        nForeColor = m.QB.UnchColor ' vbBlack ' (default)
        If .TextMatrix(lRow, lCol) <> strCellText Then
            .TextMatrix(lRow, lCol) = strCellText
            If Not bGreenRed Then
                If UseUpdatedColors Then
                    nForeColor = lColor
                    .Cell(flexcpForeColor, lRow, GDCol(eGDCol_SymbolID)) = nForeColor
                    ' save TickCount for this cell (so will turn back to black after 1 second)
                    strFlexData = .Cell(flexcpData, lRow, lCol)
                    strFlexData = Str(gdTickCount) & "|" & Parse(strFlexData, "|", 2)
                    .Cell(flexcpData, lRow, lCol) = strFlexData     'gdTickCount
                End If
            End If
        ElseIf UseUpdatedColors Then
            dTickCount = ValOfText(Parse(.Cell(flexcpData, lRow, lCol), "|", 1))
            dTickCount = gdTickCount - dTickCount
            If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                nForeColor = lColor
            End If
        End If
        If bGreenRed Then
            ' color:  Green if UP,  Red if DOWN
            ' (strip caret to see if positive or negative)
            dValue = ValOfText(StripStr(strCellText, "^"))
            If dValue > 0 Then
                nForeColor = m.QB.UpColor ' QBColor(2) 'vbGreen
            ElseIf dValue < 0 Then
                nForeColor = m.QB.DownColor ' vbRed
            End If
            If g.nColorTheme = kDarkThemeColor Then
                If IsBlueRange(nForeColor) Then
                    nForeColor = vbCyan
                ElseIf IsGreenRange(nForeColor, True) Then
                    nForeColor = vbGreen
                End If
            End If
        End If
        .Cell(flexcpForeColor, lRow, lCol) = nForeColor
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ChangeCell", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolExists
'' Description: Does the symbol exist in the quote board?
'' Inputs:      Symbol
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolExists(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:

    SymbolExists = m.QBData.FieldArray(TblField(eQbTbl_Symbol), False).BinarySearch(strSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.SymbolExists", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: Color the background of the cell for symbol and field
'' Inputs:      Symbol, Field name, Color
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ColorCell(ByVal strSymbol As String, ByVal strPeriod As String, ByVal strField As String, _
    ByVal strQbTab As String, Optional ByVal vBackColor As Variant)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row that holds the given symbol
    Dim lCol As Long                    ' Column that holds the data
    Dim lIndex As Long                  ' Index for a for loop
    Dim lSymbolID As Long               ' Symbol ID for the symbol passed in
    Dim astrRows As New cGdArray        ' Array of rows that the symbol/period appears in
    Dim lPos As Long                    ' Position of symbol/period in the table
    Dim lGridRow As Long                ' Grid row to change the color on
    Dim lCurrTab As Long                ' Current tab
    
    Dim lRedraw As Long
    
    
    If Len(strQbTab) > 0 Then
        If TabStr(eGDTabSettings_Name, vsTab.CurrTab) <> strQbTab And TabStr(eGDTabSettings_Name, m.lCurrentTab) <> strQbTab Then
            GoTo ProcessDetachTabs
        End If
    End If
    
    If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Or _
        TabStr(eGDTabSettings_Style, m.lCurrentTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                
        lCol = -1&
        lRow = -1&
        
        lSymbolID = GetSymbolID(strSymbol)
        
        With fgQuotes
            ' Find the column with the given field name...
            For lIndex = 0 To .Cols - 1
                If UCase(.TextMatrix(0, lIndex)) = UCase(strField) Then
                    lCol = lIndex
                    Exit For
                End If
            Next lIndex
    
            ' Find all rows with the given symbol/period...
            If lCol <> -1 Then
#If 0 Then
                For lRow = .FixedRows To .Rows - 1
                    strGridSymbol = UCase(Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1))
                    strGridPeriod = UCase(.TextMatrix(lRow, GDCol(eGDCol_Period)))
                    
                    If strGridSymbol = UCase(strSymbol) Then
                        If strGridPeriod = UCase(strPeriod) Then
                            If IsMissing(vBackColor) Then
                                .Cell(flexcpBackColor, lRow, lCol) = .Cell(flexcpBackColor, lRow, GDCol(eGDCol_SymbolID))
                            Else
                                .Cell(flexcpBackColor, lRow, lCol) = vBackColor
                            End If
                        End If
                    End If
                Next lRow
#End If
                lCurrTab = vsTab.CurrTab
                If TblSearch(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod, lPos) Then
                    astrRows.SplitFields TblStr(eQbTbl_Rows, lPos), ","
                    
                    lRedraw = .Redraw
                    .Redraw = flexRDNone
                    For lRow = 0 To astrRows.Size - 1
                        If (Len(astrRows(lRow)) > 0) And (vsTab.CurrTab = lCurrTab) Then
                            lGridRow = CLng(Val(astrRows(lRow)))
                            
                            DoEvents
                            If lCol >= 0 And lCol < .Cols Then
                                If lGridRow >= .FixedRows And lGridRow < .Rows Then
                                    If IsMissing(vBackColor) Then
                                        .Cell(flexcpBackColor, lGridRow, lCol) = .Cell(flexcpBackColor, lGridRow, GDCol(eGDCol_SymbolID))
                                    Else
                                        .Cell(flexcpBackColor, lGridRow, lCol) = vBackColor
                                    End If
                                End If
                            End If
                        End If
                    Next lRow
                    .Redraw = lRedraw
                End If
            End If
        End With
    End If
    
ProcessDetachTabs:
    
    Dim frm As frmDetachedQBTab
    
    For lIndex = 0 To m.aDetachedTabs.Size - 1
        Set frm = m.aDetachedTabs(lIndex)
        If Not frm Is Nothing Then
            If IsMissing(vBackColor) Then
                frm.ColorCell strSymbol, strPeriod, strField, strQbTab
            Else
                frm.ColorCell strSymbol, strPeriod, strField, strQbTab, vBackColor
            End If
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ColorCell", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsBusy
'' Description: Are we doing a total refresh?
'' Inputs:      None
'' Returns:     True if doing a TotalRefresh, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBusy() As Boolean
On Error GoTo ErrSection:

    If m.nRefreshRow > 0 Then IsBusy = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.IsBusy", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlignColumns
'' Description: Align the columns appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AlignColumns(Optional fg As VSFlexGrid = Nothing)
On Error GoTo ErrSection:

    Dim nCol&, iRedraw As Integer
    Dim fgCurrGrid As VSFlexGrid
    
    If fg Is Nothing Then
        Set fgCurrGrid = fgQuotes
    Else
        Set fgCurrGrid = fg
    End If
    
    With fgCurrGrid
        iRedraw = .Redraw
        .Redraw = flexRDNone
        For nCol = 0 To .Cols - 1
            Select Case UCase(.TextMatrix(0, nCol))
                Case "SYMBOL", "PERIOD", "DESCRIPTION", "T", "LAST TICK"
                    .ColAlignment(nCol) = flexAlignLeftCenter
                Case "EXCHANGE", "SESSION" ', "LAST TICK"
                    .ColAlignment(nCol) = flexAlignCenterCenter
                Case Else
                    .ColAlignment(nCol) = flexAlignRightCenter
            End Select
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        .Redraw = iRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.AlignColumns", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyPress
'' Description: Handle certain keystrokes special
'' Inputs:      Ascii representation of key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub KeyPress(KeyAscii As Integer, Optional Shift As Integer = -1)
On Error Resume Next

Exit Sub

    'JM 06-03-2009 - this code not used so doesn't matter
    '   just changed type of form so would not show up in searches
    
    Dim frm As Form
    Dim bLookForChart As Boolean

    If KeyAscii = 0 Then Exit Sub

    If Shift >= 0 Then ' (came from KeyDown event)
        If KeyAscii >= vbKeyF2 And KeyAscii <= vbKeyF12 Then
            bLookForChart = True
        End If
    Else ' (came from KeyPress event)
        Select Case Asc(UCase(Chr(KeyAscii)))
            
            Case 65 To 90, 48 To 57, 43, 45, 61:
                bLookForChart = True
        End Select
    End If
       
    If bLookForChart Then
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            frm.KeyPress KeyAscii, Shift
        End If
        KeyAscii = 0
    End If
       
    Set frm = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRequests
'' Description: Set up the request file for doing a refresh
'' Inputs:      Array to fill in with requests
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddRequests(aRequests As cGdArray)
On Error GoTo ErrSection:

    Dim strDates As String              ' Dates to request data for
    Dim strTemp As String               ' Symbol and security type
    Dim lIndex As Long                  ' Index into a for loop
    Dim lDelay As Long                  ' Delay for the symbol (if real-time)
    Dim dNyDate As Double

    ' Walk through the grid and add to the request file
    With fgQuotes
        ' find out what date it is in NY right now
        dNyDate = ConvertTimeZone(Now)
        dNyDate = Int(dNyDate - 1 / 24#) '(adjust by 1 hour just to make sure)
        ' don't need to request today if already did daily download for today
        If dNyDate <= LastDailyDownload Then
            dNyDate = LastDailyDownload + 1
        End If
        
        ' ask for today and tomorrow
        strDates = "@" & Format(dNyDate, "YYYYMMDD") & "-" & Format(dNyDate + 1, "YYYYMMDD")
        For lIndex = 0 To m.QBData.NumRecords - 1
            strTemp = TblStr(eQbTbl_SecType, lIndex) & ";" & TblStr(eQbTbl_Symbol, lIndex)
            
            ' If real-time is active, then get the delay
            If g.RealTime.Active Then
                lDelay = g.RealTime.SymbolDelay(TblStr(eQbTbl_Symbol, lIndex))
            Else
                lDelay = -1&
            End If
            
            ' If they are authorized for ticks, ask for both...
            If InStr(g.strAuthorizationString, "," & Left(TblStr(eQbTbl_SecType, lIndex), 1) & "T,") > 0 Then
                aRequests.Add strDates & ";B;I;" & strTemp & ";" & Str(lDelay)
                
            ' Otherwise just ask for end of day...
            Else
                aRequests.Add strDates & ";E;I;" & strTemp & ";" & Str(lDelay)
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.AddRequests", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PopulateGrid
'' Description: Populate grid with symbols data
'' Inputs:      flex grid control, string of symbols
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateGrid(fg As VSFlexGrid, ByVal strSymbols As String)
On Error GoTo ErrSection:
    
    Dim strSecType As String            ' Security Type of the current row
    Dim strSymbol As String             ' Symbol of the current row
    Dim astrSymbols As New cGdArray     ' Symbols to show for this tab
    Dim lIndex As Long                  ' Index into a for loop
    Dim strPeriod As String             ' Period for the given symbol
    Dim lRedraw As Long                 ' Current state of the grids redraw
    Dim lSymbolID As Long               ' Symbol ID for the symbol to add
    Dim lPoolRec As Long                ' Record number into the symbol pool
    
    If fg Is Nothing Then Exit Sub
    
    astrSymbols.Create eGDARRAY_Strings
    astrSymbols.SplitFields strSymbols, ","
    
    With fg
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Clear out the grid...
        .Rows = .FixedRows
        
        ' Add the symbols into the grid...
        For lIndex = 0 To astrSymbols.Size - 1
            strSymbol = Parse(astrSymbols(lIndex), ";", 1)
            strPeriod = Parse(astrSymbols(lIndex), ";", 2)
            
            If strSymbol = "Label" Then
                AddLabelRow strPeriod
            ElseIf Len(strSymbol) > 0 Then
                If IsNumeric(strSymbol) Then
                    lSymbolID = CLng(strSymbol)
                    strSymbol = GetSymbol(lSymbolID)
                Else
                    lSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
                End If
                
                If lSymbolID <> 0& Then
                    lPoolRec = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
                    strSecType = GetSecType(g.SymbolPool.SecType(lPoolRec))
                Else
                    If InStr(strSymbol, " ") <> 0 Then
                        If InStr(strSymbol, "-") <> 0 Then
                            strSecType = "FO"
                        Else
                            strSecType = "SO"
                        End If
                    End If
                End If
                
                If Len(strSymbol) > 0 Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_SecType)) = UCase(strSecType)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolID)) = Str(lSymbolID)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = strSymbol
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = strPeriod
                    .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_SecType), .Rows - 1, GDCol(eGDCol_Period)) = m.QB.UnchColor
                    .MergeRow(.Rows - 1) = False
                    UpdateCols .Rows - 1, , False, , fg
                    .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, .Cols - 1) = Nothing
                End If
            End If
        Next lIndex
        
        .Row = 0
        ' Add a blank row to the bottom of the list...
        AddLabelRow "", , fg
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbols = Nothing
    RaiseError "frmQuotes.PopulateGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowCategory
'' Description: Set up the grid according to the new current category tab
'' Inputs:      Current Tab
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowCategory(Optional ByVal iTab& = -1)
On Error GoTo ErrSection:

    Dim strSymbols As String            ' Symbols to limit to
    Dim lIndex As Long                  ' Index into a for loop
    Dim strCells As String              ' Cells that the symbols go in
    Dim lStyle As eGDQuoteStyle         ' Quote board style for this tab
    Dim lRedraw As Long                 ' Current state of the grids redraw
    Dim bTimerEnabled As Boolean


    ' No need to do this routine if there is no data installed yet...
    If g.SymbolPool.NumRecords = 0 Then Exit Sub
    
    ' temporarily turn off the timer so everything won't show in blue when switching tabs
    bTimerEnabled = tmrRealtime.Enabled
    tmrRealtime.Enabled = False
    
    ' Default to the current tab...
    If iTab < 0 Or iTab >= vsTab.NumTabs Then iTab = vsTab.CurrTab
        
    ' Set up the column order and visibility...
    SetUpColumns iTab
    
    ' Get list of symbols for this tab...
    lStyle = CLng(ValOfText(TabStr(eGDTabSettings_Style, iTab)))
    strSymbols = TabStr(eGDTabSettings_Symbols, iTab)
    
    ' Hide/Show controls as applicable for the style...
    If iTab = m.lCurrentTab Then
        ChangeStyle lStyle, iTab
    End If
        
    ' If the style of the tab is a grid, then set up the grid...
    If lStyle = QStyle(eGDQuoteStyle_Grid) Then
        cmdSortBySym.Visible = False
        
        PopulateGrid fgQuotes, strSymbols
                    
        With fgQuotes
            lRedraw = .Redraw
            .Redraw = flexRDNone
            ' Make sure that the alternate coloring is correct...
            ColorQuoteRows
                        
            ' Reset the rows field of the data table...
            DoEvents
            ResetRows iTab
            
            If m.lTotWidth = 0& Then .AutoSize 0, .Cols - 1, False, 75
            If ((.Row < .FixedRows) Or (.Row >= .Rows)) And (.Rows > .FixedRows) Then
                .Row = .FixedRows
                .RowSel = .FixedRows
                '.Col = GDCol(eGDCol_Symbol)
            End If
            .Redraw = lRedraw
        End With
            
    ' Otherwise set up the box style quote board...
    Else
        cmdSortBySym.Visible = True
        strCells = strSymbols
        If Left(strCells, 1) = "," Then strCells = Mid(strCells, 2)
        If Right(strCells, 1) = "," Then strCells = Left(strCells, Len(strCells) - 1)
        m.QB.BoardFromString strCells
        
        For lIndex = 1 To m.BarsColl.Count
            m.QB.UpdateSymbol m.BarsColl(lIndex), False
        Next lIndex
        
        ResetRows iTab
        m.QB.DrawBoard
        pbQuoteBoard.Refresh
    End If
            
ErrExit:
    tmrRealtime.Enabled = bTimerEnabled
    Exit Sub
    
ErrSection:
    tmrRealtime.Enabled = bTimerEnabled
    RaiseError "frmQuotes.ShowCategory", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowQuotesPopup
'' Description: Show the Quotes pop-up menu
'' Inputs:      Location to show the menu
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowQuotesPopup(ByVal X#, ByVal Y#, Optional frmDetTab As frmDetachedQBTab = Nothing)
On Error GoTo ErrSection:

    Dim strSymbol$, i&
    Dim eStyle As eGDQuoteStyle
    Dim strPeriod As String             ' Period for the row
    Dim nPeriod As Long                 ' Periodicity for the row
    
    Dim bFilterTab As Boolean           ' Is the current tab the filter tab?
    Dim bShowEditMenu As Boolean
    
    Dim fgCurrGrid As VSFlexGrid
    Dim QB As cQuoteCellBoard
    
    strPeriod = ""
    nPeriod = 0
    
    Set m.frmActiveDetTab = frmDetTab
    If frmDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
        Set QB = m.QB
        eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), vsTab.CurrTab)
        bFilterTab = (vsTab.TabCaption(vsTab.CurrTab) = "(Filter)")
    Else
        Set fgCurrGrid = frmDetTab.fgQuotes         'this is on detached QB form
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), frmDetTab.MyTabIndex)
        bFilterTab = frmDetTab.Caption = "(Filter)"
    End If
    
    If fgCurrGrid.Visible = True And Not fgCurrGrid Is Nothing Then
        With fgCurrGrid
            If (.Row >= .FixedRows) And (.MergeRow(.Row) = False) Then
                strSymbol = Parse(.TextMatrix(.Row, GDCol(eGDCol_Symbol)), "(", 1)
                If .Col >= .FixedCols And .Col < .Cols Then
                    If Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then bShowEditMenu = True
                End If
        
                strPeriod = .TextMatrix(.Row, GDCol(eGDCol_Period))
                nPeriod = GetPeriodicity(strPeriod)
            End If
        End With
    ElseIf Not QB Is Nothing Then
        With QB
            If .Row <> -1& And .Col <> -1& Then
                If Not .Cell(.Row, .Col) Is Nothing Then
                    strSymbol = Parse(.Cell(.Row, .Col).Symbol, "(", 1)
                Else
                    strSymbol = ""
                End If
            End If
        End With
    End If
    
    mnuAddSymbol.Enabled = Not bFilterTab
    mnuLabelRow.Enabled = Not bFilterTab
    If Len(strSymbol) = 0 Then
        strSymbol = "symbol"
        mnuRemoveSymbol.Enabled = False
        mnuToCategory.Enabled = False
        mnuBuy.Enabled = False
        mnuSell.Enabled = False
        mnuBuy.Caption = "Buy"
        mnuSell.Caption = "Sell"
    Else
        mnuRemoveSymbol.Enabled = Not bFilterTab
        mnuToCategory.Enabled = Not bFilterTab
        mnuBuy.Enabled = True
        mnuSell.Enabled = True
        mnuBuy.Caption = "Buy " & RollSymbolForDate(strSymbol, Now)
        mnuSell.Caption = "Sell " & RollSymbolForDate(strSymbol, Now)
    End If
    mnuRemoveSymbol.Caption = "Remove " & strSymbol & "   (hotkey: 'Delete')"
    mnuToCategory.Caption = "Move " & strSymbol & " to Tab"
    
    mnuChangeFilter.Visible = bFilterTab
    mnuClearFilter.Visible = bFilterTab
    mnuClearFilter.Enabled = Len(TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1)) > 0
    
    Enable mnuChangePeriod, (nPeriod >= ePRD_Days) Or (GetPeriodType(nPeriod) = ePRD_Minutes)
    
    For i = 0 To m.tblTabInfo.NumRecords - 1
        If TabStr(eGDTabSettings_Name, i) <> "(Filter)" Then
            If i + 1 > mnuCategory.UBound Then
                Load mnuCategory(i + 1)
                mnuCategory(i + 1).Visible = True
            End If
            mnuCategory(i + 1).Caption = TabStr(eGDTabSettings_Name, i)
        End If
    Next i
    For i = mnuCategory.UBound To m.tblTabInfo.NumRecords + 1 Step -1
        If i > 0 Then Unload mnuCategory(i)
    Next i
    
    mnuAddAlert.Visible = True
    
    If bFilterTab Then
        mnuQuoteBoardStyle.Visible = False      '4239
        mnuExportQuoteTab.Visible = False
        mnuChangePeriod.Visible = False
    Else
        mnuQuoteBoardStyle.Visible = True
        mnuExportQuoteTab.Visible = FileExist("C:\Common\Files32.EXE")
        
        If eStyle = eGDQuoteStyle_Grid Then
            mnuConvertToGrid.Visible = False
            mnuConvertToBox.Visible = True
            mnuConvertToForex.Visible = True
            mnuSortBySym.Visible = False
            mnuEditAlert.Visible = bShowEditMenu
            mnuFields.Visible = True
            mnuChangePeriod.Visible = True
            If Len(strPeriod) > 0 Then
                mnuChangePeriod.Caption = "Change Period for All Symbols on Tab to '" & strPeriod & "'"
            Else
                mnuChangePeriod.Caption = "Change Period for All Symbols on Tab"
            End If
        ElseIf eStyle = eGDQuoteStyle_Forex Then
            mnuConvertToGrid.Visible = True
            mnuConvertToBox.Visible = True
            mnuConvertToForex.Visible = False
            mnuSortBySym.Visible = True
            mnuEditAlert.Visible = False
            mnuFields.Visible = False           '4282
            mnuChangePeriod.Visible = False
        Else
            mnuConvertToGrid.Visible = True
            mnuConvertToBox.Visible = False
            mnuConvertToForex.Visible = True
            mnuSortBySym.Visible = True
            mnuEditAlert.Visible = False
            mnuFields.Visible = False
            mnuChangePeriod.Visible = False
        End If
    End If
    
    If bFilterTab Then
        mnuDetachQBTab.Visible = False
    Else
        mnuDetachQBTab.Visible = True
        If m.frmActiveDetTab Is Nothing Then
            mnuDetachQBTab.Caption = "Detach Quote Tab"
        Else
            mnuDetachQBTab.Caption = "Attach Quote Tab"     '5451
        End If
    End If
    
    mnuShowButtons.Visible = frmDetTab Is Nothing
    mnuShowButtons.Checked = m.bShowButtons
    
    PopupMenu mnuQuotes, vbPopupMenuLeftAlign, X, Y
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ShowQuotesPopup", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditCategories
'' Description: Allow the user to add/remove/change categories
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditCategories(Optional ByVal bCopy As Boolean = False)
On Error GoTo ErrSection:

    Dim strCat As String                ' Name of the category
    Dim strOld As String                ' Old Position of the tab
    Dim astrTabInfo As New cGdArray     ' Array of tab info (active, name, tab index, active, style, detach flag)
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lCurrTab As Long                ' Current tab
    Dim lNewTab As Long                 ' Tab to set focus to when through
    Dim lOldPos As Long                 ' Old position of the tab
    Dim strTemp As String               ' Temporary string
    Dim astrSave As New cGdArray        ' Copy of current table
    Dim strExist As String              ' String of existing indexes
    Dim lNewIndex As Long               ' New place to store the information
    Dim strFilterTab As String          ' String of info for the filter tab
    Dim eQTabNewStyle As eGDQuoteStyle
    
    Dim iTabIndex&, strCaption$         'for handling detached tab
    Dim iForm&, iNewDetachFlag&
    
    Dim aReattach As New cGdArray       'holds handle to detached tabs that user wants to be reattached
    Dim aDetach As New cGdArray         'holds semicolon separate index;0/1 (0/1 --> flag for clearing symbols)
    
    Dim bLocked As Boolean

'JM 05-12-2010: bShowNew was for when the (new) tab existed
'    If bShowNew Then
'        If Not HasGold(True, "Adding New Quote Board Tabs") Then Exit Sub
'    End If

    ' Create the string of tab info for the add/remove form...
    astrTabInfo.Create eGDARRAY_Strings
    astrSave.Create eGDARRAY_Strings
    For lIndex = 0 To m.tblTabInfo.NumRecords - 1
        If TabStr(eGDTabSettings_Name, lIndex) = "(Filter)" Then
            strFilterTab = "True" & vbTab & TabStr(eGDTabSettings_Name, lIndex) _
                & vbTab & Str(lIndex) & vbTab & "False" & vbTab & TabStr(eGDTabSettings_Style, lIndex) _
                & vbTab & TabStr(eGDTabSettings_Form, lIndex)
        Else
            astrTabInfo(lIndex) = "True" & vbTab & TabStr(eGDTabSettings_Name, lIndex) _
                & vbTab & Str(lIndex) & vbTab & "True" & vbTab & TabStr(eGDTabSettings_Style, lIndex) _
                & vbTab & TabStr(eGDTabSettings_Form, lIndex)
        End If
            
        astrSave(lIndex) = m.tblTabInfo.GetRecord(lIndex, vbTab)
    Next lIndex
    
    If m.frmActiveDetTab Is Nothing Then
        iTabIndex = vsTab.CurrTab
        strCaption = vsTab.TabCaption(vsTab.CurrTab)
    Else
        iTabIndex = m.frmActiveDetTab.MyTabIndex        'true only for detached-tab copy
        strCaption = m.frmActiveDetTab.Caption
    End If
    
    ' Call the add/remove form with the list of QB tabs...
    If frmQuoteBoardFields.ShowMe(astrTabInfo, eQbfMode_QBCat, , bCopy, strCaption) Then
        lCurrTab = iTabIndex
        lNewTab = 0&
        strExist = ","
        
        lIndex = m.tblTabInfo.NumRecords - 1
        If TabStr(eGDTabSettings_Name, lIndex) = "(Filter)" Then
            strFilterTab = "True" & vbTab & TabStr(eGDTabSettings_Name, lIndex) _
                & vbTab & Str(lIndex) & vbTab & "False" & vbTab & strTemp & vbTab & TabStr(eGDTabSettings_Form, lIndex)
        Else
            GoTo ErrExit    'something very wrong
        End If
        
        astrTabInfo.Add strFilterTab
        lNewIndex = astrTabInfo.Size
        
        'parse field num: 1:active \t 2:name \t 3:QB table index \t 4:show \t 5:style flag \t 6:detached
        ' Walk through the tab info string returned and make any changes...
        For lIndex = 0 To astrTabInfo.Size - 1
            strCat = Parse(astrTabInfo(lIndex), vbTab, 2)
            strOld = Parse(astrTabInfo(lIndex), vbTab, 3)
            iNewDetachFlag = Val(Parse(astrTabInfo(lIndex), vbTab, 6))   'this is value of check-box in frmQuoteBoardFields grid (1=checked,2=unchecked)

            lOldPos = -999999
            If Len(strOld) > 0 Then lOldPos = Val(strOld)
            
            If Left(strOld, 1) <> "-" Then
                ' This is a pre-existing tab...
                If lNewTab = 0 And lOldPos = lCurrTab Then lNewTab = lIndex
                strExist = strExist & strOld & ","
                m.tblTabInfo.SetRecord astrSave(lOldPos), lIndex, vbTab
                
                'update the detached tab
                If m.tblTabInfo(eGDTabSettings_Form, lIndex) = 0 Then
                    If iNewDetachFlag = 1 Then aDetach.Add Str(lIndex) & ";0"         '0 indicates existing tab
                Else
                    For iForm = 0 To m.aDetachedTabs.Size - 1
                        If m.aDetachedTabs(iForm).hWnd = m.tblTabInfo(eGDTabSettings_Form, lIndex) Then
                            m.aDetachedTabs(iForm).MyTabIndex = lIndex
                            m.aDetachedTabs(iForm).Caption = strCat
                            If iNewDetachFlag = 2 Then aReattach.Add m.aDetachedTabs(iForm)
                            Exit For
                        End If
                    Next
                End If
                            
                ' Do this just in case it was renamed...
                TabStr(eGDTabSettings_Name, lIndex) = strCat
            Else
                eQTabNewStyle = ValOfText(Parse(astrTabInfo(lIndex), vbTab, 5))
                If eQTabNewStyle <> eGDQuoteStyle_Forex And eQTabNewStyle <> eGDQuoteStyle_Grid Then
                    eQTabNewStyle = m.DefaultStyle      'box style can be bollinger, candle etc, use last saved default for new tabs
                End If
                
                m.tblTabInfo.AddRecord strCat, lIndex
                If lOldPos = -999999 Then
                    If iNewDetachFlag = 1 Then aDetach.Add Str(lIndex) & ";1" '1 --> clear symbols from newly created tab
                    'this is a brand new tab to be added ...
                    TabStr(eGDTabSettings_Fields, lIndex) = m.strDefaultFields
                    TabStr(eGDTabSettings_Style, lIndex) = Str(eQTabNewStyle)
                Else
                    If iNewDetachFlag = 1 Then aDetach.Add Str(lIndex) & ";0"
                    'this is a new tab to be added by copying from an existing tab
                    CopyTab astrSave, Abs(lOldPos), lIndex, strCat, eQTabNewStyle
                End If
                
                'show last newly created or copied tab if not detached
                If iNewDetachFlag <> 1 Then lNewTab = lIndex
            
                strTemp = ""
            End If
        Next
        
        ' Move any removed information to the end...
        For lIndex = 0 To astrSave.Size - 1
            If InStr(strExist, "," & Str(lIndex) & ",") = 0 Then
                m.tblTabInfo.SetRecord astrSave(lIndex), lNewIndex, vbTab
                lNewIndex = lNewIndex + 1
            End If
        Next lIndex
        
        ' Delete any tabs that no longer exist...
        For lIndex = m.tblTabInfo.NumRecords - 1 To astrTabInfo.Size Step -1
            If lIndex = lCurrTab Then
                lCurrTab = 0&
                vsTab.CurrTab = 0&
            End If
            RemoveTab lIndex
        Next lIndex
        
        'reattach any detached tabs that user unchecked check box for
        For iForm = 0 To aReattach.Size - 1
            Unload aReattach(iForm)
        Next
               
        ' Reinitialize the tab captions and re-show the current category...
        InitQuoteTabs lCurrTab
        If m.frmActiveDetTab Is Nothing Then ShowCategory
        If lCurrTab <> lNewTab Then vsTab.CurrTab = lNewTab
    
        '01-30-2008: the fix for 4392 broke tab move & deletion (4398) - moving this code here fixes both issues
        'detach any tab that were previously attached that user checked check box for
        
        bLocked = LockWindowUpdate(Me.hWnd)
        For iForm = 0 To aDetach.Size - 1
            If ValOfText(Parse(aDetach(iForm), ";", 2)) = 1 Then
                DetachTab ValOfText(Parse(aDetach(iForm), ";", 1)), True
            Else
                DetachTab ValOfText(Parse(aDetach(iForm), ";", 1))
            End If
            ShowCategory
        Next
        If bLocked Then LockWindowUpdate 0
    
    End If

ErrExit:
    Set astrSave = Nothing
    Set astrTabInfo = Nothing
    Set aReattach = Nothing
    Set aDetach = Nothing
    Exit Sub
    
ErrSection:
    Set astrSave = Nothing
    Set astrTabInfo = Nothing
    RaiseError "frmQuotes.EditCategories", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitQuoteTabs
'' Description: Initialize the category tabs
'' Inputs:      Current Tab
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitQuoteTabs(ByVal iCurTab As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strTabs As String               ' Tab names
    Dim strTemp As String               ' Temporary string variable
    
    If iCurTab < 0 Then iCurTab = vsTab.CurrTab
    
    For lIndex = 0 To m.tblTabInfo.NumRecords - 1
        strTemp = TabStr(eGDTabSettings_Name, lIndex)
        strTabs = strTabs & "|" & Replace(strTemp, "&", "&&") ', 1, 1)
    Next
    If Left(strTabs, 1) = "|" Then strTabs = Mid(strTabs, 2)
    
    ' Make sure that there is a "Filter" tab...
    If InStr(strTabs, "(Filter)") = 0 Then strTabs = strTabs & "|(Filter)"
    
    ' Make sure that there is a "new" tab...
'    strTabs = strTabs & "|(new)"
    strTabs = strTabs & "|(manage)"
    vsTab.Caption = strTabs
    
    If iCurTab >= vsTab.NumTabs - 1 Then iCurTab = 0
    
    If iCurTab >= 0 And iCurTab < vsTab.NumTabs - 1 Then
        vsTab.FirstTab = 0
        vsTab.CurrTab = iCurTab
    End If
    
    ' hide Filter tab if Extreme version
    For lIndex = 0 To vsTab.NumTabs - 1
        If lIndex = vsTab.NumTabs - 2 Then
            vsTab.TabVisible(lIndex) = (ExtremeCharts <> 1)
        ElseIf m.tblTabInfo(TabField(eGDTabSettings_Form), lIndex) = 0 Then
            vsTab.TabVisible(lIndex) = True
        Else
            vsTab.TabVisible(lIndex) = False
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.InitQuoteTabs", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorQuoteRows
'' Description: Alternate the background color on quote grid rows
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorQuoteRows()
On Error GoTo ErrSection:

'NOT using hidden rows anymore as of version 3.03
'
'    Dim nRow&, nColor&, iTab%, iSaveRedraw%
'    Dim bAltRow As Boolean
'
'    With fgQuotes
'        iTab = vsTab.CurrTab
'        iSaveRedraw = .Redraw
'        .Redraw = flexRDNone
'        ' unfortunately we must color the rows manually when hidden rows
'        .BackColorAlternate = 0 '(turns it off)
'        For nRow = .FixedRows To .Rows - 1
'            If Not .RowHidden(nRow) Then
'                If bAltRow Then
'                    nColor = ALT_GRID_ROW_COLOR
'                Else
'                    nColor = .BackColor
'                End If
'                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = nColor
'                bAltRow = Not bAltRow
'            End If
'        Next
'
'        If (UCase(vsTab.TabCaption(m.lCurrentTab)) = "(FILTER)") And (.MergeRow(.FixedRows) = True) Then
'            .Cell(flexcpBackColor, .FixedRows, GDCol(eGDCol_Symbol)) = .BackColorFixed
'            .Cell(flexcpAlignment, .FixedRows, GDCol(eGDCol_Symbol)) = flexAlignCenterCenter
'            .TextMatrix(.FixedRows, GDCol(eGDCol_Symbol)) = "Change Filter..."
'            If m.lTotWidth = 0 Then .AutoSize 0, .Cols - 1, False, 75
'        Else
'            .Cell(flexcpBackColor, .FixedRows, GDCol(eGDCol_Symbol)) = .Cell(flexcpBackColor, .FixedRows, 0)
'            .Cell(flexcpAlignment, .FixedRows, GDCol(eGDCol_Symbol)) = flexAlignLeftTop
'        End If
'
'        .Redraw = iSaveRedraw
'    End With
'
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ColorQuoteRows", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TblSearch
'' Description: Look for a symbol/period in the data table
'' Inputs:      Symbol, Period, Position where it was found/would be inserted
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TblSearch(ByVal strSymbolOrSymbolID As String, ByVal strPeriod As String, Optional lPos As Long = 0&) As Boolean
On Error GoTo ErrSection:

    Dim strSearchKey As String          ' String to search for

    strSearchKey = strSymbolOrSymbolID & vbTab & strPeriod
    TblSearch = m.QBData.FieldArray(TblField(eQbTbl_SearchKey), False).BinarySearch(strSearchKey, lPos)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.TblSearch", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColNum
'' Description: Find the column number for a given column name
'' Inputs:      Name of column to search for, Criteria ID
'' Returns:     Number of column found (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColNum(ByVal strColName As String, ByVal strCriteriaID As String) As Long
On Error GoTo ErrSection:

    Dim lCol As Long                ' Index into a for loop
    
    ColNum = -1&
    For lCol = 0 To fgQuotes.Cols - 1
        If fgQuotes.TextMatrix(0, lCol) = strColName And fgQuotes.ColData(lCol) = strCriteriaID Then
            ColNum = lCol
            Exit For
        End If
    Next lCol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.ColNum", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpColumns
'' Description: Set up the columns for a given tab
'' Inputs:      Tab to get the column information from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpColumns(Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lCol As Long                    ' Index into a for loop
    Dim astrCol As New cGdArray         ' Column information for this column
    Dim astrCols As New cGdArray        ' Column information for all columns
    Dim lPos As Long                    ' Current position of the column
    Dim lColWidth As Long               ' Column width for current column
    
    ' If no tab was passed in, use the current one...
    If lTab = -1& Then lTab = vsTab.CurrTab
    
    ' Get the column information for the tab...
    astrCols.SplitFields TabStr(eGDTabSettings_Fields, lTab), "|"
    
    With fgQuotes
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Hide all of the columns first so that in case a column doesn't show
        ' up in the column information, it gets hidden by default...
        For lCol = 0 To .Cols - 1
            .ColHidden(lCol) = True
        Next lCol
        
        ' Total custom width of all the columns...
        m.lTotWidth = 0&
               
        If UCase(Left(astrCols(0), 6)) = "DIRTY;" Then
            astrCols.Remove 0
        End If
               
        ' Walk through all of the columns and set the position, hidden, and width...
        For lCol = 0 To astrCols.Size - 1
            'saved string can get corrupted and astrcols.size ends up > .cols
            If lCol >= .Cols Then Exit For
            
            astrCol.Clear
            astrCol.SplitFields astrCols(lCol), ";"
            
            ' Move the column to the proper position...
            ' 03/05/2010 DAJ: If the column cannot be found (lPos = -1&), then don't do
            ' the rest of the block of information because it could lead to the wrong
            ' columns being shown/hidden...
            lPos = ColNum(astrCol(0), astrCol(2))
            If lPos <> -1& Then
                .ColPosition(lPos) = lCol
            
                ' Make sure that the Symbol column never gets hidden...
                If astrCol(0) = "Symbol" Then
                    astrCol(1) = "0"
                End If
                
                ' Make sure that the SymbolID, and SecType columns are hidden...
                If astrCol(0) = "SymbolID" Or astrCol(0) = "SecType" Then
                    astrCol(1) = "-1"
                End If
                
                ' Hide the column if we need to...
                If Len(astrCol(1)) = 0 Then
                    .ColHidden(lCol) = False
                Else
                    .ColHidden(lCol) = CBool(ValOfText(astrCol(1)))
                End If
                
                ' Set the column width appropriately...
                lColWidth = CLng(ValOfText(astrCol(3)))
                If lColWidth > 0& Then .ColWidth(lCol) = lColWidth
                m.lTotWidth = m.lTotWidth + lColWidth
            End If
        Next lCol
        
        ' If no column widths were specified, do an AutoSize...
        If m.lTotWidth = 0& Then .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Set astrCol = Nothing
    Set astrCols = Nothing
    Exit Sub
    
ErrSection:
    Set astrCol = Nothing
    Set astrCols = Nothing
    RaiseError "frmQuotes.SetUpColumns", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveTabInfo
'' Description: Save the row and column information for a given tab
'' Inputs:      Tab to save information for
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveTabInfo(Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Index into a for loop
    Dim strColInfo As String            ' Column information to save
    Dim astrTabInfo As New cGdArray     ' Tab information
    Dim lRow As Long                    ' Index into a for loop
    Dim strSymbols As String            ' Symbol information to save
    Dim strText As String
    
    ' Don't do anything if data has not been installed yet...
    If g.SymbolPool.NumRecords = 0 Then Exit Sub
    
    Dim fgCurrGrid As VSFlexGrid
    Dim oCurrQB As cQuoteCellBoard
    
    If m.frmActiveDetTab Is Nothing Then
        ' If no tab was passed in, use the current one...
        If lTab = -1& Then lTab = vsTab.CurrTab
        Set fgCurrGrid = fgQuotes
        Set oCurrQB = m.QB
    Else
        lTab = m.frmActiveDetTab.MyTabIndex
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        Set oCurrQB = m.frmActiveDetTab.QuoteCellBoard
    End If
    strColInfo = ""
               
    With fgCurrGrid
        ' Walk through all of the columns saving off appropriate information...
        For lCol = 0 To .Cols - 1
            If .TextMatrix(0, lCol) <> "#DELETED#" Then
                If .TextMatrix(0, lCol) = "Symbol" And .ColHidden(lCol) = True Then
                    .ColHidden(lCol) = False
                End If
                
                If m.bUserResize Or m.lTotWidth > 0& Then
                    strColInfo = strColInfo & .TextMatrix(0, lCol) & ";" & _
                        Str(CLng(.ColHidden(lCol))) & ";" & .ColData(lCol) & ";" & _
                        Str(.ColWidth(lCol)) & "|"
                Else
                    strColInfo = strColInfo & .TextMatrix(0, lCol) & ";" & _
                        Str(CLng(.ColHidden(lCol))) & ";" & .ColData(lCol) & "|"
                End If
            End If
        Next lCol
        
        If TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
            strSymbols = ","
            For lRow = .FixedRows To .Rows - 1
                If .MergeRow(lRow) = True Then
                    If lRow < .Rows - 1 Then
                        strText = Parse(.TextMatrix(lRow, GDCol(eGDCol_Period)), "(Click", 1, True)         '5711
                        strSymbols = strSymbols & "Label;" & strText & ","
                        
                        'JM 04-29-2009: original code leave awhile then remove if all okay
                        'strSymbols = strSymbols & "Label;" & .TextMatrix(lRow, GDCol(eGDCol_Period)) & ","
                    End If
                Else
                    If ValOfText(.TextMatrix(lRow, GDCol(eGDCol_SymbolID))) <> 0 Then
                        strSymbols = strSymbols & .TextMatrix(lRow, GDCol(eGDCol_SymbolID))
                    Else
                        strSymbols = strSymbols & .TextMatrix(lRow, GDCol(eGDCol_Symbol))
                    End If
                
                    strSymbols = strSymbols & ";" & .TextMatrix(lRow, GDCol(eGDCol_Period)) & ","
                End If
            Next lRow
        Else
            strSymbols = oCurrQB.SaveString
        End If
    End With

    TabStr(eGDTabSettings_Symbols, lTab) = strSymbols
    TabStr(eGDTabSettings_Fields, lTab) = strColInfo
    m.strDefaultFields = strColInfo
        
ErrExit:
    Set astrTabInfo = Nothing
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    Set astrTabInfo = Nothing
    Set m.frmActiveDetTab = Nothing
    RaiseError "frmQuotes.SaveTabInfo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolRow
'' Description: Row in the grid of the given symbol/period
'' Inputs:      Symbol/Period to look up in the grid
'' Returns:     Row containing the Symbol/Period, or -1 if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolRow(ByVal strSymbol As String, ByVal strPeriod As String, _
            Optional ByVal bOnlyIfVisible As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop

    With fgQuotes
        SymbolRow = -1&
        For lRow = .FixedRows To .Rows - 1
            'If .RowHidden(lRow) = False Or bOnlyIfVisible = False Then
            If bOnlyIfVisible = False Then
                If UCase(Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)) = UCase(strSymbol) Then
                    If UCase(.TextMatrix(lRow, GDCol(eGDCol_Period))) = UCase(strPeriod) Then
                        SymbolRow = lRow
                        Exit For
                    End If
                End If
            End If
        Next lRow
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.SymbolRow", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLabelRow
'' Description: Add a label row to the grid
'' Inputs:      Label for the row, Position to place the label row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLabelRow(ByVal strLabel As String, Optional ByVal lRow As Long = -1&, _
    Optional fg As VSFlexGrid = Nothing)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColor As Long
    Dim bBold As Boolean
    
    Dim fgCurrGrid As VSFlexGrid
    
    If Not fg Is Nothing Then
        Set fgCurrGrid = fg
    ElseIf Not m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
    Else
        Set fgCurrGrid = fgQuotes
    End If
    
    lColor = m.QB.UnchColor
    
    If Len(strLabel) = 0 Then
        strLabel = " "
    ElseIf InStr(strLabel, "Current Filter") <> 0 Then
        If InStr(strLabel, "(Click") = 0 Then strLabel = strLabel & " (Click here to change Filter)"          '5711
        lColor = vbBlue
        bBold = True
    End If
    
    With fgCurrGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .Cell(flexcpText, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = strLabel
        .Cell(flexcpAlignment, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpForeColor, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = lColor       'm.QB.UnchColor
        
        If bBold Then .Cell(flexcpFontBold, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = True
        
        .RowData(.Rows - 1) = "Label"
        .MergeRow(.Rows - 1) = True
        
        If lRow > -1& Then
            .RowPosition(.Rows - 1) = lRow
        ElseIf .Row >= .FixedRows And .Row < .Rows Then
            .RowPosition(.Rows - 1) = .Row
        End If
                
        ColorQuoteRows
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.AddLabelRow", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeStyle
'' Description: Change the quote board style of the current tab
'' Inputs:      New Style
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeStyle(ByVal Style As eGDQuoteStyle, Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim astrSettings As New cGdArray    ' Settings for the category

    If lTab = -1& Then lTab = vsTab.CurrTab

    ' Hide/Show the grid and the board accordingly...
    If Style = eGDQuoteStyle_Grid Then
        pbQuoteBoard.Visible = False
        fgQuotes.Visible = True
    Else
        fgQuotes.Visible = False
        m.QB.QuoteBoardStyle = Style
        pbQuoteBoard.Visible = True
    End If
    
    ' Save the new style in the category settings...
    TabStr(eGDTabSettings_Style, lTab) = Str(Style)
    
    vsTab.Tag = Str(Style)
    
ErrExit:
    Set astrSettings = Nothing
    Exit Sub
    
ErrSection:
    Set astrSettings = Nothing
    RaiseError "frmQuotes.ChangeStyle", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoAddSymbol
'' Description: Perform an add symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoAddSymbol(Optional ByVal lRow As Long = -1&, Optional ByVal lCol As Long = -1&, Optional ByVal strToSend As String = "")
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbols that the user selected to add
    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumSymbols As Long             ' Number of symbols on the quote board
    Dim bAdded As Boolean               ' Was the symbol added to the data table?
    Dim bRefresh As Boolean             ' Do we need to do a total refresh?
    Dim bAsk As Boolean
    Dim bAddDups As Boolean
    Dim lSymbolID As Long
    Dim lNumFilterSymbols As Long       ' Number of symbols on the filter tab
    Dim lMaxAllowed As Long             ' Maximum number of symbols allowed

    Dim eStyle As eGDQuoteStyle
    Dim fgCurrGrid As VSFlexGrid
    Dim iTabIdx As Long
    Dim strCaption As String

    Dim Bars As cGdBars
    
    ' Tell the user to wait until TotalRefresh is finished...
    If IsBusy Then
        DoEvents
        If IsBusy Then
            InfBox "Please wait until the quote board|has finished refreshing...", "t", "+-OK", "Quote Board", True, 1&
            Exit Sub
        End If
    End If
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
        iTabIdx = vsTab.CurrTab
        strCaption = vsTab.TabCaption(iTabIdx)
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        iTabIdx = m.frmActiveDetTab.MyTabIndex
        strCaption = m.frmActiveDetTab.Caption
        If m.frmActiveDetTab.MyQBStyle = eGDQuoteStyle_Grid Then
            lRow = -1
        ElseIf Not m.frmActiveDetTab.QuoteCellBoard Is Nothing Then
            lRow = m.frmActiveDetTab.QuoteCellBoard.Row
            lCol = m.frmActiveDetTab.QuoteCellBoard.Col
        End If
    End If
    
    eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), iTabIdx)
    
    If strCaption = "(Filter)" Then
        If InfBox("You cannot add symbols to the (Filter) tab.  Would you like to change the filter instead?", "!", "&Change Filter|&Abort", "Error") = "C" Then
            ChangeFilter
        End If
        Exit Sub
    End If
        
    astrSymbols.Create eGDARRAY_Strings
    
    If eStyle = eGDQuoteStyle_Grid Then
        ' Get position to add the symbol...
        If lRow = -1& Then
            If fgCurrGrid.Row >= fgCurrGrid.FixedRows And fgCurrGrid.Row < fgCurrGrid.Rows Then
                lRow = fgCurrGrid.Row
            Else
                lRow = fgCurrGrid.FixedRows
            End If
        ElseIf lRow = 0& Then
            lRow = 1&
        End If
        
'JM:11-30-2007  - original code causes extra row when user cancels from symbol selector (leave awhile then remove)
'        If lRow < fgCurrGrid.Rows - 1 Then
'            AddLabelRow "", lRow
'        End If
    End If
    
    If Len(strToSend) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", HasGold(False), True, "Symbols to Add to Quote Board", False, True, True)
    Else
        Set astrSymbols = frmSymbolSelector.ShowMe(strToSend, HasGold(False), True, "Symbols to Add to Quote Board", False, False, True)
    End If
    
    If astrSymbols.Size > 0 Then
        bRefresh = False
        bAsk = True
        
        lNumSymbols = m.QBData.NumRecords + astrSymbols.Size
        lNumFilterSymbols = NumberOfFilterSymbols
        lMaxAllowed = MaxSymbolsAllowed
        
        If lNumSymbols > lMaxAllowed Then
            If HasGold(False) Then
                If ((lNumSymbols - lNumFilterSymbols) > lMaxAllowed) Then
                    Err.Raise vbObjectError + 1000, , "There cannot be more than " & Str(lMaxAllowed) & " symbols on the quote board"
                Else
                    Err.Raise vbObjectError + 1000, , "There cannot be more than " & Str(lMaxAllowed) & " symbols on the quote board.  Try removing some symbols from the (Filter) tab to allow you to add more symbols."
                End If
            Else
                Err.Raise vbObjectError + 1000, , "You need to upgrade to Gold or Platinum in order to add more symbols on the quote board"
            End If
        Else
            Screen.MousePointer = vbHourglass
            If eStyle = eGDQuoteStyle_Grid Then
                If lRow < fgCurrGrid.Rows - 1 Then
                    AddLabelRow "", lRow
                    fgCurrGrid.RemoveItem lRow
                End If
                For lIndex = 0 To astrSymbols.Size - 1
                    lSymbolID = GetSymbolID(astrSymbols(lIndex))
                    bAdded = False
                    If InStr(TabStr(eGDTabSettings_Symbols, iTabIdx), "," & Str(lSymbolID) & ";Daily,") > 0 Then
                        If bAsk Then
                            bAsk = False
                            If InfBox("The list of symbols you have chosen to add contains symbols that are already on this tab.|Do you wish to add the duplicates?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                                bAddDups = True
                            Else
                                bAddDups = False
                            End If
                        End If
                        
                        If bAddDups Then
                            bAdded = AddSymbolToGrid(astrSymbols(lIndex), "Daily", lRow)
                            If Not m.frmActiveDetTab Is Nothing Then      '6093
                                Set Bars = GetBars(lSymbolID, "Daily")
                                fgCurrGrid.TextMatrix(lRow, GDCol(eGDCol_SymbolID)) = Bars.Prop(eBARS_SymbolID)
                                fgCurrGrid.TextMatrix(lRow, GDCol(eGDCol_Symbol)) = Bars.Prop(eBARS_Symbol)
                                UpdateCols lRow, Bars, , , m.frmActiveDetTab.fgQuotes
                            End If
                            lRow = lRow + 1
                        End If
                    Else
                        bAdded = AddSymbolToGrid(astrSymbols(lIndex), "Daily", lRow)
                        If Not m.frmActiveDetTab Is Nothing Then      '6093
                            Set Bars = GetBars(lSymbolID, "Daily")
                            UpdateCols lRow, Bars, , , m.frmActiveDetTab.fgQuotes
                        End If
                        lRow = lRow + 1
                    End If
                    
                    If bAdded = True Then bRefresh = True
                        
                Next lIndex
                
                lRow = lRow - 1
                If lRow = fgCurrGrid.Rows - 1 Then
                    AddLabelRow "", fgCurrGrid.Rows
                    fgCurrGrid.Row = fgCurrGrid.Rows - 1
                    fgCurrGrid.Col = fgCurrGrid.FrozenCols
                    fgCurrGrid.EditCell
                Else
                    fgCurrGrid.Row = lRow
                    fgCurrGrid.Col = GDCol(eGDCol_Symbol)
                End If
            Else
                For lIndex = 0 To astrSymbols.Size - 1
                    lSymbolID = GetSymbolID(astrSymbols(lIndex))
                    If InStr(TabStr(eGDTabSettings_Symbols, iTabIdx), "," & Str(lSymbolID)) > 0 Then
                        If bAsk Then
                            bAsk = False
                            If InfBox("The list of symbols you have chosen to add contains symbols that are already on this tab.|Do you wish to add the duplicates?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                                bAddDups = True
                            Else
                                bAddDups = False
                            End If
                        End If
                        
                        If bAddDups Then
                            If lIndex > 0 Then
                                bAdded = AddSymbolToBox(astrSymbols(lIndex))    '4361
                            Else
                                bAdded = AddSymbolToBox(astrSymbols(lIndex), lRow, lCol)
                            End If
                        End If
                    ElseIf lIndex > 0 Then
                        bAdded = AddSymbolToBox(astrSymbols(lIndex))
                    Else
                        bAdded = AddSymbolToBox(astrSymbols(lIndex), lRow, lCol)
                    End If

                    If bAdded Then bRefresh = True
                    
                    If Not m.frmActiveDetTab Is Nothing Then
                        Set Bars = GetBars(lSymbolID, "Daily")
                        If Not Bars Is Nothing Then
                            'when style of detached tab is box, the grid for the detached tab is not passed in
                            'new symbol is added to the bottom of grid along with a blank row so use rows-2
                            UpdateCols m.frmActiveDetTab.fgQuotes.Rows - 2, Bars, , , m.frmActiveDetTab.fgQuotes
                            UpdateSymbols m.frmActiveDetTab.QuoteCellBoard.Row, False, True, Bars
                        End If
                    End If
                
                Next lIndex
                
                If m.frmActiveDetTab Is Nothing Then
                    m.QB.NextFreeCell
                ElseIf Not m.frmActiveDetTab.QuoteCellBoard Is Nothing Then
                    m.frmActiveDetTab.QuoteCellBoard.NextFreeCell
                End If
            End If
            
            ' Reset the rows information in the data table...
            If m.frmActiveDetTab Is Nothing Then
                ResetRows
                TotalRefresh False
            Else
                TabStr(eGDTabSettings_Symbols, m.frmActiveDetTab.MyTabIndex) = m.frmActiveDetTab.MySymbols(True)
            End If
                        
            If TabStr(eGDTabSettings_Style, iTabIdx) = QStyle(eGDQuoteStyle_Grid) And m.lTotWidth = 0& Then
                fgCurrGrid.AutoSize 0, fgCurrGrid.Cols - 1, False, 75
            End If
            Screen.MousePointer = vbDefault
        End If
'    Else
'        If TabStr(eGDTabSettings_Style, iTabIdx) = QStyle(eGDQuoteStyle_Grid) Then
'            fgCurrGrid.Row = lRow
'            If fgCurrGrid.MergeRow(lRow) = True Then
'                fgCurrGrid.Col = fgCurrGrid.FrozenCols
'                fgCurrGrid.EditCell
'            Else
'                fgCurrGrid.Col = GdCol(eGDCol_Symbol)
'            End If
'        End If
    End If

ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbols = Nothing
    RaiseError "frmQuotes.DoAddSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoRemoveSymbol
'' Description: Perform a remove symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoRemoveSymbol()
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim strAsk As String
    Dim strSymbol As String
    Dim nSymbolID As Long
    Dim strPeriod As String
    Dim bRemove As Boolean
    Dim astrList As cGdArray
    Dim astrReturn As cGdArray
    Dim lIndex As Long
    Dim bRefresh As Boolean             ' Do we need to do a refresh?
    
    Dim eStyle As eGDQuoteStyle
    Dim fgCurrGrid As VSFlexGrid
    Dim QB As cQuoteCellBoard
    Dim iTabIdx As Long
    Dim strCaption As String

    ' Tell the user to wait until TotalRefresh is finished...
    If IsBusy Then
        DoEvents
        If IsBusy Then
            InfBox "Please wait until the quote board|has finished refreshing...", "t", "+-OK", "Quote Board", True, 1&
            Exit Sub
        End If
    'ElseIf vsTab.TabCaption(vsTab.CurrTab) = "(Filter)" Then
    '    If InfBox("You cannot remove symbols from the (Filter) tab.  Would you like to clear the filter instead?", "!", "&Clear Filter|&Abort", "Warning") = "C" Then
    '        bRefresh = ClearFilterTab
    '        ShowCategory vsTab.CurrTab
    '        If bRefresh Then TotalRefresh False
    '        SortFilterTab
    '    End If
    '    Exit Sub
    End If
    
    Set astrList = New cGdArray
    astrList.Create eGDARRAY_Strings
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
        Set QB = m.QB
        iTabIdx = vsTab.CurrTab
        strCaption = vsTab.TabCaption(iTabIdx)
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        iTabIdx = m.frmActiveDetTab.MyTabIndex
        strCaption = m.frmActiveDetTab.Caption
        lRow = -1
    End If
    
    eStyle = m.tblTabInfo(TabField(eGDTabSettings_Style), iTabIdx)
    
    If eStyle = eGDQuoteStyle_Grid And Not fgCurrGrid Is Nothing Then
        If fgCurrGrid.Row = fgCurrGrid.Rows - 1 Then GoTo ErrExit
        If (strCaption = "(Filter)") And (fgCurrGrid.MergeRow(fgCurrGrid.Row) = True) Then GoTo ErrExit
        With fgCurrGrid
            lRow = .Row
            
            If .MergeRow(lRow) = True Then
                strSymbol = Trim(.TextMatrix(lRow, GDCol(eGDCol_Period)))
                If Len(strSymbol) = 0 Then
                    strPeriod = "Blank Row"
                Else
                    strPeriod = "Label Row"
                End If
            Else
                strSymbol = Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)
                strPeriod = .TextMatrix(lRow, GDCol(eGDCol_Period))
            End If
            
            For lIndex = .FixedRows To .Rows - 2
                If .MergeRow(lIndex) = True Then
                    If Len(Trim(.TextMatrix(lIndex, GDCol(eGDCol_Period)))) = 0 Then
                        astrList.Add "  (Blank Row)" & vbTab & Str(lIndex)
                    Else
                        astrList.Add .TextMatrix(lIndex, GDCol(eGDCol_Period)) & "  (Label Row)" & vbTab & Str(lIndex)
                    End If
                Else
                    astrList.Add Parse(.TextMatrix(lIndex, GDCol(eGDCol_Symbol)), "(", 1) & "  (" & .TextMatrix(lIndex, GDCol(eGDCol_Period)) & ")" & vbTab & Str(lIndex)
                End If
            Next lIndex
        End With
    ElseIf Not QB Is Nothing Then
        If Not QB.Cell(QB.Row, QB.Col) Is Nothing Then
            With QB
                strSymbol = Parse(.Cell(.Row, .Col).Symbol, "(", 1)
                strPeriod = "Daily"
                
                For lIndex = 1 To .CellCount
                    astrList.Add Parse(.CellItem(lIndex).Symbol, "(", 1) & "  (Daily)" & vbTab & Str(.CellItem(lIndex).Row) & "," & Str(.CellItem(lIndex).Col)
                Next lIndex
                
                ' sort the symbol list for the box-style
                astrList.Sort
            End With
        End If
    End If
        
    If strPeriod = "Blank Row" Or strPeriod = "Label Row" Then
        astrReturn.Add Str(lRow)
    Else
        Set astrReturn = frmDelete.ShowMe(astrList, strSymbol & "  (" & strPeriod & ")")
    End If
    
    If Not astrReturn Is Nothing Then
        Screen.MousePointer = vbHourglass
        bRemove = False
        For lIndex = astrReturn.Size - 1 To 0 Step -1
            If eStyle = eGDQuoteStyle_Grid Then
                If RemoveSymbolFromGrid(CLng(Val(astrReturn(lIndex)))) = True Then
                    bRemove = True
                End If
            ElseIf RemoveSymbolFromBox(CLng(Val(Parse(astrReturn(lIndex), ",", 1))), CLng(Val(Parse(astrReturn(lIndex), ",", 2)))) = True Then
                bRemove = True
            End If
        Next lIndex
        
        ' Reset the rows information in the data table...
        If m.frmActiveDetTab Is Nothing Then
            ResetRows
        Else
            TabStr(eGDTabSettings_Symbols, iTabIdx) = m.frmActiveDetTab.MySymbols(True)
        End If
        
'        If bRemove And g.RealTime.Active Then
            'g.RealTime.UpdateSymbolList
'        End If
        Screen.MousePointer = vbDefault
    End If
    
ErrExit:
    Set astrList = Nothing
    Set astrReturn = Nothing
    Exit Sub
    
ErrSection:
    Set astrList = Nothing
    Set astrReturn = Nothing
    RaiseError "frmQuotes.DoRemoveSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowSettings
'' Description: Show the quote board settings form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowSettings()
On Error GoTo ErrSection:

    Dim strGridFont As String           ' String version of the grid font
    Dim strBoxFont As String            ' String version of the box font
    Dim UpColor As Long                 ' Color of an up day
    Dim DownColor As Long               ' Color of a down day
    Dim UnchColor As Long               ' Color of an unchanged day
    Dim UpdateColor As Long             ' Color of an updated cell
    Dim bFullHeight As Boolean          ' Does the bar need to be full height?
    Dim TmpFont As New StdFont          ' Temporary font object
    Dim Style As eGDQuoteStyle          ' Box-Style quote board style
    Dim lIndex As Long                  ' Index into a for loop
    
    Dim lForexBkColor As Long
    Dim lForexTextColor As Long
    Dim lForexUpDownBk As Long
    Dim lUseUpdateColor As Long
    Dim lColorSymbol As Long
    Dim lCompactQB As Long
    
    Dim eShowExtraInfo As eGEQbExtraInfo
    
    Dim bShowButtons As Boolean
    
    Dim frm As frmDetachedQBTab
    
    ' Get current values of the properties...
    strGridFont = FontToString(fgQuotes.Font)
    strBoxFont = FontToString(m.QB.Font)
    UpColor = m.QB.UpColor
    DownColor = m.QB.DownColor
    UnchColor = m.QB.UnchColor
    UpdateColor = m.QB.UpdateColor
    lForexBkColor = m.QB.ForexBkColor
    lForexTextColor = m.QB.ForexTextColor
    lForexUpDownBk = m.QB.ForexUpDownBk
    bFullHeight = m.QB.FullBarHeight
    lUseUpdateColor = m.QB.UseUpdateColor
    lColorSymbol = m.QB.ColorSymbol
    eShowExtraInfo = m.QB.ShowExtraInfo
    bShowButtons = m.bShowButtons
    
    If m.frmActiveDetTab Is Nothing Then
        Style = TabStr(eGDTabSettings_Style, vsTab.CurrTab)        'm.DefaultStyle
    Else
        Style = TabStr(eGDTabSettings_Style, m.frmActiveDetTab.MyTabIndex)
    End If
    lCompactQB = m.QB.CompactQB(Style)
    
    ' Show the Quote Board Settings form...
    If frmQuoteSettings.ShowMe(strGridFont, strBoxFont, UpColor, DownColor, UnchColor, UpdateColor, _
        bFullHeight, Style, lForexBkColor, lForexTextColor, lForexUpDownBk, lUseUpdateColor, lColorSymbol, _
        eShowExtraInfo, lCompactQB, bShowButtons) Then
        ' If the user clicks on OK, set the new values of the properties...
        FontFromString fgQuotes.Font, strGridFont
        fgQuotes.Font = fgQuotes.Font
        FontFromString TmpFont, strBoxFont
        
        If Style = eGDQuoteStyle_Forex Then
            m.QB.ForexBkColor = lForexBkColor
            m.QB.ForexTextColor = lForexTextColor
            m.QB.ForexUpDownBk = lForexUpDownBk
            m.QB.CompactQB(eGDQuoteStyle_Forex) = lCompactQB
        Else
            m.QB.UnchColor = UnchColor
            m.QB.QuoteBoardStyle = Style
            m.QB.FullBarHeight = bFullHeight
            m.QB.ShowExtraInfo = eShowExtraInfo
            m.QB.CompactQB(eGDQuoteStyle_OHLC) = lCompactQB
        End If
        
        m.QB.UpColor = UpColor
        m.QB.DownColor = DownColor
        m.QB.UpdateColor = UpdateColor
        m.QB.UseUpdateColor = lUseUpdateColor
        m.QB.ColorSymbol = lColorSymbol
        m.QB.Font = TmpFont
        
        If Style <> m.DefaultStyle And Style <> eGDQuoteStyle_Forex Then
            For lIndex = 0 To m.tblTabInfo.NumRecords - 1
                If TabStr(eGDTabSettings_Style, lIndex) <> Str(QStyle(eGDQuoteStyle_Grid)) _
                   And TabStr(eGDTabSettings_Style, lIndex) <> Str(QStyle(eGDQuoteStyle_Forex)) Then
                        
                        TabStr(eGDTabSettings_Style, lIndex) = Str(Style)
                        m.DefaultStyle = Style
                End If
            Next lIndex
        End If
        
        ' If the current tab is a grid style, remove the blank row since ShowCategory
        ' is going to put it back...
        If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
            fgQuotes.RemoveItem fgQuotes.Rows - 1
        End If
                
        ShowCategory
        DoEvents            'give main QB a chance to update
        
        If m.bShowButtons <> bShowButtons Then
            m.bShowButtons = bShowButtons
            PromptShowButtons
            Form_Resize
            If fgQuotes.Visible Then
                MoveFocus fgQuotes
            ElseIf pbQuoteBoard.Visible Then
                MoveFocus pbQuoteBoard
            End If
        End If
        
        If Not m.aDetachedTabs Is Nothing Then
            For lIndex = 0 To m.aDetachedTabs.Size - 1
                DoEvents
                Set frm = m.aDetachedTabs(lIndex)
                If Not frm Is Nothing Then frm.UpdateSettings m.QB, fgQuotes
                DoEvents
            Next
        End If
        
    End If

ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ShowSettings", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSettings
'' Description: Load the quote board settings from the registry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSettings()
On Error GoTo ErrSection:

    Dim strFont As String               ' Font string from the ini file
    
    ' Load font information from the ini file...
    strFont = GetIniFileProperty("QuoteCellBoard", "", "Fonts", g.strIniFile)
    If Len(strFont) > 0 Then FontFromString pbQuoteBoard.Font, strFont
    strFont = GetIniFileProperty("QuoteBoard", "", "Fonts", g.strIniFile)
    If Len(strFont) > 0 Then FontFromString fgQuotes.Font, strFont
    
    ' Load color information from the ini file...
    m.QB.UpColor = GetIniFileProperty("UpColor", RGB(0, 128, 0), "QuoteBoard", g.strIniFile)
    m.QB.DownColor = GetIniFileProperty("DownColor", vbRed, "QuoteBoard", g.strIniFile)
    m.QB.UnchColor = GetIniFileProperty("UnchColor", vbBlack, "QuoteBoard", g.strIniFile)
    m.QB.UpdateColor = GetIniFileProperty("UpdateColor", vbBlue, "QuoteBoard", g.strIniFile)
    m.QB.UseUpdateColor = GetIniFileProperty("UseUpdateColor", 1, "QuoteBoard", g.strIniFile)
    m.QB.ColorSymbol = GetIniFileProperty("ColorSymbol", 1, "QuoteBoard", g.strIniFile)
    
    'load show date time from ini file...
    m.QB.ShowExtraInfo = GetIniFileProperty("ShowExtraInfo", eGEQbExtraInfo_TimeStamp, "QuoteBoard", g.strIniFile)
    
    'load whether or not to use compact QB
    m.QB.CompactQB(eGDQuoteStyle_OHLC) = GetIniFileProperty("CompactQB", 0, "QuoteBoard", g.strIniFile)
    m.QB.CompactQB(eGDQuoteStyle_Forex) = GetIniFileProperty("CompactQBForex", 0, "QuoteBoard", g.strIniFile)
    
    'load QB style  -aardvark 6781
    m.QB.QuoteBoardStyle = GetIniFileProperty("DisplayStyle", eGDQuoteStyle_OHLC, "QuoteBoard", g.strIniFile)
    
    ' Load bar information from the ini file...
    m.QB.FullBarHeight = GetIniFileProperty("FullBarHeight", 0&, "QuoteBoard", g.strIniFile)
    m.DefaultStyle = GetIniFileProperty("BarStyle", eGDQuoteStyle_OHLC, "QuoteBoard", g.strIniFile)
    
    'Load forex settings from ini file
    m.QB.ForexUpDownBk = GetIniFileProperty("ForexUpDownBk", 0, "QuoteBoard", g.strIniFile)
    m.QB.ForexBkColor = GetIniFileProperty("ForexBkColor", RGB(128, 128, 128), "QuoteBoard", g.strIniFile)
    m.QB.ForexTextColor = GetIniFileProperty("ForexTextColor", vbWhite, "QuoteBoard", g.strIniFile)
    
    'show buttons setting
    m.bShowButtons = GetIniFileProperty("ShowButtons", 1, "QuoteBoard", g.strIniFile)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.LoadSettings", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveSettings
'' Description: Save the quote board settings to the registry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveSettings()
On Error GoTo ErrSection:

    ' Save font information to the registry...
    SetIniFileProperty "QuoteCellBoard", FontToString(pbQuoteBoard.Font), "Fonts", g.strIniFile
    SetIniFileProperty "QuoteBoard", FontToString(fgQuotes.Font), "Fonts", g.strIniFile
    
    ' Save color information to the registry...
    SetIniFileProperty "UpColor", m.QB.UpColor, "QuoteBoard", g.strIniFile
    SetIniFileProperty "DownColor", m.QB.DownColor, "QuoteBoard", g.strIniFile
    SetIniFileProperty "UnchColor", m.QB.UnchColor, "QuoteBoard", g.strIniFile
    SetIniFileProperty "UpdateColor", m.QB.UpdateColor, "QuoteBoard", g.strIniFile
    SetIniFileProperty "UseUpdateColor", m.QB.UseUpdateColor, "QuoteBoard", g.strIniFile
    SetIniFileProperty "ColorSymbol", m.QB.ColorSymbol, "QuoteBoard", g.strIniFile
    
    'save show datetime, bid/ask flag
    SetIniFileProperty "ShowExtraInfo", m.QB.ShowExtraInfo, "QuoteBoard", g.strIniFile
    
    'save whether to use compact QB
    SetIniFileProperty "CompactQB", m.QB.CompactQB(eGDQuoteStyle_OHLC), "QuoteBoard", g.strIniFile
    SetIniFileProperty "CompactQBForex", m.QB.CompactQB(eGDQuoteStyle_Forex), "QuoteBoard", g.strIniFile
    
    'save QB display style      -aardvark 6781
    SetIniFileProperty "DisplayStyle", m.QB.QuoteBoardStyle, "QuoteBoard", g.strIniFile
    
    ' Save bar information to the registry...
    SetIniFileProperty "FullBarHeight", CLng(m.QB.FullBarHeight), "QuoteBoard", g.strIniFile
    SetIniFileProperty "BarStyle", m.DefaultStyle, "QuoteBoard", g.strIniFile
    
    'Save settings for forex quoteboard to ini file
    SetIniFileProperty "ForexUpDownBk", m.QB.ForexUpDownBk, "QuoteBoard", g.strIniFile
    SetIniFileProperty "ForexBkColor", m.QB.ForexBkColor, "QuoteBoard", g.strIniFile
    SetIniFileProperty "ForexTextColor", m.QB.ForexTextColor, "QuoteBoard", g.strIniFile
    
    'show buttons setting
    SetIniFileProperty "ShowButtons", m.bShowButtons, "QuoteBoard", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.SaveSettings", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTabInfo
'' Description: Load the tab info table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTabInfo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lRow As Long                    ' Current row in the table
    Dim strTemp As String               ' Temporary string
    Dim strSymbol As String             ' Symbol to add to the string
    Dim strPeriod As String             ' Period of the symbol to add to the string
    Dim strDefaultSymbols As String     ' Default set of symbols for new user
    Dim strDefaultBoxSyms As String     ' Default set of symbols for box-style
    Dim SymGroup As New cSymbolGroup    ' Temporary symbol group object
    Dim astrSymbols As New cGdArray     ' Array of symbols
    Dim astrTemp As New cGdArray        ' Temporary string array

    ' Set up the table...
    'Symbols field is comma delimited: symbolID;periodicity,Label;LabelText,symbolID;periodicity...
    'Fields field is pipe delimited: columnName;hidden bool flag;criteriaName;columnWidth|columnName;hidden bool flag;criteriaName;columnWidth|...
    Set m.tblTabInfo = New cGdTable
    With m.tblTabInfo
        .CreateField eGDARRAY_Strings, TabField(eGDTabSettings_Name), "Name"
        .CreateField eGDARRAY_Strings, TabField(eGDTabSettings_Style), "Style"
        .CreateField eGDARRAY_Strings, TabField(eGDTabSettings_Symbols), "Symbols"
        .CreateField eGDARRAY_Strings, TabField(eGDTabSettings_Fields), "Fields"
        .CreateField eGDARRAY_Strings, TabField(eGDTabSettings_FilterID), "FilterID"
        .CreateField eGDARRAY_Longs, TabField(eGDTabSettings_Form), "DetachedTabForm", 0
    End With
        
    ' Create the string arrays...
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    astrSymbols.Create eGDARRAY_Strings
    astrTemp.Create eGDARRAY_Strings

    ' Set up the default symbols string in case we need it...
    ''strDefaultSymbols = ",50;Daily,27;Daily,207;Daily,"
    strDefaultBoxSyms = "50,27,207,"
    If HasModule("S") Then
        ''strDefaultSymbols = strDefaultSymbols & "11936;Daily,"
        strDefaultBoxSyms = strDefaultBoxSyms & "11936,15759,332,"
    End If
    If HasModule("F") Then
        ''strDefaultSymbols = strDefaultSymbols & "41180;Daily,41183;Daily,"
        strDefaultBoxSyms = strDefaultBoxSyms & "41142,42020,203182,203188,200929,"
    End If
    strDefaultSymbols = Replace(strDefaultBoxSyms, ",", ";Daily,")
       
       
    ' If the QuoteBoard.INF exists, load it...
    If FileExist(AddSlash(App.Path) & "Custom\QuoteBoard.INF") Then
        m.tblTabInfo.FromString FileToString(AddSlash(App.Path) & "Custom\QuoteBoard.INF"), vbLf, vbTab
        
    ' Otherwise, create it from existing information...
    Else
        ' If the QuoteList.GRP exists, get the symbols/ids from it...
        If SymGroup.FromFile(AddSlash(App.Path) & "Custom", "QuoteList.GRP", True) Then
            ' Add a "My Quotes" tab that contains all symbols...
            TabStr(eGDTabSettings_Name, 0) = "My Quotes"
            TabStr(eGDTabSettings_Style, 0) = Str(QStyle(eGDQuoteStyle_Grid))
            
            strTemp = ","
            For lIndex = 0 To SymGroup.SymbolIDs.Size - 1
                strTemp = strTemp & SymGroup.SymbolIDs(lIndex) & ";Daily,"
            Next lIndex
            For lIndex = 0 To SymGroup.Symbols.Size - 1
                strTemp = strTemp & SymGroup.Symbols(lIndex) & ";Daily,"
            Next lIndex
            TabStr(eGDTabSettings_Symbols, 0) = strTemp
        End If
        
        ' If the QuoteList.CAT file exists, load all existing custom categories...
        If astrTemp.FromFile(AddSlash(App.Path) & "Custom\QuoteList.CAT") Then
            For lIndex = 0 To astrTemp.Size - 1
                If Len(Trim(astrTemp(lIndex))) > 0 Then
                    m.tblTabInfo.AddRecord Parse(astrTemp(lIndex), vbTab, 1)
                    lRow = m.tblTabInfo.NumRecords - 1
                    astrSymbols.SplitFields Parse(astrTemp(lIndex), vbTab, 2), ","
                    strTemp = ","
                    For lIndex2 = 0 To astrSymbols.Size - 1
                        strTemp = strTemp & astrSymbols(lIndex2) & ";Daily,"
                    Next lIndex2
                    TabStr(eGDTabSettings_Style, lRow) = Str(QStyle(eGDQuoteStyle_Grid))
                    TabStr(eGDTabSettings_Symbols, lRow) = strTemp
                End If
            Next lIndex
            
            m.tblTabInfo.AddRecord "Box Style Sample"
            lRow = m.tblTabInfo.NumRecords - 1
            TabStr(eGDTabSettings_Style, lRow) = Str(m.DefaultStyle)
            TabStr(eGDTabSettings_Symbols, lRow) = strDefaultBoxSyms
        
        ' Otherwise create a sample grid style and a sample box style category...
        Else
            m.tblTabInfo.AddRecord "Box Style Sample"
            lRow = m.tblTabInfo.NumRecords - 1
            TabStr(eGDTabSettings_Style, lRow) = Str(m.DefaultStyle)
            TabStr(eGDTabSettings_Symbols, lRow) = strDefaultBoxSyms
            
            m.tblTabInfo.AddRecord "Grid Style Sample"
            lRow = m.tblTabInfo.NumRecords - 1
            TabStr(eGDTabSettings_Style, lRow) = Str(QStyle(eGDQuoteStyle_Grid))
            TabStr(eGDTabSettings_Symbols, lRow) = strDefaultSymbols
        End If
    End If
       
    ' Find the filter tab if it exists and set it up...
    If m.tblTabInfo(TabField(eGDTabSettings_Name), m.tblTabInfo.NumRecords - 1) <> "(Filter)" Then
        m.tblTabInfo.AddRecord "(Filter)" & vbTab & Str(QStyle(eGDQuoteStyle_Grid)) & vbTab & "" & vbTab & "" & vbTab & ""
    End If
    
    ' Get the complete set of symbol/periods out of the table...
    For lIndex = 0 To m.tblTabInfo.NumRecords - 1
        If Len(TabStr(eGDTabSettings_Style, lIndex)) = 0 Then
            Err.Raise vbObjectError + 1000, , "Error loading QuoteBoard file"
        End If

        astrTemp.SplitFields TabStr(eGDTabSettings_Symbols, lIndex), ","
        For lIndex2 = 0 To astrTemp.Size - 1
            strSymbol = Parse(astrTemp(lIndex2), ";", 1)
            If TabStr(eGDTabSettings_Style, lIndex) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                strPeriod = Parse(astrTemp(lIndex2), ";", 2)
            Else
                strPeriod = "Daily"
            End If
            If Len(strSymbol) > 0 And Len(strPeriod) > 0 Then
                m.astrSymbols.Add strSymbol & ";" & strPeriod
            End If
        Next lIndex2
    Next lIndex
    m.astrSymbols.Sort

ErrExit:
    Set SymGroup = Nothing
    Set astrTemp = Nothing
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set SymGroup = Nothing
    Set astrTemp = Nothing
    Set astrSymbols = Nothing
    RaiseError "frmQuotes.LoadTabInfo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateTable
'' Description: Update the table from the ticks that have come through on the
''              real time server
'' Inputs:      None
'' Returns:     True if need to do a total refresh
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UpdateTable() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewBar As Boolean              ' Are we starting a new bar?
    Dim Bars As cGdBars                 ' Temporary bars object
    Dim bUpdate As Boolean              ' Should we update this symbol?
    Dim nSaveUseUpdateColor As Long
    
    Dim i&, iRowFound&
    Dim frmQBTab As frmDetachedQBTab
    
    ' Walk through and update the bars in the data table...
    For lIndex = 0 To m.QBData.NumRecords - 1
        ' Only do rows before where TotalRefresh is currently operating
        ' (if TotalRefresh is even currently processing)
        If m.nRefreshRow <> 0 And lIndex >= m.nRefreshRow Then Exit For
        
        bUpdate = False
        Set Bars = m.BarsColl(TblStr(eQbTbl_SearchKey, lIndex))
        If Not Bars Is Nothing Then
            ' If tick is in a new bar then call TotalRefresh to reload...
            If g.RealTime.UpdateBars(Bars, bNewBar) = True Then
                If bNewBar Then
                    ' TLB 12/23/2009: for daily bars we'll just get a full reload, but
                    ' for intraday bars we just want a new SpliceBars to get done (in TotalRefresh)
                    If Not Bars.IsIntraday Then
                        Set m.BarsColl(TblStr(eQbTbl_SearchKey, lIndex)) = Nothing
                    End If
                    UpdateTable = True
                    Exit For
                End If
                TblNum(eQbTbl_Recalc, lIndex) = 1
                TblNum(eQbTbl_Dirty, lIndex) = 1
                bUpdate = True
            ElseIf TblNum(eQbTbl_Dirty, lIndex) = 1 Then
                bUpdate = True
            ElseIf m.QB.QuoteBoardStyle = eGDQuoteStyle_Forex Then
                bUpdate = g.RealTime.UpdateBidAsk(Bars)
            End If
           
            If bUpdate Then
                ' TLB 8/24/2009: just to avoid the "all blue" visual effect when streaming has just started,
                ' temporarily turn off the update color for a symbol with a new bar within 30 seconds of starting
                nSaveUseUpdateColor = kNullData
                If bNewBar And g.RealTime.Active Then
                    If gdTickCount <= m.dWhenStreamingStarted + 30000 Then
                        nSaveUseUpdateColor = m.QB.UseUpdateColor
                        m.QB.UseUpdateColor = 0
                    End If
                End If
            
                ' Update either the grid or box, whichever is visible...
                If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = QStyle(eGDQuoteStyle_Grid) Then
                    UpdateSymbols lIndex
                ElseIf Bars.Prop(eBARS_Periodicity) = ePRD_Days + 1 Then
                    m.QB.UpdateSymbol Bars
                End If
                TblNum(eQbTbl_Dirty, lIndex) = 0
                
                For i = 0 To m.aDetachedTabs.Size - 1
                    Set frmQBTab = m.aDetachedTabs(i)
                    If Not frmQBTab Is Nothing Then
                        iRowFound = frmQBTab.UpdateRT(Bars)
                        If iRowFound > 0 Then
                            UpdateCols iRowFound, Bars, , , frmQBTab.fgQuotes
                        End If
                    End If
                Next
                
                If nSaveUseUpdateColor <> kNullData Then
                    m.QB.UseUpdateColor = nSaveUseUpdateColor
                End If
                
''                If FormIsLoaded("frmTTSummary") Then frmTTSummary.RefreshPrices Bars
            End If
        End If
    Next lIndex
        
ErrExit:
    Set frmQBTab = Nothing
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.UpdateTable", eGDRaiseError_Raise
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbolToTable
'' Description: Add a symbol to the data table
'' Inputs:      PoolRec #, Period, Symbol
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddSymbolToTable(ByVal lRecNum As Long, ByVal strPeriod As String, _
            Optional ByVal strSymbol As String = "") As Boolean
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of something in a string
    Dim strSecType As String            ' Security type for the symbol
    Dim lSymbolID As Long               ' Symbol ID for the symbol
    Dim Bars As cGdBars
    
    ' If a pool record is specified, then get the symbol from there...
    If lRecNum <> -1 Then
        strSymbol = g.SymbolPool.Symbol(lRecNum)
        strSecType = GetSecType(g.SymbolPool.SecType(lRecNum))
        lSymbolID = g.SymbolPool.SymbolID(lRecNum)
    Else
        If InStr(strSymbol, " ") <> 0 Then
            If InStr(strSymbol, "-") <> 0 Then
                strSecType = "FO"
            Else
                strSecType = "SO"
            End If
        End If
        lSymbolID = 0&
    End If
        
    If Len(strSymbol) > 0 Then
        If TblSearch(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod, lPos) = False Then
            If m.QBData.NumRecords < MaxSymbolsAllowed Then
                'm.QBData.AddRecord strSymbol & vbTab & strPeriod, lPos, ","
                m.QBData.AddRecord SymbolOrSymbolID(lSymbolID, strSymbol) & vbTab & strPeriod, lPos, ","
                TblStr(eQbTbl_SecType, lPos) = UCase(strSecType)
                TblStr(eQbTbl_Symbol, lPos) = strSymbol
                TblNum(eQbTbl_SymbolID, lPos) = lSymbolID
                TblStr(eQbTbl_Rows, lPos) = ""
                TblNum(eQbTbl_Recalc, lPos) = 1
                TblNum(eQbTbl_Dirty, lPos) = 1
                AddSymbolToTable = True
                ' add bars to collection
                Set Bars = New cGdBars
                SetBarProperties Bars, strSymbol
                g.RealTime.AddTickBuffer Bars, False
                Bars.ArrayMask = eBARS_Ask
                Set m.BarsColl(TblStr(eQbTbl_SearchKey, lPos)) = Bars
                If Not m.frmActiveDetTab Is Nothing Then
                    'this is done in ResetRows when tab is not detached
                    Set m.alDataIndex = m.QBData.CreateSortedIndex(TblField(eQbTbl_Rows))
                    m.hDataIndex = m.alDataIndex.ArrayHandle
                    TotalRefresh False      'to get data for new symbol
                End If
            End If
        Else
            TblNum(eQbTbl_Recalc, lPos) = 1
            TblNum(eQbTbl_Dirty, lPos) = 1
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.AddSymbolToTable", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTable
'' Description: Load up the table from the unique set of symbol/period pairs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadTable() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol to add to the table
    Dim strPeriod As String             ' Period of the symbol to add

    ' Set RefreshRow so TotalRefresh and UpdateTable will not interfere...
    m.nRefreshRow = 1
                
    ' Start adding symbols...
    For lIndex = 0 To m.astrSymbols.Size - 1
        strSymbol = Parse(m.astrSymbols(lIndex), ";", 1)
        strPeriod = Parse(m.astrSymbols(lIndex), ";", 2)
        
        If strSymbol <> "Label" Then
            If Len(strPeriod) = 0 Then strPeriod = "Daily"
            
            If Val(strSymbol) <> 0 Then
                AddSymbolToTable g.SymbolPool.PoolRecForSymbolID(strSymbol), strPeriod
            ElseIf Len(strSymbol) > 0 Then
                AddSymbolToTable -1, strPeriod, strSymbol
            End If
        End If
    Next lIndex
    
    ' Reset the rows information in the data table...
    ''ResetRows
    
    AlignColumns
    ShowCategory
        
    m.nRefreshRow = 0
    'If g.RealTime.Active Then g.RealTime.UpdateSymbolList

    LoadTable = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.LoadTable", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(Optional fg As VSFlexGrid = Nothing)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim astrCols As New cGdArray        ' Column information for setting up the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lTab As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Sorted fields array
    Dim lPos As Long                    ' Position of an item in an array
    Dim strName As String               ' Field name
    Dim strCriteriaID As String         ' Criteria ID
    Dim lPos2 As Long                   ' Position of an item in an array

    Dim fgCurrent As VSFlexGrid
    
    If fg Is Nothing Then
        Set fgCurrent = fgQuotes
    Else
        Set fgCurrent = fg
    End If

    With fgCurrent
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightNever '(so colors will show for selected row)
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Cols = 0
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .Editable = flexEDKbd ' = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        ''.ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .AllowBigSelection = False
        .MergeCells = flexMergeFree
        
        ' 03/05/2010 DAJ: Since we may have imported quote board tabs that have columns that
        ' are not on the first tab, walk through all of the tabs to make sure that all of
        ' the columns are created...
        astrFields.Create eGDARRAY_Strings
        For lTab = 0 To m.tblTabInfo.NumRecords - 1
            astrCols.Create eGDARRAY_Strings
            astrCols.SplitFields TabStr(eGDTabSettings_Fields, lTab), "|"
            
            For lIndex = 0 To astrCols.Size - 1
                strName = Parse(astrCols(lIndex), ";", 1)    'field name
                If astrFields.BinarySearch(strName, lPos) = False Then
                    .Cols = .Cols + 1
                    .TextMatrix(0, .Cols - 1) = strName
                    
                    strCriteriaID = Parse(astrCols(lIndex), ";", 3)
                    .ColData(.Cols - 1) = strCriteriaID
                    If Len(strCriteriaID) > 0 Then
                        If m.astrCriteria.BinarySearch(strCriteriaID, lPos2) = False Then
                            m.astrCriteria.Add strCriteriaID, lPos2
                        End If
                    End If
                    
                    astrFields.Add strName, lPos
                    m.astrFields.Add Parse(astrCols(lIndex), ";", 1)
                End If
            Next lIndex
        Next lTab
        
        '.ColDataType(GDCol(eGDCol_Recalc)) = flexDTBoolean
        
        '.ColHidden(GDCol(eGDCol_Recalc)) = True
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        .ColHidden(GDCol(eGDCol_SecType)) = True
        
        ' TLB: keep symbol column frozen when scrolling
        .FrozenCols = GDCol(eGDCol_Symbol) + 1 'GDCol(eGDCol_NumFixed)
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Set astrCols = Nothing
    Exit Sub
    
ErrSection:
    Set astrCols = Nothing
    RaiseError "frmQuotes.InitGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadQBFs
'' Description: Load the custom quote board fields if there are any
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadQBFs()
On Error GoTo ErrSection:

    Dim astrCols As New cGdArray        ' Array of column information
    Dim Criteria As New cCriteria       ' Criteria object to add to the collection
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' Criteria ID from the string
    
    ' Create the array of column information...
    astrCols.Create eGDARRAY_Strings
    astrCols.SplitFields m.strDefaultFields, "|"
    
    ' Convert any existing Quote Board Fields to regular Criteria...
    ConvertQBF
    
    ' Initialize the collection of criteria...
    Set m.QBFs = New cGdTree
    
    ' Walk through the array looking for Criteria ID's...
    For lIndex = 0 To astrCols.Size - 1
        strID = Parse(astrCols(lIndex), ";", 3)
        
        ' If we find a Criteria ID, add the criteria to the collection and
        ' add a field to the data table...
        If Len(strID) > 0 Then
            Set Criteria = New cCriteria
            Criteria.FromFile AddSlash(App.Path) & "Custom\", strID
            
            If HasModule(Criteria.Required) Then
                m.QBFs.Add Criteria, strID
                m.QBData.CreateField eGDARRAY_Doubles, , strID
            End If
        End If
    Next lIndex

ErrExit:
    Set astrCols = Nothing
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    Set astrCols = Nothing
    Set Criteria = Nothing
    RaiseError "frmQuotes.LoadQBFs", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadQbfs2
'' Description: Load the custom quote board fields if there are any
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadQbfs2()
On Error GoTo ErrSection:

    Dim Criteria As New cCriteria       ' Criteria object to add to the collection
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' Criteria ID from the string
    
    ' Convert any existing Quote Board Fields to regular Criteria...
    ConvertQBF
    
    ' Initialize the collection of criteria...
    Set m.QBFs = New cGdTree
    
    ' Walk through the array looking for Criteria ID's...
    For lIndex = 0 To m.astrCriteria.Size - 1
        strID = m.astrCriteria(lIndex)
        
        Set Criteria = New cCriteria
        Criteria.FromFile AddSlash(App.Path) & "Custom\", strID
        
        If HasModule(Criteria.Required) Then
            m.QBFs.Add Criteria, strID
            m.QBData.CreateField eGDARRAY_Doubles, , strID
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.LoadQbfs2"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResetRows
'' Description: Reset the rows field of the data table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResetRows(Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position of the symbol/period in the table
    Dim strSymbol As String             ' Symbol from the grid
    Dim strSymbolID As String           ' Symbol ID from the grid
    Dim strPeriod As String             ' Period from the grid
    Dim strSymbols As String            ' Category symbols to store in table
    Dim lSymbolID As Long               ' Symbol ID from the grid
        
    ' Don't do anything if data has not been installed yet...
    If g.SymbolPool.NumRecords = 0 Then Exit Sub
    
    ' Default to the current tab...
    If lTab = -1& Then lTab = vsTab.CurrTab
    
    ' Clear out the rows field of the data table...
    For lIndex = 0 To m.QBData.NumRecords - 1
        TblStr(eQbTbl_Rows, lIndex) = ""
    Next lIndex
    
    If TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        strSymbols = ","
        With fgQuotes
            ' Walk through all of the rows in the grid (except for the blank row at the bottom)...
            For lIndex = .FixedRows To .Rows - 2
                ' Only worry about non-label rows...
                If .MergeRow(lIndex) = False Then
                    strSymbolID = .TextMatrix(lIndex, GDCol(eGDCol_SymbolID))
                    lSymbolID = Val(strSymbolID)
                    strSymbol = Parse(.TextMatrix(lIndex, GDCol(eGDCol_Symbol)), "(", 1)
                    strPeriod = .TextMatrix(lIndex, GDCol(eGDCol_Period))
                    
                    ' Add this row to the appropriate row in the data table...
                    If TblSearch(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod, lPos) Then
                        TblStr(eQbTbl_Rows, lPos) = TblStr(eQbTbl_Rows, lPos) & "," & Format(lIndex, "000")
                    End If
                Else
                    strSymbolID = ""
                    lSymbolID = 0&
                    strSymbol = "Label"
                    strPeriod = .TextMatrix(lIndex, GDCol(eGDCol_Period))
                End If
                
                ' Accumulate the symbols for the category...
                'If ValOfText(strSymbolID) = 0 Then
                '    strSymbols = strSymbols & strSymbol & ";" & strPeriod & ","
                'Else
                '    strSymbols = strSymbols & strSymbolID & ";" & strPeriod & ","
                'End If
                strSymbols = strSymbols & SymbolOrSymbolID(lSymbolID, strSymbol) & ";" & strPeriod & ","
                'clear out cell color from alerts being true
                .Cell(flexcpBackColor, lIndex, 1, lIndex, .Cols - 1) = .Cell(flexcpBackColor, lIndex, GDCol(eGDCol_Symbol))
            Next lIndex
            'clear out any bell icons in header row
            .Cell(flexcpPicture, 0, .FixedCols, 0, .Cols - 1) = Nothing
        End With
        
        Set m.alDataIndex = m.QBData.CreateIndex
        m.QBData.SortIndex m.alDataIndex, TblField(eQbTbl_Rows)
        m.hDataIndex = m.alDataIndex.ArrayHandle
    Else
        strSymbols = m.QB.SaveString
        For lIndex = 1 To m.QB.CellCount
            strSymbolID = m.QB.CellItem(lIndex).SymbolID
            lSymbolID = Val(strSymbolID)
            strSymbol = Parse(m.QB.CellItem(lIndex).Symbol, "(", 1)
            strPeriod = "Daily"
            
            If TblSearch(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod, lPos) Then
                TblStr(eQbTbl_Rows, lPos) = TblStr(eQbTbl_Rows, lPos) & "," & Format(lIndex, "000")
            End If
        Next lIndex
        
        Set m.alDataIndex = m.QBData.CreateIndex
        m.QBData.SortIndex m.alDataIndex, TblField(eQbTbl_Rows)
        m.hDataIndex = m.alDataIndex.ArrayHandle
    End If
    
    ' Store the new symbols information for the current category...
    TabStr(eGDTabSettings_Symbols, lTab) = strSymbols

    ' Make sure that the delay column is showing if real-time is active...
    ShowDelayColumn lTab
    
    SymbolCriteria
    
    g.Alerts.DisplayQBAlerts
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ResetRows", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBars
'' Description: Get the bars for the given symbol and period
'' Inputs:      Symbol and Period
'' Returns:     Bars for the symbol/period, Nothing if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBars(ByVal strSymbolOrSymbolID As String, ByVal strPeriod As String) As cGdBars
On Error GoTo ErrSection:
    
    Set GetBars = Nothing
    If m.BarsColl.Exists(strSymbolOrSymbolID & vbTab & strPeriod) Then
        Set GetBars = m.BarsColl(strSymbolOrSymbolID & vbTab & strPeriod)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.GetBars", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateSymbols
'' Description: Update the rows in the grid that are stored in the rows field
'' Inputs:      Row in the data table
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateSymbols(ByVal lRow As Long, Optional ByVal bOnlyIfVisible As Boolean = True, _
    Optional ByVal bRefreshSymbol As Boolean = False, Optional Bars As cGdBars = Nothing)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrRows As New cGdArray        ' Array of rows to update
    
    Dim i&, iRowFound&
    Dim frmQBTab As frmDetachedQBTab
    
    astrRows.SplitFields TblStr(eQbTbl_Rows, lRow), ","
    For lIndex = 0 To astrRows.Size - 1
        If Len(astrRows(lIndex)) > 0 Then
            If TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                UpdateCols CLng(astrRows(lIndex)), , bOnlyIfVisible, bRefreshSymbol

'JM(11-11-2009) - original code, not needed - leave awhile then remove if all okay
'            ElseIf m.frmActiveDetTab Is Nothing Then
'                m.QB.CellItem(CLng(astrRows(lIndex))).UpdateData m.BarsColl(lRow), m.QB.FullBarHeight, m.QB.QuoteBoardStyle, m.QB.CompactQB, True
            End If
        End If
    Next lIndex
        
    'update detached tabs if this is called from Total Refresh and real time is not on
    'as of 08-14-2007 Total Refresh is the function that passes Bars into this function
    If Not Bars Is Nothing And Not g.RealTime.Active Then
        For i = 0 To m.aDetachedTabs.Size - 1
            Set frmQBTab = m.aDetachedTabs(i)
            If Not frmQBTab Is Nothing Then
                iRowFound = frmQBTab.UpdateRT(Bars)
                If iRowFound > 0 Then
                    UpdateCols iRowFound, Bars, , , frmQBTab.fgQuotes
                End If
            End If
        Next
    End If
    
    TblNum(eQbTbl_Dirty, lRow) = 0

ErrExit:
    Set astrRows = Nothing
    Exit Sub
    
ErrSection:
    Set astrRows = Nothing
    RaiseError "frmQuotes.UpdateSymbols", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbolToGrid
'' Description: Add a symbol to the grid
'' Inputs:      Category to add to, Symbol/Period to add, Position to add it
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddSymbolToGrid(ByVal strSymbol As String, ByVal strPeriod As String, Optional ByVal lPos As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID of the symbol passed in
    Dim strSym As String                ' Symbol to use for comparison
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim strSecType As String            ' Security type of the symbol to add
    Dim lPoolRec As Long                ' Record number into the symbol pool
    
    Dim fgCurrGrid As VSFlexGrid
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
    End If

    ' Get SymbolID, Symbol, and Security Type information...
    lSymbolID = GetSymbolID(strSymbol)
    
    ' DAJ 06/11/2015: We are going to start allowing custom indexes on the quote board for the DTG...
    'If lSymbolID < 0 Then
    '    Beep
    '    Exit Function
    'ElseIf lSymbolID > 0 Then
    If lSymbolID <> 0 Then
        lPoolRec = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
        strSym = Str(lSymbolID)
        strSecType = GetSecType(g.SymbolPool.SecType(lPoolRec))
    Else
        lPoolRec = -1&
        strSym = strSymbol
        If InStr(strSymbol, " ") <> 0 Then
            If InStr(strSymbol, "-") <> 0 Then
                strSecType = "FO"
            Else
                strSecType = "SO"
            End If
        End If
    End If
    
    With fgCurrGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Add the new symbol/period to the grid...
        .RemoveItem .Rows - 1
        .Rows = .Rows + 1
        .MergeRow(.Rows - 1) = False
        .TextMatrix(.Rows - 1, GDCol(eGDCol_SecType)) = UCase(strSecType)
        .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolID)) = Str(lSymbolID)
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = strSymbol
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Period)) = strPeriod
        
        ' Move the new row to a given position if needed...
        If lPos >= .FixedRows Then
            .RowPosition(.Rows - 1) = lPos
            .ShowCell lPos, GDCol(eGDCol_Symbol)
        Else
            .ShowCell .Rows - 1, GDCol(eGDCol_Symbol)
        End If
        AddLabelRow "", .Rows
        
        ' Make sure that the alternate coloring is correct...
        ColorQuoteRows
        
        .Redraw = lRedraw
    End With
    
    ' Add the symbol/period to the table if not already there...
    AddSymbolToGrid = AddSymbolToTable(lPoolRec, strPeriod, strSymbol)
    
    ' Update the list of symbols...
    If Len(strSym) > 0 And Len(strPeriod) > 0 Then
        m.astrSymbols.Add strSym & ";" & strPeriod
    End If
    m.astrSymbols.Sort
    
    ' Reset the rows information in the data table...
    ''ResetRows
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.AddSymbolToGrid", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddSymbolToBox
'' Description: Add a symbol to the box style quote board
'' Inputs:      Symbol to add, Row and Column to add it to
'' Returns:     True if added to table, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddSymbolToBox(ByVal strSymbol As String, Optional ByVal lRow As Long = -1&, Optional ByVal lCol As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID of the symbol passed in
    Dim lPoolRec As Long                ' Symbol pool record for the symbol
    Dim strSym As String                ' Symbol to use for comparison
    
    Dim QB As cQuoteCellBoard
        
    ' Get SymbolID, Symbol, and Security Type information...
    lSymbolID = GetSymbolID(strSymbol)
    
    ' DAJ 06/11/2015: We are going to start allowing custom indexes on the quote board for the DTG...
    'If lSymbolID < 0 Then
    '    Beep
    '    Exit Function
    'ElseIf lSymbolID > 0 Then
    If lSymbolID <> 0 Then
        lPoolRec = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
        strSym = Str(lSymbolID)
    Else
        lPoolRec = -1&
        strSym = strSymbol
    End If
    
    If m.frmActiveDetTab Is Nothing Then
        Set QB = m.QB
    Else
        Set QB = m.frmActiveDetTab.QuoteCellBoard
        AddSymbolToGrid strSymbol, "Daily"      'do this for when user switches from box to grid on detached tab
    End If
    
    ' Add the symbol to the box style quote board...
    QB.AddSymbol strSymbol, lRow, lCol

    ' Add the symbol/period to the table if not already there...
    AddSymbolToBox = AddSymbolToTable(lPoolRec, "Daily", strSymbol)
    
    ' Update the list of symbols...
    If m.frmActiveDetTab Is Nothing Then
        'when doing detached tab this is done in the call to add symbol to grid above
        If Len(strSym) > 0 Then m.astrSymbols.Add strSym & ";Daily"
        m.astrSymbols.Sort
        TotalRefresh False
    End If
    
    ' Reset the rows information in the data table...
    ''ResetRows
        
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.AddSymbolToBox", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveSymbolFromGrid
'' Description: Remove a row from the grid (and remove symbol/period if necessary)
'' Inputs:      Row to remove, Period
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemoveSymbolFromGrid(ByVal lRow As Long, Optional ByVal strPeriod As String = "") As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' SymbolID to delete
    Dim strSymbol As String             ' Symbol to delete
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lPos As Long                    ' Position of symbol/period in array
    Dim strLookup As String             ' String to look up in array
    Dim bReturn As Boolean              ' Return value

    Dim fgCurrGrid As VSFlexGrid
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
    End If
    
    With fgCurrGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If .MergeRow(lRow) = False Then
            ' Get information from the grid before removing the row...
            lSymbolID = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_SymbolID))))
            strSymbol = Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)
            If Len(strPeriod) = 0 Then
                strPeriod = .TextMatrix(lRow, GDCol(eGDCol_Period))
            End If
            
            ' Put together the lookup string into the symbols array...
            'If lSymbolID = 0& Then
            '    strLookup = strSymbol & ";" & strPeriod
            'Else
            '    strLookup = Str(lSymbolID) & ";" & strPeriod
            'End If
            strLookup = SymbolOrSymbolID(lSymbolID, strSymbol) & ";" & strPeriod
            
            ' If we can find the lookup string, remove that element...
            If m.astrSymbols.BinarySearch(strLookup, lPos) = True Then
                m.astrSymbols.Remove lPos
            End If
            
            ' If we can't find the lookup now, delete from the data table...
            If m.astrSymbols.BinarySearch(strLookup) = False Then
                bReturn = RemoveSymbolFromTable(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod)
            End If
        End If
        
        ' Remove the row from the grid...
        .RemoveItem lRow
            
        ' Make sure that the alternate coloring is correct...
        ColorQuoteRows
            
        .Redraw = lRedraw
    End With

    ' Reset the rows information in the data table...
    ''ResetRows

    RemoveSymbolFromGrid = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.RemoveSymbolFromGrid", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveSymbolFromBox
'' Description: Remove a cell from the box (and remove symbol/period if necessary)
'' Inputs:      Row and Col of cell to remove
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemoveSymbolFromBox(ByVal lRow As Long, ByVal lCol As Long) As Boolean
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' SymbolID to delete
    Dim lPos As Long                    ' Position of symbol/period in array
    Dim strSymbol As String             ' Symbol to delete
    Dim strPeriod As String             ' Period to delete
    Dim strLookup As String             ' String to look up in array
    Dim bReturn As Boolean              ' Return value
    
    Dim QB As cQuoteCellBoard
    Dim iRow As Long                    'row from detached tab for symbol to be removed

    iRow = -1
    If m.frmActiveDetTab Is Nothing Then
        Set QB = m.QB
    Else
        Set QB = m.frmActiveDetTab.QuoteCellBoard
    End If
    
    If Not QB Is Nothing Then
        ' Get information from the grid before removing the row...
        strSymbol = Parse(QB.Cell(lRow, lCol).Symbol, "(", 1)
        strPeriod = "Daily"
        lSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
                
        ' Put together the lookup string into the symbols array...
        strLookup = SymbolOrSymbolID(lSymbolID, strSymbol) & ";" & strPeriod
        
        ' If we can find the lookup string, remove that element...
        If m.astrSymbols.BinarySearch(strLookup, lPos) = True Then
            m.astrSymbols.Remove lPos
        End If
        
        ' If we can't find the lookup now, delete from the data table...
        If m.astrSymbols.BinarySearch(strLookup) = False Then
            bReturn = RemoveSymbolFromTable(SymbolOrSymbolID(lSymbolID, strSymbol), strPeriod)
        End If
    
        ' Remove the row from the boxstyle board...
        QB.RemoveCell lRow, lCol
        
        If Not m.frmActiveDetTab Is Nothing Then
            'remove symbol from detached tab's grid for when user switches style
            With m.frmActiveDetTab
                iRow = .FindInGrid(strSymbol, strPeriod)
                If iRow = -1 Then iRow = .FindInGrid(strSymbol, "")
                If iRow >= .fgQuotes.FixedRows And iRow < .fgQuotes.Rows Then
                    RemoveSymbolFromGrid iRow
                    .fgQuotes.Row = .fgQuotes.FixedRows
                End If
            End With
        End If
                
        ' Reset the rows information in the data table...
        ''ResetRows
    End If
    
    RemoveSymbolFromBox = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.RemoveSymbolFromBox", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveSymbolFromTable
'' Description: Remove a row from the data table
'' Inputs:      Symbol and Period to remove
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemoveSymbolFromTable(ByVal strSymbolOrSymbolID As String, ByVal strPeriod As String) As Boolean
On Error GoTo ErrSection:
    
    Dim lPos As Long                    ' Position of record in the table
    
    ' Temporary band-aid fix -- if we get a symbol ID passed in, convert it to a
    ' symbol... (DAJ: 11/5/2004)
    'If Not IsAlpha(strSymbol) And Len(strSymbol) > 0 Then
    '    strSymbol = GetSymbol(CLng(strSymbol))
    'End If
    
    g.Alerts.RemoveAlertsForSymbol strSymbolOrSymbolID, strPeriod
    If TblSearch(strSymbolOrSymbolID, strPeriod, lPos) Then
        m.BarsColl.Remove TblStr(eQbTbl_SearchKey, lPos)
        m.QBData.RemoveRecords lPos
        RemoveSymbolFromTable = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.RemoveSymbolFromTable", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveTab
'' Description: Remove a tab from the data table
'' Inputs:      Tab to remove
'' Returns:     True if symbols removed from table, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemoveTab(ByVal lTab As Long) As Boolean
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols from the tab
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position of symbol/period in array
    Dim bReturn As Boolean              ' Return of the function
    Dim strLookup As String             ' String to look up in array
    Dim strSymbol As String             ' Symbol to delete
    Dim strPeriod As String             ' Period to delete
    Dim Style As eGDQuoteStyle          ' Style of the tab

    astrSymbols.Create eGDARRAY_Strings
    astrSymbols.SplitFields TabStr(eGDTabSettings_Symbols, lTab), ","
    Style = CLng(ValOfText(TabStr(eGDTabSettings_Style, lTab)))
    
    For lIndex = 0 To astrSymbols.Size - 1
        If Len(astrSymbols(lIndex)) > 0 Then
            strSymbol = Parse(astrSymbols(lIndex), ";", 1)
            If Style = eGDQuoteStyle_Grid Then
                strPeriod = Parse(astrSymbols(lIndex), ";", 2)
            Else
                strPeriod = "Daily"
            End If
            strLookup = strSymbol & ";" & strPeriod
            
            ' If we can find the lookup string, remove that element...
            If m.astrSymbols.BinarySearch(strLookup, lPos) = True Then
                m.astrSymbols.Remove lPos
            End If
            
            ' If we can't find the lookup now, delete from the data table...
            If m.astrSymbols.BinarySearch(strLookup) = False Then
                bReturn = RemoveSymbolFromTable(strSymbol, strPeriod)
                If bReturn = True Then RemoveTab = True
            End If
        End If
    Next lIndex
    
    Dim frmDetached As frmDetachedQBTab
    
    If TabStr(eGDTabSettings_Form, lTab) <> "0" Then
        lPos = ValOfText(TabStr(eGDTabSettings_Form, lTab))     'window handle saved as long in table
        For lIndex = 0 To m.aDetachedTabs.Size
            Set frmDetached = m.aDetachedTabs(lIndex)
            If Not frmDetached Is Nothing Then
                If frmDetached.Caption = vsTab.TabCaption(frmDetached.MyTabIndex) And frmDetached.hWnd = lPos Then
                    frmDetached.RemoveTab
                    Exit For
                End If
            End If
        Next
    End If
    m.tblTabInfo.RemoveRecords lTab
    
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.RemoveTab", eGDRaiseError_Raise
    
End Function

Private Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    Dim lRow&, lCol&, lColor&, dTickCount#, iSaveRedraw%
    Dim bStillColor As Boolean

    Dim frmDetach As frmDetachedQBTab
    Dim fgCurrGrid As VSFlexGrid
    Dim QB As cQuoteCellBoard
    Dim lIndex&, lEndLoop&, i&
    
    lEndLoop = m.aDetachedTabs.Size
    
    lColor = m.QB.UpdateColor
    If g.nColorTheme = kDarkThemeColor Then
        If IsBlueRange(m.QB.UpdateColor) Then
            lColor = vbCyan
        ElseIf IsGreenRange(m.QB.UpdateColor, True) Then
            lColor = vbGreen
        End If
    End If
    
    For i = 0 To lEndLoop

        If i >= lEndLoop Then
            lIndex = vsTab.CurrTab
            Set fgCurrGrid = fgQuotes
            Set QB = m.QB
        ElseIf Not m.aDetachedTabs(i) Is Nothing Then
            lIndex = m.aDetachedTabs(i).MyTabIndex
            Set fgCurrGrid = m.aDetachedTabs(i).fgQuotes
            Set QB = m.aDetachedTabs(i).QuoteCellBoard
        Else
            Set fgCurrGrid = Nothing
            Set QB = Nothing
            lIndex = -1
        End If

        If lIndex >= 0 Then
            If TabStr(eGDTabSettings_Style, lIndex) = Str(QStyle(eGDQuoteStyle_Grid)) Then
                If Not fgCurrGrid Is Nothing Then
                    With fgCurrGrid
                        iSaveRedraw = .Redraw
                        .Redraw = flexRDNone
                        For lRow = .FixedRows To .Rows - 1
                            If g.bUnloading Then Exit Sub
                            If .Cell(flexcpForeColor, lRow, GDCol(eGDCol_SymbolID)) = lColor Then
                                bStillColor = False
                                If UseUpdatedColors Then
                                    For lCol = GDCol(eGDCol_Symbol) + 1 To .Cols - 1
                                        If Not .ColHidden(lCol) Then
                                            If .Cell(flexcpForeColor, lRow, lCol) = lColor Then
                                                ' see if has been more than 1 second since colored
                                                'format for flexcpData: tickCount|alertkey
                                                dTickCount = ValOfText(Parse(.Cell(flexcpData, lRow, lCol), "|", 1))
                                                dTickCount = gdTickCount - dTickCount
                                                If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                                                    bStillColor = True
                                                Else
                                                    .Cell(flexcpForeColor, lRow, lCol) = m.QB.UnchColor
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                                ' color symbol cell only if a cell was still colored
                                If Not bStillColor Then
                                    .Cell(flexcpForeColor, lRow, GDCol(eGDCol_SymbolID)) = m.QB.UnchColor
                                End If
                            End If
                        Next
                        .Redraw = iSaveRedraw
                    End With
                End If
            ElseIf Not QB Is Nothing Then
                QB.ClearUpdatedColors
            End If
        End If
    
    Next
    ' For right now, clear the updated colors on the TradeTracker summary form as well...
    If FormIsLoaded("frmTTSummary") Then
        frmTTSummary.ClearUpdatedColors
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.ClearUpdatedColors", eGDRaiseError_Raise

End Sub

Public Sub FeedSymbol(ByVal strGenSymbol As String, ByVal strFeedSymbol As String)
On Error GoTo ErrSection:

    Dim lIndex As Long, strSymbol$
    
    For lIndex = 0 To m.QBData.NumRecords - 1
        strSymbol = TblStr(eQbTbl_Symbol, lIndex)
        If g.RealTime.ConvertContinuous Then
            strSymbol = RollSymbolForDate(strSymbol)
        End If
        If strSymbol = strGenSymbol Then
            TblStr(eQbTbl_FeedSymbol, lIndex) = strFeedSymbol
            TblNum(eQbTbl_Dirty, lIndex) = 1
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.FeedSymbol", eGDRaiseError_Raise
    
End Sub

Private Sub SymbolCriteria()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lTab As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strSearch As String             ' Symbol/Period pairing
    Dim astrCols As New cGdArray        ' Array of column information for a tab
    Dim astrCriteria As New cGdArray    ' List of criteria to add for this symbol
    'Dim lSymbolID As Long               ' Symbol ID for the symbol
    
    For lRow = 0 To m.QBData.NumRecords - 1
        'lSymbolID = g.SymbolPool.SymbolIDforSymbol(TblStr(eQbTbl_Symbol, lRow))
        'lSymbolID = TblNum(eQbTbl_SymbolID, lRow)
        'If lSymbolID = 0 Then
        '    strSearch = TblStr(eQbTbl_Symbol, lRow) & ";" & Parse(TblStr(eQbTbl_SearchKey, lRow), vbTab, 2)
        'Else
        '    strSearch = Str(lSymbolID) & ";" & Parse(TblStr(eQbTbl_SearchKey, lRow), vbTab, 2)
        'End If
        strSearch = Replace(TblStr(eQbTbl_SearchKey, lRow), vbTab, ";")
        
        astrCriteria.Size = 0
        For lTab = 0 To m.tblTabInfo.NumRecords - 1
            If InStr("," & TabStr(eGDTabSettings_Symbols, lTab) & ",", "," & strSearch) <> 0 Then
                astrCols.SplitFields TabStr(eGDTabSettings_Fields, lTab), "|"
                For lCol = 0 To astrCols.Size - 1
                    If Len(Parse(astrCols(lCol), ";", 3)) > 0 And Parse(astrCols(lCol), ";", 2) = "0" Then
                        astrCriteria.Add Parse(astrCols(lCol), ";", 3)
                    End If
                Next lCol
            End If
        Next lTab
        TblStr(eQbTbl_Criteria, lRow) = astrCriteria.JoinFields(",")
    Next lRow

ErrExit:
    Set astrCols = Nothing
    Set astrCriteria = Nothing
    Exit Sub
    
ErrSection:
    Set astrCols = Nothing
    Set astrCriteria = Nothing
    RaiseError "frmQuotes.SymbolCriteria", eGDRaiseError_Raise
    
End Sub

Public Function EditFields() As Boolean
On Error GoTo ErrSection:

    Dim astrUsed As New cGdArray        ' Array of fields from the quotes grid
    Dim lRow As Long                    ' Index into a for loop
    Dim bFound As Boolean               ' Was the column found in the grid?
    Dim lPos&, lCol&                    ' Indexes into for loops
    Dim QBF As New cCriteria            ' QBF to store in the quotes grid
    Dim strDefault As String            ' Default string
    Dim astrDefaults As New cGdArray    ' Array of defaults to hand off
    Dim strField As String              ' Field to append to other tab info
    Dim lTab As Long                    ' Index into a for loop
    Dim astrTab As New cGdArray         ' Temporary string array
    Dim astrCols As New cGdArray        ' Temporary string array
    Dim lIndex As Long                  ' Index into a for loop
    Dim bNonDaily As Boolean            ' Are any of the rows not set to Daily?
    
    Dim fgCurrGrid As VSFlexGrid
    
    strDefault = ",Date,Open,High,Low,Last,T,Prev Close,Change,% Change,Last Tick,Exchange,"
    
    ' Tell the user to wait until TotalRefresh is finished...
    If IsBusy Then
        DoEvents
        If IsBusy Then
            InfBox "Please wait until the quote board|has finished refreshing...", "t", "+-OK", "Quote Board", True, 1&
            Exit Function
        End If
    End If
    
    If m.frmActiveDetTab Is Nothing Then
        Set fgCurrGrid = fgQuotes
    Else
        Set fgCurrGrid = m.frmActiveDetTab.fgQuotes
    End If
    
    astrUsed.Create eGDARRAY_Strings
    
    With fgCurrGrid
        bNonDaily = False
        For lRow = .FixedRows To .Rows - 1
            If .MergeRow(lRow) = False Then
                If UCase(.TextMatrix(lRow, GDCol(eGDCol_Period))) <> "DAILY" Then
                    bNonDaily = True
                    Exit For
                End If
            End If
        Next lRow
        If .ColHidden(GDCol(eGDCol_Period)) = True Then
            astrUsed.Add Str(vbUnchecked) & vbTab & "Period" & vbTab & "" & vbTab & CStr(True)
        Else
            astrUsed.Add Str(vbChecked) & vbTab & "Period" & vbTab & "" & vbTab & CStr(True)
        End If
        
        For lRow = GDCol(eGDCol_NumFixed) To .Cols - 1
            If .ColHidden(lRow) = True Then lPos = vbUnchecked Else lPos = vbChecked
            If .TextMatrix(0, lRow) = "Feed Symbol" Or .TextMatrix(0, lRow) = "Delay" Then
                astrUsed.Add Str(lPos) & vbTab & .TextMatrix(0, lRow) & vbTab & .ColData(lRow) & vbTab & CStr(CBool(DirExist(AddSlash(App.Path) & "..\RealTime")))
            ElseIf .TextMatrix(0, lRow) = "T" Then
                astrUsed.Add Str(lPos) & vbTab & "Tick/Settle" & vbTab & .ColData(lRow) & vbTab & CStr(True)
            ElseIf .TextMatrix(0, lRow) <> "#DELETED#" Then
                astrUsed.Add Str(lPos) & vbTab & .TextMatrix(0, lRow) & vbTab & .ColData(lRow) & vbTab & CStr(True)
            End If
            If .ColData(lRow) = "" Then
                If InStr(strDefault, "," & .TextMatrix(0, lRow) & ",") > 0 Then
                    If .TextMatrix(0, lRow) = "T" Then
                        astrDefaults.Add Str(vbChecked) & vbTab & "Tick/Settle" & vbTab & "" & vbTab & CStr(True)
                    Else
                        astrDefaults.Add Str(vbChecked) & vbTab & .TextMatrix(0, lRow) & vbTab & "" & vbTab & CStr(True)
                    End If
                ElseIf .TextMatrix(0, lRow) = "Feed Symbol" Or .TextMatrix(0, lRow) = "Delay" Then
                    astrDefaults.Add Str(vbUnchecked) & vbTab & .TextMatrix(0, lRow) & vbTab & "" & vbTab & CStr(DirExist(AddSlash(App.Path) & "..\RealTime"))
                Else
                    astrDefaults.Add Str(vbUnchecked) & vbTab & .TextMatrix(0, lRow) & vbTab & "" & vbTab & CStr(True)
                End If
            End If
        Next lRow
        
        If frmQuoteBoardFields.ShowMe(astrUsed, eQbfMode_QBFld, astrDefaults) = True Then
            .Redraw = flexRDNone
            m.astrFields.Clear
            For lPos = 0 To astrUsed.Size - 1
                If Parse(astrUsed(lPos), vbTab, 2) = "Period" Then
                    If Parse(astrUsed(lPos), vbTab, 1) <> flexChecked And bNonDaily Then
                        InfBox "All periods must be set to Daily before the column can be hidden", "!", , "Error"
                        'astrUsed(lPos) = Str(vbChecked) & Mid(astrUsed(lPos), 2)
                        .ColHidden(GDCol(eGDCol_Period)) = False
                    End If
                    'astrUsed.MoveItems lPos, 1, -lPos
                    If CLng(Parse(astrUsed(lPos), vbTab, 1)) = flexChecked Then
                        .ColHidden(GDCol(eGDCol_Period)) = False
                    Else
                        .ColHidden(GDCol(eGDCol_Period)) = True
                    End If
                    astrUsed.Remove lPos
                    
                    Exit For
                End If
            Next lPos
            For lPos = 0 To astrUsed.Size - 1
                If Parse(astrUsed(lPos), vbTab, 2) = "Tick/Settle" Then
                    astrUsed(lPos) = Replace(astrUsed(lPos), "Tick/Settle", "T")
                End If
                m.astrFields.Add Parse(astrUsed(lPos), vbTab, 2)
                bFound = False
                For lCol = GDCol(eGDCol_NumFixed) To .Cols - 1
                    If .TextMatrix(0, lCol) = Parse(astrUsed(lPos), vbTab, 2) Then
                        If .ColData(lCol) = Parse(astrUsed(lPos), vbTab, 3) Then
                            bFound = True
                            .ColPosition(lCol) = lPos + GDCol(eGDCol_NumFixed)
                            If CLng(Parse(astrUsed(lPos), vbTab, 1)) = flexChecked Then
                                .ColHidden(lPos + GDCol(eGDCol_NumFixed)) = False
                            Else
                                .ColHidden(lPos + GDCol(eGDCol_NumFixed)) = True
                            End If
                            
                            Set QBF = New cCriteria
                            If Parse(astrUsed(lPos), vbTab, 3) = "" Then
                                .ColData(lPos + GDCol(eGDCol_NumFixed)) = ""
                            Else
                                QBF.FromFile App.Path & "\Custom\", Parse(astrUsed(lPos), vbTab, 3)
                                .ColData(lPos + GDCol(eGDCol_NumFixed)) = QBF.ID
                                m.QBFs.Add QBF, QBF.ID
                            End If
                            Exit For
                        End If
                    End If
                Next lCol
                
                If bFound = False Then
                    .Cols = .Cols + 1
                    If CLng(Parse(astrUsed(lPos), vbTab, 1)) = flexChecked Then
                        .ColHidden(.Cols - 1) = False
                    Else
                        .ColHidden(.Cols - 1) = True
                    End If
                    .TextMatrix(0, .Cols - 1) = Parse(astrUsed(lPos), vbTab, 2)
                    
                    Set QBF = New cCriteria
                    If Parse(astrUsed(lPos), vbTab, 3) = "" Then
                        .ColData(.Cols - 1) = ""
                    Else
                        QBF.FromFile App.Path & "\Custom\", Parse(astrUsed(lPos), vbTab, 3)
                        .ColData(.Cols - 1) = QBF.ID
                        m.QBFs.Add QBF, QBF.ID
                        
                        ' Add a field to the data table for this as well...
                        m.QBData.CreateField eGDARRAY_Doubles, , QBF.ID
                    End If
                    
                    ' Since this is a new field, add it to the other tabs as a hidden
                    ' column...
                    .AutoSize .Cols - 1, , False, 75
                    strField = Parse(astrUsed(lPos), vbTab, 2) & ";-1;" & Parse(astrUsed(lPos), vbTab, 3) & ";" & Str(.ColWidth(.Cols - 1))
                    For lTab = 0 To m.tblTabInfo.NumRecords - 1
                        If lTab <> vsTab.CurrTab Then
                            If Right(TabStr(eGDTabSettings_Fields, lTab), 1) <> "|" Then
                                TabStr(eGDTabSettings_Fields, lTab) = TabStr(eGDTabSettings_Fields, lTab) & "|" & strField & "|"
                            Else
                                TabStr(eGDTabSettings_Fields, lTab) = TabStr(eGDTabSettings_Fields, lTab) & strField & "|"
                            End If
                        End If
                    Next lTab
                    
                    .ColPosition(.Cols - 1) = lPos + GDCol(eGDCol_NumFixed) - 1
                End If
            Next lPos
            
            For lCol = GDCol(eGDCol_NumFixed) To .Cols - 1
                bFound = False
                For lPos = 0 To astrUsed.Size - 1
                    If .TextMatrix(0, lCol) = Parse(astrUsed(lPos), vbTab, 2) Then
                        bFound = True
                        Exit For
                    End If
                Next lPos
                
                If bFound = False Then
                    .ColHidden(lCol) = True
                    
                    bFound = False
                    For lRow = 0 To m.tblTabInfo.NumRecords - 1
                        If lRow <> vsTab.CurrTab Then
                            If InStr(TabStr(eGDTabSettings_Fields, lRow), .TextMatrix(0, lCol) & ";0") <> 0 Then
                                bFound = True
                                Exit For
                            End If
                        End If
                    Next lRow
                                        
                    If bFound = False Then
                        g.Alerts.RemoveAlertsForField .TextMatrix(0, lCol)
                        
                        For lRow = 0 To m.tblTabInfo.NumRecords - 1
                            If lRow <> vsTab.CurrTab Then
                                astrCols.SplitFields TabStr(eGDTabSettings_Fields, lRow), "|"
                                For lIndex = 0 To astrCols.Size - 1
                                    If Parse(astrCols(lIndex), ";", 1) = .TextMatrix(0, lCol) Then
                                        astrCols.Remove lIndex
                                        Exit For
                                    End If
                                Next lIndex
                                TabStr(eGDTabSettings_Fields, lRow) = astrCols.JoinFields("|")
                            End If
                        Next lRow
                        
                        If Len(.ColData(lCol)) > 0 Then
                            m.QBData.ClearField m.QBData.FieldNum(.ColData(lCol))
                            m.QBFs.Remove .ColData(lCol)
                        End If
                        
                        Set QBF = New cCriteria
                        .TextMatrix(0, lCol) = "#DELETED#"
                        .ColData(lCol) = ""
                    End If
                End If
            Next lCol
    
            ColorQuoteRows
            AlignColumns
            .Redraw = flexRDBuffered
            
            SaveTabInfo
            SymbolCriteria
            TotalRefresh True
            
            EditFields = True
        End If
    End With
    
ErrExit:
    Set astrUsed = Nothing
    Exit Function
    
ErrSection:
    Set astrUsed = Nothing
    RaiseError "frmQuotes.EditFields", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbolOnGrid
'' Description: Allow the user to change the symbol of the given row
'' Inputs:      Row, String to send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbolOnGrid(ByVal Row As Long, Optional ByVal strToSend As String = "")
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) returned from symbol selector
    Dim strSymbol As String             ' Current symbol from the grid
    Dim strPeriod As String             ' Current period from the grid
    Dim bRemoved As Boolean             ' Was the old one removed from the table?
    Dim bAdded As Boolean               ' Was the new one added to the table?
    Dim lRow As Long                    ' Current row in the grid
    Dim bLastRow As Boolean

    If vsTab.TabCaption(vsTab.CurrTab) = "(Filter)" Then Exit Sub

    If Row = fgQuotes.Rows - 1 Then bLastRow = True

    astrSymbols.Create eGDARRAY_Strings
    ' Ask the user for a new symbol...
    If Len(strToSend) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe(Parse(fgQuotes.TextMatrix(Row, GDCol(eGDCol_Symbol)), "(", 1), False, , , , , True)
    Else
        Set astrSymbols = frmSymbolSelector.ShowMe(strToSend, False, , , , False, True)
    End If
    
    ' If the user chose a symbol...
    If astrSymbols.Size > 0 Then
        ' and it was different from the previous symbol...
        If astrSymbols(0) <> m.strSaveSymbol Then
            ' if adding a symbol (instead of replacing), check # symbols
            'If Len(m.strSaveSymbol) = 0 And m.QBData.NumRecords + astrSymbols.Size > MaxSymbolsAllowed Then
            If m.QBData.NumRecords + astrSymbols.Size > MaxSymbolsAllowed Then
                If HasGold(False) Then
                    Err.Raise vbObjectError + 1000, , "There cannot be more than " & MaxSymbolsAllowed & " symbols on the quote board"
                Else
                    Err.Raise vbObjectError + 1000, , "You need to upgrade to Gold or Platinum in order to add more symbols on the quote board"
                End If
            End If
            
            lRow = Row
            strSymbol = fgQuotes.TextMatrix(Row, GDCol(eGDCol_Symbol))
            strPeriod = fgQuotes.TextMatrix(Row, GDCol(eGDCol_Period))
            If fgQuotes.MergeRow(lRow) = True Then strPeriod = "Daily"
            
            If Row < fgQuotes.Rows - 1 Then
                bRemoved = RemoveSymbolFromGrid(Row)
            End If
            bAdded = AddSymbolToGrid(astrSymbols(0), strPeriod, lRow)
            
            If (bRemoved Or bAdded) And g.RealTime.Active Then
                'g.RealTime.UpdateSymbolList
            End If
            
            If Row = fgQuotes.Rows - 1 Then
                AddLabelRow "", fgQuotes.Rows
            End If
                            
            ResetRows
            TotalRefresh False
            m.strSaveSymbol = ""
            
            ColorQuoteRows
            
            ' if started on the last row, set to new blank last row
            If bLastRow Then
                fgQuotes.Row = fgQuotes.Rows - 1
                fgQuotes.ShowCell fgQuotes.Row, 0
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ChangeSymbolOnGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbolOnBox
'' Description: Allow the user to change the symbol of the given cell
'' Inputs:      Row, Column, String to send
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbolOnBox(ByVal lRow As Long, ByVal lCol As Long, Optional ByVal strToSend As String = "")
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) returned from symbol selector
    Dim strSymbol As String             ' Current symbol from the grid
    Dim strPeriod As String             ' Current period from the grid
    Dim bRemoved As Boolean             ' Was the old one removed from the table?
    Dim bAdded As Boolean               ' Was the new one added to the table?

    If vsTab.TabCaption(vsTab.CurrTab) <> "(Filter)" Then
        astrSymbols.Create eGDARRAY_Strings
        
        ' Ask the user for a new symbol...
        If Len(strToSend) = 0 Then
            Set astrSymbols = frmSymbolSelector.ShowMe(Parse(m.QB.Cell(lRow, lCol).Symbol, "(", 1), False, , , , , True)
        Else
            Set astrSymbols = frmSymbolSelector.ShowMe(strToSend, False, , , , False, True)
        End If
        
        ' If the user chose a symbol...
        If astrSymbols.Size > 0 Then
            ' and it was different from the previous symbol...
            If astrSymbols(0) <> m.strSaveSymbol Then
                If m.QBData.NumRecords + astrSymbols.Size > MaxSymbolsAllowed Then
                    If HasGold(False) Then
                        Err.Raise vbObjectError + 1000, , "There cannot be more than " & MaxSymbolsAllowed & " symbols on the quote board"
                    Else
                        Err.Raise vbObjectError + 1000, , "You need to upgrade to Gold or Platinum in order to add more symbols on the quote board"
                    End If
                End If
                
                bRemoved = RemoveSymbolFromBox(lRow, lCol)
                bAdded = AddSymbolToBox(astrSymbols(0), lRow, lCol)
                
                ResetRows
                TotalRefresh False
                m.strSaveSymbol = ""
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ChangeSymbolOnBox"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumSymbolsForTab
'' Description: Calculate the number of symbols on a given tab
'' Inputs:      Tab to calculate for
'' Returns:     Number of symbols on that tab
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumSymbolsForTab(ByVal lTab As Long) As Long
On Error Resume Next

    Dim astrTemp As New cGdArray        ' Temporary array
    Dim lCount As Long                  ' Number of symbols
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbols As String            ' Symbols for the tab
    
    strSymbols = TabStr(eGDTabSettings_Symbols, lTab)
    If Len(strSymbols) > 2 Then
        astrTemp.SplitFields Mid(strSymbols, 2, Len(strSymbols) - 2), ","
        For lIndex = 0 To astrTemp.Size - 1
            If Len(astrTemp(lIndex)) > 0 And UCase(Parse(astrTemp(lIndex), ";", 1)) <> "LABEL" Then
                lCount = lCount + 1
            End If
        Next lIndex
    End If
    
    NumSymbolsForTab = lCount
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshQBF
'' Description: If the Criteria with the given ID is on the quote board,
''              refresh it
'' Inputs:      ID of the Criteria to refresh
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshQBF(ByVal strID As String)
On Error GoTo ErrSection:

    If m.QBFs.Exists(strID) Then
        m.QBFs(strID).FromFile AddSlash(App.Path) & "Custom\", strID
        TotalRefresh True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.RefreshQBF", eGDRaiseError_Raise
    
End Sub

' to move focus to the quote board
Private Sub MoveFocusToQb()

    On Error Resume Next
    
    If Not m.frmActiveDetTab Is Nothing Then
        If m.frmActiveDetTab.MyQBStyle = eGDQuoteStyle_Grid Then
            MoveFocus m.frmActiveDetTab.fgQuotes
        Else
            MoveFocus m.frmActiveDetTab.QuoteCellBoard
        End If
    ElseIf TabStr(eGDTabSettings_Style, vsTab.CurrTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        MoveFocus fgQuotes
    Else
        MoveFocus pbQuoteBoard
    End If

End Sub

Public Property Get UseUpdatedColors() As Boolean
    
    UseUpdatedColors = tmrRealtime.Enabled And m.QB.UseUpdateColor
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolOrSymbolID
'' Description: Pass back the Symbol ID if non-zero or else the symbol
'' Inputs:      Symbol ID, Symbol
'' Returns:     Symbol ID if non-zero, otherwise Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolOrSymbolID(ByVal lSymbolID As Long, ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    If lSymbolID = 0& Then
        SymbolOrSymbolID = strSymbol
    Else
        SymbolOrSymbolID = Str(lSymbolID)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.SymbolOrSymbolID", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentTabStyle
'' Description: Determine the style of the current quote board tab
'' Inputs:      None
'' Returns:     Style of the Current Tab
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentTabStyle(Optional ByVal bUsePrivateMember As Boolean = False) As eGDQuoteStyle
On Error GoTo ErrSection:

    Dim strStyle As String              ' Style of the current tab
    
    If bUsePrivateMember Then
        strStyle = TabStr(eGDTabSettings_Style, m.lCurrentTab)
    Else
        strStyle = TabStr(eGDTabSettings_Style, vsTab.CurrTab)
    End If
    CurrentTabStyle = Val(strStyle)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.CurrentTabStyle", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFilterTab
'' Description: Load up the filter tab from the appropriate filter id
'' Inputs:      None
'' Returns:     Symbol list change?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadFilterTab() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrSymbols As New cGdArray     ' Array of symbols on the filter tab
    Dim strFilterID As String           ' ID of the filter to apply
    Dim lFieldNum As Long               ' Field number of the filter in the pool
    Dim lSymbolID As Long               ' Symbol ID for the current symbol
    Dim lPoolRec As Long                ' Pool record for the symbol ID
    Dim lPos As Long                    ' Position of the symbol in array
    Dim lMaxSymbols As Long             ' Maximum number of symbols that can be loaded
    Dim bReturn As Boolean              ' Return value from the function
    Dim lIndex2 As Long                 ' Index into a for loop
    
    bReturn = False
    astrSymbols.SplitFields m.tblTabInfo(TabField(eGDTabSettings_Symbols), m.tblTabInfo.NumRecords - 1), ","
    strFilterID = m.tblTabInfo(TabField(eGDTabSettings_FilterID), m.tblTabInfo.NumRecords - 1)
    If Len(strFilterID) > 0 Then
        lFieldNum = g.SymbolPool.FieldNumForID(strFilterID)
        If lFieldNum >= 0 Then
            astrSymbols.Sort
            lMaxSymbols = MaxSymbolsAllowed - m.QBData.NumRecords

            ' If current symbols aren't in the filter anymore, delete them...
            For lIndex = astrSymbols.Size - 1 To 0 Step -1
                lSymbolID = Val(Parse(astrSymbols(lIndex), ";", 1))
                lPoolRec = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
                If g.SymbolPool.ArrayTable(lFieldNum, lPoolRec) <> 1 Then
                    astrSymbols.Remove lIndex
                    If m.astrSymbols.BinarySearch(Str(lSymbolID) & ";" & m.strFilterPeriod, lPos) = True Then
                        m.astrSymbols.Remove lPos
                    End If
                    If m.astrSymbols.BinarySearch(Str(lSymbolID) & ";" & m.strFilterPeriod) = False Then
                        lMaxSymbols = lMaxSymbols + 1&
                        If RemoveSymbolFromTable(Str(lSymbolID), m.strFilterPeriod) = True Then bReturn = True
                    End If
                End If
            Next lIndex
            
            BuildFilterVols lFieldNum
            
            If g.RealTime.IsServerActive Then frmMain.SuspendNewSymbolCheck 15
        
            ' Walk through the symbol pool and add any new symbols that need to be added...
            For lIndex2 = 0 To g.SymbolPool.NumRecords - 1
                lIndex = m.alFilterIdx(lIndex2)
                If g.SymbolPool.ArrayTable(lFieldNum, lIndex) = 1 Then
                    If astrSymbols.BinarySearch(Str(g.SymbolPool.SymbolID(lIndex)) & ";", lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                        astrSymbols.Add Str(g.SymbolPool.SymbolID(lIndex)) & ";" & m.strFilterPeriod, lPos
                        m.astrSymbols.Add Str(g.SymbolPool.SymbolID(lIndex)) & ";" & m.strFilterPeriod
                        If TblSearch(Str(g.SymbolPool.SymbolID(lIndex)), m.strFilterPeriod) = False Then
                            If AddSymbolToTable(lIndex, m.strFilterPeriod) = True Then bReturn = True
                            
                            lMaxSymbols = lMaxSymbols - 1&
                        End If
                        
                        If lMaxSymbols <= 0 Then
                            Exit For
                        End If
                    End If
                End If
            Next lIndex2
            
            m.astrSymbols.Sort
            
            If g.RealTime.IsServerActive Then frmMain.SuspendNewSymbolCheck
        Else
            astrSymbols.Size = 0
        End If
    Else
        astrSymbols.Size = 0
    End If
    
    m.tblTabInfo(TabField(eGDTabSettings_Symbols), m.tblTabInfo.NumRecords - 1) = astrSymbols.JoinFields(",")
    LoadFilterTab = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.LoadFilterTab"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFilter
'' Description: Update the filter if it is the one assigned to the filter tab
'' Inputs:      Filter ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateFilter(ByVal strFilterID As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRefresh As Boolean             ' Do we need to do a TotalRefresh?
    
    If Not m.tblTabInfo Is Nothing Then
        For lIndex = 0 To m.tblTabInfo.NumRecords - 1
            If TabStr(eGDTabSettings_FilterID, lIndex) = strFilterID Then
                bRefresh = LoadFilterTab
                If vsTab.TabCaption(vsTab.CurrTab) = "(Filter)" Then
                    ShowCategory vsTab.CurrTab
                End If
                If bRefresh Then
                    TotalRefresh False
                End If
                SortFilterTab
                Exit For
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.UpdateFilter"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildFilterVols
'' Description: Build the filter volume array
'' Inputs:      Field Number of filter in pool
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildFilterVols(ByVal lFieldNum As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As New cGdBars             ' Bars object
    
    Screen.MousePointer = vbHourglass
    m.alFilterVols.Size = g.SymbolPool.NumRecords
    
    For lIndex = 0 To g.SymbolPool.NumRecords - 1
        If g.SymbolPool.ArrayTable(lFieldNum, lIndex) = 1 Then
            If DM_GetBars(Bars, g.SymbolPool.SymbolID(lIndex), "Daily", LastDailyDownload - 5, , , , , False) Then
                m.alFilterVols(lIndex) = Bars(eBARS_Vol, Bars.Size - 1)
            Else
                m.alFilterVols(lIndex) = m.alFilterVols.NullValue
            End If
        Else
            m.alFilterVols(lIndex) = m.alFilterVols.NullValue
        End If
    Next lIndex
    
    Set m.alFilterIdx = m.alFilterVols.CreateSortedIndex(eGdSort_Descending)

ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "frmQuotes.BuildFilterVols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortFilterTab
'' Description: Sort the filter tab based on the last sort column and order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortFilterTab(Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim strFilterID As String           ' ID of the current filter

    With fgQuotes
        If lTab = -1& Then lTab = vsTab.CurrTab
        If vsTab.TabCaption(lTab) = "(Filter)" Then
            strFilterID = TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1)
            If .MergeRow(.FixedRows) = True Then .RemoveItem .FixedRows
            
            If Len(strFilterID) > 0 Then
                If .MergeRow(.Rows - 1) = True Then .RemoveItem .Rows - 1
                
                If m.lFilterSortCol = -1& Then
                    .Col = GDCol(eGDCol_Symbol)
                    .Sort = flexSortGenericAscending
                Else
                    .Col = m.lFilterSortCol
                    .Sort = m.lFilterSortDir
                End If
                
                AddLabelRow "Current Filter = " & g.SymbolPool.ArrayTable.FieldName(g.SymbolPool.FieldNumForID(strFilterID)), .FixedRows
                AddLabelRow "", .Rows
                
            Else
                AddLabelRow "No Filter Selected", .FixedRows
                AddLabelRow "", .Rows
            End If
            
            ColorQuoteRows
            ResetRows lTab
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.SortFilterTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayAlert
'' Description: Display an icon in the appropriate field for an alert
'' Inputs:      Alert
''              nPreviouslyTabAlert
''              - -1: don't care, don't check
''              -  0: was previously a single QB cell alert
''              -  1: was previously a QB tab alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayAlert(ByVal Alert As cAlert, Optional ByVal bRemove As Boolean = False, _
    Optional ByVal bSetRedrawMode As Boolean = True, Optional ByVal nPreviouslyTabAlert As Long = -1, _
    Optional ByVal bInit As Boolean = False, Optional ByVal bAdd As Boolean = False)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Appropriate row in the table
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Appropriate column in the grid
    Dim astrRows As New cGdArray        ' Array of rows for the symbol/period combo
    Dim strFlexData$                    'format: tickCount|alertKey
    Dim i&
    
    If (Alert.AlertType = eGDAlertType_QuoteBoard) And (TabStr(eGDTabSettings_Style, m.lCurrentTab) = Str(eGDQuoteStyle_Grid)) Then
        If Alert.IsSymbol = True Then
            ' Find the column in the grid with the given field name...
            lCol = -1&
            For lIndex = 0 To fgQuotes.Cols - 1
                If fgQuotes.TextMatrix(0, lIndex) = Alert.field Then
                    lCol = lIndex
                    Exit For
                End If
            Next lIndex
        
            ' Find the row in the table with the given symbol...
            If TblSearch(SymbolOrSymbolID(Alert.SymbolID, Alert.Symbol), Alert.Period, lRow) And (lCol > -1&) Then
                astrRows.SplitFields TblStr(eQbTbl_Rows, lRow), ","
                With fgQuotes
                    If bSetRedrawMode Then .Redraw = flexRDNone
                    For lIndex = 0 To astrRows.Size - 1
                        lRow = CLng(ValOfText(astrRows(lIndex)))
                        If lRow >= .FixedRows And lRow < .Rows Then
                            strFlexData = Parse(.Cell(flexcpData, lRow, lCol), "|", 1)
                            If bRemove = True Then
                                .Cell(flexcpPicture, lRow, lCol) = Nothing
                                .Cell(flexcpBackColor, lRow, lCol) = .Cell(flexcpBackColor, lRow, 0)        'clear out highlight color
                                If Len(strFlexData) > 0 Then .Cell(flexcpData, lRow, lCol) = strFlexData
                            Else
                                If Alert.Active = True Then
                                    .Cell(flexcpPicture, lRow, lCol) = Picture16(ToolbarIcon(kActiveAlertIcon))
                                Else
                                    .Cell(flexcpPicture, lRow, lCol) = Picture16(ToolbarIcon(kInactiveAlertIcon))
                                End If
                                .Cell(flexcpPictureAlignment, lRow, lCol) = flexAlignLeftTop
                                If Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 1) <> 0 Then
                                    If Alert.LastCheckedFLag = True Then
                                        If Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 2) > 0 Then
                                            .Cell(flexcpBackColor, lRow, lCol) = ValOfText(Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 2))
                                        End If
                                    End If
                                End If
                                strFlexData = strFlexData & "|" & Alert.AlertKey
                                .Cell(flexcpData, lRow, lCol) = strFlexData
                            End If
                        End If
                    Next lIndex
                    ''.AutoSize lCol, , False, 75
                    If bSetRedrawMode Then .Redraw = flexRDDirect
                End With
            End If
        ElseIf TabStr(eGDTabSettings_Name, m.lCurrentTab) = Alert.TabName Then
            lCol = -1&
            For lIndex = 0 To fgQuotes.Cols - 1
                If fgQuotes.TextMatrix(0, lIndex) = Alert.field Then
                    lCol = lIndex
                    Exit For
                End If
            Next lIndex
            
            If lCol > -1& Then
                If bRemove = True Then
                    fgQuotes.Cell(flexcpPicture, 0, lCol) = Nothing
                    If Alert.HasColorAction Then
                        For i = fgQuotes.FixedRows To fgQuotes.Rows - 1
                            fgQuotes.Cell(flexcpBackColor, i, lCol) = fgQuotes.Cell(flexcpBackColor, i, 0)
                        Next
                    End If
                ElseIf Alert.Active = True Then
                    fgQuotes.Cell(flexcpPicture, 0, lCol) = Picture16(ToolbarIcon(kActiveAlertIcon))
                Else
                    fgQuotes.Cell(flexcpPicture, 0, lCol) = Picture16(ToolbarIcon(kInactiveAlertIcon))
                End If
                fgQuotes.Cell(flexcpPictureAlignment, 0, lCol) = flexAlignLeftTop
                fgQuotes.AutoSize lCol, , False, 75
                fgQuotes.Cell(flexcpData, 0, lCol) = "|" & Alert.AlertKey
            End If
        End If
    End If
    
    Dim frm As frmDetachedQBTab
    
    If Not bInit Then
        If bAdd And TabStr(eGDTabSettings_Style, m.lCurrentTab) <> Str(eGDQuoteStyle_Grid) Then
            m.QB.BoxQbAlertUpdate Alert, False, True
        End If
        For i = 0 To m.aDetachedTabs.Size - 1
            Set frm = m.aDetachedTabs(i)
            If Not frm Is Nothing Then
                frm.DisplayAlert Alert, bRemove, bAdd
            End If
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.DisplayAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearFilterTab
'' Description: Clear the filter tab
'' Inputs:      None
'' Returns:     True if need to do a Total Refresh, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClearFilterTab() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bReturn As Boolean              ' Return value from the function
    Dim astrSymbols As New cGdArray     ' Array of symbols on the filter tab
    Dim lSymbolID As Long               ' Symbol ID for the current symbol
    Dim lPos As Long                    ' Position of the symbol in array
    
    If m.aDetachedTabs.Size = m.tblTabInfo.NumRecords - 1 Then
        InfBox "One of the detached tab must be reattached before clearing the filter tab.", "I"
    ElseIf InfBox("Are you sure you want to remove the filter?", "?", "+Yes|-No", "Confirmation") = "Y" Then
        bReturn = False
        astrSymbols.SplitFields m.tblTabInfo(TabField(eGDTabSettings_Symbols), m.tblTabInfo.NumRecords - 1), ","
        
        ' If current symbols aren't in the filter anymore, delete them...
        For lIndex = astrSymbols.Size - 1 To 0 Step -1
            lSymbolID = Val(Parse(astrSymbols(lIndex), ";", 1))
            astrSymbols.Remove lIndex
            If m.astrSymbols.BinarySearch(Str(lSymbolID) & ";" & m.strFilterPeriod, lPos) = True Then
                m.astrSymbols.Remove lPos
            End If
            If m.astrSymbols.BinarySearch(Str(lSymbolID) & ";" & m.strFilterPeriod) = False Then
                If RemoveSymbolFromTable(Str(lSymbolID), m.strFilterPeriod) = True Then bReturn = True
            End If
        Next lIndex
        
        TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1) = ""
        TabStr(eGDTabSettings_Symbols, m.tblTabInfo.NumRecords - 1) = ""
        ClearFilterTab = bReturn
        
        If bReturn Then
            ShowCategory vsTab.CurrTab
            TotalRefresh False
            SortFilterTab
        End If
        
        'set focus to some other tab
        For lIndex = 0 To vsTab.NumTabs - 2
            If vsTab.TabVisible(lIndex) Then
                m.lCurrentTab = lIndex
                vsTab.CurrTab = lIndex
                Exit For
            End If
        Next
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.ClearFilterTab"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeFilter
'' Description: Allow the user to change the filter on the filter tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeFilter()
On Error GoTo ErrSection:

    Dim lFilterIndex As Long
    Dim strOldFilterID As String        ' Current Filter ID of the selected filter
    Dim strNewFilterID As String        ' New Filter ID of the selected filter
    Dim bRefresh As Boolean             ' Do we need to do a refresh?
    
    Dim iSaveTab As Long
    
    m.nMouseDownRow = kNullData
    
    If HasGold(True, , False) = True Then
        strOldFilterID = TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1)
        strNewFilterID = frmSelectFilter.ShowMe(strOldFilterID)
        If (Len(strNewFilterID) > 0) And (strNewFilterID <> strOldFilterID) Then
            If strNewFilterID = "<clear>" Then
                bRefresh = ClearFilterTab
            Else
                m.tblTabInfo(TabField(eGDTabSettings_FilterID), m.tblTabInfo.NumRecords - 1) = strNewFilterID
                bRefresh = LoadFilterTab
                lFilterIndex = m.tblTabInfo.NumRecords - 1
                If vsTab.CurrTab <> lFilterIndex Then vsTab.CurrTab = lFilterIndex
                ShowCategory vsTab.CurrTab
                If bRefresh Then TotalRefresh False
                SortFilterTab
            End If
        End If
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ChangeFilter"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NumberOfFilterSymbols
'' Description: Determine the current number of symbols on the filter tab
'' Inputs:      None
'' Returns:     Number of Symbols on the Filter tab
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumberOfFilterSymbols() As Long
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Array of symbols on the filter tab
    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumSymbols As Long             ' Number of symbols
    Dim lPos As Long                    ' Position in the array
    
    astrSymbols.SplitFields TabStr(eGDTabSettings_Symbols, m.tblTabInfo.NumRecords - 1), ","
    For lIndex = 0 To astrSymbols.Size - 1
        If m.astrSymbols.BinarySearch(astrSymbols(lIndex), lPos) = True Then
            If m.astrSymbols(lPos + 1) <> astrSymbols(lIndex) Then
                lNumSymbols = lNumSymbols + 1
            End If
        End If
    Next lIndex
    
    NumberOfFilterSymbols = lNumSymbols
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.NumberOfFilterSymbols"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BarsExist
'' Description: Determine if bars exist on the quote board for the given symbol
'' Inputs:      Symbol or Symbol ID, Period
'' Returns:     True if Bars exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BarsExist(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String) As Boolean
On Error GoTo ErrSection:

    BarsExist = m.BarsColl.Exists(Str(vSymbolOrSymbolID) & vbTab & strPeriod)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.BarsExist"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateTab
'' Description: Create a tab by the given name if it does not already exist,
''              set the style to the given style, make sure that the given
''              symbols/symbol ID's are on the tab, then set that tab current
'' Inputs:      Name, Style, Symbols, Fields
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateTab(ByVal strTabName As String, Optional ByVal nTabStyle As eGDQuoteStyle = eGDQuoteStyle_Grid, Optional ByVal strSymbols As String = "", Optional ByVal strFields As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Index into the tab info table
    Dim astrSymbols As New cGdArray     ' Array of symbols to add
    Dim lSymbolID As Long               ' Symbol ID for the symbol
    Dim strSymbol As String             ' Symbol to add
    Dim strTabSymbols As String         ' Tab symbols
    Dim bRefresh As Boolean             ' Do we need to do a TotalRefresh?
    
    lPos = -1&
    For lIndex = 0 To m.tblTabInfo.NumRecords - 1
        If UCase(TabStr(eGDTabSettings_Name, lIndex)) = UCase(strTabName) Then
            lPos = lIndex
            Exit For
        End If
    Next lIndex
    
    If lPos = -1& Then
        m.tblTabInfo.AddRecord "", m.tblTabInfo.NumRecords - 1
        lPos = m.tblTabInfo.NumRecords - 2
        
        TabStr(eGDTabSettings_Name, lPos) = strTabName
    End If

    TabNum(eGDTabSettings_Style, lPos) = nTabStyle
    
    If Len(TabStr(eGDTabSettings_Fields, lPos)) = 0 Then
        If Len(strFields) = 0 Then
            TabStr(eGDTabSettings_Fields, lPos) = m.strDefaultFields
        Else
            TabStr(eGDTabSettings_Fields, lPos) = strFields
        End If
    End If

    InitQuoteTabs lPos
    ShowCategory lPos

    bRefresh = False
    astrSymbols.SplitFields strSymbols, ","
    For lIndex = 0 To astrSymbols.Size - 1
        If Len(astrSymbols(lIndex)) > 0 Then
            lSymbolID = Val(astrSymbols(lIndex))
            If lSymbolID > 0 Then
                strSymbol = GetSymbol(lSymbolID)
            Else
                strSymbol = astrSymbols(lIndex)
            End If
            
            If InStr(TabStr(eGDTabSettings_Symbols, lPos), "," & strSymbol & ";") = 0 Then
                If nTabStyle = eGDQuoteStyle_Grid Then
                    If AddSymbolToGrid(strSymbol, "Daily") Then bRefresh = True
                Else
                    If AddSymbolToBox(strSymbol) Then bRefresh = True
                End If
            End If
        End If
    Next lIndex
    
    If bRefresh Then
        ResetRows lPos
        TotalRefresh False
    End If

    ShowCategory lPos

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.CreateTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowDelayColumn
'' Description: Show/Hide the delay column as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowDelayColumn(Optional ByVal lTab As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Did we find a delay?
    
    If lTab = -1& Then lTab = vsTab.CurrTab

    If TabStr(eGDTabSettings_Style, lTab) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        bFound = False
        For lIndex = 0 To m.QBData.NumRecords - 1
            If Len(TblStr(eQbTbl_Rows, lIndex)) > 0 Then
                If Len(TblStr(eQbTbl_Delay, lIndex)) > 0 Then
                    bFound = True
                    Exit For
                End If
            End If
        Next lIndex
        fgQuotes.ColHidden(GDCol(eGDCol_Delay)) = Not bFound
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ShowDelayColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPeriodAndDelay
'' Description: Make sure that the period and delay columns are in the right
''              place and that there is only one of each of them.
'' Inputs:      Column Info
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixPeriodAndDelay(astrColumnInfo As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bPeriodFound As Boolean         ' Has the period been found in the array?
    Dim bDelayFound As Boolean          ' Has the delay been found in the array?
    Dim strPeriod As String             ' Period column information
    Dim strDelay As String              ' Delay column information
    
    bPeriodFound = False
    bDelayFound = False
    strPeriod = ""
    strDelay = ""
    
    For lIndex = astrColumnInfo.Size - 1 To 0 Step -1
        If Parse(astrColumnInfo(lIndex), ";", 1) = "Period" Then
            If (lIndex <> GDCol(eGDCol_Period)) Then
                If bPeriodFound = True Then
                    astrColumnInfo.Remove lIndex
                Else
                    If Parse(astrColumnInfo(GDCol(eGDCol_Period)), ";", 1) <> "Period" Then
                        strPeriod = astrColumnInfo(lIndex)
                    End If
                    astrColumnInfo.Remove lIndex
                End If
            End If
            bPeriodFound = True
        ElseIf Parse(astrColumnInfo(lIndex), ";", 1) = "Delay" Then
            If (lIndex <> GDCol(eGDCol_Delay)) Then
                If bDelayFound = True Then
                    astrColumnInfo.Remove lIndex
                Else
                    If Parse(astrColumnInfo(GDCol(eGDCol_Delay)), ";", 1) <> "Delay" Then
                        strDelay = astrColumnInfo(lIndex)
                    End If
                    astrColumnInfo.Remove lIndex
                End If
            End If
            bDelayFound = True
        End If
    Next lIndex
        
    If Parse(astrColumnInfo(GDCol(eGDCol_Period)), ";", 1) <> "Period" Then
        If Len(strPeriod) > 0 Then
            astrColumnInfo.Add strPeriod, GDCol(eGDCol_Period)
        Else
            astrColumnInfo.Add "Period;0", GDCol(eGDCol_Period)
        End If
    End If
        
    If Parse(astrColumnInfo(GDCol(eGDCol_Delay)), ";", 1) <> "Delay" Then
        If Len(strDelay) > 0 Then
            astrColumnInfo.Add strDelay, GDCol(eGDCol_Delay)
        Else
            astrColumnInfo.Add "Delay;0", GDCol(eGDCol_Delay)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.FixPeriodAndDelay"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidLabel
'' Description: Make sure that the label that the user entered is valid
'' Inputs:      Label
'' Returns:     True if Valid, False otherwise
''
''JM 01-17-2008:
''made this public so frmDetachQBTab can use it as well for custom bar type (4365)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidLabel(ByVal strLabel As String) As Boolean
On Error GoTo ErrSection:

    If InStr(strLabel, ",") <> 0 Then
        InfBox "Commas (,) are not allowed in a label in the quote board", "!", , "Quote Board Label Error"
    ElseIf InStr(strLabel, vbTab) <> 0 Then
        InfBox "Tabs are not allowed in a label in the quote board", "!", , "Quote Board Label Error"
    ElseIf InStr(strLabel, ";") <> 0 Then
        InfBox "Semicolons (;) are not allowed in a label in the quote board", "!", , "Quote Board Label Error"
    ElseIf InStr(strLabel, "|") <> 0 Then
        InfBox "Pipes are not allowed in a label in the quote board", "!", , "Quote Board Label Error"
    Else
        ValidLabel = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.ValidLabel"

End Function

Private Sub DetachTab(ByVal iIndex&, Optional ByVal bClearSymbols As Boolean = False)
On Error GoTo ErrSection:

    Dim iNew&, i&
    Dim frm As frmDetachedQBTab
    
    iNew = -1
    
    If iIndex >= 0 And iIndex < vsTab.NumTabs Then
        vsTab.TabVisible(iIndex) = False

        If iIndex = m.lCurrentTab Then
            For i = iIndex + 1 To vsTab.NumTabs - 1
                'try to set current tab to the next tab
                If vsTab.TabVisible(i) Then
'                    If vsTab.TabCaption(i) <> "(new)" And vsTab.TabCaption(i) <> "(Filter)" Then
                    If vsTab.TabCaption(i) <> "(manage)" And vsTab.TabCaption(i) <> "(Filter)" Then
                        iNew = i
                        Exit For
                    End If
                End If
            Next
            If iNew = -1 Then
                'try to set current tab to previous tab
                For i = iIndex To 0 Step -1
                    If vsTab.TabVisible(i) Then
                        iNew = i
                        Exit For
                    End If
                Next
            End If
            If iNew = -1 And HasFilterTab Then
                iNew = m.tblTabInfo.NumRecords - 1
            End If
        Else
            iNew = m.lCurrentTab
            ShowCategory iIndex                 '5641, 5642
        End If
        
        If iNew >= 0 And iNew < vsTab.NumTabs Then      '4711
            Set frm = New frmDetachedQBTab
            
            InitGrid frm.fgQuotes
            
            m.tblTabInfo(TabField(eGDTabSettings_Form), iIndex) = frm.hWnd
            m.aDetachedTabs.Add frm
            
            AlignColumns frm.fgQuotes
            
            frm.ShowMe fgQuotes, m.QB, iIndex
            frm.fgQuotes.Redraw = flexRDBuffered
            
            If bClearSymbols Then
                frm.RemoveAllSymbols
                TabStr(eGDTabSettings_Symbols, frm.MyTabIndex) = frm.MySymbols(True)
            End If
                    
'JM 01-06-2010: this line of code is causing aardvark 5086
'   cannot duplicate issue 4339 even with this line of code commented out
'   will revisit 4339 and put in alternate fix if issue reoccurs
'            If frm.MyQBStyle = eGDQuoteStyle_Grid Then PopulateGrid frm.fgQuotes, TabStr(eGDTabSettings_Symbols, iIndex)    '4339
            
            vsTab.CurrTab = iNew
        Else
            InfBox "You cannot detach all quote tabs.", "I"
        End If
                
    End If

ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    Set frm = Nothing
    RaiseError "frmQuotes.DetachTab"

End Sub

Public Sub ReAttachTab(frm As frmDetachedQBTab)
On Error GoTo ErrSection:

    Dim iIndex&, i&
    
    If Not frm Is Nothing Then
        iIndex = frm.MyTabIndex
        If iIndex >= 0 And iIndex < m.tblTabInfo.NumRecords Then
            If frm.hWnd = m.tblTabInfo(TabField(eGDTabSettings_Form), iIndex) Then
                For i = m.aDetachedTabs.Size - 1 To 0 Step -1
                    If Not m.aDetachedTabs(i) Is Nothing Then
                        If frm.hWnd = m.aDetachedTabs(i).hWnd Then
                            m.aDetachedTabs.Remove i
                            m.tblTabInfo(TabField(eGDTabSettings_Form), iIndex) = 0
                            If iIndex >= 0 And iIndex < vsTab.NumTabs Then
                                vsTab.TabVisible(iIndex) = True
                                If Not g.RealTime.Active Then TotalRefresh False
                            End If
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If

ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ReAttachTab"
    Resume ErrExit

End Sub

Public Sub TabFuncWrappers(frmDetTab As frmDetachedQBTab, eFunction As eGDTabFuncWrapper, _
    Optional ByVal strChar As String)
On Error GoTo ErrSection:

    Dim strSymbol$, strNewPeriod$, strOldPeriod$
        
    Set m.frmActiveDetTab = frmDetTab
    If m.frmActiveDetTab Is Nothing Then Exit Sub
    
    Select Case eFunction
        Case eGDTabFuncWrapper_AddSymbol
            If frmDetTab.MyQBStyle = eGDQuoteStyle_Grid Then
                DoAddSymbol , , strChar
            ElseIf Not frmDetTab.QuoteCellBoard Is Nothing Then
                DoAddSymbol frmDetTab.QuoteCellBoard.Row, frmDetTab.QuoteCellBoard.Col, strChar
            End If
            TabStr(eGDTabSettings_Symbols, frmDetTab.MyTabIndex) = frmDetTab.MySymbols(True)
        Case eGDTabFuncWrapper_AddLabel
            With frmDetTab.fgQuotes
                If .Row = .Rows - 1 Then
                    .Redraw = flexRDNone
                    .Rows = .Rows + 1
                    .Cell(flexcpText, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = ""
                    .Cell(flexcpAlignment, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = flexAlignCenterTop
                    .Cell(flexcpForeColor, .Rows - 1, .FrozenCols, .Rows - 1, .Cols - 1) = m.QB.UnchColor
                    .RowData(.Rows - 1) = "Label"
                    .MergeRow(.Rows - 1) = True
                    .Redraw = flexRDBuffered
                End If
            End With
            TabStr(eGDTabSettings_Symbols, m.frmActiveDetTab.MyTabIndex) = m.frmActiveDetTab.MySymbols(True)
        Case eGDTabFuncWrapper_ChangePeriod
            With frmDetTab.fgQuotes
                strSymbol = Parse(.TextMatrix(.Row, GDCol(eGDCol_Symbol)), "(", 1)
                strNewPeriod = .TextMatrix(.Row, GDCol(eGDCol_Period))
                strOldPeriod = strChar
                .TextMatrix(.Row, GDCol(eGDCol_Period)) = strOldPeriod
                RemoveSymbolFromGrid .Row
                AddSymbolToGrid strSymbol, strNewPeriod, .Row
                TabStr(eGDTabSettings_Symbols, frmDetTab.MyTabIndex) = frmDetTab.MySymbols(True)
                TotalRefresh False      'this is to get data for newly added symbol
                If .TextMatrix(.Row, GDCol(eGDCol_Symbol)) = strSymbol And .TextMatrix(.Row, GDCol(eGDCol_Period)) = strNewPeriod Then
                    UpdateCols .Row, , , , frmDetTab.fgQuotes
                ElseIf .Row + 1 < .Rows Then
                    UpdateCols .Row + 1, , , , frmDetTab.fgQuotes       'periodicity of symbol at bottom of grid was changed
                End If
            End With
        Case eGDTabFuncWrapper_RemoveSymbol
            DoRemoveSymbol
            TabStr(eGDTabSettings_Symbols, frmDetTab.MyTabIndex) = frmDetTab.MySymbols(True)
        Case eGDTabFuncWrapper_SaveTab
            SaveTabInfo frmDetTab.MyTabIndex
    End Select
    
ErrExit:
    Set m.frmActiveDetTab = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.TabFuncWrappers"
    Resume ErrExit

End Sub

Public Function GetBarsTree() As cGdTree
    Set GetBarsTree = m.BarsColl
End Function

Public Property Get SymbolCol() As Long
    SymbolCol = GDCol(eGDCol_Symbol)
End Property

Public Property Get PeriodCol() As Long
    PeriodCol = GDCol(eGDCol_Period)
End Property

Public Property Get CanMoveCol(ByVal nCol&) As Boolean
    If nCol >= GDCol(eGDCol_NumFixed) Then CanMoveCol = True
End Property

Public Property Get HasFilterTab() As Boolean
On Error GoTo ErrSection:

'JM: 01-05-2009 Original code. Leave awhile then remove if all okay
'If HasGold(True, , False) = True Then      '4710
    
    If HasGold(False, , False) Then
        If Len(TabStr(eGDTabSettings_FilterID, m.tblTabInfo.NumRecords - 1)) > 0 Then
            HasFilterTab = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.HasFilterTab.Get"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    QbTabToString
'' Description: Convert the quote board tab information to a string to dump
''              to a file
'' Inputs:      Tab to Convert
'' Returns:     Converted String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function QbTabToString(ByVal lTab As Long) As String
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' File to dump
    Dim strInfo As String               ' Information about the current quote board tab
    Dim astrInfo As cGdArray            ' Array of quote board tab information
    Dim astrCols As cGdArray            ' Array of columns for this tab
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' Criteria ID
    Dim astrCriteria As cGdArray        ' Array of criteria
    Dim astrAlerts As cGdArray          ' Array of alerts for this quote board tab
    Dim astrCustom As cGdArray          ' Array of custom criteria that the user needs to change
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Set astrFile = New cGdArray
    Set astrInfo = New cGdArray
    Set astrCols = New cGdArray
    Set astrCriteria = New cGdArray
    Set astrAlerts = New cGdArray
    Set astrCustom = New cGdArray
    
    strInfo = m.tblTabInfo.GetRecord(lTab, vbTab)
    astrInfo.SplitFields strInfo, vbTab
    
    If astrInfo(1) = Str(QStyle(eGDQuoteStyle_Grid)) Then
        astrCols.SplitFields astrInfo(3), "|"
        
        For lIndex = astrCols.Size - 1 To 0 Step -1
            strID = Parse(astrCols(lIndex), ";", 3)
            If Len(strID) > 0 Then
                If Parse(astrCols(lIndex), ";", 2) = "-1" Then
                    astrCols.Remove lIndex
                ElseIf FileExist(AddSlash(App.Path) & "Custom\" & strID) Then
                    astrCriteria.Add "File=" & "Custom\" & strID & vbTab & Replace(FileToString(AddSlash(App.Path) & "Custom\" & strID), vbCrLf, vbTab)
                    If (UCase(Left(strID, 3)) = "CUS") And (IsNumeric(Mid(strID, 4, 5))) Then
                        astrCustom.Add strID
                    End If
                End If
            End If
        Next lIndex
        
        astrInfo(3) = astrCols.JoinFields("|")
    End If
    
    If astrCustom.Size > 0 Then
        InfBox "The following criteria have default filenames that need to be renamed before you can export:||" & astrCustom.JoinFields("|") & "|", "!", , "Quote Board Tab Export Error"
        strReturn = ""
    Else
        For lIndex = 1 To g.Alerts.Count
            With g.Alerts.Item(lIndex)
                If (.AlertType = eGDAlertType_QuoteBoard) And (.TabName = astrInfo(0)) Then
                    astrAlerts.Add .ToFileString
                End If
            End With
        Next lIndex
        
        astrFile.Add Str(CurrentTime("GMT"))
        astrFile.Add "[Version]"
        astrFile.Add "1"
        astrFile.Add "[Tab Info]"
        astrFile.Add astrInfo.JoinFields(vbTab)
        astrFile.Add "[Criteria]"
        For lIndex = 0 To astrCriteria.Size - 1
            astrFile.Add astrCriteria(lIndex)
        Next lIndex
        astrFile.Add "[Alerts]"
        For lIndex = 0 To astrAlerts.Size - 1
            astrFile.Add astrAlerts(lIndex)
        Next lIndex
        
        strReturn = astrFile.JoinFields(vbCrLf)
    End If
    
    QbTabToString = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.QbTabToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    QbTabFromString
'' Description: Create a quote board tab from the given string
'' Inputs:      Quote Board Tab Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub QbTabFromString(ByVal strTabString As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Information broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim lVersion As Long                ' File format version
    Dim strTabInfo As String            ' Tab information
    Dim astrCriteria As cGdArray        ' Array of criteria information
    Dim astrAlerts As cGdArray          ' Array of alert information
    Dim bFound As Boolean               ' Is the tab already in the table information?
    Dim bContinue As Boolean            ' Do we want to continue?
    Dim lPos As Long                    ' Position in a string
    Dim strFileName As String           ' Filename
    Dim strCriteria As String           ' Criteria
    Dim bReloadPool As Boolean          ' Do we need to reload the symbol pool?
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strName As String               ' Name of the criteria
    Dim Alert As cAlert                 ' Alert to add to the collection
    Dim lExistingTab As Long            ' Existing quote board tab
    
    ' Initialize the arrays...
    Set astrFile = New cGdArray
    Set astrCriteria = New cGdArray
    Set astrAlerts = New cGdArray
    
    ' Read in the information from the string...
    astrFile.SplitFields strTabString, vbCrLf
    For lIndex = 0 To astrFile.Size - 1
        Select Case UCase(astrFile(lIndex))
            Case "[VERSION]"
                lIndex = lIndex + 1
                lVersion = CLng(Val(astrFile(lIndex)))
                
            Case "[TAB INFO]"
                lIndex = lIndex + 1
                strTabInfo = astrFile(lIndex)
                
            Case "[CRITERIA]"
                Do While (Left(astrFile(lIndex + 1), 1) <> "[") And (lIndex + 1 < astrFile.Size)
                    lIndex = lIndex + 1
                    astrCriteria.Add astrFile(lIndex)
                Loop
                
            Case "[ALERTS]"
                Do While (Left(astrFile(lIndex + 1), 1) <> "[") And (lIndex + 1 < astrFile.Size)
                    lIndex = lIndex + 1
                    astrAlerts.Add astrFile(lIndex)
                Loop
                
        End Select
    Next lIndex
    
    ' Verify the version of the file...
    If lVersion < 1 Then
        InfBox "The version of the file that you are trying to import is too old", "!", , "Quote Board Tab Import Error"
        bContinue = False
    ElseIf lVersion > 1 Then
        InfBox "The version of the file that you are trying to import is too new", "!", , "Quote Board Tab Import Error"
        bContinue = False
    Else
        bContinue = True
    End If
    
    ' See if the quote board tab already exists (there is a tab by the same name)...
    If bContinue Then
        bFound = False
        lExistingTab = -1&
        
        For lIndex = 0 To m.tblTabInfo.NumRecords - 1
            If TabStr(eGDTabSettings_Name, lIndex) = Parse(strTabInfo, vbTab, 1) Then
                lExistingTab = lIndex
                bFound = True
                Exit For
            End If
        Next lIndex
        
        bContinue = True
        If bFound Then
            If InfBox("There is already a quote board tab named '" & Parse(strTabInfo, vbTab, 1) & "'.  Would you like to overwrite it?", "?", "+Yes|-No", "Quote Board Tab Import") = "Y" Then
                bContinue = True
            Else
                bContinue = False
            End If
        End If
    End If
    
    ' If the quote board tab doesn't exist, make sure none of the criteria exist...
    If (bFound = False) And (bContinue = True) Then
        For lIndex = 0 To astrCriteria.Size - 1
            lPos = InStr(astrCriteria(lIndex), vbTab)
            strFileName = Parse(Left(astrCriteria(lIndex), lPos - 1), "=", 2)
            strCriteria = Mid(astrCriteria(lIndex), lPos + 1)
            lPos = InStr(UCase(astrCriteria(lIndex)), "NAME=")
            If lPos > 0 Then
                lPos = lPos + 5
                strName = Mid(astrCriteria(lIndex), lPos)
                strName = Parse(strName, vbTab, 1)
            Else
                strName = ""
            End If
            
            If FileExist(AddSlash(App.Path) & strFileName) = True Then
                bFound = True
            Else
                For lIndex2 = 1 To g.SymbolPool.Criterias.Count
                    If g.SymbolPool.Criterias(lIndex2).Name = strName Then
                        bFound = True
                        Exit For
                    End If
                Next lIndex2
            End If
            
            If bFound = True Then
                Exit For
            End If
        Next lIndex
    
        If bFound Then
            If InfBox("One or more of the criteria for this quote board tab already exist.  Would you like to overwrite them?", "?", "+Yes|-No", "Quote Board Tab Import") = "Y" Then
                bContinue = True
            Else
                bContinue = False
            End If
        End If
    End If
    
    ' Import the criteria...
    If bContinue Then
        bReloadPool = False
        For lIndex = 0 To astrCriteria.Size - 1
            lPos = InStr(astrCriteria(lIndex), vbTab)
            strFileName = Parse(Left(astrCriteria(lIndex), lPos - 1), "=", 2)
            strCriteria = Mid(astrCriteria(lIndex), lPos + 1)
            
            FileFromString AddSlash(App.Path) & strFileName, Replace(strCriteria, vbTab, vbCrLf)
            bReloadPool = True
        Next lIndex
        
        If bReloadPool Then
            g.SymbolPool.Load
        End If
    End If
    
    ' Import the quote board tab...
    If bContinue Then
        If lExistingTab = -1& Then
            m.tblTabInfo.AddRecord strTabInfo, , vbTab
        Else
            m.tblTabInfo.SetRecord strTabInfo, lExistingTab, vbTab
        End If
    End If
    
    ' Import the alerts...
    If bContinue Then
        For lIndex = 0 To astrAlerts.Size - 1
            If g.Alerts.AlertExists(astrAlerts(lIndex)) = False Then
                Set Alert = New cAlert
                Alert.FromFileString astrAlerts(lIndex)
                g.Alerts.Add Alert
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuotes.QbTabFromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportCurrentTab
'' Description: Export the current tab to a file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportCurrentTab()
On Error GoTo ErrSection:

    Dim strQbTab As String              ' Quote board tab converted to string
    Dim strName As String               ' Name of the quote board tab
    Dim lTab As Long                    ' Current tab
    
    lTab = CurrentTab
    strName = vsTab.TabCaption(lTab)
    
    InfBox "Exporting '" & strName & "'. Please Wait...", , , "Quote Board Tab Export", True
    strQbTab = QbTabToString(lTab)
    If Len(strQbTab) > 0 Then
        FileFromString AddSlash(App.Path) & "QBT\" & strName & ".QBT", strQbTab
        FileCopy AddSlash(App.Path) & "QBT\" & strName & ".QBT", AddSlash(App.Path) & "QBT\" & strName & ".DON", True
        InfBox strName & " has been exported to|" & AddSlash(App.Path) & "QBT\" & strName & ".QBT", "i", "+-OK", "Quote Board Tab Export"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ExportCurrentTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentTab
'' Description: Determine the current tab whether or not it is detached
'' Inputs:      None
'' Returns:     Current Tab
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CurrentTab() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    If m.frmActiveDetTab Is Nothing Then
        lReturn = vsTab.CurrTab
    Else
        lReturn = m.frmActiveDetTab.MyTabIndex
    End If
    
    CurrentTab = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.CurrentTab"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameCriteriaFile
'' Description: Change any references to the old filename to the new filename
'' Inputs:      Old Filename (without path), New Filename (without path)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RenameCriteriaFile(ByVal strOldFileBase As String, ByVal strNewFileBase As String)
On Error GoTo ErrSection:

    Dim lTab As Long                    ' Index into a for loop
    Dim lField As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As cGdArray          ' Array of field information
    Dim QBF As cCriteria                ' Quote Board Field object
    
    ' Update the Quote Board Field collection if necessary...
    If m.QBFs.Exists(strOldFileBase) Then
        Set QBF = New cCriteria
        QBF.FromFile AddSlash(App.Path) & "Custom\", strNewFileBase
        m.QBFs.Add QBF, QBF.ID
        m.QBFs.Remove strOldFileBase
    End If
    
    ' Update the QB Data table...
    For lCol = 0 To m.QBData.NumFields - 1
        If UCase(m.QBData.FieldName(lCol)) = UCase(strOldFileBase) Then
            m.QBData.FieldName(lCol) = strNewFileBase
        End If
    Next lCol
    
    ' Update the current grid if necessary...
    With fgQuotes
        For lCol = 0 To .Cols - 1
            If UCase(.ColData(lCol)) = UCase(strOldFileBase) Then
                .ColData(lCol) = strNewFileBase
            End If
        Next lCol
    End With
    
    ' Update each of the quote board tabs if necessary...
    Set astrFields = New cGdArray
    For lTab = 0 To m.tblTabInfo.NumRecords - 1
        astrFields.SplitFields m.tblTabInfo(TabField(eGDTabSettings_Fields), lTab), "|"
        For lField = 0 To astrFields.Size - 1
            If UCase(Parse(astrFields(lField), ";", 3)) = UCase(strOldFileBase) Then
                astrFields(lField) = Parse(astrFields(lField), ";", 1) & ";" & Parse(astrFields(lField), ";", 2) & ";" & strNewFileBase
            End If
        Next lField
        
        m.tblTabInfo(TabField(eGDTabSettings_Fields), lTab) = astrFields.JoinFields("|")
    Next lTab
    
    ' Reload the current quote board tab...
    ShowCategory CurrentTab
    
    ' Save the tab info for the current tab which will also reset the default display string...
    SaveTabInfo CurrentTab
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.RenameCriteriaFile"
    
End Sub

' return array of symbols currently visible (on the active QB tab, current chart page, etc)
Public Function GetVisibleSymbols() As cGdArray
On Error GoTo ErrSection:

    Dim i&, strSymbol$
    Dim aSymbols As New cGdArray
    Dim frm As Form

    aSymbols.Create eGDARRAY_Strings
    If CurrentTabStyle = eGDQuoteStyle_Grid Then
        ' get symbols in the visible rows
        For i = fgQuotes.TopRow To fgQuotes.BottomRow
            If i >= fgQuotes.FixedRows And i < fgQuotes.Rows Then
                strSymbol = Trim(fgQuotes.TextMatrix(i, GDCol(eGDCol_Symbol)))
                If Len(strSymbol) > 0 Then
                    aSymbols.Add RollSymbolForDate(strSymbol)
                End If
            End If
        Next
    Else
        ' get symbols in the visible box-style cells
        Set aSymbols = m.QB.CellSymbols(True, True)
    End If
    
    ' get symbols from the current chart page
    For i = 0 To Forms.Count - 1
        If IsFrmChart(Forms(i)) Then
            Set frm = Forms(i)
            If Not IsIntraday(frm.Chart.Periodicity) Then
                strSymbol = GetSymbol(frm.SymbolID)
                If Len(strSymbol) > 0 Then
                    aSymbols.Add RollSymbolForDate(strSymbol)
                End If
            End If
        End If
    Next
    Set frm = Nothing
    
    ' request EOD for these symbols now (to get a head-start on them)
    aSymbols.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues

ErrExit:
    Set GetVisibleSymbols = aSymbols
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.GetVisibleSymbols"
End Function

Private Sub CheckStreamStart()
On Error GoTo ErrSection:

    Dim i&, bEnable As Boolean

    If Not g.RealTime.Active Then
        m.dWhenStreamingStarted = 0
        EnableButtons
    ElseIf m.dWhenStreamingStarted = 0 Then
        m.dWhenStreamingStarted = gdTickCount
        EnableButtons
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.CheckStreamStart"
End Sub

Private Sub PromptShowButtons()
On Error GoTo ErrSection:

    Dim strPrompt$

    If Not m.bShowButtons Then
        strPrompt = GetIniFileProperty("ShowButtons", "", "DontAsk", g.strIniFile)
        If Len(strPrompt) = 0 Then
            strPrompt = "You can right click the quoteboard to turn show buttons on again."
            strPrompt = InfBox(strPrompt, "I", "Ok", "Quoteboard", , , , , , , , , True)
            If InStr(strPrompt, "-") > 0 Then
                ' don't show message anymore
                Call SetIniFileProperty("ShowButtons", Left(strPrompt, 1), "DontAsk", g.strIniFile)
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.PromptShowButtons"
End Sub

Public Sub InitBoxQbAlerts()
On Error GoTo ErrSection:

    If Not m.QB Is Nothing Then m.QB.BoxQbAlertInit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.InitBoxQbAlerts"

End Sub

Public Property Get CurrentTabNum(Optional ByVal bUsePrivateMember As Boolean = False)
On Error GoTo ErrSection:

    If bUsePrivateMember Then
        CurrentTabNum = m.lCurrentTab
    Else
        CurrentTabNum = vsTab.CurrTab
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmQuotes.CurrentTabNum.Get"

End Property

Public Property Get CurrentTabName(Optional ByVal nTabNum As Long = -1)
On Error GoTo ErrSection:

    If nTabNum <> -1 Then
        CurrentTabName = TabStr(eGDTabSettings_Name, nTabNum)
    Else
        CurrentTabName = TabStr(eGDTabSettings_Name, vsTab.CurrTab)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmQuotes.CurrentTabName.Get"

End Property

Public Sub WebPageCheck()
    
    On Error Resume Next
    Static nPrevSecond&, bInProgress As Boolean
   
    If bInProgress Then Exit Sub
    bInProgress = True
   
    ' only check once a second
    If Second(Now) <> nPrevSecond Then
        nPrevSecond = Second(Now)
        QuotesTableWebPage
        AcctStatusCheck
        AcctStatusCheck2
    End If
    
    bInProgress = False

End Sub

' JUST FOR INTERNAL PURPOSES: to create a simple web page with a few updated quotes
Private Sub QuotesWebPage_OLD()
On Error Resume Next

    Dim i&, d#, nSymbolID&, dClose#, dPrev#, dDate#, dTime#
    Dim strFile$, strSymbol$, strText$, strTitle$, strAcct$
    Dim aStrings As New cGdArray
    Dim aHTML As New cGdArray
    Dim aAcctInfo As New cGdArray
    Dim Bars As cGdBars
    Dim Acct As cPtAccount
    Dim Positions As cAccountPositions
    Dim Position As cAccountPosition
    Dim Orders As cPtOrders
    Dim Order As cPtOrder
    Static dNextTime As Double
    
    If dNextTime < 0 Then Exit Sub
    If gdTickCount < dNextTime Then Exit Sub
    
    ' only if flag file exists
    strFile = App.Path & "\QuotesWebPage.flg"
    If dNextTime = 0 Then
        If Not FileExist(strFile) Then
            dNextTime = -1
            Exit Sub
        End If
    End If
       
    ' read contents of flag file
    aStrings.FromFile strFile
    strFile = aStrings(0) ' web page file
    ' calc next refresh time
    i = Val(aStrings(1)) ' # seconds to refresh
    If i <= 0 Then i = 60 ' default
    dNextTime = gdTickCount + i * 1000#
    strAcct = Trim(aStrings(3))
    ' get list of symbols into an array
    aStrings.SplitFields aStrings(2), vbTab ' symbols
    If Len(strFile) = 0 Or aStrings.Size = 0 Then Exit Sub
    
    ' HTML header (auto-refresh every "i" seconds)
    aHTML.Size = 0
    aHTML.Add "<HTML><HEAD><TITLE>Quotes</TITLE><META HTTP-EQUIV=""REFRESH"" CONTENT=""" & Str(i) & """></HEAD>"
    aHTML.Add "<BODY TEXT=BLACK LINK=BLUE VLINK=PURPLE ALINK=RED><FONT face=""Arial, sans-serif"">"
    aHTML.Add "<A href=""."">TradeNavigator.com</A><BR>"
    aHTML.Add Format(Date, "mmm d") & " <B>" & Format(Now, "h:mm:ssa/p") & "</B><BR>"
       
    ' which symbol goes in the title depends on time of day
    i = Hour(Now) * 100 + Minute(Now)
    If i < 730 Or (i >= 1430 And Weekday(Date) <> vbFriday) Then
        strTitle = "YM-067" ' "$EUR-USD"
    Else
        strTitle = "$DJIA"
    End If
    
    ' acct info
    aAcctInfo.Size = 0
    If Len(strAcct) > 0 Then
        Set Acct = g.Broker.Account(strAcct)
        If Not Acct Is Nothing Then
            ' ACCT info
            If Acct.ConnectionStatus = eGDConnectionStatus_Connected Then
                strText = Str(Round(Acct.OpenProfit))
                If Left(strText, 1) = "-" Then
                    strText = " <Font Color=RED> - " & Mid(strText, 2)
                Else
                    strText = " <Font Color=GREEN> + " & strText
                End If
                strText = strAcct & " = $" & Str(Round(Acct.CurrentClosedBalance)) & strText
            Else
                strText = strAcct & " is Disconnected!"
            End If
            aHTML.Add "<B>" & strText & "</Font></B><BR>"
            aHTML.Add "- - - - - - - - - - - - - - - - - -<BR>"
            
            ' Open POSITIONS
            Set Positions = g.Broker.FillSummariesForAccount(strAcct)
            If Not Positions Is Nothing Then
                For i = 1 To Positions.Count
                    Set Position = Positions(i)
                    If Not Position Is Nothing Then
                        ' .AutoTradeItem is -1 for total (0 for manual)
                        If Position.CurrentPositionSnapshot <> 0 And Position.AutoTradeItemID < 0 Then
                            ' e.g. GC3-201203: Short 1 = -xxxx
                            strSymbol = Position.Symbol
                            If Position.CurrentPositionSnapshot > 0 Then
                                strText = "Long " & Str(Position.CurrentPositionSnapshot)
                            Else
                                strText = "Short " & Str(Abs(Position.CurrentPositionSnapshot))
                            End If
                            If Position.OpenProfit <= 0 Then
                                strText = " = <Font Color=RED> " & Str(Round(Position.OpenProfit))
                            Else
                                strText = " = <Font Color=GREEN> +" & Str(Round(Position.OpenProfit))
                            End If
                            strText = "<B>" & strSymbol & ": " & strText & "</Font></B><BR>"
                            aAcctInfo.Add strSymbol & vbTab & strText
                        End If
                    End If
                Next
            End If
            
            ' Open ORDERS
            Set Orders = g.Broker.OrdersForAccount(strAcct, True)
            If Not Orders Is Nothing Then
                For i = 1 To Orders.Count
                    Set Order = Orders(i)
                    If Not Order Is Nothing Then
                        If IsOpenOrder(Order.Status, True) Then
                            ' e.g. GC3-201202 Working:
                            '      ... Sell 1 at 1647.2 STOP
                            strSymbol = Order.Symbol
                            If Order.Buy Then
                                strText = "<Font Color=GREEN>"
                            Else
                                strText = "<Font Color=RED>"
                            End If
                            strText = strSymbol & " " & OrderStatus(Order.Status) & ":<BR>... " _
                                & strText & Order.OrderText(False) & "</Font><BR>"
                            aAcctInfo.Add strSymbol & " " & vbTab & strText
                        End If
                    End If
                Next
            End If
            
            ' Sort by symbol, and append current price
            If aAcctInfo.Size > 0 Then
                aAcctInfo.Sort
                For i = 0 To aAcctInfo.Size - 1
                    aHTML.Add Parse(aAcctInfo(i), vbTab, 2)
                    strSymbol = Parse(aAcctInfo(i), vbTab, 1)
                    If strSymbol <> Parse(aAcctInfo(i + 1), vbTab, 1) And g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                        Set Bars = frmTTSummary.GetBars(strSymbol)
                        If Not Bars Is Nothing Then
                            dClose = Bars(eBARS_Close, Bars.Size - 1)
                            If dClose <> kNullData Then
                                ' for this context, show up/down from the Open
                                dPrev = Bars(eBARS_Open, Bars.Size - 1)
                                If dClose >= dPrev Then
                                    strText = " (<Font Color=GREEN>+"
                                Else
                                    strText = " (<Font Color=RED>"
                                End If
                                strText = Bars.PriceDisplay(dClose) & strText _
                                    & Bars.PriceDisplay(dClose - dPrev) & "</Font>)"
                                ' time of last trade
                                dDate = Bars(eBARS_DateTime, Bars.Size - 1)
                                If dDate > 0 Then
                                    dTime = Bars.Prop(eBARS_LastTickTime)
                                    If dTime <= 0 Then
                                        ' session date
                                        strText = strText & " at " & Format(Int(dDate), "mm/dd")
                                    Else
                                        ' time
                                        dTime = Int(dDate) + dTime / 1440#
                                        If g.bShowInLocalTimeZone Then
                                            dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                                        End If
                                        strText = strText & " at " & Format(dTime, "h:mma/p")
                                    End If
                                End If
                                aHTML.Add strText & "<BR>"
                            End If
                        End If
                        aHTML.Add "- - - - - - - - - - - - - - - - - -<BR>"
                    End If
                Next
            End If
            Set Acct = Nothing
            Set Position = Nothing
            Set Positions = Nothing
            Set Order = Nothing
            Set Orders = Nothing
            Set Bars = Nothing
            aAcctInfo.Size = 0
        End If
    End If
    
    
    ' for each symbol
    For i = 0 To aStrings.Size - 1
        strText = ""
        strSymbol = Trim(aStrings(i))
        If Len(strSymbol) > 0 Then
            nSymbolID = GetSymbolID(strSymbol)
            Set Bars = GetBars(SymbolOrSymbolID(nSymbolID, strSymbol), "Daily")
            If Not Bars Is Nothing Then
                dClose = Bars(eBARS_Close, Bars.Size - 1)
                If dClose <> kNullData Then
                    ' get Delta = Close - prior Settle
                    dPrev = GetPrevCloseForQB(Bars)
                    If dPrev = kNullData Then
                        dPrev = dClose
                    End If
                    If dClose < dPrev Then
                        strText = " <Font Color=RED>" & Bars.PriceDisplay(dClose - dPrev)
                    ElseIf dClose > dPrev Then
                        strText = " <Font Color=GREEN>+" & Bars.PriceDisplay(dClose - dPrev)
                    Else
                        strText = " <Font Color=BLACK>+" & Bars.PriceDisplay(dClose - dPrev)
                    End If
                    strText = "<B>" & strSymbol & strText & "</Font></B><BR>" & Bars.PriceDisplay(dClose)
                        
                    ' time of last trade
                    dDate = Bars(eBARS_DateTime, Bars.Size - 1)
                    If dDate > 0 Then
                        dTime = Bars.Prop(eBARS_LastTickTime)
                        If dTime <= 0 Then
                            ' session date
                            strText = strText & " at " & Format(Int(dDate), "mm/dd")
                        Else
                            ' time
                            dTime = Int(dDate) + dTime / 1440#
                            If g.bShowInLocalTimeZone Then
                                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                            End If
                            strText = strText & " at " & Format(dTime, "h:mma/p")
                            
                            ' web page title
                            If strSymbol = strTitle Then
                                If dClose >= dPrev Then
                                    strTitle = "+" & Bars.PriceDisplay(dClose - dPrev)
                                Else
                                    strTitle = Bars.PriceDisplay(dClose - dPrev)
                                End If
                                strTitle = strTitle & " " & Format(dTime, "h:mma/p") & " " & strSymbol
                                aHTML(0) = Replace(aHTML(0), ">Quotes<", ">" & strTitle & "<")
                                strTitle = ""
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Len(strText) > 0 Then
            aHTML.Add strText & "<BR>"
        End If
        Set Bars = Nothing
    Next
    
    ' HTML footer
    aHTML.Add "</BODY></HTML>"
    aHTML.ToFile strFile

End Sub

' JUST FOR INTERNAL PURPOSES: to create a simple web page with a few updated quotes
Private Sub QuotesWebPage()
On Error Resume Next

    Dim i&, d#, nSymbolID&, dClose#, dPrev#, dDate#, dTime#
    Dim strFile$, strSymbol$, strText$, strTitle$
    Dim aStrings As New cGdArray
    Dim aHTML As New cGdArray
    Dim Bars As cGdBars
    Static dNextTime As Double
    
Exit Sub
    If dNextTime < 0 Then Exit Sub
    If gdTickCount < dNextTime Then Exit Sub
    
    ' only if flag file exists
    strFile = App.Path & "\QuotesWebPage.flg"
    If dNextTime = 0 Then
        If Not FileExist(strFile) Then
            dNextTime = -1
            Exit Sub
        End If
    End If
       
    ' read contents of flag file
    aStrings.FromFile strFile
    strFile = aStrings(0) ' web page file
    ' calc next refresh time
    i = Val(aStrings(1)) ' # seconds to refresh
    If i <= 0 Then i = 60 ' default
    dNextTime = gdTickCount + i * 1000#
    ' get list of symbols into an array
    aStrings.SplitFields aStrings(2), vbTab ' symbols
    If Len(strFile) = 0 Or aStrings.Size = 0 Then Exit Sub
    
    ' HTML header (auto-refresh every "i" seconds)
    aHTML.Size = 0
    aHTML.Add "<HTML><HEAD><TITLE>Quotes</TITLE><META HTTP-EQUIV=""REFRESH"" CONTENT=""" & Str(i) & """></HEAD>"
    aHTML.Add "<BODY TEXT=BLACK LINK=BLUE VLINK=PURPLE ALINK=RED><FONT face=""Arial, sans-serif"">"
    aHTML.Add "<A href=""."">TradeNavigator.com</A><BR>"
    aHTML.Add Format(Date, "mmm d") & " <B>" & Format(Now, "h:mm:ssa/p") & "</B><BR>"
       
    ' which symbol goes in the title depends on time of day
    i = Hour(Now) * 100 + Minute(Now)
    If i < 730 Or (i >= 1430 And Weekday(Date) <> vbFriday) Then
        strTitle = "YM-067" ' "$EUR-USD"
    Else
        strTitle = "$DJIA"
    End If
    
    ' for each symbol
    For i = 0 To aStrings.Size - 1
        strText = ""
        strSymbol = Trim(aStrings(i))
        If Len(strSymbol) > 0 Then
            nSymbolID = GetSymbolID(strSymbol)
            Set Bars = GetBars(SymbolOrSymbolID(nSymbolID, strSymbol), "Daily")
            If Not Bars Is Nothing Then
                dClose = Bars(eBARS_Close, Bars.Size - 1)
                If dClose <> kNullData Then
                    ' get Delta = Close - prior Settle
                    dPrev = GetPrevCloseForQB(Bars)
                    If dPrev = kNullData Then
                        dPrev = dClose
                    End If
                    If dClose < dPrev Then
                        strText = " <Font Color=RED>" & Bars.PriceDisplay(dClose - dPrev)
                    ElseIf dClose > dPrev Then
                        strText = " <Font Color=GREEN>+" & Bars.PriceDisplay(dClose - dPrev)
                    Else
                        strText = " <Font Color=BLACK>+" & Bars.PriceDisplay(dClose - dPrev)
                    End If
                    strText = "<B>" & RollSymbolForDate(strSymbol) & strText & "</Font></B><BR>" & Bars.PriceDisplay(dClose)
                    'strText = "<B>" & RollSymbolForDate(strSymbol) & strText & "</Font></B> " & Bars.PriceDisplay(dClose)
                        
                    ' time of last trade
                    dDate = Bars(eBARS_DateTime, Bars.Size - 1)
                    If dDate > 0 Then
                        dTime = Bars.Prop(eBARS_LastTickTime)
                        If dTime <= 0 Then
                            ' session date
                            strText = strText & " at " & Format(Int(dDate), "mm/dd")
                        Else
                            ' time
                            dTime = Int(dDate) + dTime / 1440#
                            If g.bShowInLocalTimeZone Then
                                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                            End If
                            strText = strText & " at " & Format(dTime, "h:mma/p")
                            
                            ' web page title
                            If strSymbol = strTitle Then
                                If dClose >= dPrev Then
                                    strTitle = "+" & Bars.PriceDisplay(dClose - dPrev)
                                Else
                                    strTitle = Bars.PriceDisplay(dClose - dPrev)
                                End If
                                strTitle = strTitle & " " & Format(dTime, "h:mma/p") & " " & strSymbol
                                aHTML(0) = Replace(aHTML(0), ">Quotes<", ">" & strTitle & "<")
                                strTitle = ""
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Len(strText) > 0 Then
            aHTML.Add strText & "<BR>"
        End If
        Set Bars = Nothing
    Next
    
    ' HTML footer
    aHTML.Add "</BODY></HTML>"
    aHTML.ToFile strFile

End Sub

' JUST FOR INTERNAL PURPOSES: to create a web page with a few updated quotes
Private Sub QuotesTableWebPage()
On Error Resume Next

    Dim i&, iLine&, iRow&, d#, nRefresh&, nSymbolID&, dClose#, dPrev#, dDate#, dTime#, nFontSize&
    Dim strFile$, strSymbol$, strText$, strTitle$, strSymbolInTitle$, strHdr$
    Dim aStrings As New cGdArray, aSymbols As New cGdArray
    Dim aHTML As New cGdArray
    Dim Table As New cGdTable
    Dim Bars As cGdBars
    Static dNextTime As Double
    
    If dNextTime < 0 Then Exit Sub
    If gdTickCount < dNextTime Then Exit Sub
    If g.RealTime.ConnectionStatus <> eGDConnectionStatus_Connected Then Exit Sub
    
    ' only if flag file exists
    strFile = App.Path & "\QuotesWebPage.flg"
    If dNextTime = 0 Then
        If Not FileExist(strFile) Then
            dNextTime = -1
            Exit Sub
        End If
    End If
       
    ' read contents of flag file
    aStrings.FromFile strFile
    strFile = aStrings(0) ' web page file
    If Len(strFile) = 0 Then Exit Sub
    
    ' calc next refresh time
    nRefresh = Val(Parse(aStrings(1), vbTab, 1)) ' # seconds to refresh
    If nRefresh <= 0 Then nRefresh = 60 ' default
    dNextTime = gdTickCount + nRefresh * 1000#
    
    nFontSize = Val(Parse(aStrings(1), vbTab, 2)) ' font size
    If nFontSize = 0 Then nFontSize = 14
    
    ' which symbol goes in the title depends on time of day
    i = Hour(Now) * 100 + Minute(Now)
    If i < 730 Or (i >= 1430 And Weekday(Date) <> vbFriday) Then
        strSymbolInTitle = "YM-067" ' "$EUR-USD"
    Else
        strSymbolInTitle = "$DJIA"
    End If
    strTitle = "Quotes"
    
    ' create table to hold the display info
    Table.CreateField eGDARRAY_Strings, 0, "Symbol"
    Table.CreateField eGDARRAY_Strings, 1, "Price"
    Table.CreateField eGDARRAY_Strings, 2, "Change"
    Table.CreateField eGDARRAY_Strings, 3, "As of"
    iRow = 0
    
    ' a line for each category
    For iLine = 2 To aStrings.Size - 1
        strHdr = Parse(aStrings(iLine), vbTab, 2)
        aSymbols.SplitFields Parse(aStrings(iLine), vbTab, 1), ","
        
        ' for each symbol in this table row
        For i = 0 To aSymbols.Size - 1
            strText = ""
            strSymbol = Trim(aSymbols(i))
            If Len(strSymbol) > 0 Then
                nSymbolID = GetSymbolID(strSymbol)
                Set Bars = GetBars(SymbolOrSymbolID(nSymbolID, strSymbol), "Daily")
                If Not Bars Is Nothing Then
                    dClose = Bars(eBARS_Close, Bars.Size - 1)
                    If dClose <> kNullData Then
                        ' if first symbol in row, add category hdr
                        If i = 0 Then
                            Table.NumRecords = iRow + 1
                            Table(0, iRow) = "|" & strHdr
                            iRow = iRow + 1
                        End If
                        
                        ' add symbol and price
                        Table.NumRecords = iRow + 1
                        Table(0, iRow) = RollSymbolForDate(strSymbol)
                        Table(1, iRow) = Bars.PriceDisplay(dClose)

                        ' get Delta = Close - prior Settle
                        dPrev = GetPrevCloseForQB(Bars)
                        If dPrev = kNullData Then
                            dPrev = dClose
                        End If
                        If dClose < dPrev Then
                            strText = "<Font Color=RED>" & Bars.PriceDisplay(dClose - dPrev)
                        ElseIf dClose > dPrev Then
                            strText = "<Font Color=GREEN>+" & Bars.PriceDisplay(dClose - dPrev)
                        Else
                            strText = "<Font Color=BLACK>+" & Bars.PriceDisplay(dClose - dPrev)
                        End If
                        Table(2, iRow) = strText & "</Font>"
                            
                        ' time of last trade
                        dDate = Bars(eBARS_DateTime, Bars.Size - 1)
                        If dDate > 0 Then
                            dTime = Bars.Prop(eBARS_LastTickTime)
                            If dTime <= 0 Then
                                ' session date
                                Table(3, iRow) = Format(Int(dDate), "mm/dd")
                            Else
                                ' time
                                dTime = Int(dDate) + dTime / 1440#
                                If g.bShowInLocalTimeZone Then
                                    dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                                End If
                                Table(3, iRow) = Format(dTime, "h:mma/p")
                                
                                ' web page title
                                If strSymbol = strSymbolInTitle Then
                                    If dClose >= dPrev Then
                                        strTitle = "+" & Bars.PriceDisplay(dClose - dPrev)
                                    Else
                                        strTitle = Bars.PriceDisplay(dClose - dPrev)
                                    End If
                                    strTitle = strTitle & " " & Format(dTime, "h:mma/p") & " " & strSymbol
                                    strSymbolInTitle = ""
                                End If
                            End If
                        End If
                        
                        iRow = iRow + 1
                    End If
                End If
            End If
        Next
    Next
    Set Bars = Nothing
        
    ' consolidate into table rows
    iLine = 0
    For iRow = 0 To Table.NumRecords - 1
        strText = Table(0, iRow)
        If Left(strText, 1) = "|" Then
            iLine = iRow + 1
        ElseIf iRow > iLine Then
            For i = 0 To 3
                Table(i, iLine) = Table(i, iLine) & "<BR>" & Table(i, iRow)
            Next
            Table(0, iRow) = "" ' to clear this out
        End If
    Next
    
    ' HTML header (auto-refresh every "i" seconds)
    strText = "{font-family:Arial; font-size:" & Str(nFontSize) & "pt;}"
    aHTML.Size = 0
    aHTML.Add "<HTML><HEAD><TITLE>" & strTitle & "</TITLE><META HTTP-EQUIV=""REFRESH"" CONTENT=""" & Str(nRefresh) & """>"
    aHTML.Add "<style type='text/css'>"
    aHTML.Add "TH" & strText
    aHTML.Add "TR" & strText
    aHTML.Add "EM" & strText
    aHTML.Add "</style></HEAD>"
    aHTML.Add "<BODY TEXT=BLACK LINK=BLUE VLINK=PURPLE ALINK=RED><FONT face=""Arial, sans-serif"">"
    
    'aHTML.Add "<A href=""."">TradeNavigator.com</A><BR>"
    aHTML.Add "<EM>Last updated: <B>" & Format(Date, "mmm d") & "</B> at <B>" & Format(Now, "h:mm:ssa/p") & "</B></EM><BR>"
    
    aHTML.Add "<TABLE cellSpacing=0 cellPadding=4 border=3 bgcolor=WHITE BORDERCOLORLIGHT=#D8D8D8 BORDERCOLORDARK=#808080><TBODY>"
    
    ' now format each row
    For iRow = 0 To Table.NumRecords - 1
        strText = Table(0, iRow)
        If Left(strText, 1) = "|" Then
            ' table header
            strText = Trim(Mid(strText, 2))
            If Len(strText) = 0 Then
                strText = "Symbol"
            Else
                strText = "<Font Color=BLUE>" & strText & "</FONT>"
            End If
            aHTML.Add "<TR BGCOLOR=#FFFCC0><TH>" & strText & "</TH><TH>Price</TH><TH>Change</TH><TH>As of</TH></TR>"
        ElseIf Len(strText) > 0 Then
            ' table row
            aHTML.Add "<TR align=middle><TD><B>" & Table(0, iRow) & "</B></TD><TD>" & Table(1, iRow) _
                & "</TD><TD><B>" & Table(2, iRow) & "</B></TD><TD>" & Table(3, iRow) & "</TD></TR>"
        End If
    Next
    
    ' HTML footer
    aHTML.Add "</TBODY></TABLE></FONT></BODY></HTML>"
    aHTML.ToFile strFile

End Sub

' OLD flag file method -- but need to keep this active for backwards-compatibility and extra options:
' To create web page with current Acct Status (open positions, open orders, etc)
Private Sub AcctStatusCheck2()

If Not IsIDE Then
    On Error Resume Next
End If

    Dim i&, d#, nFontSize&
    Dim s$, strFile$, strAccts$, strSymbols$
    Dim bIncludeOpenEquity As Boolean
    Dim aStrings As New cGdArray
    Static dNextTime As Double, strWeb$, strWeb2$
    
    If gdTickCount < dNextTime Then Exit Sub
    dNextTime = gdTickCount + 15000# ' default time to wait
    
    ' if auto-create is implemented
    s = GetMidCmd
    If Len(s) > 0 Then
        strAccts = UCase(Parse(s, vbTab, 2))
        If Len(strAccts) > 0 Then
            s = CreateAcctStatusWebPage(strAccts, "", -1, False)
            If s <> strWeb2 Then ' only if changed
                strWeb2 = s
                SendWebPage RI_GetMachineID & "-!.txt", strWeb2
            End If
        End If
        strAccts = ""
    End If
    
    ' only if flag file exists
    strFile = App.Path & "\AcctStatusWebPage.flg"
    If Not FileExist(strFile) Then Exit Sub
       
    ' read contents of flag file:
    ' - filename (either local filename, or just ".HTM" or ".TXT" to send to our web server)
    ' - #Seconds to refresh/resend, IncludeOpenEquity, FontSize
    ' - accounts (comma-delimited list)
    ' - symbols (comma-delimited list)
    aStrings.FromFile strFile
    strFile = aStrings(0) ' web page file
    If Len(strFile) = 0 Then Exit Sub
    strAccts = aStrings(2)
    strSymbols = aStrings(3)
    
    ' calc next refresh time
    i = Val(Parse(aStrings(1), vbTab, 1)) ' # seconds to refresh
    If i <= 0 Then i = 60 ' default
    dNextTime = gdTickCount + i * 1000#
    
    ' include open equity?
    bIncludeOpenEquity = (ValOfText(Parse(aStrings(1), vbTab, 2)) <> 0)
    
    ' font size for HTML pages
    nFontSize = Val(Parse(aStrings(1), vbTab, 3))
    If InStr(UCase(strFile), ".HTM") = 0 Then
        nFontSize = -1
    End If
    
    s = CreateAcctStatusWebPage(strAccts, strSymbols, nFontSize, bIncludeOpenEquity)
    If s <> strWeb Then ' only if changed
        strWeb = s
        If UCase(strFile) = ".HTM" Or UCase(strFile) = ".TXT" Then
            strFile = RI_GetMachineID & strFile
            SendWebPage strFile, strWeb
        Else
            FileFromString strFile, strWeb
        End If
    End If

End Sub

' NEW method to create web page with current Acct Status (open positions, open orders, etc)
Public Function AcctStatusCheck(Optional ByVal bForceRefreshNow As Boolean = False) As Boolean

If Not IsIDE Then
    On Error Resume Next
End If

    Dim s$, strAccts$
    Static dNextTime#, strWeb$
    
    ' only once every 30 seconds (or 5 minutes if no live broker account)
    If gdTickCount < dNextTime And Not bForceRefreshNow Then Exit Function
    If HasModule("BRKRLIVE") And HasModule("B_*") And HasModule("RTG") Then
        dNextTime = gdTickCount + 30000 ' 30 seconds
    Else
        dNextTime = gdTickCount + 300000 ' 5 minutes
    End If
    
    If CanDoAcctStatusWebPage Then
        strAccts = Trim(g.Broker.WebAccounts)
        If Len(strAccts) > 0 Then
            s = CreateAcctStatusWebPage(strAccts, "$DJIA,ES-067", g.Broker.WebFontSize, True)
            If s <> strWeb Or bForceRefreshNow Then ' only if changed
                strWeb = s
                SendWebPage RI_GetMachineID & ".HTM", strWeb
                AcctStatusCheck = True
            End If
        End If
    End If

End Function

Private Function CreateAcctStatusWebPage(ByVal strAccts$, Optional ByVal strSymbols$ = "", _
        Optional ByVal nFontSize& = 5, Optional ByVal bIncludeOpenEquity As Boolean = True) As String

' when in production, we just want to ignore any errors related to this routine ...
If Not IsIDE Then
    On Error Resume Next
End If

    Dim i&, iItem&, nSymbolID&, nActive&, nCount&
    Dim d#, dClose#, dPrev#, dDate#, dTime#, dTotalOpenEq#, dTotalClosed#
    Dim strSymbol$, strText$, strTitle$, strAcct$
    Dim bCustomAcctTotal As Boolean
    Dim aAccts As New cGdArray
    Dim aSymbols As New cGdArray
    Dim aHTML As New cGdArray
    Dim aAcctInfo As New cGdArray
    Dim Bars As cGdBars
    Dim Acct As cPtAccount
    Dim Positions As cAccountPositions
    Dim Position As cAccountPosition
    Dim Orders As cPtOrders
    Dim Order As cPtOrder
    Dim TradeItem As cAutoTradeItem
    
    aAccts.SplitFields strAccts, ","
    aSymbols.SplitFields strSymbols, ","
    If aAccts.Size = 0 And aSymbols.Size = 0 Then Exit Function
        
    If nFontSize = 0 Then nFontSize = 5
        
    ' get time of data stream
    dTime = 0
    If Not g.RealTime Is Nothing Then
        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
            dTime = g.RealTime.FeedTime
        End If
    End If
        
    ' HTML header (auto-refresh every "i" seconds)
    i = 30 '??
    aHTML.Size = 0
    If nFontSize > 0 Then
        aHTML.Add "<HTML><HEAD><TITLE>Account Info</TITLE><META HTTP-EQUIV=""REFRESH"" CONTENT=""" & Str(i) & """></HEAD>"
        aHTML.Add "<BODY TEXT=BLACK LINK=BLUE VLINK=PURPLE ALINK=RED><FONT face=""Arial, sans-serif"" size=" & Str(nFontSize) & ">"
        'aHTML.Add "<A href=""."">TradeNavigator.com</A><BR>"
        If dTime > 0 Then
            If g.bShowInLocalTimeZone Then
                dTime = ConvertTimeZone(dTime, "NY", "")
            End If
            aHTML.Add Format(dTime, "mmm d") & " <B>" & Format(dTime, "h:mm:ssa/p") & "</B><BR>"
        Else
            aHTML.Add "<B>Data stream is DISCONNECTED!</B><BR>"
        End If
    Else
        If dTime = 0 Then
            dTime = LastDailyDownload
        End If
        aHTML.Add Format(dTime, "yyyy-mm-dd") & vbTab & Format(RI_GetDataServiceID / 1000, "#000000") & ":" & Format(RI_GetDataServiceID Mod 1000, "000")
        aSymbols.Size = 0
    End If
             
    ' for each symbol (this is optional -- just testing for now)
    For iItem = 0 To aSymbols.Size - 1
        strText = ""
        strSymbol = Trim(aSymbols(iItem))
        If Len(strSymbol) > 0 Then
            nSymbolID = GetSymbolID(strSymbol)
            Set Bars = frmQuotes.GetBars(nSymbolID, "Daily") '(SymbolOrSymbolID(nSymbolID, strSymbol), "Daily")
            If Not Bars Is Nothing Then
                dClose = Bars(eBARS_Close, Bars.Size - 1)
                If dClose <> kNullData Then
                    ' get Delta = Close - prior Settle
                    dPrev = GetPrevCloseForQB(Bars)
                    If dPrev = kNullData Then
                        dPrev = Bars(eBARS_Open, Bars.Size - 1)
                        If dPrev = kNullData Then
                            dPrev = dClose
                        End If
                    End If
                    If dClose < dPrev Then
                        strText = " (<Font Color=RED>" & Bars.PriceDisplay(dClose - dPrev)
                    ElseIf dClose > dPrev Then
                        strText = " (<Font Color=GREEN>+" & Bars.PriceDisplay(dClose - dPrev)
                    Else
                        strText = " (<Font Color=BLACK>+" & Bars.PriceDisplay(dClose - dPrev)
                    End If
                    'strText = "<B>" & strSymbol & strText & "</Font></B><BR>" & Bars.PriceDisplay(dClose)
                    strText = strSymbol & " = " & Bars.PriceDisplay(dClose) & strText & "</Font>)"
                        
                    ' time of last trade
                    dDate = Bars(eBARS_DateTime, Bars.Size - 1)
                    If dDate > 0 Then
                        dTime = Bars.Prop(eBARS_LastTickTime)
                        If dTime <= 0 Then
                            ' session date
                            strText = strText & " at " & Format(Int(dDate), "mm/dd")
                        Else
                            ' time
                            dTime = Int(dDate) + dTime / 1440#
                            If g.bShowInLocalTimeZone Then
                                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                            End If
                            strText = strText & " at " & Format(dTime, "h:mma/p")
                        End If
                    End If
                End If
            End If
        End If
        If Len(strText) > 0 Then
            aHTML.Add strText & "<BR>"
        End If
        Set Bars = Nothing
    Next
    
    ' # of enabled auto-trading items (but ignoring any strategy basket "parent")
    nActive = 0
    nCount = 0
    For i = 1 To g.TradingItems.Count
        Set TradeItem = g.TradingItems(i)
        If Not TradeItem Is Nothing Then
            If TradeItem.ParentID >= 0 Then
                nCount = nCount + 1
                If TradeItem.Active Then
                    nActive = nActive + 1
                End If
            End If
        End If
        Set TradeItem = Nothing
    Next
    If nCount > 0 Then
        strText = Str(nActive) & " of " & Str(nCount) & " auto-trading items enabled"
        If nFontSize > 0 Then
            aHTML.Add strText & "<BR>"
        Else
            aHTML.Add strText
        End If
    End If
    
    ' for each account
    For iItem = 0 To aAccts.Size - 1
        ' acct info
        strAcct = aAccts(iItem)
        aAcctInfo.Size = 0
        Set Acct = g.Broker.Account(strAcct)
        If Not Acct Is Nothing Then
            ' ACCT info -- e.g. My Pfg Acct = $12345 - $199
            If nFontSize > 0 Then
                If Acct.ConnectionStatus = eGDConnectionStatus_Connected Then
                    If Not bIncludeOpenEquity Then
                        strText = ""
                    Else
                        strText = Str(Round(Acct.OpenProfit))
                        If Left(strText, 1) = "-" Then
                            strText = " <Font Color=RED> - $" & Mid(strText, 2)
                        Else
                            strText = " <Font Color=GREEN> + $" & strText
                        End If
                        strText = strText & "</Font> = $" & Str(Round(Acct.CurrentClosedBalance + Acct.OpenProfit))
                    End If
                    strText = " = $" & Str(Round(Acct.CurrentClosedBalance)) & strText
                    If Left(Acct.Name, 1) = "$" Then
                        ' if any Account name starts with a '$' then only use those in the Totals
                        If Not bCustomAcctTotal Then
                            bCustomAcctTotal = True
                            dTotalClosed = 0
                            dTotalOpenEq = 0
                        End If
                        dTotalClosed = dTotalClosed + Acct.CurrentClosedBalance
                        dTotalOpenEq = dTotalOpenEq + Acct.OpenProfit
                    ElseIf Acct.TypeOfAccount = eGDTypeOfAccount_BrokerLive And Not bCustomAcctTotal Then
                        dTotalClosed = dTotalClosed + Acct.CurrentClosedBalance
                        dTotalOpenEq = dTotalOpenEq + Acct.OpenProfit
                    End If
                Else
                    strText = " is Disconnected!"
                    strTitle = "Disconnected!"
                End If
                aHTML.Add "=========================<BR>"
                aHTML.Add "<B>" & Acct.Name & strText & "</B><BR>"
            Else
                aHTML.Add ""
                If Acct.ConnectionStatus = eGDConnectionStatus_Connected Then
                    If Not bIncludeOpenEquity Then
                        strText = ""
                    ElseIf Acct.OpenProfit < 0 Then
                        strText = Str(Round(Acct.OpenProfit))
                    Else
                        strText = "+" & Str(Round(Acct.OpenProfit))
                    End If
                    strText = strAcct & vbTab & "= $" & Str(Round(Acct.CurrentClosedBalance)) & vbTab & strText
                Else
                    strText = strAcct & vbTab & "= Disconnected"
                End If
                aHTML.Add strText
            End If
            
            ' Open POSITIONS
            'Set Positions = g.Broker.FillSummariesForAccount(strAcct)
            Set Positions = g.Broker.OpenPositionsForAccount(strAcct, True)
            If Not Positions Is Nothing Then
                For i = 1 To Positions.Count
                    Set Position = Positions(i)
                    If Not Position Is Nothing Then
                        ' .AutoTradeItemID is -1 for total (0 for manual)
                        If Position.CurrentPositionSnapshot <> 0 And Position.AutoTradeItemID < 0 Then
                            ' e.g. GC3-201203: Short 1 = -$xxxx
                            strSymbol = Position.Symbol
                            If Position.CurrentPositionSnapshot > 0 Then
                                strText = "Long " & Str(Position.CurrentPositionSnapshot)
                            Else
                                strText = "Short " & Str(Abs(Position.CurrentPositionSnapshot))
                            End If
                            If nFontSize > 0 Then
                                If bIncludeOpenEquity Then
                                    If Position.OpenProfit < 0 Then
                                        strText = strText & " = <Font Color=RED> -$" & Str(Abs(Round(Position.OpenProfit)))
                                    Else
                                        strText = strText & " = <Font Color=GREEN> $" & Str(Round(Position.OpenProfit))
                                    End If
                                    strText = "<B>" & strSymbol & ": " & strText & "</Font></B><BR>"
                                Else
                                    strText = "<B>" & strSymbol & ": " & strText & "</B><BR>"
                                End If
                            ElseIf Not bIncludeOpenEquity Then
                                strText = "= " & strText
                            ElseIf Position.OpenProfit < 0 Then
                                strText = "= " & strText & vbTab & Str(Round(Position.OpenProfit))
                            Else
                                strText = "= " & strText & vbTab & "+" & Str(Round(Position.OpenProfit))
                            End If
                            aAcctInfo.Add strSymbol & vbTab & strText
                        End If
                    End If
                Next
            End If
            
            ' Open ORDERS
            Set Orders = g.Broker.OrdersForAccount(strAcct, True)
            If Not Orders Is Nothing Then
                For i = 1 To Orders.Count
                    Set Order = Orders(i)
                    If Not Order Is Nothing Then
                        If IsOpenOrder(Order.Status, True) Then
                            ' e.g. Working: Sell 1 at 1647.2 STOP
                            strSymbol = Order.Symbol
                            If nFontSize > 0 Then
                                If Order.Buy Then
                                    strText = "<Font Color=GREEN>"
                                Else
                                    strText = "<Font Color=RED>"
                                End If
                                strText = OrderStatus(Order.Status) & ": " _
                                    & strText & Order.OrderText(False) & "</Font><BR>"
                            Else
                                strText = OrderStatus(Order.Status) & ": " & Order.OrderText(False)
                            End If
                            aAcctInfo.Add strSymbol & " " & vbTab & strText
                        End If
                    End If
                Next
            End If
            
            ' Sort by symbol, and append current price
            If aAcctInfo.Size > 0 Then
                aAcctInfo.Sort
                For i = 0 To aAcctInfo.Size - 1
                    strText = ""
                    strSymbol = aAcctInfo(i)
                    d = InStr(strSymbol, vbTab)
                    If d > 0 Then
                        strText = Trim(Mid(strSymbol, d + 1))
                        strSymbol = Trim(Left(strSymbol, d - 1))
                    End If
                    'strSymbol = Parse(aAcctInfo(i), vbTab, 1)
                    'strText = Parse(aAcctInfo(i), vbTab, 2)
                    If strSymbol <> Parse(aAcctInfo(i - 1), vbTab, 1) Then
                        If nFontSize > 0 Then
                            aHTML.Add "- - - - - - - - - - - - - - - - - - - - -<BR>"
                        End If
                        ' if no position, then add a "Flat" for this symbol
                        If InStr(strText, "= ") = 0 Then
                            If nFontSize > 0 Then
                                aHTML.Add "<B>" & strSymbol & "</B>: Flat<BR>"
                            Else
                                aHTML.Add strSymbol & vbTab & "= Flat"
                            End If
                        End If
                    End If
                    If nFontSize > 0 Then
                        aHTML.Add strText
                    Else
                        aHTML.Add strSymbol & vbTab & strText
                    End If
                    If strSymbol <> Parse(aAcctInfo(i + 1), vbTab, 1) And g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected And nFontSize > 0 Then
                        Set Bars = frmTTSummary.GetBars(strSymbol)
                        If Not Bars Is Nothing Then
                            dClose = Bars(eBARS_Close, Bars.Size - 1)
                            If dClose <> kNullData Then
                                dPrev = kNullData
                                If 1 Then ' TLB 9/29/2014:
                                    ' get Delta = Close - prior Settle
                                    dPrev = GetPrevCloseForQB(Bars)
                                End If
                                If dPrev = kNullData Then
                                    ' or show up/down from the Open
                                    dPrev = Bars(eBARS_Open, Bars.Size - 1)
                                    If dPrev = kNullData Then
                                        dPrev = dClose
                                    End If
                                End If
                                If dClose >= dPrev Then
                                    strText = " (<Font Color=GREEN>+"
                                Else
                                    strText = " (<Font Color=RED>"
                                End If
                                strText = Bars.PriceDisplay(dClose) & strText _
                                    & Bars.PriceDisplay(dClose - dPrev) & "</Font>)"
                                ' time of last trade
                                dDate = Bars(eBARS_DateTime, Bars.Size - 1)
                                If dDate > 0 Then
                                    dTime = Bars.Prop(eBARS_LastTickTime)
                                    If dTime <= 0 Then
                                        ' session date
                                        'strText = strText & " at " & Format(Int(dDate), "mm/dd")
                                        strText = "As of " & Format(Int(dDate), "mm/dd") & " = " & strText
                                    Else
                                        ' time
                                        dTime = Int(dDate) + dTime / 1440#
                                        If g.bShowInLocalTimeZone Then
                                            dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                                        End If
                                        'strText = strText & " at " & Format(dTime, "h:mma/p")
                                        strText = "As of " & Format(dTime, "h:mma/p") & " = " & strText
                                    End If
                                End If
                                aHTML.Add strText & "<BR>"
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Set Acct = Nothing
        Set Position = Nothing
        Set Positions = Nothing
        Set Order = Nothing
        Set Orders = Nothing
        Set Bars = Nothing
        aAcctInfo.Size = 0
    Next
    
    If nFontSize > 0 Then
        ' Title
        If Len(strTitle) = 0 Then
            If bIncludeOpenEquity Then
                ' for web page title
                If dTotalOpenEq > 0 Then
                    strTitle = "+" & Str(Round(dTotalOpenEq))
                ElseIf dTotalOpenEq < 0 Then
                    strTitle = Str(Round(dTotalOpenEq))
                End If
                strTitle = strTitle & " = " & Str(Round(dTotalClosed + dTotalOpenEq))
                ' for total of live accounts
                strText = Str(Round(dTotalOpenEq))
                If Left(strText, 1) = "-" Then
                    strText = " <Font Color=RED> - $" & Mid(strText, 2)
                Else
                    strText = " <Font Color=GREEN> + $" & strText
                End If
                strText = strText & "</Font> = $" & Str(Round(dTotalClosed + dTotalOpenEq))
            Else
                strTitle = "$" & Str(Round(dTotalClosed))
                strText = ""
            End If
            ' add total for all live accounts
            If dTotalClosed + dTotalOpenEq <> 0 Then
                strText = "Total LIVE = $" & Str(Round(dTotalClosed)) & strText
                For i = 0 To aHTML.Size - 1
                    If InStr(aHTML(i), "==========<BR>") > 0 Then
                        aHTML.Add "<B>" & strText & "</B><BR>", i
                        Exit For
                    End If
                Next
            End If
        End If
        aHTML(0) = Replace(aHTML(0), ">Account Info<", ">" & strTitle & "<")
        
        ' HTML footer
        aHTML.Add "</BODY></HTML>"
    End If
    
    CreateAcctStatusWebPage = aHTML.JoinFields(vbCrLf)
End Function

Private Sub CopyTab(ByRef astrSave As cGdArray, ByVal idxFrom&, ByVal idxTo&, _
    ByVal strName$, ByVal eTabStyleTo As eGDQuoteStyle)
On Error GoTo ErrSection:

    Dim i&, strTemp$, strSym$, strPeriod$
    
    Dim eTabStyleFrom As eGDQuoteStyle
    Dim aCopy As New cGdArray
    
    If idxFrom < 0 Or idxFrom >= astrSave.Size Then Exit Sub

    strTemp = astrSave(idxFrom)
    eTabStyleFrom = Parse(strTemp, vbTab, 2)
    
    aCopy.SplitFields strTemp, vbTab
    If UCase(aCopy(0)) = "(FILTER)" Then
        'remove label text when copying filter tab  - 4238
        If InStr(UCase(aCopy(2)), "LABEL;CURRENT FILTER =") <> 0 Then
            strSym = Parse(aCopy(2), ";", 1)
            aCopy(2) = Replace(aCopy(2), strSym, "")
            strSym = Parse(aCopy(2), ",", 1)
            aCopy(2) = Replace(aCopy(2), strSym, "")
        End If
    End If
    
    aCopy(eGDTabSettings_Name) = strName
    aCopy(eGDTabSettings_Style) = eTabStyleTo
    
    strTemp = aCopy.JoinFields(vbTab)
    
    m.tblTabInfo.SetRecord strTemp, idxTo, vbTab
    m.tblTabInfo(TabField(eGDTabSettings_Form), idxTo) = 0         'reset
    TabStr(eGDTabSettings_FilterID, idxTo) = ""
                    
    If eTabStyleTo = eGDQuoteStyle_Grid And eTabStyleFrom <> eGDQuoteStyle_Grid Then
        aCopy.SplitFields TabStr(eGDTabSettings_Symbols, idxTo), ","
        For i = 0 To aCopy.Size - 1
            If Len(aCopy(i)) > 0 Then
                aCopy(i) = Parse(aCopy(i), ";", 1) & ";Daily"
                m.astrSymbols.Add aCopy(i)
            End If
        Next
        TabStr(eGDTabSettings_Symbols, idxTo) = aCopy.JoinFields(",")
        TabStr(eGDTabSettings_Fields, idxTo) = m.strDefaultFields
    Else
        aCopy.SplitFields aCopy(2), ","
        For i = 0 To aCopy.Size - 1
            strTemp = aCopy(i)
            strSym = Parse(strTemp, ";", 1)
            strPeriod = Parse(strTemp, ";", 2)
            If Len(strSym) > 0 Then
                If Len(strPeriod) > 0 And strPeriod <> "0" Then
                    m.astrSymbols.Add strTemp
                Else
                    m.astrSymbols.Add strSym & ";Daily"
                End If
            End If
        Next
    End If
    
    'aardvark 4225
    m.astrSymbols.Sort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.CopyTab"

End Sub

Public Sub GridTextIncrease()
On Error GoTo ErrSection:

    Dim FontNew As StdFont
    
    If fgQuotes.Visible Then
        StatusMsg "Grid font changed: " & Str(fgQuotes.Font.Size) & " - " & Str(fgQuotes.Font.Size + 1)
        fgQuotes.Font.Size = fgQuotes.Font.Size + 1
    Else
        Set FontNew = Me.pbQuoteBoard.Font
        FontNew.Size = FontNew.Size + 1
        m.QB.Font = FontNew
        ShowCategory
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.GridTextIncrease"

End Sub

Public Sub GridTextDecrease()
On Error GoTo ErrSection:

    Dim FontNew As StdFont
    
    If fgQuotes.Visible Then
        StatusMsg "Grid font changed: " & Str(fgQuotes.Font.Size) & " - " & Str(fgQuotes.Font.Size - 1)
        fgQuotes.Font.Size = fgQuotes.Font.Size - 1
    Else
        Set FontNew = Me.pbQuoteBoard.Font
        FontNew.Size = FontNew.Size - 1
        m.QB.Font = FontNew
        ShowCategory
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.GridTextDecrease"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangePeriodForRow
'' Description: Change the bar period for the given row
'' Inputs:      Row, New Period, Old Period
'' Returns:     True if changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangePeriodForRow(ByVal lRow As Long, ByVal strNewPeriod As String, Optional ByVal strOldPeriod As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim nRedraw As RedrawSettings       ' State of the grids redraw property
    Dim strSymbol As String             ' Symbol for the row
    Dim bRemoved As Boolean             ' Was removed from the data table
    Dim bAdded As Boolean               ' Was added to the data table
    Dim QbGrid As VSFlexGrid            ' Grid to modify
    
    bReturn = False
    If m.frmActiveDetTab Is Nothing Then
        Set QbGrid = fgQuotes
    Else
        Set QbGrid = m.frmActiveDetTab.fgQuotes
    End If
    
    With QbGrid
        If .MergeRow(lRow) = False Then
            strSymbol = Parse(.TextMatrix(lRow, GDCol(eGDCol_Symbol)), "(", 1)
            If Len(strSymbol) > 0 Then
                If Len(strOldPeriod) = 0 Then
                    strOldPeriod = .TextMatrix(lRow, GDCol(eGDCol_Period))
                End If
                
                If strNewPeriod <> strOldPeriod Then
                    nRedraw = .Redraw
                    .Redraw = flexRDNone
                    
                    bRemoved = RemoveSymbolFromGrid(lRow, strOldPeriod)
                    bAdded = AddSymbolToGrid(strSymbol, strNewPeriod, lRow)
                
                    'If ((bRemoved = True) Or (bAdded = True)) And (g.RealTime.Active = True) Then
                        'g.RealTime.UpdateSymbolList
                    'End If
                
                    .Redraw = nRedraw
                    
                    bReturn = True
                End If
            End If
        End If
    End With
    
    ChangePeriodForRow = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuotes.ChangePeriodForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangePeriodForAllRows
'' Description: Change the bar period for all of the rows in the current tab
'' Inputs:      New Period, Old Period
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangePeriodForAllRows(ByVal strNewPeriod As String, Optional ByVal strOldPeriod As String = "")
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' State of the grids redraw property
    Dim lIndex As Long                  ' Index into a for loop
    Dim QbGrid As VSFlexGrid            ' Grid to modify
    Dim alChanged As cGdArray           ' Array of rows that changed
    
    If m.frmActiveDetTab Is Nothing Then
        Set QbGrid = fgQuotes
    Else
        Set QbGrid = m.frmActiveDetTab.fgQuotes
    End If
    
    Set alChanged = New cGdArray
    alChanged.Create eGDARRAY_Longs
        
    With QbGrid
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If ChangePeriodForRow(lIndex, strNewPeriod, strOldPeriod) = True Then
                alChanged.Add lIndex
            End If
        Next lIndex
        
        .Redraw = nRedraw
        
        If m.frmActiveDetTab Is Nothing Then
            ResetRows
        Else
            TabStr(eGDTabSettings_Symbols, m.frmActiveDetTab.MyTabIndex) = m.frmActiveDetTab.MySymbols(True)
        End If
        TotalRefresh False
        If Not m.frmActiveDetTab Is Nothing Then
            For lIndex = 0 To alChanged.Size - 1
                UpdateCols alChanged(lIndex), , , , QbGrid
            Next lIndex
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuotes.ChangePeriodForAllRows"

End Sub

