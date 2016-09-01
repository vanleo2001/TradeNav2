VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDateJournals 
   Caption         =   "Date Journals"
   ClientHeight    =   5940
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgChecklist 
      Height          =   375
      Left            =   900
      TabIndex        =   7
      Top             =   5460
      Width           =   2895
      _cx             =   5106
      _cy             =   661
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
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
   Begin VSFlex7LCtl.VSFlexGrid fgPerformanceReport 
      Height          =   375
      Left            =   3900
      TabIndex        =   6
      Top             =   5460
      Width           =   2895
      _cx             =   5106
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
   Begin VB.Timer tmrMenu 
      Left            =   6900
      Top             =   5460
   End
   Begin gdOCX.gdScrollBar sbYear 
      Height          =   360
      Left            =   8280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin HexUniControls.ctlUniTextBoxXP txtYear 
      Height          =   315
      Left            =   7740
      TabIndex        =   1
      Top             =   90
      Width           =   540
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmDateJournals.frx":0000
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
      Tip             =   "frmDateJournals.frx":0028
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDateJournals.frx":0048
   End
   Begin vsOcx6LibCtl.vsIndexTab tabMonths 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   9128
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
      Caption         =   "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Order|Money Code"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
      FirstTab        =   0
      Style           =   1
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
      Begin HexUniControls.ctlUniFrameWL fraDateJournal 
         Height          =   4800
         Left            =   45
         TabIndex        =   3
         Top             =   330
         Width           =   8265
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
         Caption         =   "frmDateJournals.frx":0064
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDateJournals.frx":0090
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDateJournals.frx":00B0
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgJournal 
            Height          =   4515
            Left            =   2580
            TabIndex        =   5
            Top             =   180
            Width           =   4815
            _cx             =   8493
            _cy             =   7964
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
         Begin VSFlex7LCtl.VSFlexGrid fgDates 
            Height          =   4515
            Left            =   180
            TabIndex        =   4
            Top             =   180
            Width           =   2235
            _cx             =   3942
            _cy             =   7964
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
         Begin VSFlex7LCtl.VSFlexGrid fgSymbols 
            Height          =   4515
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   2235
            _cx             =   3942
            _cy             =   7964
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
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   7380
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "frmDateJournals.frx":00CC
      ToolBars        =   "frmDateJournals.frx":0158
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuCreateNew 
         Caption         =   "Create New Journal"
      End
      Begin VB.Menu mnuEditJournal 
         Caption         =   "Edit Journal"
      End
      Begin VB.Menu mnuDeleteJournal 
         Caption         =   "Delete Journal"
      End
      Begin VB.Menu mnuViewImage 
         Caption         =   "View Image"
      End
      Begin VB.Menu mnuInsertPerformanceReport 
         Caption         =   "Insert Performance Report"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintJournal 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmDateJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDateJournals.frm
'' Description: Form that displays the date journals
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/20/2011   DAJ         Changed icon
'' 09/22/2011   DAJ         Added the ability to show a chart image attached to a journal
'' 09/22/2011   DAJ         Added day of week to the dates grid
'' 09/23/2011   DAJ         Handle year change, bold dates with journals, auto insert order journal
'' 09/27/2011   DAJ         Added print capabilities, reports button, popup menu
'' 09/27/2011   DAJ         Don't allow selection in grids
'' 09/28/2011   DAJ         Handle user delete of order journal outside of this form
'' 09/28/2011   DAJ         Enhanced sort key on the journals grid
'' 09/29/2011   DAJ         Fix for rows getting goofed up after an edit of a journal
'' 09/30/2011   DAJ         Added code for grabbing a performance report image
'' 09/30/2011   DAJ         Added the performance report stuff to the context menu
'' 01/03/2012   DAJ         Only handle txtYear_Change if form is visible
'' 01/25/2012   DAJ         Money Code journal
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 03/13/2012   DAJ         New version of the Money Code journal
'' 03/19/2012   DAJ         Added Order Journal and Money Code Journal tabs
'' 03/20/2012   DAJ         Fixed "Object Variable..." error in InsertPerformanceReport
'' 02/20/2013   DAJ         Changed argument for frmTradeReportFilter.ShowMe
'' 07/30/2013   DAJ         Automatic Journal for a Fill
'' 08/08/2013   DAJ         Custom checklist journal
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/08/2014   DAJ         Use NavCore Image List; Use newer place/save form
'' 10/24/2014   DAJ         Core Application functions for DLL's
'' 05/18/2015   DAJ         Pass frmPrintPreview.vp to DoPrintHeader
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kNumClickHereRows As Long = 2

Private Enum eGDTabs
    eGDTabs_January = 0
    eGDTabs_February
    eGDTabs_March
    eGDTabs_April
    eGDTabs_May
    eGDTabs_June
    eGDTabs_July
    eGDTabs_August
    eGDTabs_September
    eGDTabs_October
    eGDTabs_November
    eGDTabs_December
    eGDTabs_Orders
    eGDTabs_MoneyCode
End Enum

Private Type mPrivate
    lSelectDay As Long                  ' Day to select in the dates grid
    dSelectedDate As Double             ' Selected date in the UI
    alJournalDates As cGdArray          ' Array of dates for which there is a journal
    lEditOrderJournalID As Long         ' Identifier of the order journal entry being edited
    lDeleteOrderJournalID As Long       ' Identifier of the order journal entry being deleted
    bHasMoneyCode As Boolean            ' Is the user enabled for the Money Code?

    Year As cPriceEditor                ' Editor for year
End Type
Private m As mPrivate

Private Property Get GDTab(ByVal nTab As eGDTabs) As Long
    GDTab = nTab
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    InitDatesGrid
    InitSymbolsGrid
    InitJournalGrid
    
    Set m.alJournalDates = g.JournalDB.GetJournalDates
    
    GoToDate Date

    If g.bAppIsIde = True Then
        mGenesis.ShowForm Me, eForm_Nonmodal
    Else
        g.TnCore.ShowForm Me, eForm_Nonmodal
        FixFormControls Me, ALT_GRID_ROW_COLOR
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the journals for the selected day
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "DateJournals", Me, 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.PrintMe"
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOrderJournal
'' Description: Update the given order journal
'' Inputs:      Order Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateOrderJournal(ByVal OrderJournal As cJournal)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid for the journal
    Dim lPos As Long                    ' Position for the date in the array

    If OrderJournal.JournalID <> m.lEditOrderJournalID Then
        If tabMonths.CurrTab <> GDTab(eGDTabs_MoneyCode) Then
            If (tabMonths.CurrTab = GDTab(eGDTabs_Orders)) Or (OrderJournal.JournalDate = m.dSelectedDate) Then
                With fgJournal
                    lRow = -1&
                    For lIndex = .FixedRows To .Rows - kNumClickHereRows
                        If TypeOf .RowData(lIndex) Is cJournal Then
                            If OrderJournal.JournalID = .RowData(lIndex).JournalID Then
                                lRow = lIndex
                                Exit For
                            End If
                        End If
                    Next lIndex
                    
                    AddOrderJournalEntry OrderJournal, , lRow
            
                    If m.alJournalDates.BinarySearch(CLng(OrderJournal.JournalDate), lPos) = False Then
                        m.alJournalDates.Add CLng(OrderJournal.JournalDate), lPos
                    End If
                    
                    SetDateBold
                End With
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.UpdateOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrderJournal
'' Description: Update the given order journal
'' Inputs:      Order Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteOrderJournal(ByVal OrderJournal As cJournal)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position for the date in the array

    If OrderJournal.JournalID <> m.lDeleteOrderJournalID Then
        If tabMonths.CurrTab <> GDTab(eGDTabs_MoneyCode) Then
            If (tabMonths.CurrTab = GDTab(eGDTabs_Orders)) Or (OrderJournal.JournalDate = m.dSelectedDate) Then
                With fgJournal
                    For lIndex = .FixedRows To .Rows - kNumClickHereRows
                        If TypeOf .RowData(lIndex) Is cJournal Then
                            If OrderJournal.JournalID = .RowData(lIndex).JournalID Then
                                .RemoveItem lIndex + 1
                                .RemoveItem lIndex
                                
                                Exit For
                            End If
                        End If
                    Next lIndex
                    
                    If fgJournal.Rows = fgJournal.FixedRows + 1 Then
                        If m.alJournalDates.BinarySearch(CLng(OrderJournal.JournalDate), lPos) = True Then
                            m.alJournalDates.Remove lPos
                        End If
                        
                        SetDateBold
                    End If
                End With
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.DeleteOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateDateJournal
'' Description: Update the given date journal
'' Inputs:      Date Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateDateJournal(ByVal journalEntry As cDateJournal)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position for the date in the array
    
    If journalEntry.JournalDate = m.dSelectedDate Then
        AddDateJournalEntry journalEntry
    
        If m.alJournalDates.BinarySearch(CLng(m.dSelectedDate), lPos) = False Then
            m.alJournalDates.Add CLng(m.dSelectedDate), lPos
        End If
        
        SetDateBold
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.UpdateDateJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the print preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim DateJournal As cDateJournal     ' Date journal entry
    Dim OrderJournal As cJournal        ' Order journal entry
    Dim ChartImage As IPictureDisp      ' Chart image
    Dim dClientWidth As Double          ' Client width
    Dim bExtraLineFeed As Boolean       ' Include an extra line feed?
    Dim dPicHeight As Double            ' Picture height
    Dim JournalImage As cJournalImage   ' Journal image
    
    With frmPrintPreview.vp
        .StartDoc
        g.TnCore.DoPrintHeader , frmPrintPreview.vp
        
        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
        .FontUnderline = False
        .TextAlign = taCenterMiddle
        If tabMonths.CurrTab = GDTab(eGDTabs_Orders) Then
            .Text = "Order Journal Entries"
        ElseIf tabMonths.CurrTab = GDTab(eGDTabs_MoneyCode) Then
            .Text = "Money Code Journal Entries"
        Else
            .Text = "Journal Entries for " & DateFormat(m.dSelectedDate, MM_DD_YYYY)
        End If
        .FontBold = False
        .FontSize = 12
        .TextAlign = taLeftMiddle
        
        .Text = vbLf & vbLf
        
        For lIndex = fgJournal.FixedRows To fgJournal.Rows - kNumClickHereRows
            If TypeOf fgJournal.RowData(lIndex) Is cDateJournal Then
                Set DateJournal = fgJournal.RowData(lIndex)

                .FontBold = True
                .Text = DateFormat(DateJournal.JournalTime, NO_DATE, HH_MM, AMPM_UPPER) & vbTab & fgJournal.TextMatrix(lIndex, 2) & vbLf
                .FontBold = False
                
                If DateJournal.JournalCategoryID > 0& Then
                    Select Case g.JournalCategories.TypeForId(DateJournal.JournalCategoryID)
                        Case eGDJournalCategoryType_Note
                            .Text = DateJournal.Text & vbLf & vbLf
                        
                        Case Else
                            LoadCheckListGrid DateJournal
                            .RenderControl = fgChecklist.hWnd
                            .Text = vbLf
                        
                    End Select
                Else
                    .Text = vbLf
                End If
                                
                Set JournalImage = DateJournal.JournalImage(eGDJournalImageType_Chart)
                If Not JournalImage Is Nothing Then
                    dClientWidth = .PageWidth - .MarginRight - .MarginLeft
                    Set ChartImage = LoadPicture(JournalImage.FileName)
                    
                    If ChartImage.Width >= dClientWidth Then
                        dPicHeight = (dClientWidth / ChartImage.Width) * ChartImage.Height
                        .DrawPicture ChartImage, .MarginLeft, .CurrentY, dClientWidth, dPicHeight
                    Else
                        .DrawPicture ChartImage, .MarginLeft, .CurrentY, ChartImage.Width
                    End If
                    .CurrentY = .CurrentY + ChartImage.Height
                    
                    .Text = vbLf
                End If
                
                Set JournalImage = DateJournal.JournalImage(eGDJournalImageType_SummaryReport)
                If Not JournalImage Is Nothing Then
                    fgPerformanceReport.LoadGrid JournalImage.FileName, flexFileAll
                    .RenderControl = fgPerformanceReport.hWnd
                    
                    .Text = vbLf
                End If
                
                lIndex = lIndex + 1
            ElseIf TypeOf fgJournal.RowData(lIndex) Is cJournal Then
                Set OrderJournal = fgJournal.RowData(lIndex)
                
                .FontBold = True
                .Text = DateFormat(OrderJournal.NoteDate, NO_DATE, HH_MM, AMPM_UPPER) & vbTab & fgJournal.TextMatrix(lIndex, 2) & vbLf
                .FontBold = False
                
                bExtraLineFeed = False
                If Len(OrderJournal.Action) > 0 Then
                    If UCase(OrderJournal.AccountID) = "REVERSAL" Then
                        .Text = "This order was a " & OrderJournal.Action & vbLf
                    Else
                        .Text = "This order was an " & OrderJournal.Action & vbLf
                    End If
                    bExtraLineFeed = True
                End If
                
                If OrderJournal.EmotionNumber >= 0 Then
                    .FontUnderline = True
                    .Text = "Emotions:"
                    .FontUnderline = False
                    .Text = " " & Str(OrderJournal.EmotionNumber) & vbLf
                    bExtraLineFeed = True
                End If
                
                If Len(OrderJournal.Feelings) > 0 Then
                    .FontUnderline = True
                    .Text = "Feelings:"
                    .FontUnderline = False
                    .Text = " " & OrderJournal.Feelings & vbLf
                    bExtraLineFeed = True
                End If
                
                If Len(OrderJournal.WhyTrade) > 0 Then
                    .FontUnderline = True
                    .Text = "Reasons:"
                    .FontUnderline = False
                    .Text = " " & OrderJournal.WhyTrade & vbLf
                    bExtraLineFeed = True
                End If
                
                If Len(OrderJournal.Thoughts) > 0 Then
                    .FontUnderline = True
                    .Text = "Thoughts:"
                    .FontUnderline = False
                    .Text = " " & OrderJournal.Thoughts & vbLf
                    bExtraLineFeed = True
                End If
                
                If Len(OrderJournal.Note) > 0 Then
                    .FontUnderline = True
                    .Text = "Notes:"
                    .FontUnderline = False
                    .Text = " " & OrderJournal.Note & vbLf
                    bExtraLineFeed = True
                End If
                
                If bExtraLineFeed = True Then
                    .Text = vbLf
                End If
                
                Set JournalImage = OrderJournal.JournalImage(eGDJournalImageType_Chart)
                If Not JournalImage Is Nothing Then
                    If Len(JournalImage.FileName) > 0 Then
                        dClientWidth = .PageWidth - .MarginRight - .MarginLeft
                        Set ChartImage = LoadPicture(JournalImage.FileName)
                        
                        If ChartImage.Width >= dClientWidth Then
                            dPicHeight = (dClientWidth / ChartImage.Width) * ChartImage.Height
                            .DrawPicture ChartImage, .MarginLeft, .CurrentY, dClientWidth, dPicHeight
                        Else
                            .DrawPicture ChartImage, .MarginLeft, .CurrentY
                        End If
                        .CurrentY = .CurrentY + ChartImage.Height
                        
                        .Text = vbLf
                    End If
                End If
                
                Set JournalImage = OrderJournal.JournalImage(eGDJournalImageType_OptionNavOrder)
                If Not JournalImage Is Nothing Then
                    If Len(JournalImage.FileName) > 0 Then
                        dClientWidth = .PageWidth - .MarginRight - .MarginLeft
                        Set ChartImage = LoadPicture(JournalImage.FileName)
                        
                        If ChartImage.Width >= dClientWidth Then
                            dPicHeight = (dClientWidth / ChartImage.Width) * ChartImage.Height
                            .DrawPicture ChartImage, .MarginLeft, .CurrentY, dClientWidth, dPicHeight
                        Else
                            .DrawPicture ChartImage, .MarginLeft, .CurrentY
                        End If
                        .CurrentY = .CurrentY + ChartImage.Height
                        
                        .Text = vbLf
                    End If
                End If
                
                lIndex = lIndex + 1
            End If
        Next lIndex
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDates_AfterRowColChange
'' Description: After a row change, load the journals for the new date
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDates_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If NewRow <> OldRow Then
        If (NewRow >= fgDates.FixedRows) And (NewRow < fgDates.Rows) Then
            LoadJournalForDate fgDates.RowData(NewRow)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.fgDates_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgJournal_BeforeMouseDown
'' Description: Bring up the popup menu on a right click by the user
'' Inputs:      Button, Shift/Ctrl/Alt Status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgJournal_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row that the user clicked in
    Dim lMouseCol As Long               ' Column that the user clicked in
    Dim bValidRow As Boolean            ' Is the current mouse row valid?
    
    lMouseRow = fgJournal.MouseRow
    lMouseCol = fgJournal.MouseCol

    If Button = vbRightButton Then
        bValidRow = JournalRowValid(lMouseRow, False)
        If bValidRow Then
            If Len(fgJournal.TextMatrix(lMouseRow, 1)) > 0 Then
                fgJournal.Row = lMouseRow
            Else
                fgJournal.Row = lMouseRow - 1
            End If
        End If
        
        mnuCreateNew.Enabled = True
        mnuEditJournal.Enabled = AllowEdit(lMouseRow)
        mnuDeleteJournal.Enabled = bValidRow
        mnuViewImage.Enabled = RowHasValidImage(lMouseRow)
        mnuInsertPerformanceReport.Enabled = True
        mnuPrintJournal.Enabled = True
        mnuChangeFont.Enabled = True
        
        PopupMenu mnuPopUp
    ElseIf Button = vbLeftButton Then
        bValidRow = JournalRowValid(lMouseRow, True)
        
        If bValidRow Then
            If fgJournal.MergeRow(lMouseRow) = False Then
                If lMouseCol = 3 Then
                    EditJournalEntry lMouseRow
                ElseIf lMouseCol = 4 Then
                    DeleteJournalEntry lMouseRow
                ElseIf lMouseCol = 5 Then
                    ShowImage lMouseRow
                End If
            ElseIf lMouseRow = fgJournal.Rows - 2 Then
                NewJournalEntry
            ElseIf lMouseRow = fgJournal.Rows - 1 Then
                InsertPerformanceReport
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.fgJournal_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgJournal_MouseMove
'' Description: Change the mouse pointer if over a click here spot
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Location of the Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgJournal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseRow As Long               ' Row that the user clicked in
    Dim lMouseCol As Long               ' Column that the user clicked in
    
    lMouseRow = fgJournal.MouseRow
    lMouseCol = fgJournal.MouseCol

    If JournalRowValid(lMouseRow) Then
        If fgJournal.Cell(flexcpForeColor, lMouseRow, lMouseCol) = vbBlue Or fgJournal.Cell(flexcpForeColor, lMouseRow, lMouseCol) = vbCyan Then
            MousePointer = vbCustom
            MouseIcon = g.CoreBridge.Picture16(g.TnCore.ToolbarIcon("kHand"))
        Else
            MousePointer = vbDefault
        End If
    Else
        MousePointer = vbDefault
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgSymbols_AfterRowColChange
'' Description: After a row change, load the journals for the correct tab
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgSymbols_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If NewRow <> OldRow Then
        If (NewRow >= fgDates.FixedRows) And (NewRow < fgDates.Rows) Then
            If tabMonths.CurrTab = GDTab(eGDTabs_Orders) Then
                LoadOrderJournalForSymbol fgSymbols.TextMatrix(NewRow, 0)
            ElseIf tabMonths.CurrTab = GDTab(eGDTabs_MoneyCode) Then
                LoadMoneyCodeJournalForSymbol fgSymbols.TextMatrix(NewRow, 0)
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.fgDates_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    tabMonths.BoldCurrent = True

    m.lSelectDay = -1&
    
    g.Styler.StyleForm Me
    
    PlaceTheForm Me, g.strIniFile
    
    Icon = g.CoreBridge.Picture16(g.TnCore.ToolbarIcon("kScroll"))

    Set m.Year = New cPriceEditor
    m.Year.Init sbYear, txtYear, Nothing, Year(g.TnCore.CurrentTime)
    txtYear.Locked = True
    
    Set m.alJournalDates = New cGdArray
    m.alJournalDates.Create eGDARRAY_Longs

    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    With tbToolbar
        .Tools("ID_Reports").Picture = g.AppBridge.ReportsPicture
        .Tools("ID_Print").Picture = g.CoreBridge.Picture16(g.TnCore.ToolbarIcon("kPrint"))
    End With
    
    fgChecklist.Visible = False
    fgPerformanceReport.Visible = False
    mnuPopUp.Visible = False
    
    m.bHasMoneyCode = g.TnCore.HasModule("LWMC")
    tabMonths.TabVisible(13) = m.bHasMoneyCode
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_MouseMove
'' Description: Change the mouse pointer back to default
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Location of the Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    MousePointer = vbDefault
    
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

    Dim lMinScaleWidth As Long          ' Minimum scale width for the form
    
    If m.bHasMoneyCode Then
        lMinScaleWidth = 9930
    Else
        lMinScaleWidth = 8625
    End If

    If LimitFormSize(Me, lMinScaleWidth, 5475) = False Then
        With tabMonths
            .Move .Left, .Top, ScaleWidth - (.Left * 2), ScaleHeight - (.Top * 2)
        End With
        
        With fgDates
            .Move .Left, .Top, .Width, tabMonths.ClientHeight - (.Top * 2)
            fgSymbols.Move .Left, .Top, .Width, .Height
        End With
        
        With fgJournal
            .Move .Left, .Top, tabMonths.ClientWidth - .Left - fgDates.Left, tabMonths.ClientHeight - (.Top * 2)
        End With
        
        With txtYear
            If m.bHasMoneyCode Then
                .Move 9060
            Else
                .Move 7740
            End If
        End With
        
        With sbYear
            If m.bHasMoneyCode Then
                .Move 9600
            Else
                .Move 8280
            End If
        End With
        
        AutoSizeJournalGrid
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings upon the form unloading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SaveTheFormPlacement Me, g.strIniFile

    Set m.Year = Nothing
    
    tmrMenu.Enabled = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change fonts in the journals grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    StartMenuTimer "CHANGEFONT"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCreateNew_Click
'' Description: Allow the user to create a new journal entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCreateNew_Click()
On Error GoTo ErrSection:

    StartMenuTimer "NEWJOURNAL"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuCreateNew_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeleteJournal_Click
'' Description: Allow the user to delete the selected journal
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeleteJournal_Click()
On Error GoTo ErrSection:

    StartMenuTimer "DELETEJOURNAL"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuDeleteJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditJournal_Click
'' Description: Allow the user to edit the selected journal
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditJournal_Click()
On Error GoTo ErrSection:

    StartMenuTimer "EDITJOURNAL"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuEditJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuInsertPerformanceReport_Click
'' Description: Allow the user to insert a performance report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuInsertPerformanceReport_Click()
On Error GoTo ErrSection:

    StartMenuTimer "INSERTREPORT"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuInsertPerformanceReport_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrintJournal_Click
'' Description: Allow the user to print the journals for the selected day
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrintJournal_Click()
On Error GoTo ErrSection:

    StartMenuTimer "PRINT"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuPrintJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuViewImage_Click
'' Description: Allow the user to view the image for the selected journal
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuViewImage_Click()
On Error GoTo ErrSection:

    StartMenuTimer "VIEWIMAGE"

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.mnuViewImage_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabMonths_Switch
'' Description: Handle a tab switch
'' Inputs:      Old Tab, New Tab, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabMonths_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    Dim lColor As Long
    
    If GetAppBackColor = kDarkThemeColor Then
        lColor = GetAppBackColor
    Else
        lColor = vbWhite
    End If

    fgDates.BackColor = lColor
    fgDates.BackColorAlternate = ALT_GRID_ROW_COLOR
    fgSymbols.BackColor = lColor
    fgSymbols.BackColorAlternate = ALT_GRID_ROW_COLOR
    fgJournal.BackColor = ALT_GRID_ROW_COLOR
    fgJournal.BackColorAlternate = lColor
    
    Select Case NewTab
        Case 12, 13:
            fgDates.Visible = False
            fgSymbols.Visible = True
            LoadSymbolsGrid NewTab
        
        Case Else:
            fgDates.Visible = True
            fgSymbols.Visible = False
            LoadDatesGrid NewTab + 1
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.tabMonths_Switch"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle a user selection on the toolbar
'' Inputs:      Tool
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Select Case UCase(Tool.ID)
        Case "ID_REPORTS"
            StartMenuTimer "REPORTS"
            
        Case "ID_PRINT"
            StartMenuTimer "PRINT"
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Perform a specified action when the timer goes off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Action to perform

    strTag = tmrMenu.Tag
    tmrMenu.Tag = ""
    tmrMenu.Enabled = False
    
    Select Case UCase(strTag)
        Case "CHANGEFONT"
            g.TnCore.ChangeGridFont fgJournal
            
        Case "DELETEJOURNAL"
            DeleteJournalEntry fgJournal.Row
        
        Case "EDITJOURNAL"
            EditJournalEntry fgJournal.Row
            
        Case "INSERTREPORT"
            InsertPerformanceReport
    
        Case "NEWJOURNAL"
            NewJournalEntry
            
        Case "PRINT"
            PrintMe
            
        Case "REPORTS"
            g.AppBridge.ShowTradeFilter
            
        Case "VIEWIMAGE"
            ShowImage fgJournal.Row
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtYear_Change
'' Description: Handle the user changing the date
'' Inputs:      Old Tab, New Tab, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtYear_Change()
On Error GoTo ErrSection:

    If Visible Then
        If m.Year.Price = Year(Date) Then
            GoToDate Date
        Else
            GoToDate JulFromLong((CLng(m.Year.Price) * 10000) + 101)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.txtYear_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitDatesGrid
'' Description: Initialize the dates grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitDatesGrid()
On Error GoTo ErrSection:

    With fgDates
        .Redraw = flexRDNone
        
        SetupGrid fgDates, eGridMode_List
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColor = vbWhite
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 1
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.InitDatesGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDatesGrid
'' Description: Load the dates grid based on the current tab
'' Inputs:      Month
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadDatesGrid(ByVal lMonth As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw setting for the grid
    Dim iMonth As Integer               ' Month
    Dim lYear As Long                   ' Year
    Dim lIndex As Long                  ' Index into a for loop
    Dim lDate As Long                   ' Date
    Dim lMaxDay As Long                 ' Max day for the month
    Dim lRow As Long                    ' Previously selected row

    With fgDates
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = .Row
        
        .Rows = .FixedRows
        
        lYear = CLng(Val(txtYear.Text))
        
        Select Case lMonth
            Case 1, 3, 5, 7, 8, 10, 12
                lMaxDay = 31
            Case 4, 6, 9, 11
                lMaxDay = 30
            Case 2
                If IsLeapYear(lYear) Then
                    lMaxDay = 29
                Else
                    lMaxDay = 28
                End If
        End Select
        
        For lIndex = 1 To lMaxDay
            lDate = JulFromLong((lYear * 10000) + (lMonth * 100) + lIndex)
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = DateFormat(lDate, MM_DD_YYYY) & " " & Format(lDate, "ddd")
            .RowData(.Rows - 1) = lDate
            
            SetDateBold lDate, lMonth
        Next lIndex
        
        If m.lSelectDay > -1& Then
            .Row = m.lSelectDay
            m.lSelectDay = -1&
            
            .ShowCell .Row, 0
        Else
            .Row = .FixedRows
        End If
        
        If (lRow = .Row) And (lRow >= .FixedRows) And (lRow < .Rows) Then
            LoadJournalForDate .RowData(lRow)
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadDatesGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitJournalGrid
'' Description: Initialize the journal grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitJournalGrid()
On Error GoTo ErrSection:

    With fgJournal
        .Redraw = flexRDNone
        
        SetupGrid fgDates, eGridMode_List
        
        .AllowSelection = False
        .AllowBigSelection = False
        .ExtendLastCol = True
        .BackColor = ALT_GRID_ROW_COLOR
        .BackColorAlternate = vbWhite
        .BackColorBkg = vbWhite
        .GridLines = flexGridNone
        .MergeCells = flexMergeFree
        .SheetBorder = vbWhite
        .WordWrap = True
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 6
        .Rows = 0
        
        .ColHidden(0) = True
        
        AddClickHereRows
                
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.InitJournalGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereRows
'' Description: Add the "Click Here" rows to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereRows()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim lColor As Long
    
    If GetAppBackColor = kDarkThemeColor Then
        lColor = vbCyan
    Else
        lColor = vbBlue
    End If
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "999998"
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Click here to add a new journal entry"
        .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = lColor
        .Cell(flexcpFontUnderline, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .MergeRow(.Rows - 1) = True
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "999999"
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Click here to add a performance summary"
        .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = lColor
        .Cell(flexcpFontUnderline, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .MergeRow(.Rows - 1) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.AddClickHereRows"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortJournalGrid
'' Description: Sort the journal grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortJournalGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Col = 0
        .Sort = flexSortStringAscending
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.SortJournalGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddDateJournalEntry
'' Description: Add a date journal entry to the grid
'' Inputs:      Journal entry, Auto size grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddDateJournalEntry(ByVal journalEntry As cDateJournal, Optional ByVal bAutoSize As Boolean = True, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim lRow2 As Long                   ' Second row of the journal entry
    Dim JournalImage As cJournalImage   ' Journal image information
    Dim astrChecklist As cGdArray       ' Check list broken out into an array
    Dim moneyCodeFields As cGdTree      ' Dictionary of Money Code fields
    Dim customChecklistFields As cGdTree ' Dictionary of Custom Checklist fields
    Dim lColor As Long
    
    If GetAppBackColor = kDarkThemeColor Then
        lColor = vbCyan
    Else
        lColor = vbBlue
    End If
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
            
        If lRow = -1& Then
            .Rows = .Rows + 2
            lRow = .Rows - 2
            lRow2 = .Rows - 1
        Else
            lRow2 = lRow + 1
        End If
        
        .RowData(lRow) = journalEntry
        .TextMatrix(lRow, 0) = Str(journalEntry.JournalDate + journalEntry.JournalTime) + "_D" + Str(journalEntry.DateJournalID) + "_1"
        .TextMatrix(lRow, 1) = DateFormat(journalEntry.JournalTime, NO_DATE, HH_MM, AMPM_UPPER)
        If g.JournalCategories.TypeForId(journalEntry.JournalCategoryID) = eGDJournalCategoryType_CustomChecklist Then
            .TextMatrix(lRow, 2) = JournalCategoryName(journalEntry.JournalCategoryID) & " for " & journalEntry.Symbol
        Else
            .TextMatrix(lRow, 2) = JournalCategoryName(journalEntry.JournalCategoryID)
        End If
        .TextMatrix(lRow, 3) = "Edit"
        .TextMatrix(lRow, 4) = "Delete"
        .TextMatrix(lRow, 5) = "Image"
        
        If journalEntry.JournalCategoryID > 0 Then
            .Cell(flexcpForeColor, lRow, 3) = lColor
        Else
            .Cell(flexcpForeColor, lRow, 3) = RGB(128, 128, 128)
        End If
        .Cell(flexcpForeColor, lRow, 4) = lColor
        If journalEntry.HasValidJournalImages Then
            .Cell(flexcpForeColor, lRow, 5) = lColor
        Else
            .Cell(flexcpForeColor, lRow, 5) = RGB(128, 128, 128)
        End If
        
        .Cell(flexcpFontUnderline, lRow, 3, lRow, 5) = True
        .MergeRow(lRow) = False
        
        .RowData(lRow2) = journalEntry
        .TextMatrix(lRow2, 0) = Str(journalEntry.JournalDate + journalEntry.JournalTime) + "_D" + Str(journalEntry.DateJournalID) + "_2"
        
        Select Case g.JournalCategories.TypeForId(journalEntry.JournalCategoryID)
            Case eGDJournalCategoryType_Note
                .Cell(flexcpText, lRow2, 2, lRow2, 2) = Replace(journalEntry.Text, vbCrLf, vbLf)
                
            Case eGDJournalCategoryType_MoneyCode
                If UCase(Left(journalEntry.Text, 8)) <> "VERSION=" Then
                    Set astrChecklist = New cGdArray
                    astrChecklist.SplitFields journalEntry.Text, "|"
                    .Cell(flexcpText, lRow2, 2, lRow2, 2) = Replace(Parse(astrChecklist(9), ";", 2), vbCrLf, vbLf)
                Else
                    Set moneyCodeFields = New cGdTree
                    moneyCodeFields.FromKeyValueString journalEntry.Text, "|", "="
                    
                    If moneyCodeFields.Exists("Version") Then
                        Select Case moneyCodeFields("Version")
                            Case "1"
                                If moneyCodeFields.Exists("Notes") Then
                                    .Cell(flexcpText, lRow2, 2, lRow2, 2) = moneyCodeFields("Notes")
                                End If
                                
                        End Select
                    End If
                End If
            
            Case eGDJournalCategoryType_CustomChecklist
                Set customChecklistFields = New cGdTree
                customChecklistFields.FromKeyValueString journalEntry.Text, "|", "="
            
                If customChecklistFields.Exists("Version") Then
                    Select Case customChecklistFields("Version")
                        Case "1"
                            If customChecklistFields.Exists("Notes") Then
                                .Cell(flexcpText, lRow2, 2, lRow2, 2) = customChecklistFields("Notes")
                            End If
                            
                    End Select
                End If
        
        End Select
        
        .MergeRow(lRow2) = True
        
        If bAutoSize Then
            SortJournalGrid
            SetBackColors fgJournal
            AutoSizeJournalGrid
        End If
        
        If fgDates.Cell(flexcpFontBold, fgDates.Row, 0) = False Then
            fgDates.Cell(flexcpFontBold, fgDates.Row, 0) = True
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.AddDateJournalEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrderJournalEntry
'' Description: Add an order journal entry to the grid
'' Inputs:      Journal entry, Auto size grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddOrderJournalEntry(ByVal journalEntry As cJournal, Optional ByVal bAutoSize As Boolean = True, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim lRow2 As Long                   ' Second row of the journal entry
    Dim lColor As Long
    
    If GetAppBackColor = kDarkThemeColor Then
        lColor = vbCyan
    Else
        lColor = vbBlue
    End If
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
            
        If lRow = -1& Then
            .Rows = .Rows + 2
            lRow = .Rows - 2
            lRow2 = .Rows - 1
        Else
            lRow2 = lRow + 1
        End If
        
        .RowData(lRow) = journalEntry
        .TextMatrix(lRow, 0) = Str(journalEntry.NoteDate) + "_O" + Str(journalEntry.JournalID) + "_1"
        .TextMatrix(lRow, 1) = DateFormat(journalEntry.NoteDate, NO_DATE, HH_MM, AMPM_UPPER)
        .TextMatrix(lRow, 2) = g.AppBridge.OrderTextForId(journalEntry.OrderID, True, True)
        .TextMatrix(lRow, 3) = "Edit"
        .TextMatrix(lRow, 4) = "Delete"
        .TextMatrix(lRow, 5) = "Image"
        
        .Cell(flexcpForeColor, lRow, 3, lRow, 4) = lColor
        If journalEntry.HasValidJournalImages Then
            .Cell(flexcpForeColor, lRow, 5) = lColor
        Else
            .Cell(flexcpForeColor, lRow, 5) = RGB(128, 128, 128)
        End If
        
        .Cell(flexcpFontUnderline, lRow, 3, lRow, 5) = True
        .MergeRow(lRow) = False
        
        .RowData(lRow2) = journalEntry
        .TextMatrix(lRow2, 0) = Str(journalEntry.NoteDate) + "_O" + Str(journalEntry.JournalID) + "_2"
        .Cell(flexcpText, lRow2, 2, lRow2, 2) = journalEntry.DisplayString
        .MergeRow(lRow2) = True
        
        If bAutoSize Then
            SortJournalGrid
            SetBackColors fgJournal
            AutoSizeJournalGrid
        End If
        
        If fgDates.Cell(flexcpFontBold, fgDates.Row, 0) = False Then
            fgDates.Cell(flexcpFontBold, fgDates.Row, 0) = True
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.AddOrderJournalEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lTotal As Long                  ' Total of the column widths
    Dim lExtCol As Long                 ' Column to extend
    
    lExtCol = 2
    
    With fgJournal
        ' size the custom extended column in order to fill the client width
        .ColHidden(lExtCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        
        lTotal = 0
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal <> .ColWidth(lExtCol) Then
            .ColWidth(lExtCol) = lTotal
        End If
        .ColHidden(lExtCol) = False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
       
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadJournalForDate
'' Description: Load the journal entries for the given date
'' Inputs:      Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadJournalForDate(ByVal lDate As Long)
On Error GoTo ErrSection:

    Dim DateJournals As cDateJournals   ' Date journals for the given date
    Dim OrderJournals As cJournals      ' Order journals for the given date
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    Set DateJournals = New cDateJournals
    g.JournalDB.LoadDateJournalsForDate DateJournals, lDate
    
    Set OrderJournals = New cJournals
    g.JournalDB.LoadOrderJournalsForDate OrderJournals, lDate
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To DateJournals.Count
            AddDateJournalEntry DateJournals(lIndex), False
        Next lIndex
        For lIndex = 1 To OrderJournals.Count
            AddOrderJournalEntry OrderJournals(lIndex), False
        Next lIndex
        
        AddClickHereRows
        
        SortJournalGrid
        SetBackColors fgJournal
        AutoSizeJournalGrid
        
        m.dSelectedDate = CDbl(lDate)
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadJournalsForDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GoToDate
'' Description: Go to the given date
'' Inputs:      Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GoToDate(ByVal lDate As Long)
On Error GoTo ErrSection:

    m.lSelectDay = Day(lDate) - 1
    txtYear.Text = Str(Year(lDate))
    If tabMonths.CurrTab = Month(lDate) - 1 Then
        LoadDatesGrid Month(lDate)
    Else
        tabMonths.CurrTab = Month(lDate) - 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.GoToDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewJournalEntry
'' Description: Create a new journal entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewJournalEntry()
On Error GoTo ErrSection:

    Dim journalEntry As cDateJournal    ' Journal entry

    Set journalEntry = New cDateJournal
    If frmDateJournal.ShowMe(journalEntry, m.dSelectedDate) = True Then
        g.JournalDB.SaveDateJournal journalEntry
        UpdateDateJournal journalEntry
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.NewJournalEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditJournalEntry
'' Description: Allow the user to edit a journal entry
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditJournalEntry(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim DateJournal As cDateJournal     ' Journal entry to edit
    Dim OrderJournal As cJournal        ' Journal entry to edit
    
    If JournalRowValid(lRow, False) Then
        If TypeOf fgJournal.RowData(lRow) Is cDateJournal Then
            Set DateJournal = fgJournal.RowData(lRow)
            If DateJournal.JournalCategoryID > 0 Then
                If frmDateJournal.ShowMe(DateJournal, DateJournal.JournalDate) Then
                    g.JournalDB.SaveDateJournal DateJournal
                    If (tabMonths.CurrTab = GDTab(eGDTabs_MoneyCode)) Or (DateJournal.JournalDate = m.dSelectedDate) Then
                        AddDateJournalEntry DateJournal, , lRow
                    Else
                        fgJournal.RemoveItem lRow + 1
                        fgJournal.RemoveItem lRow
                    End If
                End If
            End If
        ElseIf TypeOf fgJournal.RowData(lRow) Is cJournal Then
            Set OrderJournal = fgJournal.RowData(lRow)
            m.lEditOrderJournalID = OrderJournal.JournalID
            If frmJournal.ShowMeForOrderID(OrderJournal.OrderID, OrderJournal.JournalID, OrderJournal) Then
                g.JournalDB.LoadOrderJournal OrderJournal.JournalID, OrderJournal
                If (tabMonths.CurrTab = GDTab(eGDTabs_Orders)) Or (OrderJournal.JournalDate = m.dSelectedDate) Then
                    AddOrderJournalEntry OrderJournal, , lRow
                Else
                    fgJournal.RemoveItem lRow + 1
                    fgJournal.RemoveItem lRow
                End If
            End If
            m.lEditOrderJournalID = 0&
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.EditJournalEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteJournalEntry
'' Description: Allow the user to delete a journal entry
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteJournalEntry(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim DateJournal As cDateJournal     ' Journal entry to edit
    Dim OrderJournal As cJournal        ' Journal entry to edit
    Dim lPos As Long                    ' Position of the date in the array
    Dim dJournalDate As Double          ' Journal date
    
    If JournalRowValid(lRow, False) Then
        If TypeOf fgJournal.RowData(lRow) Is cDateJournal Then
            Set DateJournal = fgJournal.RowData(lRow)
            If InfBox("Are you sure you want to delete this journal entry?", "?", "+Yes|-No", "Delete Journal Entry") = "Y" Then
                g.JournalDB.DeleteDateJournal DateJournal
                
                dJournalDate = DateJournal.JournalDate
                fgJournal.RemoveItem lRow + 1
                fgJournal.RemoveItem lRow
            End If
        ElseIf TypeOf fgJournal.RowData(lRow) Is cJournal Then
            Set OrderJournal = fgJournal.RowData(lRow)
            If InfBox("Are you sure you want to delete this journal entry?", "?", "+Yes|-No", "Delete Journal Entry") = "Y" Then
                m.lDeleteOrderJournalID = OrderJournal.JournalID
                g.JournalDB.DeleteOrderJournal OrderJournal
                
                dJournalDate = OrderJournal.JournalDate
                fgJournal.RemoveItem lRow + 1
                fgJournal.RemoveItem lRow
                m.lDeleteOrderJournalID = 0&
            End If
        End If
        
        If fgJournal.Rows = fgJournal.FixedRows + 1 Then
            If m.alJournalDates.BinarySearch(CLng(dJournalDate), lPos) = True Then
                m.alJournalDates.Remove lPos
            End If
            
            SetDateBold
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.DeleteJournalEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowImage
'' Description: Show the appropriate image
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowImage(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim DateJournal As cDateJournal     ' Journal entry to edit
    Dim OrderJournal As cJournal        ' Journal entry to edit
    Dim lIndex As Long                  ' Index into a for loop
    Dim JournalImage As cJournalImage   ' Journal image
    Dim JournalImages As cGdTree        ' Collection of journal images
    Dim frm As frmDisplayImage          ' Display image form

    If JournalRowValid(lRow, False) Then
        If TypeOf fgJournal.RowData(lRow) Is cDateJournal Then
            Set DateJournal = fgJournal.RowData(lRow)
            Set JournalImages = DateJournal.JournalImages
        ElseIf TypeOf fgJournal.RowData(lRow) Is cJournal Then
            Set OrderJournal = fgJournal.RowData(lRow)
            Set JournalImages = OrderJournal.JournalImages
        End If
        
        For lIndex = 1 To JournalImages.Count
            Set JournalImage = JournalImages(lIndex)
            If Len(JournalImage.FileName) > 0 Then
                If JournalImage.ImageType = eGDJournalImageType_SummaryReport Then
                    frmShowSavedGrid.ShowMe JournalImage.Caption, JournalImage.FileName
                Else
                    Set frm = New frmDisplayImage
                    frm.ShowMe JournalImage.Caption, JournalImage.FileName
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.ShowImage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoSizeJournalGrid
'' Description: Auto size the journal grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AutoSizeJournalGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        ExtendCustomColumn
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, .Cols - 1, False
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.AutoSizeJournalGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowHasValidImage
'' Description: Does the given row have a valid image?
'' Inputs:      Row
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowHasValidImage(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim DateJournal As cDateJournal     ' Journal entry to edit
    Dim OrderJournal As cJournal        ' Journal entry to edit
    
    bReturn = False
    If JournalRowValid(lRow, False) Then
        If TypeOf fgJournal.RowData(lRow) Is cDateJournal Then
            Set DateJournal = fgJournal.RowData(lRow)
            bReturn = DateJournal.HasValidJournalImages
        ElseIf TypeOf fgJournal.RowData(lRow) Is cJournal Then
            Set OrderJournal = fgJournal.RowData(lRow)
            bReturn = OrderJournal.HasValidJournalImages
        End If
    End If
    
    RowHasValidImage = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.RowHasValidImage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDateBold
'' Description: Set the given date bold or not based on if it has journals
'' Inputs:      Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetDateBold(Optional ByVal lDate As Long = kNullData, Optional ByVal lMonth As Long = kNullData)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current setting of the redraw on the grid
    
    With fgDates
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lDate = kNullData Then
            lDate = CLng(m.dSelectedDate)
        End If
        If lMonth = kNullData Then
            lMonth = tabMonths.CurrTab + 1
        End If
        
        If Year(lDate) = m.Year.Price Then
            If Month(lDate) = lMonth Then
                .Cell(flexcpFontBold, Day(lDate) - 1, 0) = m.alJournalDates.BinarySearch(lDate)
            End If
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.SetDateBold"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartMenuTimer
'' Description: Start the menu timer with the given command
'' Inputs:      Command
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartMenuTimer(ByVal strCommand As String)
On Error GoTo ErrSection:

    tmrMenu.Tag = strCommand
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalRowValid
'' Description: Determine if the given row is valid
'' Inputs:      Row, Include "Click Here" Rows as Valid?
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function JournalRowValid(ByVal lRow As Long, Optional ByVal bIncludeClickHereRowsAsValid As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    If (bIncludeClickHereRowsAsValid = True) Or (tabMonths.CurrTab >= GDTab(eGDTabs_Orders)) Then
        bReturn = ((lRow >= fgJournal.FixedRows) And (lRow < fgJournal.Rows))
    Else
        bReturn = ((lRow >= fgJournal.FixedRows) And (lRow < fgJournal.Rows - kNumClickHereRows))
    End If
    
    JournalRowValid = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.JournalRowValid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InsertPerformanceReport
'' Description: Insert a performance report into the journals
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InsertPerformanceReport()
On Error GoTo ErrSection:

    Dim strCaptureFile As String        ' Filename of the capture file
    Dim DateJournal As cDateJournal     ' Date journal entry to hold report information
    Dim dCurrentTime As Double          ' Current time
    Dim JournalImage As cJournalImage   ' Journal image object
    
    If g.AppBridge.GetPerformanceReport(strCaptureFile, dCurrentTime) = True Then
        If FileExist(strCaptureFile) Then
            Set DateJournal = New cDateJournal
            
            With DateJournal
                .JournalCategoryID = -1&
                .JournalDate = CDbl(Int(dCurrentTime))
                .JournalTime = dCurrentTime - .JournalDate
                .Text = "Performance Report captured at " & Format(dCurrentTime, "YYYY-MM-DD HH:MM:SS")
                
                Set JournalImage = New cJournalImage
                JournalImage.ImageType = eGDJournalImageType_SummaryReport
                JournalImage.Caption = "Performance Report " & Format(dCurrentTime, "YYYY-MM-DD HH:MM:SS")
                JournalImage.FileName = strCaptureFile
                .JournalImage(eGDJournalImageType_SummaryReport) = JournalImage
            End With
            g.JournalDB.SaveDateJournal DateJournal
            
            AddDateJournalEntry DateJournal
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.InsertPerformanceReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalCategoryName
'' Description: Return the category name for the given category ID
'' Inputs:      Category ID
'' Returns:     Category Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function JournalCategoryName(ByVal lJournalCategoryID As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If lJournalCategoryID = -1& Then
        strReturn = "Performance Report"
    ElseIf g.JournalCategories.Exists(Str(lJournalCategoryID)) Then
        strReturn = g.JournalCategories(Str(lJournalCategoryID)).Text
    End If
    
    JournalCategoryName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.JournalCategoryName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllowEdit
'' Description: Determine if the user can edit the journal on the given row
'' Inputs:      Row
'' Returns:     True if can edit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllowEdit(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim DateJournal As cDateJournal     ' Journal entry to edit
    
    bReturn = False
    If JournalRowValid(lRow, False) Then
        If TypeOf fgJournal.RowData(lRow) Is cDateJournal Then
            Set DateJournal = fgJournal.RowData(lRow)
            bReturn = (DateJournal.JournalCategoryID > 0)
        ElseIf TypeOf fgJournal.RowData(lRow) Is cJournal Then
            bReturn = True
        End If
    End If
    
    AllowEdit = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.AllowEdit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCheckListGrid
'' Description: Initialize and load the check list grid
'' Inputs:      Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCheckListGrid(DateJournal As cDateJournal)
On Error GoTo ErrSection:

    Dim checklistFields As cGdTree      ' Dictionary of fields from the journal text

    Select Case g.JournalCategories.TypeForId(DateJournal.JournalCategoryID)
        Case eGDJournalCategoryType_MoneyCode
            If UCase(Left(DateJournal.Text, 8)) <> "VERSION=" Then
                LoadMoneyCodeCheckListGrid0 DateJournal.Text
            Else
                Set checklistFields = New cGdTree
                checklistFields.FromKeyValueString DateJournal.Text, "|", "="
                
                If checklistFields.Exists("Version") Then
                    Select Case checklistFields("Version")
                        Case "1"
                            LoadMoneyCodeCheckListGrid1 checklistFields
                        
                    End Select
                End If
            End If
            
        Case eGDJournalCategoryType_CustomChecklist
            Set checklistFields = New cGdTree
            checklistFields.FromKeyValueString DateJournal.Text, "|", "="
            
            If checklistFields.Exists("Version") Then
                Select Case checklistFields("Version")
                    Case "1"
                        LoadCustomCheckListGrid1 checklistFields
                    
                End Select
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadCheckListGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMoneyCodeCheckListGrid0
'' Description: Initialize and load the check list grid
'' Inputs:      Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMoneyCodeCheckListGrid0(ByVal strMoneyCodeFields As String)
On Error GoTo ErrSection:

    Dim astrChecklist As cGdArray       ' Check list broken out into an array
    Dim lIndex As Long                  ' Index into a for loop

    Set astrChecklist = New cGdArray
    astrChecklist.SplitFields strMoneyCodeFields, "|"

    With fgChecklist
        .Redraw = flexRDNone
        
        .Clear
        
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        .HighLight = flexHighlightNever
        .SelectionMode = flexSelectionFree
        .SheetBorder = .BackColor
        
        .Cols = 2
        .Rows = 20
        .FixedCols = 0
        .FixedRows = 0
        
        .TextMatrix(0, 1) = "Download Data, Recalculate Criteria, and Run Filters"
        .TextMatrix(2, 1) = "Analyze COT Commercial Conditional Set Ups"
        .TextMatrix(4, 1) = "Analyze Larger Traders/Advisor Sentiment"
        .TextMatrix(6, 1) = "Analyze Premium of Agricultural Futures"
        .TextMatrix(8, 1) = "Analyze Seasonal Trend"
        .TextMatrix(10, 1) = "Analyze Open Interest"
        .TextMatrix(12, 1) = "Analyze Accumulation and Distribution"
        .TextMatrix(14, 1) = "Analyze Williams %R with Trend"
        .TextMatrix(16, 1) = "Analyze Technical Entry Triggers and Exit Triggers"
        .TextMatrix(18, 1) = "Notes"
        
        For lIndex = 0 To 9
            CheckedCell(fgChecklist, lIndex * 2, 0) = (Val(Parse(astrChecklist(lIndex), ";", 1)) <> 0)
            .TextMatrix((lIndex * 2) + 1, 1) = Parse(astrChecklist(lIndex), ";", 2)
        Next lIndex
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, 1, False
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, 1, False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadMoneyCodeCheckListGrid0"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMoneyCodeCheckListGrid1
'' Description: Initialize and load the check list grid
'' Inputs:      Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMoneyCodeCheckListGrid1(ByVal moneyCodeFields As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With fgChecklist
        .Redraw = flexRDNone
        
        .Clear
        
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        .MergeCells = flexMergeFree
        .HighLight = flexHighlightNever
        .SelectionMode = flexSelectionFree
        .SheetBorder = .BackColor
        
        .Cols = 3
        .Rows = 15
        .FixedCols = 0
        .FixedRows = 0
        
        .TextMatrix(0, 0) = "Symbol"
        If moneyCodeFields.Exists("Symbol") Then
            .Cell(flexcpText, 0, 1, 0, 2) = moneyCodeFields("Symbol")
        Else
            .Cell(flexcpText, 0, 1, 0, 2) = ""
        End If
        .MergeRow(0) = True
        
        .TextMatrix(1, 0) = "COT Commercial Conditional Set Ups"
        If moneyCodeFields.Exists("Cot") Then
            .Cell(flexcpText, 1, 1, 1, 2) = moneyCodeFields("Cot")
        Else
            .Cell(flexcpText, 1, 1, 1, 2) = ""
        End If
        .MergeRow(1) = True
        
        .TextMatrix(2, 0) = "Large Traders Sentiment"
        If moneyCodeFields.Exists("LargeTraders") Then
            .Cell(flexcpText, 2, 1, 2, 2) = moneyCodeFields("LargeTraders")
        Else
            .Cell(flexcpText, 2, 1, 2, 2) = ""
        End If
        .MergeRow(2) = True
        
        .TextMatrix(3, 0) = "Advisor Sentiment"
        If moneyCodeFields.Exists("Advisor") Then
            .Cell(flexcpText, 3, 1, 3, 2) = moneyCodeFields("Advisor")
        Else
            .Cell(flexcpText, 3, 1, 3, 2) = ""
        End If
        .MergeRow(3) = True
        
        .TextMatrix(4, 0) = "Premium of Agricultural Futures"
        If moneyCodeFields.Exists("Agricultural") Then
            .Cell(flexcpText, 4, 1, 4, 2) = moneyCodeFields("Agricultural")
        Else
            .Cell(flexcpText, 4, 1, 4, 2) = ""
        End If
        .MergeRow(4) = True
        
        .TextMatrix(5, 0) = "Seasonal Trend Conditional Setup"
        If moneyCodeFields.Exists("SeasonalSetup") Then
            .Cell(flexcpText, 5, 1, 5, 2) = moneyCodeFields("SeasonalSetup")
        Else
            .Cell(flexcpText, 5, 1, 5, 2) = ""
        End If
        .MergeRow(5) = True
        
        .TextMatrix(6, 0) = "Seasonal Trend Direction"
        If moneyCodeFields.Exists("SeasonalDirection") Then
            .Cell(flexcpText, 6, 1, 6, 2) = moneyCodeFields("SeasonalDirection")
        Else
            .Cell(flexcpText, 6, 1, 6, 2) = ""
        End If
        .MergeRow(6) = True
        
        .TextMatrix(7, 0) = "Open Interest Conditional Setup"
        If moneyCodeFields.Exists("OpenInterest") Then
            .Cell(flexcpText, 7, 1, 7, 2) = moneyCodeFields("OpenInterest")
        Else
            .Cell(flexcpText, 7, 1, 7, 2) = ""
        End If
        .MergeRow(7) = True
        
        .TextMatrix(8, 0) = "Accumulation and Distribution"
        If moneyCodeFields.Exists("Accumulation") Then
            .Cell(flexcpText, 8, 1, 8, 2) = moneyCodeFields("Accumulation")
        Else
            .Cell(flexcpText, 8, 1, 8, 2) = ""
        End If
        .MergeRow(8) = True
        
        .TextMatrix(9, 0) = "Williams %R with Trend"
        If moneyCodeFields.Exists("PercentR") Then
            .Cell(flexcpText, 9, 1, 9, 2) = moneyCodeFields("PercentR")
        Else
            .Cell(flexcpText, 9, 1, 9, 2) = ""
        End If
        .MergeRow(9) = True
        
        .TextMatrix(10, 0) = "Overall Trend"
        If moneyCodeFields.Exists("Overall") Then
            .Cell(flexcpText, 10, 1, 10, 2) = moneyCodeFields("Overall")
        Else
            .Cell(flexcpText, 10, 1, 10, 2) = ""
        End If
        .MergeRow(10) = True
        
        .TextMatrix(11, 0) = "Technical Entry Triggers"
        If moneyCodeFields.Exists("EntryDirection") Then
            .TextMatrix(11, 1) = moneyCodeFields("EntryDirection")
        End If
        If moneyCodeFields.Exists("EntryTrigger") Then
            .TextMatrix(11, 2) = moneyCodeFields("EntryTrigger")
        End If
        .MergeRow(11) = False
        
        .TextMatrix(12, 0) = "Exit Triggers"
        If moneyCodeFields.Exists("ExitTrigger") Then
            .Cell(flexcpText, 12, 1, 12, 2) = moneyCodeFields("ExitTrigger")
        Else
            .Cell(flexcpText, 12, 1, 12, 2) = ""
        End If
        .MergeRow(12) = True
        
        .Cell(flexcpText, 13, 0, 13, 2) = "Notes"
        .MergeRow(13) = True
        If moneyCodeFields.Exists("Notes") Then
            .Cell(flexcpText, 14, 0, 14, 2) = moneyCodeFields("Notes")
        Else
            .Cell(flexcpText, 14, 0, 14, 2) = ""
        End If
        .MergeRow(14) = True
                
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, 1, False
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, 1, False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadMoneyCodeCheckListGrid1"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCustomCheckListGrid1
'' Description: Initialize and load the check list grid
'' Inputs:      Journal
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCustomCheckListGrid1(ByVal checklistFields As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumWeekly As Long              ' Number of weekly setups
    Dim lNumDaily As Long               ' Number of daily setups
    Dim lRow As Long                    ' Row in the grid
    Dim strSetup As String              ' Setup text

    With fgChecklist
        .Redraw = flexRDNone
        
        .Clear
        
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        .MergeCells = flexMergeFree
        .HighLight = flexHighlightNever
        .SelectionMode = flexSelectionFree
        .SheetBorder = .BackColor
        
        lNumWeekly = 0
        lNumDaily = 0
        
        For lIndex = 0 To 5
            If checklistFields.Exists("Weekly" & Str(lIndex)) Then
                If Len(checklistFields("Weekly" & Str(lIndex))) > 0 Then
                    lNumWeekly = lNumWeekly + 1
                End If
            End If
            If checklistFields.Exists("Daily" & Str(lIndex)) Then
                If Len(checklistFields("Daily" & Str(lIndex))) > 0 Then
                    lNumDaily = lNumDaily + 1
                End If
            End If
        Next lIndex
        
        If lNumWeekly > 0 Then
            lNumWeekly = lNumWeekly + 1
        End If
        If lNumDaily > 0 Then
            lNumDaily = lNumDaily + 1
        End If
        
        .Cols = 3
        .Rows = 2 + lNumWeekly + lNumDaily
        .FixedCols = 0
        .FixedRows = 0
        
'        .TextMatrix(0, 0) = "Symbol"
'        If checklistFields.Exists("Symbol") Then
'            .Cell(flexcpText, 0, 1, 0, 2) = checklistFields("Symbol")
'        Else
'            .Cell(flexcpText, 0, 1, 0, 2) = ""
'        End If
'        .MergeRow(0) = True
        
        lRow = 0
        
        If lNumWeekly > 0 Then
            .Cell(flexcpText, lRow, 0, lRow, 2) = "Weekly Setups"
            .Cell(flexcpFontUnderline, lRow, 0, lRow, 2) = True
            .MergeRow(lRow) = True
            lRow = lRow + 1
        End If
        
        For lIndex = 0 To 5
            If checklistFields.Exists("Weekly" & Str(lIndex)) Then
                strSetup = checklistFields("Weekly" & Str(lIndex))
                If Len(strSetup) > 0 Then
                    .TextMatrix(lRow, 0) = Parse(strSetup, ";", 1)
                    .Cell(flexcpText, lRow, 1, lRow, 2) = Parse(strSetup, ";", 2)
                    .MergeRow(lRow) = True
                    
                    lRow = lRow + 1
                End If
            End If
        Next lIndex
        
        If lNumDaily > 0 Then
            .Cell(flexcpText, lRow, 0, lRow, 2) = "Daily Setups"
            .Cell(flexcpFontUnderline, lRow, 0, lRow, 2) = True
            .MergeRow(lRow) = True
            lRow = lRow + 1
        End If
        
        For lIndex = 0 To 5
            If checklistFields.Exists("Daily" & Str(lIndex)) Then
                strSetup = checklistFields("Daily" & Str(lIndex))
                If Len(strSetup) > 0 Then
                    .TextMatrix(lRow, 0) = Parse(strSetup, ";", 1)
                    .Cell(flexcpText, lRow, 1, lRow, 2) = Parse(strSetup, ";", 2)
                    .MergeRow(lRow) = True
                    
                    lRow = lRow + 1
                End If
            End If
        Next lIndex
        
        .Cell(flexcpText, lRow, 0, lRow, 2) = "Notes"
        .Cell(flexcpFontUnderline, lRow, 0, lRow, 2) = True
        .MergeRow(lRow) = True
        lRow = lRow + 1
        
        If checklistFields.Exists("Notes") Then
            .Cell(flexcpText, lRow, 0, lRow, 2) = checklistFields("Notes")
        Else
            .Cell(flexcpText, lRow, 0, lRow, 2) = ""
        End If
        .MergeRow(lRow) = True
                
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, 1, False
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, 1, False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadCustomCheckListGrid1"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitSymbolsGrid
'' Description: Initialize the symbols grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitSymbolsGrid()
On Error GoTo ErrSection:

    With fgSymbols
        .Redraw = flexRDNone
        
        SetupGrid fgSymbols, eGridMode_List
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColor = vbWhite
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 1
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.InitSymbolsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbolsGrid
'' Description: Load the symbols grid
'' Inputs:      Current Tab
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSymbolsGrid(ByVal lTab As Long)
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray         ' Unique array of symbols
    Dim lIndex As Long                  ' Index into a for loop

    Set astrSymbols = New cGdArray
    astrSymbols.Create eGDARRAY_Strings

    With fgSymbols
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        If lTab = GDTab(eGDTabs_Orders) Then
            Set astrSymbols = g.JournalDB.GetSymbolsForOrderJournals
        ElseIf lTab = GDTab(eGDTabs_MoneyCode) Then
            Set astrSymbols = g.JournalDB.GetSymbolsForMoneyCodeJournals
        End If
        
        For lIndex = 0 To astrSymbols.Size - 1
            .AddItem astrSymbols(lIndex)
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            If lTab = GDTab(eGDTabs_Orders) Then
                LoadOrderJournalForSymbol .TextMatrix(.Row, 0)
            ElseIf lTab = GDTab(eGDTabs_MoneyCode) Then
                LoadMoneyCodeJournalForSymbol .TextMatrix(.Row, 0)
            End If
        Else
            fgJournal.Rows = fgJournal.FixedRows
        End If
        
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadSymbolsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournalForSymbol
'' Description: Load the order journal entries for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOrderJournalForSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim OrderJournals As cJournals      ' Order journals for the symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    Set OrderJournals = New cJournals
    g.JournalDB.LoadOrderJournalsForSymbol OrderJournals, strSymbol
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To OrderJournals.Count
            AddOrderJournalEntry OrderJournals(lIndex), False
        Next lIndex
        
        'AddClickHereRows
        
        SortJournalGrid
        SetBackColors fgJournal
        AutoSizeJournalGrid
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadOrderJournalForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMoneyCodeJournalForSymbol
'' Description: Load the Money Code journal entries for the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMoneyCodeJournalForSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim MoneyCodeJournals As cDateJournals ' Money Code journals for the symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    Set MoneyCodeJournals = New cDateJournals
    g.JournalDB.LoadDateJournalsForSymbol MoneyCodeJournals, strSymbol
    
    With fgJournal
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To MoneyCodeJournals.Count
            AddDateJournalEntry MoneyCodeJournals(lIndex), False
        Next lIndex
        
        'AddClickHereRows
        
        SortJournalGrid
        SetBackColors fgJournal
        AutoSizeJournalGrid
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDateJournals.LoadMoneyCodeJournalForSymbol"
    
End Sub

