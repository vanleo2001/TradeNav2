VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmTTSummary 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGridDump 
      Left            =   7740
      Top             =   60
   End
   Begin VB.Timer tmrTbCaption 
      Left            =   8220
      Top             =   480
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   60
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   14
      DisplayContextMenu=   0   'False
      Tools           =   "frmTTSummary.frx":0000
      ToolBars        =   "frmTTSummary.frx":0387
   End
   Begin VB.Timer tmrMenu 
      Left            =   7260
      Top             =   480
   End
   Begin VB.Timer tmrFlatten 
      Left            =   7260
      Top             =   60
   End
   Begin VB.Timer tmrBrokers 
      Left            =   8700
      Top             =   60
   End
   Begin VB.Timer tmrExitAll 
      Left            =   7740
      Top             =   480
   End
   Begin VB.Timer tmrRealTime 
      Left            =   8220
      Top             =   60
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrders 
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2025
      _cx             =   3572
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
   Begin VSFlex7LCtl.VSFlexGrid fgPositions 
      Height          =   810
      Left            =   2100
      TabIndex        =   1
      Top             =   0
      Width           =   1995
      _cx             =   3519
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
      Rows            =   2
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
   Begin VSFlex7LCtl.VSFlexGrid fgAccounts 
      Height          =   810
      Left            =   4200
      TabIndex        =   2
      Top             =   0
      Width           =   1980
      _cx             =   3492
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
   Begin VB.Image imgLog 
      Height          =   240
      Left            =   7800
      Picture         =   "frmTTSummary.frx":06B3
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSell 
      Height          =   240
      Left            =   7500
      Picture         =   "frmTTSummary.frx":07FD
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBuy 
      Height          =   240
      Left            =   7260
      Picture         =   "frmTTSummary.frx":0D87
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgGreen 
      Height          =   195
      Left            =   6960
      Picture         =   "frmTTSummary.frx":1311
      Top             =   180
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgYellow 
      Height          =   195
      Left            =   6660
      Picture         =   "frmTTSummary.frx":1597
      Top             =   180
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRed 
      Height          =   195
      Left            =   6360
      Picture         =   "frmTTSummary.frx":181D
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      Begin VB.Menu mnuAccountConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuAccountDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuAccountSwitch 
         Caption         =   "Switch"
      End
      Begin VB.Menu mnuAccountSwitchMode 
         Caption         =   "Switch Mode"
      End
      Begin VB.Menu mnuAccountConnection 
         Caption         =   "Connection Info"
      End
      Begin VB.Menu mnuAccountChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuAccountRefresh 
         Caption         =   "Refresh Account"
      End
      Begin VB.Menu mnuAccountActivity 
         Caption         =   "Activity View"
      End
      Begin VB.Menu mnuAccountBrokerView 
         Caption         =   "Broker View"
      End
      Begin VB.Menu mnuAccountOnline 
         Caption         =   "View Online"
      End
      Begin VB.Menu mnuAccountVerifyPositions 
         Caption         =   "Verify Positions"
      End
      Begin VB.Menu mnuAccountDetails 
         Caption         =   "View Account Details"
      End
      Begin VB.Menu mnuAccountSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountNew 
         Caption         =   "New Account"
      End
      Begin VB.Menu mnuAccountEdit 
         Caption         =   "Edit Account"
      End
      Begin VB.Menu mnuAccountDelete 
         Caption         =   "Delete Account"
      End
      Begin VB.Menu mnuAccountPerformanceReports 
         Caption         =   "Performance Reports"
      End
      Begin VB.Menu mnuAccountSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuAccountSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuAccountViewJournals 
         Caption         =   "View Journals"
      End
      Begin VB.Menu mnuAccountsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountsAutoSizeColumns 
         Caption         =   "Auto Size Columns"
      End
      Begin VB.Menu mnuAccountsDefaultColumns 
         Caption         =   "Default Columns"
      End
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "Orders"
      Begin VB.Menu mnuOrdersBuy 
         Caption         =   "BUY a Security"
      End
      Begin VB.Menu mnuOrdersSell 
         Caption         =   "SELL a Security"
      End
      Begin VB.Menu mnuOrdersOrderGroups 
         Caption         =   "Order Groups"
         Begin VB.Menu mnuOrdersOrderGroup 
            Caption         =   "<Manage>"
            Index           =   0
         End
      End
      Begin VB.Menu mnuOrdersSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersEdit 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuOrdersCancel 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuOrdersPark 
         Caption         =   "Park Order"
      End
      Begin VB.Menu mnuOrdersSubmit 
         Caption         =   "Submit Order"
      End
      Begin VB.Menu mnuOrdersSubmitAll 
         Caption         =   "Submit All Parked Orders"
      End
      Begin VB.Menu mnuOrdersOrderHistory 
         Caption         =   "Order History"
      End
      Begin VB.Menu mnuOrdersNewJournal 
         Caption         =   "New Journal for Order"
      End
      Begin VB.Menu mnuOrdersSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersManageXOS 
         Caption         =   "Manage Exit Order Strategies"
      End
      Begin VB.Menu mnuOrdersSelectXOS 
         Caption         =   "Select Exit Order Strategy"
      End
      Begin VB.Menu mnuOrdersRemoveXOS 
         Caption         =   "Remove Exit Order Strategy"
      End
      Begin VB.Menu mnuOrdersSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuOrdersTradeHistory 
         Caption         =   "Trade History"
      End
      Begin VB.Menu mnuOrdersSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuOrdersCheckStatus 
         Caption         =   "Check Status"
      End
      Begin VB.Menu mnuOrdersViewJournals 
         Caption         =   "View Journals"
      End
      Begin VB.Menu mnuOrdersSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersAutoSizeColumns 
         Caption         =   "Auto Size Columns"
      End
      Begin VB.Menu mnuOrdersDefaultColumns 
         Caption         =   "Default Columns"
      End
   End
   Begin VB.Menu mnuPositions 
      Caption         =   "Positions"
      Begin VB.Menu mnuPositionsFlatten 
         Caption         =   "Flatten Position"
      End
      Begin VB.Menu mnuPositionsReverse 
         Caption         =   "Reverse Position"
      End
      Begin VB.Menu mnuPositionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPositionsManageXOS 
         Caption         =   "Manage Exit Order Strategies"
      End
      Begin VB.Menu mnuPositionsSelectXOS 
         Caption         =   "Select Exit Order Strategies"
      End
      Begin VB.Menu mnuPositionsRemoveXOS 
         Caption         =   "Remove Exit Order Strategy"
      End
      Begin VB.Menu mnuPositionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPositionsPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPositionsPerformance 
         Caption         =   "Performance Report"
      End
      Begin VB.Menu mnuPositionsTradeHistory 
         Caption         =   "Trade History"
      End
      Begin VB.Menu mnuPositionsSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuPositionsCheckStatus 
         Caption         =   "Check Status"
      End
      Begin VB.Menu mnuPositionsViewJournals 
         Caption         =   "View Journals"
      End
      Begin VB.Menu mnuPositionsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPositionsAutoSizeColumns 
         Caption         =   "Auto Size Columns"
      End
      Begin VB.Menu mnuPositionsDefaultColumns 
         Caption         =   "Default Columns"
      End
   End
End
Attribute VB_Name = "frmTTSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTSummary.frm
'' Description: Display trade summary information to the user
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/06/2009   DAJ         Display "Mismatch" for position if in a mismatch
'' 02/23/2009   DAJ         Prepend WeekNum to PFG orders instead of calendar date
'' 05/06/2009   DAJ         Fail on enable of auto trade item if no data in the bars
'' 06/02/2009   DAJ         Use Bid/Ask instead of Last for Options P&L
'' 08/21/2009   DAJ         Set UserCancel flag on CancelOrder call
'' 08/25/2009   DAJ         Added support for PFG account information
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 10/08/2009   DAJ         Fix for advanced order trailing stop
'' 12/16/2009   DAJ         Remove old simulated "Sent" and "Trigger Pending" orders
'' 01/04/2009   DAJ         Keep trying to activate auto exits until we get data
'' 01/08/2009   DAJ         Added calls for working orders and open positions
'' 03/11/2010   DAJ         Moved to external forms, moved collections global
'' 03/15/2010   DAJ         Fixed the grid column information persistence
'' 03/16/2010   DAJ         Fixed FixSummaryDisplay when clean INI file
'' 03/17/2010   DAJ         Added routines for updating the toolbar captions
'' 03/17/2010   DAJ         Removed HasOpenPositions and HasWorkingOrders calls
'' 03/17/2010   DAJ         Toggle show/hide of auto trade column in orders, positions
'' 03/18/2010   DAJ         Don't color accounts, only color auto trade if active
'' 03/26/2010   DAJ         Changed Summary to Dashboard, no number on accounts button
'' 03/26/2010   DAJ         Changed icons for Buy, Sell, and Activity Log
'' 06/03/2010   DAJ         Changes for new Trade Sense Order Groups
'' 06/15/2010   DAJ         Added TradeSense orders as new Trade Console form
'' 07/01/2010   DAJ         Mods for the Extreme Charts versions of the software
'' 08/04/2010   DAJ         Added flag file for DanielCode/TradeSense Orders/Groups
'' 09/13/2010   DAJ         Show TradeSense order groups in working orders grids
'' 10/26/2010   DAJ         Logging enhancements, changed interval for broker timer
'' 10/26/2010   DAJ         Implemented timers for grid dumps and toolbar caption updates
'' 12/01/2010   DAJ         Require Gold for TradeSense order groups instead of flag file
'' 03/07/2011   DAJ         Added Change Password to Accounts context menu
'' 06/21/2011   DAJ         Separate out Simulated trading types, DoBrokerTimer enhancements
'' 06/28/2011   DAJ         Setup clickable cells like hyperlinks
'' 07/27/2011   DAJ         Trade Console toolbar button customization, Added Reports and Journal
'' 09/08/2011   DAJ         Changed call to update bars in trading items
'' 09/20/2011   DAJ         Changed journal icon
'' 10/04/2011   DAJ         Call the ShowJournals function instead of calling the form direct
'' 02/14/2012   DAJ         Added multi-leg order support
'' 11/28/2012   DAJ         Speed enhancements for the Trade Console
'' 01/07/2013   DAJ         Profiling for trade stuff ( for Brady and Tim )
'' 01/08/2013   DAJ         Only check for grid difference once per second
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 02/20/2013   DAJ         Added "Actual Performance" menu item
'' 05/29/2013   DAJ         Moved g.TradingItems.UpdateBars to a dedicated timer
'' 06/24/2013   DAJ         Timer Logging
'' 07/30/2013   DAJ         Moved tmrUpdateCharts to frmOnlineBroker
'' 06/11/2014   DAJ         Dump the automated trading items grid to a file if it changes
'' 08/13/2014   DAJ         Changed flatten timer interval
'' 08/28/2014   DAJ         Changed profile numbers for real-time timer
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/29/2014   DAJ         Remove old synthetic order/MIT code
'' 12/21/2015   MJM         Changed ForecolorSource of tool buttons ID_ActivityLog, ID_TodaysFills to "UseControl".
''                          This was the only way to get the text to show up white in Charcoal theme
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kAspectRatio = 0.33
Private Const kMenuPrefix = "C:"

Public Enum eGDConsoleTabs
    eGDConsoleTab_Summary = 0
    eGDConsoleTab_OpenOrders
    eGDConsoleTab_Positions
    eGDConsoleTab_Accounts
    eGDConsoleTab_AutoTrading
    eGDConsoleTab_ActivityLog
    eGDConsoleTab_TodaysFills
End Enum

Private Type mPrivate
    Orders As cWorkingOrdersUI          ' Working orders object
    Positions As cPositionsUI           ' Open Positions object (non-Summary)
    Accounts As cAccountsUI             ' Accounts object (non-Summary)
    
    TbCaption As cGdTree                ' Collection of Toolbar Caption updates
    GridDump As cGdTree                 ' Collection of Grid dumps to do

    BarsColl As cGdTree                 ' Collection of Bars for Real Time
    SpreadData As cGdTree               ' Collection of Spread Data objects
    
    JournalOrder As cPtOrder            ' Order to show journal for
            
    adLastChanged As cGdArray           ' Array of Last Changed information by broker
    
    nDockState As DockState             ' Dockable state of the form
    nDockAlign As HostAlign             ' Dockable alignment of the form
    lScaleHeight As Long                ' Last known non-zero scale height of the form
    lOrderGroupIndex As Long            ' Menu index of the selected order group item
    bShowTabOnAttach As Boolean         ' Show tab when attaching it?
End Type
Private m As mPrivate

Public DumpProfile As Boolean           ' Dump the trade console profiling?

Private Function Tabs(ByVal nTab As eGDConsoleTabs)
    Tabs = nTab
End Function

Public Property Get BarsExist(ByVal vSymbolOrSymbolID As Variant) As Boolean
    BarsExist = m.BarsColl.Exists(Str(vSymbolOrSymbolID))
End Property

Public Property Get TabAttached(ByVal nTab As eGDConsoleTabs) As Boolean
End Property
Public Property Let TabAttached(ByVal nTab As eGDConsoleTabs, ByVal bAttached As Boolean)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow as outside caller to print the grid information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    PrintMe = frmPrintPreview.ShowMe("CNV TradeConsole", Me, , , , , , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.PrintMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader

        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = taCenterMiddle
        .Text = "Trade Console - Summary"
        .TextAlign = taLeftMiddle
        .Font.Bold = False
        
        .Paragraph = ""
        
        .Text = "Open Orders" & vbLf
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgOrders
        Else
            .RenderControl = fgOrders.hWnd
        End If
        
        .Paragraph = ""
        
        .Text = "Open Positions" & vbLf
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgPositions
        Else
            .RenderControl = fgPositions.hWnd
        End If
        
        .Paragraph = ""
        
        .Text = "Accounts" & vbLf
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgAccounts
        Else
            .RenderControl = fgAccounts.hWnd
        End If
        
        .EndDoc
    End With
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors on both grids if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    m.Orders.ClearUpdatedColors
    m.Positions.ClearUpdatedColors
    
    If FormIsLoaded("frmWorkingOrders") Then
        frmWorkingOrders.ClearUpdatedColors
    End If
    If FormIsLoaded("frmOpenPositions") Then
        frmOpenPositions.ClearUpdatedColors
    End If
    If FormIsLoaded("frmTTPositions") Then
        frmTTPositions.ClearUpdatedColors
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTSummary.ClearUpdatedColors"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBars
'' Description: Get the Bars for the given symbol from the collection (add it
''              to the collection if not there)
'' Inputs:      Symbol to get Bars for, Add to RT?
'' Returns:     Bars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBars(ByVal vSymbolOrSymbolID As Variant, Optional ByVal bAddToRT As Boolean = True) As cGdBars
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary Bars structure
    Dim lSymbolID As Long               ' Symbol ID for variant passed in
    Dim strSymbol As String             ' Symbol for variant passed in
    Dim strLeadContract As String       ' Lead contract for the spread
    
    lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If lSymbolID <> 0 Then vSymbolOrSymbolID = lSymbolID Else vSymbolOrSymbolID = strSymbol
    
    If InStr(strSymbol, "|") Then
        Set Bars = GetSpreadData(strSymbol, bAddToRT).Bars
    Else
        If m.BarsColl.Exists(Str(vSymbolOrSymbolID)) = False Then
            LoadBars Bars, vSymbolOrSymbolID, bAddToRT
            m.BarsColl.Add Bars, Str(vSymbolOrSymbolID)
        End If
        Set Bars = m.BarsColl(Str(vSymbolOrSymbolID))
        
        ' If this is a future calendar spread, then make sure to add the "base" contract
        ' to the stream as well...
        strLeadContract = LeadContractForSpread(strSymbol)
        If Len(strLeadContract) > 0 Then
            GetBars strLeadContract, bAddToRT
        End If
    End If
    
    Set GetBars = Bars

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.GetBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSpreadData
'' Description: Get the spread data for the given symbol from the collection (add it
''              to the collection if not there)
'' Inputs:      Symbol to get spread data for, Add to RT?
'' Returns:     Spread data
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSpreadData(ByVal strSymbol As String, Optional ByVal bAddToRT As Boolean = True) As cSpreadData
On Error GoTo ErrSection:

    Dim SpreadData As cSpreadData       ' Spread data object to return
    
    If m.SpreadData.Exists(strSymbol) = False Then
        LoadSpreadData SpreadData, strSymbol, bAddToRT
        m.SpreadData.Add SpreadData, strSymbol
    End If
    
    Set GetSpreadData = m.SpreadData(strSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.GetSpreadData"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshData
'' Description: Refresh the data for each of the bars in the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshData()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As cGdBars                 ' Temporary bars object
    Dim bWorkingOrders As Boolean       ' Is the working orders form loaded?
    Dim bTradeSummary As Boolean        ' Is the trade summary form loaded?
    
    bWorkingOrders = FormIsLoaded("frmWorkingOrders")
    bTradeSummary = FormIsLoaded("frmTTPositions")
    
    For lIndex = 1 To m.BarsColl.Count
        Set Bars = m.BarsColl(lIndex)
        
        If LoadBars(Bars, Bars.SymbolOrSymbolID) Then
            If Bars.BarsHandle <> m.BarsColl(lIndex).BarsHandle Then Set m.BarsColl(lIndex) = Bars
            RefreshPrices m.BarsColl(lIndex), bWorkingOrders, bTradeSummary
        End If
    Next lIndex
    
    For lIndex = 1 To m.SpreadData.Count
        RefreshSpreadData m.SpreadData(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.RefreshData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshForm
'' Description: Refresh the form after a download in case module code changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshForm()
On Error GoTo ErrSection:

    Form_Resize
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.RefreshForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrids
'' Description: Re-Filter all of the grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FilterGrids()
On Error GoTo ErrSection:

    m.Orders.FilterOrdersGrid
    m.Positions.FilterPositionsGrid
    m.Accounts.FilterAccountsGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.FilterGrids"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisableTimers
'' Description: Disable all of the timers on the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisableTimers()
On Error GoTo ErrSection:

    tmrRealtime.Enabled = False
    tmrExitAll.Enabled = False
    tmrBrokers.Enabled = False
    tmrFlatten.Enabled = False
    tmrMenu.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.DisableTimers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TempBrokerAccount
'' Description: Create a temporary broker account for purposes of connection
'' Inputs:      Broker Type, Broker User
'' Returns:     Temporary Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TempBrokerAccount(ByVal nBroker As eTT_AccountType, ByVal bBrokerUser As Boolean)
On Error GoTo ErrSection:

    m.Accounts.TempBrokerAccount nBroker, bBrokerUser
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.TempBrokerAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoBrokerTimer
'' Description: Perform the actions for the broker timer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoBrokerTimer()
On Error GoTo ErrSection:

gdResetProfiles 620, 629
gdStartProfile 620
gdStartProfile 621

    Dim adBrokers As cGdArray           ' List of brokers back from call
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOrdersAfter As cGdArray     ' Orders grid after update
    Dim astrPositionsAfter As cGdArray  ' Positions grid after update
    Dim bLastStatusChanged As Boolean   ' Did the time of the last connection status change?
    Static bInProgress As Boolean       ' Are we still processing?
    Static dLastStatus As Double        ' Time of the last known connection status change
    Static dLastCheck As Double         ' Last time that the grids were checked
    Static astrOrdersBefore As cGdArray ' Orders grid before update
    Static astrPositionsBefore As cGdArray ' Positions grid before update
    
gdStopProfile 621
    
    If g.bUnloading = False Then
        If bInProgress = False Then
            bInProgress = True
            
gdStartProfile 622
            bLastStatusChanged = (dLastStatus <> g.Broker.LastStatusChange)
            If bLastStatusChanged Then
                dLastStatus = g.Broker.LastStatusChange
            End If

gdStopProfile 622
gdStartProfile 623

            If (g.bStarting = False) And ((bLastStatusChanged = True) Or (g.Broker.PositionsToVerify = True)) Then
                If (g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected) Or (g.RealTime.ConnectionStatus = eGDConnectionStatus_Disconnected) Then
                    ' Check to see if the user needs to fix any positions...
                    g.Broker.FixPositions True

                    ' Check to see if any broker needs their positions verified...
                    g.Broker.VerifyBrokerPositions
                    
                    ' Check to see if we need to notify user about new exchanges...
                    g.Broker.NotifyNewExchanges
                End If
            End If
    
gdStopProfile 623
gdStartProfile 624

            Set adBrokers = g.Broker.LastChangedForAll
            
gdStopProfile 624
gdStartProfile 625

            If Not adBrokers Is Nothing Then
                For lIndex = 1 To adBrokers.Size - 1
                    If m.adLastChanged(lIndex) < adBrokers(lIndex) Then
                        m.Orders.Update lIndex
                        m.Positions.Update lIndex
                        m.Accounts.Update lIndex
                        
                        m.adLastChanged(lIndex) = adBrokers(lIndex)
                        g.Broker.BrokerDebug lIndex, vbTab & "Trade Console Broker Loop Done"
                    End If
                Next lIndex
                
                ' 01/08/2013 DAJ: The broker timer is going off every 100ms, but we don't want to do
                ' this check every time -- only do it every second...
                If gdTickCount > dLastCheck + 1000 Then
                    If astrOrdersBefore Is Nothing Then
                        Set astrOrdersBefore = New cGdArray
                        astrOrdersBefore.Create eGDARRAY_Strings
                    End If
                    If astrPositionsBefore Is Nothing Then
                        Set astrPositionsBefore = New cGdArray
                        astrPositionsBefore.Create eGDARRAY_Strings
                    End If
                    
                    Set astrOrdersAfter = GridToArray(fgOrders)
                    Set astrPositionsAfter = GridToArray(fgPositions)
                
                    DumpGridIfDifferent astrOrdersBefore, astrOrdersAfter, "Trade Console Orders"
                    DumpGridIfDifferent astrPositionsBefore, astrPositionsAfter, "Trade Console Positions"
                    
                    Set astrOrdersBefore = astrOrdersAfter.MakeCopy
                    Set astrPositionsBefore = astrPositionsAfter.MakeCopy
                    
                    dLastCheck = gdTickCount
                    
                    If FormIsLoaded("frmTradeItems") Then
                        frmTradeItems.DumpGridIfDifferent
                    End If
                End If
            End If
    
gdStopProfile 625
gdStartProfile 626

            ' Update the TradeSense order groups on the open orders grid...
            m.Orders.UpdateTsOrders
    
gdStopProfile 626
gdStartProfile 627

            ' Set the account pictures accordingly on the accounts grid...
            If bLastStatusChanged = True Then
                m.Accounts.SetAccountPictures
            End If
    
gdStopProfile 627
gdStartProfile 628

            ' See if there are any auto exits that need to be deleted because they are
            ' inactive and have no working orders...
            g.OrderStrategies.DeleteInactiveExits
            
gdStopProfile 628
gdStartProfile 629

            ' See if there are any old pending orders for any brokers that will cause us
            ' to do a refresh for that broker...
            g.Broker.CheckPendingOrders
            
gdStopProfile 629

            bInProgress = False
        End If
    End If
    
gdStopProfile 620

If DumpProfile Then
    DebugLog "=================" & vbCrLf & gdGetProfiles(620, 629, vbCrLf) & vbCrLf & "================="
End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "frmTTSummary.DoBrokerTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateConsoleSettings
'' Description: Update the console settings from the configuration form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateConsoleSettings()
On Error GoTo ErrSection:
    
    m.Orders.UpdateConsoleSettings
    m.Positions.UpdateConsoleSettings
    m.Accounts.UpdateConsoleSettings
    
    If FormIsLoaded("frmWorkingOrders") Then
        frmWorkingOrders.UpdateConsoleSettings
    End If
    If FormIsLoaded("frmOpenPositions") Then
        frmOpenPositions.UpdateConsoleSettings
    End If
    If FormIsLoaded("frmAccounts") Then
        frmAccounts.UpdateConsoleSettings
    End If
    If FormIsLoaded("frmTradeItems") Then
        frmTradeItems.UpdateConsoleSettings
    End If
    
    ShowToolbarButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.UpdateConsoleSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowJournalForOrder
'' Description: Show the journal form for the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowJournalForOrder(Order As cPtOrder)
On Error GoTo ErrSection:

    Set m.JournalOrder = Order
    StartMenuTimer "JOURNAL"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.ShowJournalForOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormShown
'' Description: Set the user interface based on if the given form is shown
'' Inputs:      Form, Shown?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FormShown(ByVal nForm As eGDConsoleForms, ByVal bShown As Boolean)
On Error GoTo ErrSection:

    If bShown Then
        tbToolbar.Tools(ToolForForm(nForm)).State = ssChecked
    Else
        tbToolbar.Tools(ToolForForm(nForm)).State = ssUnchecked
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.FormShown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateToolbarCaptions
'' Description: Update the toolbar captions with the number of visible items
'' Inputs:      Form, Number of visible items
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateToolbarCaption(ByVal nForm As eGDConsoleForms, ByVal lNumVisible As Long)
On Error GoTo ErrSection:

    If m.TbCaption.Exists(Str(nForm)) Then
        m.TbCaption(Str(nForm)) = lNumVisible
    Else
        m.TbCaption.Add lNumVisible, Str(nForm)
    End If
    tmrTbCaption.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.UpdateToolbarCaptions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Upon the activation of the form, if the user does not have any
''              accounts, allow them to create one
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean      ' Have we done some first time stuff?
    Static bInProgress As Boolean       ' Are we already handling this?
    Dim lIndex As Long                  ' Index into a for loop

    If (bAlreadyDone = False) And (frmMain.Visible = True) Then
        bAlreadyDone = True
        g.ConsoleForms.ShowForms
    End If

    If (bInProgress = False) And (Visible = True) Then
        bInProgress = True
        
        If fgAccounts.Rows = fgAccounts.FixedRows Then
            If frmTTEditAccount.ShowMe(0&) = False Then
                frmMain.tbToolbar.Tools("ID_TradeTracker").State = ssUnchecked
            End If
        End If
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Deactivate
'' Description: Upon the deactivation of the form, set the previus active form
''              so that printing from the toolbar will work properly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me '(so will print from toolbar button)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.Form_Deactivate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form and controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim i&, strText$
    Dim lAccountID As Long              ' Default SimTrade Account ID
    Dim strDisplay As String            ' Display string from the ini file
    Dim rs As Recordset                 ' Recordset into the database
    Dim lTab As Long                    ' Tab to set the tabs to
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrDisplay As New cGdArray     ' Display string split into an array
    Dim astrLine As New cGdArray        ' Display line
    Dim lOrders As Long                 ' Index of the orders grid in array
    Dim lPositions As Long              ' Index of the positions grid in array
    Dim lAccounts As Long               ' Index of the accounts grid in array
    Dim OrdersUI As cWorkingOrdersControls  ' Working order controls object
    Dim PositionUI As cPositionsControls    ' Position controls object
    Dim AccountUI As cAccountsControls  ' Account controls object
    Dim Tool As SSTool
    
    g.Styler.StyleForm Me
    
    ' 03/03/2010 DAJ: Put this in because Tim ran into an instance of this on the
    ' installs machine.  We need to see what is causing the form to reload on unload...
    If g.bUnloading And IsIDE Then
        MsgBox "Come and get Dave because we shouldn't be here", , "Console Load"
    End If
    
    DumpProfile = FileExist(AddSlash(App.Path) & "TTProfile.FLG")
    m.nDockState = -1
    m.nDockAlign = -1
    
    If Not DirExist(AddSlash(App.Path) & "TradeConsole") Then MakeDir AddSlash(App.Path) & "TradeConsole"
    KillFile AddSlash(App.Path) & "TradeConsole\*.LOG /o=-30"
    
    ' Make sure to do this before loading the online broker form which will load the SimTrade object
    ' which will load the SimTrade Orders and Fills, however it needs g.Broker to be alive...
    PerformStartupFixes
        
    If Not FormIsLoaded("frmOnlineBroker") Then
        Load frmOnlineBroker
        StartupLog "Online Broker Form Loaded"
    End If
    
    Set m.BarsColl = New cGdTree
    Set m.SpreadData = New cGdTree
    
    Set m.adLastChanged = New cGdArray
    m.adLastChanged.Create eGDARRAY_Doubles, kNumBrokers
    
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)

    tmrExitAll.Interval = 1000
    tmrExitAll.Enabled = False
    
    tmrBrokers.Interval = 100
    tmrBrokers.Enabled = False
    
    ' DAJ 08/13/2014: We received a complaint from a customer saying that the flatten was taking four
    ' seconds which was too long in his opinion.  I noticed in looking at his logs that most of the time
    ' was waiting for the next timer interval, so I am going to change the interval here...
    tmrFlatten.Interval = 100 ' 1000
    tmrFlatten.Enabled = False
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    tmrTbCaption.Interval = 10
    tmrTbCaption.Enabled = False
    Set m.TbCaption = New cGdTree
    
    tmrGridDump.Interval = 10
    tmrGridDump.Enabled = False
    Set m.GridDump = New cGdTree
    
    g.FlattenQueue.Init tmrFlatten
    g.ExitAllOrders.Init tmrExitAll
    
    'cmdSettings.Picture = Picture16(ToolbarIcon("ID_Settings"))
    'cmdSettings.ToolTipText = "Trade Console Settings"
    'cmdTradeTracker.Picture = Picture16(ToolbarIcon("ID_TradeTracker"))
    'cmdTradeTracker.ToolTipText = "View Historical Trades and Details in the Trade Tracker"
    'cmdModify.ToolTipText = "Edit, Cancel, or Submit an Order"
        
    mnuAccounts.Visible = False
    mnuOrders.Visible = False
    mnuPositions.Visible = False
    
    If Len(g.Broker.GridFont) > 0 Then
        SetAllGridFonts g.Broker.GridFont
    End If

    FixSummaryDisplay
    
    fgOrders.BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbButtonFace
    
    Set OrdersUI = New cWorkingOrdersControls
    With OrdersUI
        Set .frm = Me
        Set .fgGrid = fgOrders
        Set .tmrRealtime = tmrRealtime
        Set .tmrMenu = tmrMenu
        Set .mnuOrders = mnuOrders
        Set .mnuBuy = mnuOrdersBuy
        Set .mnuSell = mnuOrdersSell
        Set .mnuOrderGroups = mnuOrdersOrderGroups
        Set .mnuOrderGroup = mnuOrdersOrderGroup
        Set .mnuEditOrder = mnuOrdersEdit
        Set .mnuCancelOrder = mnuOrdersCancel
        Set .mnuParkOrder = mnuOrdersPark
        Set .mnuSubmitOrder = mnuOrdersSubmit
        Set .mnuSubmitAll = mnuOrdersSubmitAll
        Set .mnuOrderHistory = mnuOrdersOrderHistory
        Set .mnuNewJournal = mnuOrdersNewJournal
        Set .mnuManageXOS = mnuOrdersManageXOS
        Set .mnuSelectXOS = mnuOrdersSelectXOS
        Set .mnuRemoveXOS = mnuOrdersRemoveXOS
        Set .mnuPrint = mnuOrdersPrint
        Set .mnuTradeHistory = mnuOrdersTradeHistory
        Set .mnuSettings = mnuOrdersSettings
        Set .mnuCheckStatus = mnuOrdersCheckStatus
        Set .mnuViewJournals = mnuOrdersViewJournals
        Set .mnuAutoSizeColumns = mnuOrdersAutoSizeColumns
        Set .mnuDefaultColumns = mnuOrdersDefaultColumns
    End With
    
    Set PositionUI = New cPositionsControls
    With PositionUI
        Set .frm = Me
        Set .fgGrid = fgPositions
        Set .tmrRealtime = tmrRealtime
        Set .tmrMenu = tmrMenu
        Set .mnuPositions = mnuPositions
        Set .mnuFlatten = mnuPositionsFlatten
        Set .mnuReverse = mnuPositionsReverse
        Set .mnuManageXOS = mnuPositionsManageXOS
        Set .mnuSelectXOS = mnuPositionsSelectXOS
        Set .mnuRemoveXOS = mnuPositionsRemoveXOS
        Set .mnuPrint = mnuPositionsPrint
        Set .mnuActualPerformance = mnuPositionsPerformance
        Set .mnuTradeHistory = mnuPositionsTradeHistory
        Set .mnuSettings = mnuPositionsSettings
        Set .mnuCheckStatus = mnuPositionsCheckStatus
        Set .mnuViewJournals = mnuPositionsViewJournals
        Set .mnuAutoSizeColumns = mnuPositionsAutoSizeColumns
        Set .mnuDefaultColumns = mnuPositionsDefaultColumns
    End With
    
    Set AccountUI = New cAccountsControls
    With AccountUI
        Set .frm = Me
        Set .fgGrid = fgAccounts
        Set .tmrRealtime = tmrRealtime
        Set .tmrMenu = tmrMenu
        Set .mnuAccounts = mnuAccounts
        Set .mnuConnect = mnuAccountConnect
        Set .mnuDisconnect = mnuAccountDisconnect
        Set .mnuSwitchAccounts = mnuAccountSwitch
        Set .mnuSwitchAccountsMode = mnuAccountSwitchMode
        Set .mnuConnectInfo = mnuAccountConnection
        Set .mnuChangePassword = mnuAccountChangePassword
        Set .mnuRefresh = mnuAccountRefresh
        Set .mnuViewActivity = mnuAccountActivity
        Set .mnuBrokerView = mnuAccountBrokerView
        Set .mnuViewOnline = mnuAccountOnline
        Set .mnuVerifyPositions = mnuAccountVerifyPositions
        Set .mnuAccountDetails = mnuAccountDetails
        Set .mnuSep1 = mnuAccountSep1
        Set .mnuNewAccount = mnuAccountNew
        Set .mnuEditAccount = mnuAccountEdit
        Set .mnuDeleteAccount = mnuAccountDelete
        Set .mnuSep2 = mnuAccountSep2
        Set .mnuReports = mnuAccountPerformanceReports
        Set .mnuPrint = mnuAccountPrint
        Set .mnuSettings = mnuAccountSettings
        Set .mnuViewJournals = mnuAccountViewJournals
        Set .mnuAutoSizeColumns = mnuAccountsAutoSizeColumns
        Set .mnuDefaultColumns = mnuAccountsDefaultColumns
    End With
    
    Set m.Orders = New cWorkingOrdersUI
    m.Orders.Init "Trade Console", OrdersUI, True
    
    Set m.Positions = New cPositionsUI
    m.Positions.Init "Trade Console", PositionUI, True
    
    Set m.Accounts = New cAccountsUI
    m.Accounts.Init "Trade Console", AccountUI, True
    
    Set m.JournalOrder = Nothing
    m.lOrderGroupIndex = -1&
    
    strText = "Classic"
    If g.nTbIconStyle = 1 Then
        If g.nColorTheme = kDarkThemeColor Then
            strText = "Light"
        Else
            strText = "Dark"
        End If
    End If
    With tbToolbar
        .Tools("ID_Summary").Picture = Picture16(ToolbarIcon("kDashboard"))
        .Tools("ID_OpenOrders").Picture = Picture16(ToolbarIcon("kOrders"))
        .Tools("ID_Positions").Picture = Picture16(ToolbarIcon("kPositions"))
        .Tools("ID_Accounts").Picture = Picture16(ToolbarIcon("kDollarSign"))
        .Tools("ID_AutoTrading").Picture = Picture16(ToolbarIcon("kSystem"))
        .Tools("ID_TradeSenseOrders").Picture = Picture16(ToolbarIcon("kTradeSenseOrders"))
        .Tools("ID_ActivityLog").Picture = Picture16(ToolbarIcon("kNews2"))
        .Tools("ID_TodaysFills").Picture = Picture16(ToolbarIcon("kFills"))
        
        .Tools("ID_Buy").Picture = Picture16(ToolbarIcon("kBuy"))
        .Tools("ID_Sell").Picture = Picture16(ToolbarIcon("kSell"))
        .Tools("ID_Tracking").Picture = Picture16(ToolbarIcon("kTracking"))
        If IsWoodiesVersion Then
            .Tools("ID_Reports").Picture = Picture16(ToolbarIcon("ID_TradeFilter"))
        Else
            .Tools("ID_Reports").Picture = Picture16(ToolbarIcon("kPerformance"))
        End If
        .Tools("ID_Journals").Picture = Picture16(ToolbarIcon("kScroll"))
        .Tools("ID_Settings").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Settings"), "", 16).ExtractIcon
        
        For Each Tool In .Tools
            strText = TranslatedText(Tool.ID, Tool.Name)
            If strText <> Tool.Name Then
                Tool.ChangeAll ssChangeAllName, strText
            End If
        Next
    End With
    ShowToolbarButtons
    
    EnableControls
    
    tmrBrokers.Enabled = True
    StartupLog "Trade Console Loaded"
    
    ' Tell the SimTrade object to do a refresh since we are "connected"...
    'If Not g.SimTradeStream Is Nothing Then
    '    g.SimTradeStream.Refresh
    'End If
    
    StartupLog "SimTrade Information Loaded"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_MouseMove
'' Description: If the mouse cursor has been set somewhere else, reset it
'' Inputs:      Button pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Me.MousePointer = vbCustom Then
        Me.MousePointer = vbDefault
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Set the state of the toolbar icon on the main toolbar
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_TradeTracker").State = ssUnchecked
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.Form_QueryUnload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Form_Resize()
On Error Resume Next

    Dim nDockState As DockState         ' Dockable state of the form
    Dim nDockAlign As HostAlign         ' Dockable alignment of the form
    Dim bDisableSummary As Boolean      ' Disable the summary button?

    If Visible Then
        nDockState = frmMain.DockPro.State("frmTTSummary")
        nDockAlign = frmMain.DockPro.AlignWhenDocked("frmTTSummary")
        bDisableSummary = False
        
        If nDockState <> m.nDockState Then
            g.ConsoleForms.SummaryChangedState m.nDockState
            m.nDockState = nDockState
        ElseIf nDockState = DPDocked Then
            If nDockAlign = HAlignTop Then
                g.ConsoleForms.SummaryHeight = frmMain.DockPro.TopEdgeHeight
                If m.nDockAlign <> nDockAlign Then
                    nDockAlign = m.nDockAlign
                    If frmMain.DockPro.DockedCount(HAlignTop) > 1 Then
                        bDisableSummary = True
                    End If
                End If
            ElseIf nDockAlign = HAlignBottom Then
                g.ConsoleForms.SummaryHeight = frmMain.DockPro.BottomEdgeHeight
                If m.nDockAlign <> nDockAlign Then
                    nDockAlign = m.nDockAlign
                    If frmMain.DockPro.DockedCount(HAlignBottom) > 1 Then
                        bDisableSummary = True
                    End If
                End If
            End If
        ElseIf nDockState = DPUndocked Then
            g.ConsoleForms.SummaryHeight = Height
        End If
        
        tbToolbar.Tools("ID_Summary").Enabled = Not bDisableSummary
    End If

    PlaceGrids
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Do some cleanup when the form gets unloaded
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    SetIniFileProperty "TTSummary", FontToString(fgAccounts.Font), "Fonts", g.strIniFile
    
    DisableTimers
    
    Set m.Orders = Nothing
    Set m.Positions = Nothing
    Set m.Accounts = Nothing
    
    frmMain.DockPro.RemoveForm Me.Name
    
    For lIndex = 1 To m.BarsColl.Count
        g.RealTime.RemoveTickBuffer m.BarsColl(lIndex)
    Next lIndex
    Set m.BarsColl = Nothing
    
    Set m.adLastChanged = Nothing
        
    If FormIsLoaded("frmOnlineBroker") Then
        Unload frmOnlineBroker
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersOrderGroup_Click
'' Description: The user has chosen an order group
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersOrderGroup_Click(Index As Integer)
On Error GoTo ErrSection:

    m.Orders.SelectOrderGroup Index

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.mnuOrdersOrderGroup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Allow the user to show a particular form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim bShow As Boolean                ' Show the form?
    Dim nForm As eGDConsoleForms        ' Form to show or hide
    
    bShow = (Tool.State = ssChecked)
    nForm = -1&
    
    Select Case UCase(Tool.ID)
        Case "ID_SUMMARY"
            nForm = eGDConsoleForm_Summary
            
        Case "ID_OPENORDERS"
            nForm = eGDConsoleForm_OpenOrders
            
        Case "ID_POSITIONS"
            nForm = eGDConsoleForm_Positions
            
        Case "ID_ACCOUNTS"
            nForm = eGDConsoleForm_Accounts
            
        Case "ID_AUTOTRADING"
            nForm = eGDConsoleForm_AutoTrading
            
        Case "ID_TRADESENSEORDERS"
            If bShow Then
                If HasLevel(eTN4_Gold, True, "TradeSense Orders") Then
                    nForm = eGDConsoleForm_TradeSenseOrders
                Else
                    Tool.State = ssUnchecked
                End If
            Else
                nForm = eGDConsoleForm_TradeSenseOrders
            End If
            
        Case "ID_ACTIVITYLOG"
            nForm = eGDConsoleForm_ActivityLog
            
        Case "ID_TODAYSFILLS"
            nForm = eGDConsoleForm_TodaysFills
            
        Case "ID_BUY"
            CreateOrder "", 0&, 1, , , "Trade Console"
            
        Case "ID_SELL"
            CreateOrder "", 0&, 0&, , , "Trade Console"
            
        Case "ID_SETTINGS"
            StartMenuTimer "SETTINGS"
        
        Case "ID_TRACKING"
            StartMenuTimer "TRADEHISTORY"
            
        Case "ID_REPORTS"
            StartMenuTimer "REPORTS"
            
        Case "ID_JOURNALS"
            StartMenuTimer "JOURNALS"
                        
    End Select
    
    If nForm <> -1& Then
        g.ConsoleForms.ShowForm(nForm) = bShow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrBrokers_Timer
'' Description: Check to see if there is any new information for any symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrBrokers_Timer()
On Error GoTo ErrSection:

    TimerStart "frmTTSummary.tmrBrokers"
    DoBrokerTimer
    TimerEnd "frmTTSummary.tmrBrokers", tmrBrokers.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.tmrBrokers_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrGridDump_Timer
'' Description: Dump whatever grid information is in the queue
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrGridDump_Timer()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    TimerStart "frmTTSummary.tmrGridDump"
    tmrGridDump.Enabled = False

    For lIndex = 1 To m.GridDump.Count
        DumpDebug m.GridDump(lIndex)
    Next lIndex
    m.GridDump.Clear
    TimerEnd "frmTTSummary.tmrGridDump", tmrGridDump.Interval
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTSummary.tmrGridDump_Timer"
    
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
    Dim Account As cPtAccount           ' Selected account

    TimerStart "frmTTSummary.tmrMenu"
    strTag = tmrMenu.Tag
    If Len(strTag) > 2 Then
        If Left(strTag, 2) = kMenuPrefix Then
            strTag = Mid(strTag, 3)
            tmrMenu.Tag = ""
            tmrMenu.Enabled = False
            
            Select Case UCase(strTag)
                Case "JOURNAL"
                    g.TnJournal.ShowOrderJournal m.JournalOrder
                    
                Case "JOURNALS"
                    g.TnJournal.ShowJournals
                    
                Case "REPORTS"
                    ShowTradeFilter
                
                Case "SETTINGS"
                    frmTTSummaryCfg.ShowMe
                    
                Case "TRADEHISTORY"
                    Set Account = m.Accounts.SelectedAccount
                    If Account Is Nothing Then
                        frmTTAccounts.ShowMe True
                    Else
                        frmTTPositions.ShowMe Account.AccountID, Account.AccountType
                    End If
                    
            End Select
        End If
    End If
    TimerEnd "frmTTSummary.tmrMenu", tmrMenu.Interval
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRealTime_Timer
'' Description: Update the real time information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRealTime_Timer()
On Error GoTo ErrSection:

gdResetProfiles 500, 529

gdStartProfile 500
gdStartProfile 501

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewBar As Boolean              ' There is a new bar, so we need to reload data
    Dim Bars As cGdBars                 ' Temporary bars object
    Dim Order As cPtOrder               ' Temporary order object
    Static sdLastExp As Double          ' Last check for expiring orders
    Dim bUpdateBars As Boolean          ' Were the bars updated?
    Dim bWorkingOrders As Boolean       ' Is the working orders form loaded?
    Dim bTradeTracker As Boolean        ' Is the trade tracker form loaded?
    
    TimerStart "frmTTSummary.tmrRealTime"
    bWorkingOrders = FormIsLoaded("frmWorkingOrders")
    bTradeTracker = FormIsLoaded("frmTTPositions")
    
gdStopProfile 501
gdStartProfile 502
    
    For lIndex = 1 To m.BarsColl.Count
        bUpdateBars = g.RealTime.UpdateBars(m.BarsColl(lIndex), bNewBar)
        If bUpdateBars = True Then
            If bNewBar Then
                Set Bars = m.BarsColl(lIndex)
                
                LoadBars Bars, Bars.SymbolOrSymbolID
                If Bars.BarsHandle <> m.BarsColl(lIndex).BarsHandle Then Set m.BarsColl(lIndex) = Bars
                
                g.OrderStrategies.UpdateSessionDate Bars.SymbolOrSymbolID, Bars.SessionDate(Bars.Size - 1)
            End If

            RefreshPrices m.BarsColl(lIndex), bWorkingOrders, bTradeTracker
        End If
    Next lIndex
    
gdStopProfile 502
gdStartProfile 503

    m.Positions.RefreshPrices2 True
    If FormIsLoaded("frmOpenPositions") Then
        frmOpenPositions.RefreshPrices2 False
    End If

gdStopProfile 503
gdStartProfile 504
    
    m.Accounts.RefreshPrices
    If FormIsLoaded("frmAccounts") Then
        frmAccounts.RefreshPrices
    End If
    
gdStopProfile 504
gdStartProfile 505
    
    For lIndex = 1 To m.SpreadData.Count
        If m.SpreadData(lIndex).UpdateData Then
            RefreshSpreadData m.SpreadData(lIndex)
        End If
    Next lIndex
    
gdStopProfile 505
gdStartProfile 506
    
    If Not g.OrderStrategies Is Nothing Then
        ' 01/04/2010 DAJ: In the new world of Salmon, streaming could be completely connected, but
        ' we don't necessarily have data yet.  Keep calling activate exits here so that when we
        ' do have the data, the exit can be activated...
        g.OrderStrategies.ActivateExits True
        
        ' Make sure that any order strategies with break-even or trailing stops get updated...
        g.OrderStrategies.UpdateBars
        
        ' See if there are any auto exit orders that need to be modified...
        g.OrderStrategies.CheckModifyOrders
    End If
    
gdStopProfile 506
gdStartProfile 507
    
    ' Make sure that any active Trade Sense order groups get updated...
    If Not g.TsoGroups Is Nothing Then
        g.TsoGroups.UpdateBars
    End If
    
gdStopProfile 507
gdStartProfile 508
    
    ' Have each of the automated trading strategies update their data...
    'If Not g.TradingItems Is Nothing Then
    '    g.TradingItems.UpdateBars
    'End If
    
gdStopProfile 508
gdStartProfile 509
    
    g.CondOrders.CheckOrders
    
gdStopProfile 509
gdStartProfile 510
    
    ClearUpdatedColors
    
gdStopProfile 510
gdStartProfile 511
    
    If g.nReplaySession = 0 Then
        g.SimTradeStream.CheckQuickFillOrders
    Else
        g.SimTradeReplay.CheckQuickFillOrders
    End If
    
gdStopProfile 511
gdStartProfile 512
    
    ' Update the bars objects in the alerts object as well...
    g.Alerts.UpdateBars
    
gdStopProfile 512
gdStopProfile 500

    If TimerEnd("frmTTSummary.tmrRealTime", tmrRealtime.Interval) Then

'If DumpProfile Then
        DebugLog "=================" & vbCrLf & gdGetProfiles(500, 529, vbCrLf) & vbCrLf & "================="
'End If

    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.tmrRealTime_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrTbCaption_Timer
'' Description: Update the toolbar captions if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrTbCaption_Timer()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nForm As eGDConsoleForms        ' Form to update
    
    TimerStart "frmTTSummary.tmrTbCaption"
    tmrTbCaption.Enabled = False

    If m.TbCaption.Count > 0 Then
        tbToolbar.Redraw = False
        
        For lIndex = 1 To m.TbCaption.Count
            nForm = CLng(Val(m.TbCaption.Key(lIndex)))
            DoUpdateToolbarCaption nForm, m.TbCaption(lIndex)
        Next lIndex
        m.TbCaption.Clear
        
        tbToolbar.Redraw = True
    End If
    TimerEnd "frmTTSummary.tmrTbCaption", tmrTbCaption.Interval
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.tmrTbCaption_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls on the form as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrices
'' Description: Refresh the prices in the grids with the info in the Bars
'' Inputs:      Bars, Working Orders Loaded?, Trade Tracker Loaded?, Last Good Bar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshPrices(Bars As cGdBars, ByVal bWorkingOrders As Boolean, ByVal bTradeTracker As Boolean, Optional ByVal nLastGoodBar As Long = -1&)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the Bars structure
    Dim lSymbolID As Long               ' Symbol ID from the Bars structure
    Dim dPrice As Double                ' Current price from the Bars
    Dim dBid As Double                  ' Current bid from the Bars
    Dim dAsk As Double                  ' Current ask from the Bars
    Dim strPrice As String              ' Formatted version of the current price
    Dim strBid As String                ' Formatted version of the current bid
    Dim strAsk As String                ' Formatted version of the current ask
    
    ' Get the information out of the Bars structure once...
    strSymbol = Bars.Prop(eBARS_Symbol)
    lSymbolID = Bars.Prop(eBARS_SymbolID)
    If nLastGoodBar = -1& Then
        nLastGoodBar = Bars.Size - 1
        dPrice = Bars(eBARS_Close, Bars.Size - 1)
        dBid = Bars(eBARS_Bid, Bars.Size - 1)
        dAsk = Bars(eBARS_Ask, Bars.Size - 1)
    Else
        dPrice = Bars(eBARS_Close, nLastGoodBar)
        dBid = Bars(eBARS_Bid, nLastGoodBar)
        dAsk = Bars(eBARS_Ask, nLastGoodBar)
    End If
    
    strPrice = Bars.PriceDisplay(dPrice)
    strBid = Bars.PriceDisplay(dBid)
    strAsk = Bars.PriceDisplay(dAsk)
    
    m.Orders.RefreshPrices2 Bars.SymbolOrSymbolID, strPrice, strBid, strAsk
    
    If bWorkingOrders Then
        frmWorkingOrders.RefreshPrices2 Bars.SymbolOrSymbolID, strPrice, strBid, strAsk
    End If
        
    ' If the Trade Tracker form is up, refresh the prices there as well...
    If bTradeTracker Then
        frmTTPositions.RefreshPrices Bars
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.RefreshPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshSpreadData
'' Description: Refresh the spread prices in the grids
'' Inputs:      Spread data structure of current prices
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshSpreadData(SpreadData As cSpreadData)
On Error GoTo ErrSection:

    m.Orders.RefreshPrices SpreadData.Symbol, kNullData, SpreadData.Bid, SpreadData.Ask
    If FormIsLoaded("frmWorkingOrders") Then
        frmWorkingOrders.RefreshPrices SpreadData.Symbol, kNullData, SpreadData.Bid, SpreadData.Ask
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.RefreshSpreadData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAllGridFonts
'' Description: Set the font on all of the grids
'' Inputs:      String version of the Font
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAllGridFonts(ByVal strGridFont As String)
On Error GoTo ErrSection:

    FontFromString fgAccounts.Font, strGridFont
    fgAccounts.Font = fgAccounts.Font
    FontFromString fgPositions.Font, strGridFont
    fgPositions.Font = fgPositions.Font
    FontFromString fgOrders.Font, strGridFont
    fgOrders.Font = fgOrders.Font

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.SetAllGridFonts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoSizeGrid
'' Description: Figure out the perfect width for the given grid
'' Inputs:      Grid to size
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoSizeGrid(Grid As VSFlexGrid) As Long
On Error GoTo ErrSection:

    Dim lNonClient As Long              ' Non-client area of the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lColWidths As Long              ' Sum of the visible column widths
    Dim lNumVisible As Long             ' Number of visible columns
    
    With Grid
        lNonClient = .Width - .ClientWidth
        
        lColWidths = 0&
        lNumVisible = 0&
        
        For lIndex = 0 To .Cols - 1
            If .ColHidden(lIndex) = False Then
                lColWidths = lColWidths + .ColWidth(lIndex)
                lNumVisible = lNumVisible + 1
            End If
        Next lIndex
        
        AutoSizeGrid = lNonClient + lColWidths + (.GridLineWidth * lNumVisible)
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.AutoSizeGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceGrids
'' Description: Place and size the grids accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PlaceGrids()
On Error Resume Next ' (since called from Form_Resize)

    Dim lLeft As Long                   ' Left of the next grid
    Dim lRight As Long
    Dim lWidth As Long
    Dim lOrdersWidth As Long
    Dim lPosWidth As Long
    Dim lAccountsWidth As Long
    Dim lTotalWidth As Long
    Dim lAvailable As Long
    Dim lExtra As Long

    LockWindowUpdate Me.hWnd

    lLeft = 60
    lRight = ScaleWidth - 60
    lAvailable = lRight - lLeft
    
    lOrdersWidth = AutoSizeGrid(fgOrders)
    lPosWidth = AutoSizeGrid(fgPositions)
    lAccountsWidth = AutoSizeGrid(fgAccounts)
    lTotalWidth = lOrdersWidth + lPosWidth + lAccountsWidth + 120
    lExtra = lTotalWidth - lAvailable
    
    If lTotalWidth <> 0& Then
        If ScaleWidth * kAspectRatio > ScaleHeight Then
            With fgOrders
                .Redraw = flexRDNone
                .ScrollBars = flexScrollBarVertical
                .Left = lLeft
                If lTotalWidth <= lAvailable Then
                    lWidth = lOrdersWidth + (Abs(lExtra) / 2)
                Else
                    lWidth = lOrdersWidth - (lExtra * (lOrdersWidth / lTotalWidth))
                End If
                If lWidth < 0 Then lWidth = 0
                
                .Move lLeft, 60, lWidth, ScaleHeight - 120
                .Redraw = flexRDBuffered
                
                .Refresh
            End With
            
            lLeft = lLeft + lWidth + 60
            
            With fgPositions
                .Redraw = flexRDNone
                .ScrollBars = flexScrollBarVertical
                .Left = lLeft
                If lTotalWidth <= lAvailable Then
                    lWidth = lPosWidth + (Abs(lExtra) / 2)
                Else
                    lWidth = lPosWidth - (lExtra * (lPosWidth / lTotalWidth))
                End If
                If lWidth < 0 Then lWidth = 0
                
                .Move lLeft, 60, lWidth, ScaleHeight - 120
                .Redraw = flexRDBuffered
                
                .Refresh
            End With
        
            lLeft = lLeft + lWidth + 60
            
            With fgAccounts
                .Redraw = flexRDNone
                .ScrollBars = flexScrollBarVertical
                .Left = lLeft
                If lTotalWidth <= lAvailable Then
                    lWidth = lAccountsWidth
                Else
                    lWidth = lRight - lLeft
                End If
                If lWidth < 0 Then lWidth = 0
                
                .Move lLeft, 60, lWidth, ScaleHeight - 120
                .Redraw = flexRDBuffered
                
                .Refresh
            End With
        Else
            With fgOrders
                .ScrollBars = flexScrollBarBoth
                .Move lLeft, 60, ScaleWidth - (lLeft * 2), (ScaleHeight - (60 * 2)) / 3
            End With
            
            With fgPositions
                .ScrollBars = flexScrollBarBoth
                .Move lLeft, fgOrders.Top + fgOrders.Height, ScaleWidth - (lLeft * 2), fgOrders.Height
            End With
            
            With fgAccounts
                .ScrollBars = flexScrollBarBoth
                .Move lLeft, fgPositions.Top + fgPositions.Height, ScaleWidth - (lLeft * 2), fgOrders.Height
            End With
        End If
    End If

    LockWindowUpdate 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load and splice the bars for a particular symbol or symbol ID
'' Inputs:      Bars to Load, Symbol or Symbol ID to load
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadBars(Bars As cGdBars, ByVal vSymbolOrSymbolID As Variant, Optional ByVal bAddToRT As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value
    Dim Bars2 As cGdBars                ' Bars from the quote board if exist
    
    ' Check to see first if the symbol exists on the quote board...
    Set Bars2 = frmQuotes.GetBars(Str(vSymbolOrSymbolID), "Daily")
    If Bars2 Is Nothing Then
        Bars.ArrayMask = eBARS_EodBidAsk
        bReturn = DM_GetBars(Bars, vSymbolOrSymbolID, , LastDailyDownload - 5, , , False)
    ElseIf Bars2.Size > 0 Then
        Set Bars = Bars2.MakeCopy
        
        ' Make sure to return True here, otherwise the RefreshData routine doesn't update the
        ' bars collection correctly (Aardvark Issue #4419)...
        bReturn = True
    Else
        Bars.ArrayMask = eBARS_EodBidAsk
        bReturn = DM_GetBars(Bars, vSymbolOrSymbolID, , LastDailyDownload - 5, , , False)
    End If
    
    If bAddToRT = True Then
        g.RealTime.AddTickBuffer Bars
        g.RealTime.SpliceBars Bars, , True
    End If
    
    LoadBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.LoadBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSpreadData
'' Description: Load and splice the bars for each component of a spread
'' Inputs:      Spread Data to Load, Spread Symbol to load
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadSpreadData(SpreadData As cSpreadData, ByVal strSymbol As String, Optional ByVal bAddToRT As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars object for each component
    Dim lIndex As Long                  ' Index into a for loop

    Set SpreadData = New cSpreadData
    SpreadData.Symbol = strSymbol
    
    For lIndex = 0 To SpreadData.NumLegs - 1
        Set Bars = GetBars(SpreadData.LegSymbols(lIndex), bAddToRT)
    Next lIndex
    
    SpreadData.UpdateData

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.LoadSpreadData"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridToArray
'' Description: Dump a grid to an array
'' Inputs:      Grid
'' Returns:     Array
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridToArray(Grid As VSFlexGrid) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array to return from the function
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strLine As String               ' Line in the grid
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    With Grid
        If .Rows = .FixedRows Then
            astrReturn.Add "No Rows"
        Else
            For lRow = 0 To .Rows - 1
                If .RowHidden(lRow) = True Then
                    strLine = "Hidden"
                Else
                    strLine = "Visible"
                End If
                
                For lCol = 0 To .FrozenCols - 1
                    strLine = strLine & vbTab & .TextMatrix(lRow, lCol)
                Next lCol
                If UCase(Grid.Name) = "FGORDERS" Then
                    strLine = strLine & vbTab & .TextMatrix(lRow, m.Orders.StatusCol)
                End If
                
                astrReturn.Add strLine
            Next lRow
        End If
    End With
    
    Set GridToArray = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.GridToArray"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpGridIfDifferent
'' Description: Dump the grid to the appropriate log if it changed
'' Inputs:      Grid Before Update, Grid After Update, Grid Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpGridIfDifferent(ByVal astrBefore As cGdArray, ByVal astrAfter As cGdArray, ByVal strGridName As String)
On Error GoTo ErrSection:

    Dim bDump As Boolean                ' Do we want to dump the grid?
    Dim lIndex As Long                  ' Index into a for loop
    
    If astrBefore.Size <> astrAfter.Size Then
        bDump = True
    Else
        bDump = False
        For lIndex = 0 To astrBefore.Size - 1
            If astrBefore(lIndex) <> astrAfter(lIndex) Then
                bDump = True
                Exit For
            End If
        Next lIndex
    End If
    
    If bDump = True Then
        m.GridDump.Add strGridName & ": " & vbCrLf & vbTab & astrAfter.JoinFields(vbCrLf & vbTab)
        tmrGridDump.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.DumpGridIfDifferent"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with
    fh = FreeFile
    Open AddSlash(App.Path) & "TradeConsole\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(App.Path) & "TradeConsole\TN*.LOG"
    End If
    LogFile.WriteText strMessage

#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixSummaryDisplay
'' Description: Fix the summary display and split up
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixSummaryDisplay()
On Error GoTo ErrSection:

    Dim strDisplay As String            ' Display string from the INI file
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrDisplay As New cGdArray     ' Display string split into an array
    Dim astrLine As New cGdArray        ' Display line
    Dim lOrders As Long                 ' Index of the orders grid in array
    Dim lPositions As Long              ' Index of the positions grid in array
    Dim lAccounts As Long               ' Index of the accounts grid in array
    Dim strOrders As String             ' Orders display string
    Dim strPositions As String          ' Positions display string
    Dim strAccounts As String           ' Accounts display string
    Dim lGrid As Long                   ' Grid currently working on
    
    strDisplay = GetIniFileProperty("SummaryOrdersDisplay", "", "TTSummary", g.strIniFile)
    If Len(strDisplay) = 0 Then
        Set astrDisplay = New cGdArray
        Set astrLine = New cGdArray
    
        strDisplay = GetIniFileProperty("Display", "", "TTSummary", g.strIniFile)
        If Len(strDisplay) = 511 Then
            If GetIniFileProperty("Reset", 0&, "TTSummary", g.strIniFile) = 0& Then
                strDisplay = ""
                SetIniFileProperty "Reset", 1&, "TTSummary", g.strIniFile
            End If
        End If
        
        If Len(strDisplay) > 0 Then
            strDisplay = Replace(strDisplay, ";Open Position;", ";Position;")
            If InStr(strDisplay, ";X;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Working Order" Then
                        astrDisplay.Add "1;X;1;" & Str(lIndex + 2 + 1), lIndex + 1
                        Exit For
                    End If
                Next lIndex
                For lIndex = 1 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Open Positions" Then Exit For
                    astrLine(3) = Str(lIndex) + 2
                    astrDisplay(lIndex) = astrLine.JoinFields(";")
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Auto Exit;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Accounts" Then
                        astrDisplay.Add "1;Auto Exit;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Auto Exit;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Accounts" Then
                        astrDisplay.Add "1;Auto Exit;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Last Traded;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Accounts" Then
                        astrDisplay.Add "1;Last Traded;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Session Date;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Accounts" Then
                        astrDisplay.Add "1;Session Date;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Session Qty;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Accounts" Then
                        astrDisplay.Add "1;Session Qty;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Session Profit;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Accounts" Then
                        astrDisplay.Add "1;Session Profit;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            If InStr(strDisplay, ";Link;") = 0 Then
                astrDisplay.SplitFields strDisplay, ","
                For lIndex = 0 To astrDisplay.Size - 1
                    astrLine.SplitFields astrDisplay(lIndex), ";"
                    If astrLine(1) = "Open Positions" Then
                        astrDisplay.Add "1;Link;1;" & Str(CLng(Parse(astrDisplay(lIndex - 1), ";", 4)) + 1), lIndex
                        Exit For
                    End If
                Next lIndex
                strDisplay = astrDisplay.JoinFields(",")
            End If
            
            astrDisplay.SplitFields strDisplay, ","
            For lIndex = 0 To astrDisplay.Size - 1
                astrLine.SplitFields astrDisplay(lIndex), ";"
                Select Case astrLine(1)
                    Case "Open Orders"
                        lOrders = lIndex
                    Case "Open Positions"
                        lPositions = lIndex
                    Case "Accounts"
                        lAccounts = lIndex
                End Select
            Next lIndex
            
            For lIndex = astrDisplay.Size - 1 To lAccounts + 1 Step -1
                Select Case UCase(Parse(astrDisplay(lIndex), ";", 2))
                    Case "ACCOUNT ID", "ACCOUNT TYPE", "REMOVE", "ON"
                        astrDisplay.Remove lIndex
                End Select
            Next lIndex
            
            For lIndex = lAccounts - 1 To lPositions + 1 Step -1
                Select Case UCase(Parse(astrDisplay(lIndex), ";", 2))
                    Case "POSITION ID", "ACCOUNT ID", "SYMBOL ID", "SYMBOL", "ACCOUNT", "AUTO TRADE ITEM", "POSITION", "ACCOUNT TYPE", "REMOVE"
                        astrDisplay.Remove lIndex
                End Select
            Next lIndex
            
            For lIndex = lPositions - 1 To lOrders + 1 Step -1
                Select Case UCase(Parse(astrDisplay(lIndex), ";", 2))
                    Case "ORDER ID", "ACCOUNT ID", "SYMBOL ID", "OPEN ORDERS", "SYMBOL", "WORKING ORDER", "X", "ACCOUNT TYPE", "REMOVE"
                        astrDisplay.Remove lIndex
                End Select
            Next lIndex
            
            For lIndex = lPositions + 1 To astrDisplay.Size - 1
                astrLine.SplitFields astrDisplay(lIndex), ";"
                If astrLine(1) = "Accounts" Then
                    Exit For
                Else
                    astrLine(3) = Str(7 + (lIndex - (lPositions + 1)))
                    astrDisplay(lIndex) = astrLine.JoinFields(";")
                End If
            Next lIndex
            
            lGrid = -1&
            For lIndex = 0 To astrDisplay.Size - 1
                If UCase(Parse(astrDisplay(lIndex), ";", 2)) = "OPEN ORDERS" Then
                    lGrid = 0&
                ElseIf UCase(Parse(astrDisplay(lIndex), ";", 2)) = "OPEN POSITIONS" Then
                    lGrid = 1&
                ElseIf UCase(Parse(astrDisplay(lIndex), ";", 2)) = "ACCOUNTS" Then
                    lGrid = 2&
                ElseIf lGrid = 0& Then
                    strOrders = strOrders & astrDisplay(lIndex) & ","
                ElseIf lGrid = 1& Then
                    strPositions = strPositions & astrDisplay(lIndex) & ","
                ElseIf lGrid = 2& Then
                    strAccounts = strAccounts & astrDisplay(lIndex) & ","
                End If
            Next lIndex
            
            SetIniFileProperty "SummaryOrdersDisplay", Left(strOrders, Len(strOrders) - 1), "TTSummary", g.strIniFile
            SetIniFileProperty "SummaryPositionsDisplay", Left(strPositions, Len(strPositions) - 1), "TTSummary", g.strIniFile
            SetIniFileProperty "SummaryAccountsDisplay", Left(strAccounts, Len(strAccounts) - 1), "TTSummary", g.strIniFile
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTSummary.FixSummaryDisplay"
    
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

    tmrMenu.Tag = kMenuPrefix & strCommand
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToolForForm
'' Description: Get the toolbar identifier for the given form
'' Inputs:      Form
'' Returns:     Tool ID (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToolForForm(ByVal nForm As eGDConsoleForms) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nForm
        Case eGDConsoleForm_Summary
            strReturn = "ID_Summary"
        Case eGDConsoleForm_OpenOrders
            strReturn = "ID_OpenOrders"
        Case eGDConsoleForm_Positions
            strReturn = "ID_Positions"
        Case eGDConsoleForm_Accounts
            strReturn = "ID_Accounts"
        Case eGDConsoleForm_AutoTrading
            strReturn = "ID_AutoTrading"
        Case eGDConsoleForm_ActivityLog
            strReturn = "ID_ActivityLog"
        Case eGDConsoleForm_TodaysFills
            strReturn = "ID_TodaysFills"
        Case eGDConsoleForm_TradeSenseOrders
            strReturn = "ID_TradeSenseOrders"
    End Select
    
    ToolForForm = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.ToolForForm"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CaptionForForm
'' Description: Get the toolbar caption for the given form
'' Inputs:      Form
'' Returns:     Caption (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CaptionForForm(ByVal nForm As eGDConsoleForms) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nForm
        Case eGDConsoleForm_Summary
            strReturn = "Dashboard"
        Case eGDConsoleForm_OpenOrders
            strReturn = "Open Orders"
        Case eGDConsoleForm_Positions
            strReturn = "Positions"
        Case eGDConsoleForm_Accounts
            strReturn = "Accounts"
        Case eGDConsoleForm_AutoTrading
            strReturn = "Auto Trading"
        Case eGDConsoleForm_ActivityLog
            strReturn = "Activity Log"
        Case eGDConsoleForm_TodaysFills
            strReturn = "Todays Fills"
        Case eGDConsoleForm_TradeSenseOrders
            strReturn = "TradeSense Orders"
    End Select
    
    strReturn = TranslatedText(strReturn, strReturn)

    CaptionForForm = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTSummary.CaptionForForm"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoUpdateToolbarCaptions
'' Description: Update the toolbar captions with the number of visible items
'' Inputs:      Form, Number of visible items
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoUpdateToolbarCaption(ByVal nForm As eGDConsoleForms, ByVal lNumVisible As Long)
On Error GoTo ErrSection:

    Dim bColorRed As Boolean            ' Color the caption red?
    Dim bRedraw As Boolean              ' Current state of the toolbar's redraw

    bRedraw = tbToolbar.Redraw
    tbToolbar.Redraw = False
    
    If lNumVisible > 0 Then
        bColorRed = True
        If nForm = eGDConsoleForm_Accounts Then
            bColorRed = False
        ElseIf nForm = eGDConsoleForm_AutoTrading Then
            If g.TradingItems Is Nothing Then
                bColorRed = False
            ElseIf g.TradingItems.HasActiveAutoTradeItems = False Then
                bColorRed = False
            End If
        End If
        
        ' Don't change the caption for the summary or accounts buttons...
        If (nForm <> eGDConsoleForm_Summary) And (nForm <> eGDConsoleForm_Accounts) Then
            tbToolbar.Tools(ToolForForm(nForm)).ChangeAll ssChangeAllName, CaptionForForm(nForm) & " (" & Str(lNumVisible) & ")"
        End If
    Else
        bColorRed = False
        
        ' Don't change the caption for the summary or accounts buttons...
        If (nForm <> eGDConsoleForm_Summary) And (nForm <> eGDConsoleForm_Accounts) Then
            tbToolbar.Tools(ToolForForm(nForm)).ChangeAll ssChangeAllName, CaptionForForm(nForm)
        End If
    End If
    
    If bColorRed Then
        tbToolbar.ToolBars("General").Tools(ToolForForm(nForm)).ForeColor = vbRed
    ElseIf g.nColorTheme = kDarkThemeColor Then
        tbToolbar.ToolBars("General").Tools(ToolForForm(nForm)).ForeColor = vbWhite
    Else
        tbToolbar.ToolBars("General").Tools(ToolForForm(nForm)).ForeColor = &H80000012
    End If
    
    tbToolbar.Redraw = bRedraw
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.DoUpdateToolbarCaptions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowToolbarButtons
'' Description: Show/Hide the toolbar buttons as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowToolbarButtons()
On Error GoTo ErrSection:

    Dim TcBtns As cTradeConsoleButtons  ' Trade console button settings
    
    Set TcBtns = New cTradeConsoleButtons
    TcBtns.Load
    
    With tbToolbar
        .Tools("ID_Summary").Visible = TcBtns.Show(eGDTcButtons_Dashboard)
        .Tools("ID_OpenOrders").Visible = TcBtns.Show(eGDTcButtons_OpenOrders)
        .Tools("ID_Positions").Visible = TcBtns.Show(eGDTcButtons_Positions)
        .Tools("ID_Accounts").Visible = TcBtns.Show(eGDTcButtons_Accounts)
        .Tools("ID_AutoTrading").Visible = TcBtns.Show(eGDTcButtons_AutoTrading)
        .Tools("ID_TradeSenseOrders").Visible = TcBtns.Show(eGDTcButtons_TradeSenseOrders)
        .Tools("ID_ActivityLog").Visible = TcBtns.Show(eGDTcButtons_ActivityLog)
        .Tools("ID_TodaysFills").Visible = TcBtns.Show(eGDTcButtons_TodaysFills)
        .Tools("ID_Buy").Visible = TcBtns.Show(eGDTcButtons_BuySell)
        .Tools("ID_Sell").Visible = TcBtns.Show(eGDTcButtons_BuySell)
        .Tools("ID_Settings").Visible = TcBtns.Show(eGDTcButtons_Settings)
        .Tools("ID_Tracking").Visible = TcBtns.Show(eGDTcButtons_Tracking)
        .Tools("ID_Reports").Visible = TcBtns.Show(eGDTcButtons_Reports)
        .Tools("ID_Journals").Visible = TcBtns.Show(eGDTcButtons_Journals)
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummary.ShowToolbarButtons"
    
End Sub

