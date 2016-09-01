VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmLots 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrItemsSent 
      Left            =   11040
      Top             =   1440
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   11040
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   22
      Tools           =   "frmLots.frx":0000
      ToolBars        =   "frmLots.frx":205E
   End
   Begin VB.Timer tmrRealtime 
      Left            =   11040
      Top             =   960
   End
   Begin VB.Timer tmrMenu 
      Left            =   11040
      Top             =   480
   End
   Begin VSFlex7LCtl.VSFlexGrid fgLots 
      Height          =   1515
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   2895
      _cx             =   5106
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
   Begin HexUniControls.ctlUniFrameWL fraAccounts 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
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
      Caption         =   "frmLots.frx":22BD
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLots.frx":22E9
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLots.frx":2309
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdAddFeedYard 
         Height          =   315
         Left            =   4500
         TabIndex        =   1
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLots.frx":2325
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLots.frx":2361
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":2381
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddCustomer 
         Height          =   315
         Left            =   8760
         TabIndex        =   2
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLots.frx":239D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLots.frx":23D9
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":23F9
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboView 
         Height          =   315
         Left            =   9780
         TabIndex        =   4
         Top             =   0
         Width           =   1875
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
         Tip             =   "frmLots.frx":2415
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":2435
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   315
         Left            =   8340
         TabIndex        =   6
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLots.frx":2451
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLots.frx":2483
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":24A3
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboCustomers 
         Height          =   315
         Left            =   5820
         TabIndex        =   5
         Top             =   0
         Width           =   2475
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
         Tip             =   "frmLots.frx":24BF
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":24DF
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboFeedYards 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   0
         Width           =   2115
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
         Tip             =   "frmLots.frx":24FB
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":251B
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblView 
         Height          =   195
         Left            =   9300
         Top             =   60
         Width           =   495
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
         Caption         =   "frmLots.frx":2537
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLots.frx":2563
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":2583
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStatus 
         Height          =   195
         Left            =   240
         Top             =   60
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
         Caption         =   "frmLots.frx":259F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLots.frx":25D9
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":25F9
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblCustomers 
         Height          =   195
         Left            =   4980
         Top             =   60
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
         Caption         =   "frmLots.frx":2615
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLots.frx":264B
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":266B
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFeedYards 
         Height          =   195
         Left            =   1500
         Top             =   60
         Width           =   915
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
         Caption         =   "frmLots.frx":2687
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLots.frx":26BF
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLots.frx":26DF
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Image imgStatus 
         Height          =   195
         Left            =   0
         Picture         =   "frmLots.frx":26FB
         Top             =   60
         Width           =   195
      End
   End
   Begin VB.Menu mnuLots 
      Caption         =   "Lots"
      Begin VB.Menu mnuPriceLadder 
         Caption         =   "Price Ladder"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddLot 
         Caption         =   "Add Lot"
      End
      Begin VB.Menu mnuEditLot 
         Caption         =   "Edit Lot"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseOrder 
         Caption         =   "Close Order"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandAll 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuCollapseAll 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuEditColumns 
         Caption         =   "Edit Columns"
      End
      Begin VB.Menu mnuShowClosedLots 
         Caption         =   "Show Closed Lots"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExportToCsv 
         Caption         =   "Export to CSV"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "frmLots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLots.frm
'' Description: Form for allowing user to setup and view Turnkey information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/05/2012   DAJ         Set Lot Number & Symbol to links, use front month LE as default for orders
'' 06/05/2012   DAJ         Call mTradeTracker.CreateOrder instead of frmTTEditOrder.ShowMe
'' 06/05/2012   DAJ         Fix for clicking anywhere in lot row, persisting column info
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/12/2012   DAJ         Fix for market order w/ fill waiting for account and order
'' 06/14/2012   DAJ         Don't underline profits on symbols row, handle account ID in delete order/fill
'' 06/25/2012   DAJ         Column sorting, Moving columns, Feedyard Hidden Columns, Bug fixes
'' 08/16/2012   DAJ         Added display templates for Turnkey form
'' 09/11/2012   DAJ         Associate accounts, associate parts of fills with lot
'' 09/13/2012   DAJ         When get a fill from broker, don't wait for Turnkey Order confirmation
'' 09/13/2012   DAJ         Added a Refresh button to the toolbar
'' 09/14/2012   DAJ         Visible Lot Column by Genesis Customer not Feedyard Customer
'' 09/18/2012   DAJ         Don't allow associations while associations pending
'' 09/20/2012   DAJ         Changed DeleteTrades to DeleteTrade (only delete one trade)
'' 09/26/2012   DAJ         Added T+/T- toolbar buttons to control font size
'' 10/22/2012   DAJ         Rename Turnkey to HedgeLinc, UI mods
'' 10/23/2012   DAJ         When updating account because of an order, set the FeedYardID
'' 10/26/2012   DAJ         Fix for scrolling/streaming bug, show lot number as tooltip text
'' 10/26/2012   DAJ         Fix for export file not having CSV extension, focus issue when click in non-client grid
'' 10/26/2012   DAJ         Re-filter grid after Expand All or Collapse All
'' 01/09/2013   DAJ         Added right-click option for price ladder
'' 01/30/2013   DAJ         Live/Demo/Test modes for Turnkey
'' 01/31/2013   DAJ         Show all accounts from broker view form when associating
'' 02/05/2013   DAJ         Left justify Lot Number column, clear InfBoxes when form unloaded
'' 10/14/2013   DAJ         Timer logs; Don't accept event if unloading; Print defaults
'' 10/15/2013   DAJ         Multi-row header
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
'' 11/15/2013   DAJ         Mods for adding/editing feedyard/lots/customers
'' 11/20/2013   DAJ         Reports
'' 11/22/2013   DAJ         Import historical fills for Turnkey
'' 11/25/2013   DAJ         Show error for reports if feedyard not selected or no lots
'' 11/26/2013   DAJ         CanEditLots per feedyard; Lot column display formats
'' 12/03/2013   DAJ         Expand/Collapse level; Turnkey Mode; Use frmTurnkeySelect for accounts
'' 12/04/2013   DAJ         Detail Options
'' 12/19/2013   DAJ         "Lauren List" tweaks
'' 01/16/2014   DAJ         Fix for multiple lots with same Lot Number
'' 01/23/2014   DAJ         Multiple owners per lot
'' 01/27/2014   DAJ         Implemented "End" flag to have server tell us not to continue state machine
'' 01/31/2014   DAJ         Manage feedyards and feedyard customers
'' 02/10/2014   DAJ         Associate fill label notification
'' 02/25/2014   DAJ         Rations/Ingredients; Commissions on fills; Remove associate label
'' 02/26/2014   DAJ         Double click in grid brings up lot editor on correct field
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/10/2014   DAJ         Use SetWindowPos instead of ShowForm now that we are in DLL
'' 03/13/2014   DAJ         Allow brokers to associate carried fill information
'' 03/14/2014   DAJ         Added support for Boolean lot column type
'' 03/19/2014   DAJ         Changed back to non-modal; Fix for first customer of new feedyard;
''                          Right-click/Edit lot the same as double-click; Allow hide for open
''                          equity/closed profit; Change background color for closed lots
'' 03/20/2014   DAJ         Tweaked print header for all customers; Fixed open equity and closed
''                          profit not being totalled up; Blank out open equity if current price null
'' 03/21/2014   DAJ         Fix for fields showing up in edit columns that shouldn't
'' 04/08/2014   DAJ         Copied the Grid Scroll fix into DLL from NavSuite project to fix error
'' 04/15/2014   DAJ         Feedyard Customers; New owner lookup form; Syncrhonize ingredients;
''                          Clicking on lot closed column; Persist open equity/closed profit visiblity
'' 04/28/2014   DAJ         Fix for infinite loop checking for parent in SelectedLot; Added
''                          RowType column; Ability to Close order; Refresh orders after connect
'' 05/08/2014   DAJ         Possible fix for 'Invalid property array index' error in LotColumnForCol
'' 05/15/2014   DAJ         Implemented DeletePosition callback from server
'' 05/22/2014   DAJ         Renamed frmTurnkeyEditLot to frmEditLot; Renamed frmTurnkeyLotContentDetails
''                          to frmEditLotContentDetails; Renamed frmTurnkeyReport to frmCattleReport;
''                          Renamed frmTurnkeyReport to frmCattleReport; Renamed frmTurnkeyManage to
''                          frmCattleManage
'' 05/22/2014   DAJ         Renamed frmTurnkey to frmLots; Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Allow user to create a new manual fill and associate it to a lot; Utilize
''                          new accounts object; Position per symbol and account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDLotCols
    eGDLotCol_LotNumber                 ' Lot number
    eGDLotCol_LotStatus                 ' Lot status
    eGDLotCol_OpenEquity                ' Open equity
    eGDLotCol_ClosedProfit              ' Closed profit
    eGDLotCol_AscSortKey                ' Ascending sort key
    eGDLotCol_DescSortKey               ' Descending sort key
    eGDLotCol_SortKey                   ' Sort key
    eGDLotCol_OutlineLevel              ' Outline level for the row
    eGDLotCol_IsClosed                  ' Is the lot closed?
    eGDLotCol_RowType                   ' Type for the row
    eGDLotCol_NumCols
End Enum

Private Enum eGDLotItemType
    eGDLotItemType_Position = 1
    eGDLotItemType_Order
    eGDLotItemType_Fill
    eGDLotItemType_Trade
End Enum

Private Enum eGDRowType
    eGDRowType_Lot = 1
    eGDRowType_Order
    eGDRowType_Fill
    eGDRowType_Trade
    eGDRowType_Position
    eGDRowType_Total
    eGDRowType_Associate
    eGDRowType_Symbol
    eGDRowType_NoPosition
End Enum

Private Enum eGDLotView
    eGDLotView_AllLots = 0
    eGDLotView_LotsWithHedge = 1
    eGDLotView_LotsWithoutHedge = 2
End Enum

Private Type mPrivate
    nStatus As eGDConnectionStatus      ' Connection status to the Genesis Turnkey server
    bDoneOnce As Boolean                ' Have we done the one-time code?
    lSortedCol As Long                  ' Sorted column
    nSortedDir As SortSettings          ' Sort direction for the column
    nBackColor As OLE_COLOR             ' Normal back color of the grid
    bLoadingColumns As Boolean          ' Are we currently moving columns?
    lTemplateNumber As Long             ' Number of the currently selected template
    dLastItemSent As Double             ' Tick count when the last item was sent
    lGridFontSize As Long               ' Grid font size
    bFormUnloading As Boolean           ' Is the form unloading?
    bCanEditLots As Boolean             ' Can the user edit lots?
    lExpandLevel As Long                ' Current expand level
    iButtonPressed As Integer           ' Mouse button pressed
    bHasDetails As Boolean              ' Do we have lot content details?
    bHasIngredients As Boolean          ' Do we have ingredients?
    
    LotColumns As cGdTree               ' Collection of lot columns
    KeyValueToIndex As cGdTree          ' Index going from key value name to collection index
    KeyValueToCol As cGdTree            ' Index going from key value name to column number
    alLotCol As cGdArray                ' Index for specific lot columns
    
    astrCustomers As cGdArray           ' List of customers
    Details As cGdTree                  ' Collection of lot content details
    
    Accounts As cGdTree                 ' Collection of accounts
    AssociatedOrders As cGdTree         ' Collection of associated orders
    AssociatedFills As cGdTree          ' Collection of associated fills
    Fills As cGdTree                    ' Collection of all fills for the customer
    Trades As cGdTree                   ' Collection of all trades for the customer
    WaitAccounts As cGdTree             ' Collection of accounts waiting for confirmation fro Tunkey servers
    NotConnected As cGdTree             ' Collection of accounts that are not connected
    WaitAccountOrders As cGdTree        ' Collection of orders waiting for Turnkey Account ID
    WaitAccountFills As cGdTree         ' Collection of fills waiting for Turnkey Account ID
    WaitOrderOrders As cGdTree          ' Collection of orders waiting for Turnkey Order ID
    WaitOrderFills As cGdTree           ' Collection of fills waiting for Turnkey Order ID
    WaitFillFills As cGdTree            ' Collection of fills waiting for Turnkey Fill ID
    NewOrders As cGdTree                ' Collection of new orders waiting for a broker ID
    OrdersSent As cGdTree               ' Collection of orders sent to Genesis Turnkey server for association
    AssociatedOrdersSent As cGdTree     ' Collection of associated orders sent to Genesis Turnkey server for association
    AssociatedFillsSent As cGdTree      ' Collection of associated fills sent to Genesis Turnkey server for association
    VisibleLots As cGdTree              ' Collection of visible lots per customer
    WaitDetails As cGdTree              ' Collection of lot details waiting to be sent
End Type
Private m As mPrivate

Private Property Get LotCol(ByVal nCol As eGDLotCols) As Long
    LotCol = m.alLotCol(nCol)
End Property

Private Property Get LotView(ByVal nView As eGDLotView) As Long
    LotView = nView
End Property

Private Property Get NumberOfCols() As Long
    NumberOfCols = m.LotColumns.Count
End Property

Public Property Get LotColumns() As cGdTree
    Set LotColumns = m.LotColumns
End Property

Public Property Get SelectedFeedYard() As Long
    If cboFeedYards.ListIndex > -1& Then
        SelectedFeedYard = cboFeedYards.ItemData(cboFeedYards.ListIndex)
    Else
        SelectedFeedYard = -1&
    End If
End Property

Private Property Get SelectedCustomer() As Long
    If cboCustomers.ListIndex > -1& Then
        SelectedCustomer = cboCustomers.ItemData(cboCustomers.ListIndex)
    Else
        SelectedCustomer = -2&
    End If
End Property

Public Property Get Status() As eGDConnectionStatus
    Status = m.nStatus
End Property
Public Property Let Status(ByVal nStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            imgStatus.Picture = frmCattleAM.imgRed
            lblStatus.Caption = "Disconnected"
            
            If m.nStatus <> eGDConnectionStatus_Disconnected Then
                ' Clear form?
            End If
            
        Case eGDConnectionStatus_Disconnecting
            imgStatus.Picture = frmCattleAM.imgYellow
            lblStatus.Caption = "Disconnecting"
            
        Case eGDConnectionStatus_Connecting
            imgStatus.Picture = frmCattleAM.imgYellow
            lblStatus.Caption = "Connecting"
            
        Case eGDConnectionStatus_Connected
            imgStatus.Picture = frmCattleAM.imgGreen
            lblStatus.Caption = "Connected"
            
            If m.nStatus <> eGDConnectionStatus_Connected Then
                GetLotColumnCategories
                RefreshTurnkey
            End If
                
    End Select

    If nStatus <> m.nStatus Then
        g.Cattle.DumpDebug "Connection status changed to " & lblStatus.Caption
        m.nStatus = nStatus
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmLots.Status.Let"
    
End Property

Private Property Get TemplateNumber() As Long
    TemplateNumber = m.lTemplateNumber
End Property
Private Property Let TemplateNumber(ByVal lTemplateNumber As Long)
    If lTemplateNumber <> m.lTemplateNumber Then
        m.lTemplateNumber = lTemplateNumber
        SetIniFileProperty "TemplateNumber", lTemplateNumber, "Turnkey", g.strIniFile
    End If
End Property

Private Property Get NumTemplates() As Long
    NumTemplates = GetIniFileProperty("NumTemplates", 0&, "Turnkey", g.strIniFile)
End Property
Private Property Let NumTemplates(ByVal lNumTemplates As Long)
    SetIniFileProperty "NumTemplates", lNumTemplates, "Turnkey", g.strIniFile
End Property

Private Property Get TemplateName(ByVal lIndex As Long) As String
    TemplateName = GetIniFileProperty("Name" & Str(lIndex), "", "Turnkey", g.strIniFile)
End Property
Private Property Let TemplateName(ByVal lIndex As Long, ByVal strTemplateName As String)
    SetIniFileProperty "Name" & Str(lIndex), Trim(strTemplateName), "Turnkey", g.strIniFile
End Property

Private Property Get TemplateDisplay(ByVal lIndex As Long) As String
    TemplateDisplay = GetIniFileProperty("Display" & Str(lIndex), "", "Turnkey", g.strIniFile)
End Property
Private Property Let TemplateDisplay(ByVal lIndex As Long, ByVal strTemplateDisplay As String)
    SetIniFileProperty "Display" & Str(lIndex), Trim(strTemplateDisplay), "Turnkey", g.strIniFile
End Property

Private Property Get Dirty() As Boolean
    Dirty = tbToolbar.Tools("ID_SaveTemplate").Enabled
End Property
Private Property Let Dirty(ByVal bDirty As Boolean)
    tbToolbar.Tools("ID_SaveTemplate").Enabled = bDirty
End Property

Private Property Get AccountKey(ByVal strBroker As String, ByVal strAccountNumber As String) As String
    AccountKey = strBroker & "|" & strAccountNumber
End Property

Private Property Get OrderKey(ByVal strBroker As String, ByVal strBrokerOrderID As String) As String
    OrderKey = strBroker & "|" & strBrokerOrderID
End Property

Private Property Get FillKey(ByVal strBroker As String, ByVal strBrokerFillID As String) As String
    FillKey = strBroker & "|" & strBrokerFillID
End Property

Private Property Get AssociatedFillKey(ByVal strBroker As String, ByVal strBrokerFillID As String, ByVal strFeedYardLotID As String) As String
    AssociatedFillKey = strBroker & "|" & strBrokerFillID & "|" & strFeedYardLotID
End Property

Private Property Get GridFontSize() As Long
    GridFontSize = m.lGridFontSize
End Property
Private Property Let GridFontSize(ByVal lGridFontSize As Long)
On Error GoTo ErrSection:

    If lGridFontSize <= 8 Then
        m.lGridFontSize = 8
    Else
        m.lGridFontSize = lGridFontSize
    End If
    
    With fgLots
        .Redraw = flexRDNone
        
        .FontSize = lGridFontSize
        AutoSizeGrid
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmLots.GridFontSize.Let"
    
End Property

Private Property Get RowType(ByVal lRow As Long) As eGDRowType
    If ValidGridRow(fgLots, lRow) Then
        RowType = CLng(Val(fgLots.TextMatrix(lRow, LotCol(eGDLotCol_RowType))))
    Else
        RowType = -1&
    End If
End Property
Private Property Let RowType(ByVal lRow As Long, ByVal nRowType As eGDRowType)
    If ValidGridRow(fgLots, lRow) Then
        fgLots.TextMatrix(lRow, LotCol(eGDLotCol_RowType)) = Str(nRowType)
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Run as Turnkey?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal bTurnkey As Boolean)
On Error GoTo ErrSection:

    Dim strDisplay As String            ' Display string
    
    g.Cattle.Turnkey = bTurnkey
    Caption = g.Cattle.ProductName & " Feed Lot Information"
    InitLotsGrid
    
    If Not g.Cattle Is Nothing Then
        Status = g.Cattle.ConnectionStatus
        
        If Status = eGDConnectionStatus_Disconnected Then
            g.Cattle.Connect
        End If
    End If

    ' DAJ 03/10/2014: Showing a non-modal form with ShowForm in a DLL seems to cause the form
    ' not to show up if running Trade Navigator through the IDE without the NavCattle DLL
    ' loaded ( if you run compiled or with a group with NavCattle in it, it works fine ).  Using
    ' SetWindowPos instead seems to work...
    ShowForm Me, eForm_Nonmodal, g.frmMain
    'SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ShowMe"
    
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

    ' Need more margin at the top for the header to show up...
    frmPrintPreview.ShowMe "Turnkey", Me, 0, 0.45, 0.25, 0.25, 0.25, True, False
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.PrintMe"
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_Order
'' Description: Handle an order received in from the broker
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_Order(ByVal turnkeyOrder As cBrokerMessage)
On Error GoTo ErrSection:

    If Not m.bFormUnloading Then
        HandleOrderUpdate turnkeyOrder, False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Broker_Order"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_Fill
'' Description: Handle a fill received in from the broker
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_Fill(ByVal turnkeyFill As cBrokerMessage)
On Error GoTo ErrSection:

    If Not m.bFormUnloading Then
        HandleFillUpdate turnkeyFill
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Broker_Fill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_FeedYard
'' Description: Feed yard record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_FeedYard(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                cboFeedYards.Clear
                cboCustomers.Clear
                tmrRealtime.Enabled = False
                fgLots.Rows = fgLots.FixedRows
            
            Case "END"
                InfBox ""
                If cboFeedYards.ListCount = 1 Then
                    cboFeedYards.ListIndex = 0
                ElseIf cboFeedYards.ListCount > 1 Then
                    ShowDropDown cboFeedYards
                End If
                Enable cboFeedYards, (cboFeedYards.ListCount > 1)
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                cboFeedYards.AddItem turnkeyMessage("Name")
                cboFeedYards.ItemData(cboFeedYards.NewIndex) = CLng(Val(turnkeyMessage("ID")))
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_FeedYard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_FeedyardCustomer
'' Description: Customer record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_FeedyardCustomer(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strEndFlag As String            ' Flag on the end call

    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetVisibleCustomers
                End If
            
            Case Else
            
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_FeedyardCustomer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Customer
'' Description: Customer record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Customer(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strCustomer As String           ' Customer string to add to the array
    Dim lPos As Long                    ' Position of the customer in the array
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                cboCustomers.Clear
                
                tmrRealtime.Enabled = False
                fgLots.Rows = fgLots.FixedRows
                m.astrCustomers.Clear
            
            Case "END"
                If cboCustomers.ListCount > 1 Then
                    cboCustomers.AddItem "All Customers", 0
                    cboCustomers.ItemData(0) = -1&
                End If
                If cboCustomers.ListCount >= 1 Then
                    cboCustomers.ListIndex = 0
                End If
                Enable cboCustomers, cboCustomers.ListCount > 1
                
                InfBox ""
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetLots
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                If Len(turnkeyMessage("Name")) > 0 Then
                    cboCustomers.AddItem turnkeyMessage("Number") & " (" & turnkeyMessage("Name") & ")"
                Else
                    cboCustomers.AddItem turnkeyMessage("Number")
                End If
                cboCustomers.ItemData(cboCustomers.NewIndex) = CLng(Val(turnkeyMessage("ID")))
                
                If (turnkeyMessage("InRefresh") = "0") And (cboCustomers.ListCount = 1) Then
                    cboCustomers.AddItem "All Customers", 0
                    cboCustomers.ItemData(0) = -1&
                    
                    cboCustomers.ListIndex = 0&
                    Enable cboCustomers, True
                End If
                
                strCustomer = turnkeyMessage("Number") & vbTab & vbTab & turnkeyMessage("Name")
                If m.astrCustomers.BinarySearch(strCustomer, lPos) = False Then
                    m.astrCustomers.Add strCustomer, lPos
                End If
                
                'g.AppBridge.Cattle_Customer turnkeyMessage
                If FormIsLoaded("frmOwnerLookup") Then
                    frmOwnerLookup.Cattle_Customer turnkeyMessage
                End If
                If FormIsLoaded("frmEditLot") Then
                    frmEditLot.Turnkey_Customer turnkeyMessage
                End If
                If FormIsLoaded("frmEditLotContentDetails") Then
                    frmEditLotContentDetails.Turnkey_Customer turnkeyMessage
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Customer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Lot
'' Description: Lot record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Lot(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strRequestID As String          ' Request ID
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                tmrRealtime.Enabled = False
                fgLots.Rows = fgLots.FixedRows
            
            Case "END"
                AddTotalsRow
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetVisibleLots
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                LotToGrid turnkeyMessage
                
                If turnkeyMessage("InRefresh") = "0" Then
                    SortOnCol
                    FilterGrid
                End If
                
                'If DirExist(AddSlash(App.Path) & "Lots") = False Then
                '    MkDir AddSlash(App.Path) & "Lots"
                'End If
                'FileFromString AddSlash(App.Path) & "Lots\" & turnkeyMessage("Number") & ".LOT", strMessage
                
                strRequestID = turnkeyMessage("RequestID")
                If Len(strRequestID) > 0 Then
                    If m.WaitDetails.Exists(strRequestID) Then
                        g.Cattle.UpdateLotContentDetails m.WaitDetails(strRequestID), Str(SelectedFeedYard), turnkeyMessage("FeedYardLotID")
                        m.WaitDetails.Remove strRequestID
                    End If
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Lot", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Account
'' Description: Account record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Account(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strKey As String                ' Key into collection
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.Accounts.Clear
                
            Case "END"
                CheckWaitAccountItems
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetTurnkeyOrders
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                If turnkeyMessage("Deleted") = "1" Then
                    GetAccounts
                Else
                    turnkeyMessage.Add "HasFills", "0"
                    m.Accounts.Add turnkeyMessage, turnkeyMessage("ID")
                    
                    If g.AppBridge.ConnectionStatusForAccount(turnkeyMessage("Number"), True) = eGDConnectionStatus_Disconnected Then
                        m.NotConnected.Add turnkeyMessage
                    End If
                    
                    strKey = AccountKey(turnkeyMessage("Broker"), turnkeyMessage("Number"))
                    If m.WaitAccounts.Exists(strKey) Then
                        m.WaitAccounts.Remove strKey
                        If m.WaitAccounts.Count = 0 Then
                            If m.NotConnected.Count > 0 Then
                            End If
                            
                            GetAllBrokerFills
                        End If
                    End If
                    
                    If turnkeyMessage("InRefresh") = "0" Then
                        CheckWaitAccountItems
                    End If
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Account", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Order
'' Description: Order record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Order(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim PrevOrder As cBrokerMessage     ' Previous order
    Dim strBrokerOrderID As String      ' Broker order ID
    Dim Account As cBrokerMessage       ' Account
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.AssociatedOrders.Clear
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetAllTurnkeyFills
                    'GetAssociatedTurnkeyFills
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                If turnkeyMessage("NumberOfLegs") = "1" Then
                    turnkeyMessage.Add "Symbol", Parse(turnkeyMessage("Leg1"), ",", 3)
                End If
                
                turnkeyMessage.Add "Broker", g.Cattle.Accounts.BrokerForAccountID(turnkeyMessage("BrokerAccountID"))
                
                OrderToGrid turnkeyMessage
                
                strBrokerOrderID = OrderKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerOrderID"))
                If m.OrdersSent.Exists(strBrokerOrderID) Then
                    g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' removed from list of orders sent"
                    m.OrdersSent.Remove strBrokerOrderID
                End If
                If m.AssociatedOrdersSent.Exists(strBrokerOrderID) Then
                    g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' removed from list of associated orders sent"
                    m.AssociatedOrdersSent.Remove strBrokerOrderID
                    
                    ItemReceived
                End If
                
                If m.AssociatedOrders.Exists(strBrokerOrderID) Then
                    g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' overwritten in list of associated orders"
                    Set m.AssociatedOrders(strBrokerOrderID) = turnkeyMessage
                Else
                    g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' added to list of associated orders"
                    m.AssociatedOrders.Add turnkeyMessage, strBrokerOrderID
                End If
                
                CheckWaitOrderItems turnkeyMessage
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Order", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Fill
'' Description: Fill record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Fill(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim Account As cBrokerMessage       ' Account
    Dim strKey As String                ' Key into the fills collection
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.Fills.Clear
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetAssociatedTurnkeyFills
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                turnkeyMessage.Add "Broker", g.Cattle.Accounts.BrokerForAccountID(turnkeyMessage("BrokerAccountID"))
                
                If m.Accounts.Exists(turnkeyMessage("BrokerAccountID")) Then
                    Set Account = m.Accounts(turnkeyMessage("BrokerAccountID"))
                    Account.Add "HasFills", "1"
                    Set m.Accounts(turnkeyMessage("BrokerAccountID")) = Account
                End If
                
                strKey = FillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"))
                If m.Fills.Exists(strKey) Then
                    g.Cattle.DumpDebug "Fill '" & strKey & "' overwritten in list of fills"
                    Set m.Fills(strKey) = turnkeyMessage
                Else
                    g.Cattle.DumpDebug "Fill '" & strKey & "' added to list of fills"
                    m.Fills.Add turnkeyMessage, strKey
                End If
                
                CheckWaitFillItems turnkeyMessage
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Fill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_AssociatedFill
'' Description: Associated Fill record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_AssociatedFill(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim Account As cBrokerMessage       ' Account
    Dim strKey As String                ' Key into the associated fill collection
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.AssociatedFills.Clear
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetTurnkeyTrades
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                turnkeyMessage.Add "Broker", g.Cattle.Accounts.BrokerForAccountID(turnkeyMessage("BrokerAccountID"))
                
                strKey = AssociatedFillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"), turnkeyMessage("FeedYardLotID"))
                If m.AssociatedFillsSent.Exists(strKey) Then
                    g.Cattle.DumpDebug "Fill '" & strKey & "' removed from list of associated fills sent"
                    m.AssociatedFillsSent.Remove strKey
                    
                    ItemReceived
                End If
                
                If m.AssociatedFills.Exists(strKey) Then
                    g.Cattle.DumpDebug "Fill '" & strKey & "' overwritten in list of associated fills"
                    Set m.AssociatedFills(strKey) = turnkeyMessage
                Else
                    g.Cattle.DumpDebug "Fill '" & strKey & "' added to list of associated fills"
                    m.AssociatedFills.Add turnkeyMessage, strKey
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_AssociatedFill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Trade
'' Description: Trade record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Trade(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim Account As cBrokerMessage       ' Account
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.Trades.Clear
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetTurnkeyPositions
                End If
                
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                turnkeyMessage.Add "AccountNumber", g.Cattle.Accounts.AccountNumberForAccountID(turnkeyMessage("BrokerAccountID"))
                
                TradeToGrid turnkeyMessage
                
                If m.Trades.Exists(turnkeyMessage("ID")) Then
                    Set m.Trades(turnkeyMessage("ID")) = turnkeyMessage
                Else
                    m.Trades.Add turnkeyMessage, turnkeyMessage("ID")
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Trade", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_DeleteTrade
'' Description: Event telling us to delete the given trades
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_DeleteTrade(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                End If
                
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                DeleteTrade turnkeyMessage
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_DeleteTrade", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_DeleteOrder
'' Description: Event telling us to delete the given order
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_DeleteOrder(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim Account As cBrokerMessage       ' Account
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                End If
                
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                turnkeyMessage.Add "Broker", g.Cattle.Accounts.BrokerForAccountID(turnkeyMessage("BrokerAccountID"))
                
                DeleteOrder turnkeyMessage
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_DeleteOrder", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_DeleteAssociatedFill
'' Description: Event telling us to delete the given associated fill
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_DeleteAssociatedFill(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim Account As cBrokerMessage       ' Account
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                End If
                
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                            
                turnkeyMessage.Add "Broker", g.Cattle.Accounts.BrokerForAccountID(turnkeyMessage("BrokerAccountID"))
                
                DeleteAssociatedFill turnkeyMessage
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_DeleteAssociatedFill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Position
'' Description: Position record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Position(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"

            Case "END"
                InfBox "Calculating Totals.  Please wait...", , , g.Cattle.ProductName, True
                CalcTotals
                InfBox ""
                MoveFocus fgLots
                
                SortOnCol
                
                tmrRealtime.Interval = g.lStreamInterval
                tmrRealtime.Enabled = g.bStreamActive
                
                strEndFlag = Parse(strMessage, vbTab, 2)
                If m.Accounts.Count = 0 Then
                    AssociateAccounts
                ElseIf (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetAllBrokerOrders
                    GetAllBrokerFills
                End If
                
                ExpandAll m.lExpandLevel
                'FilterGrid
                
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetLotDetails
                End If
        
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                PositionToGrid turnkeyMessage
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_Position", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_DeletePosition
'' Description: Event telling us to delete the given position
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_DeletePosition(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                End If
                
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                DeletePosition turnkeyMessage
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_DeletePosition", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_LotColumn
'' Description: Handle a lot column coming from the Turnkey server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_LotColumn(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column information
    Dim strEndFlag As String            ' Flag on the end call
    Dim lIndex As Long                  ' Index into a for loop
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.LotColumns.Clear
                m.KeyValueToIndex.Clear
                
                For lIndex = 0 To m.alLotCol.Size - 1
                    m.alLotCol(lIndex) = kNullData
                Next lIndex
            
            Case "END"
                SetupGridColumns
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetVisibleLotColumns
                End If
            
            Case Else
                Set LotColumn = New cLotColumn
                LotColumn.FromString strMessage
                m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField
                
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.Turnkey_LotColumn", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_VisibleLotColumn
'' Description: Handle a visible lot column coming from the Turnkey server
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_VisibleLotColumn(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLotColumnIds As cGdArray    ' Array of visible lot column ID's
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                For lIndex = 1 To m.LotColumns.Count
                    If m.LotColumns(lIndex).ID < 10000 Then
                        m.LotColumns(lIndex).FeedyardHidden = True
                    End If
                Next lIndex

            Case "END"
                ShowGridColumns
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    'GetFeedyards
                    GetFeedyardCustomers
                    'GetVisibleCustomers
                End If
            
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                Set astrLotColumnIds = New cGdArray
                astrLotColumnIds.SplitFields turnkeyMessage("LotColumnIds"), ","
                
                For lIndex = 0 To astrLotColumnIds.Size - 1
                    If m.LotColumns.Exists(astrLotColumnIds(lIndex)) = True Then
                        m.LotColumns(astrLotColumnIds(lIndex)).FeedyardHidden = False
                    End If
                Next lIndex
                
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.Turnkey_VisibleLotColumn", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_VisibleLots
'' Description: Visible lot record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_VisibleLots(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strEndFlag As String            ' Flag on the end call
    
    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                m.VisibleLots.Clear
            
            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetAccounts
                End If
                
            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                If m.VisibleLots.Exists(turnkeyMessage("FeedYardCustomerID")) Then
                    m.VisibleLots(turnkeyMessage("FeedYardCustomerID")) = turnkeyMessage("LotIds")
                Else
                    m.VisibleLots.Add turnkeyMessage("LotIds"), turnkeyMessage("FeedYardCustomerID")
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey_VisibleLots", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_LotContentsDetail
'' Description: Lot contents detail record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_LotContentsDetail(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strEndFlag As String            ' Flag on the end call

    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                RemoveDetails turnkeyMessage("FeedYardID"), turnkeyMessage("FeedYardLotID"), turnkeyMessage("LotColumnID")

            Case "END"
                strEndFlag = Parse(strMessage, vbTab, 2)
                If (Len(strEndFlag) = 0) Or (UCase(strEndFlag) = "GO") Then
                    GetLotDetailOptions
                End If
                LotDetailsToGrid
                
                m.bHasDetails = True
                SynchronizeIngredients

            Case Else
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strMessage
                
                If m.Details.Exists(turnkeyMessage("ID")) Then
                    Set m.Details(turnkeyMessage("ID")) = turnkeyMessage
                Else
                    m.Details.Add turnkeyMessage, turnkeyMessage("ID")
                End If

        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey.LotContentsDetail", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Ingredient
'' Description: Ingredient record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Ingredient(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message object
    Dim strEndFlag As String            ' Flag on the end call

    If Not m.bFormUnloading Then
        Select Case UCase(Parse(strMessage, vbTab, 1))
            Case "BEGIN"
            Case "END"
                m.bHasIngredients = True
                SynchronizeIngredients
                
            Case Else
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Turnkey.Turnkey_Ingredient", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpNewOrder
'' Description: Set up a new order to be associated
'' Inputs:      Genesis Order ID, Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetUpNewOrder(ByVal strGenesisOrderID As String, ByVal strFeedYardLotID As String)
On Error GoTo ErrSection:

    Dim strLotId As String              ' Lot ID for the given Lot number

    If Not m.bFormUnloading Then
        If Len(strFeedYardLotID) > 0 Then
            If m.NewOrders.Exists(strGenesisOrderID) Then
                m.NewOrders(strGenesisOrderID) = strFeedYardLotID
            Else
                g.Cattle.DumpDebug "Order " & strGenesisOrderID & " added to list of new orders"
                m.NewOrders.Add strFeedYardLotID, strGenesisOrderID
            End If
        ElseIf m.NewOrders.Exists(strGenesisOrderID) Then
            g.Cattle.DumpDebug "Order " & strGenesisOrderID & " removed from list of new orders because no lot number specified"
            m.NewOrders.Remove strGenesisOrderID
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SetUpNewOrder"
    
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

    Dim strCustomer As String           ' Customer information

    With frmPrintPreview.vp
        .StartDoc
        g.AppBridge.DoPrintHeader
        
        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
        .FontUnderline = False
        .TextAlign = taCenterMiddle
        
        strCustomer = cboCustomers.Text
        If UCase(Left(strCustomer, 3)) = "ALL" Then
            .Text = g.Cattle.ProductName & " Report for All Customers"
        Else
            .Text = g.Cattle.ProductName & " Report for Customer #" & strCustomer
        End If
        
        .FontBold = False
        .FontSize = 12
        .TextAlign = taLeftMiddle
        
        .Text = vbLf & vbLf
        
        .RenderControl = fgLots.hWnd
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedLot
'' Description: Determine the currently selected lot
'' Inputs:      None
'' Returns:     Selected Lot ( Nothing if no lot selected )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectedLot() As cBrokerMessage
On Error GoTo ErrSection:

    Dim Lot As cBrokerMessage           ' Lot object
    Dim lLotNumberCol As Long           ' Lot number column
    Dim lParentRow As Long              ' Parent row in the grid
    
    With fgLots
        Set Lot = Nothing
        
        lLotNumberCol = LotCol(eGDLotCol_LotNumber)
        If ValidGridRow(fgLots) Then
            If .RowOutlineLevel(.Row) = 1 Then
                If RowType(.Row) <> eGDRowType_Total Then
                    Set Lot = .RowData(.Row)
                End If
            Else
                lParentRow = .Row
                Do
                    lParentRow = .GetNodeRow(lParentRow, flexNTParent)
                Loop While (RowType(lParentRow) <> eGDRowType_Lot)
                
                Set Lot = .RowData(lParentRow)
            End If
        End If
    End With
    
    Set SelectedLot = Lot

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.SelectedLot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CattleFillExists
'' Description: Determine if a fill exists with the given ID and broker
'' Inputs:      Broker Fill ID, Broker
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CattleFillExists(ByVal strBrokerFillID As String, ByVal strBroker As String) As Boolean
On Error GoTo ErrSection:

    CattleFillExists = m.Fills.Exists(FillKey(strBroker, strBrokerFillID))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.CattleFillExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCustomers_Click
'' Description: Handle the selection of a customer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCustomers_Click()
On Error GoTo ErrSection:

    FilterGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.cboCustomers_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboFeedYards_Click
'' Description: Handle the selection of a feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboFeedyards_Click()
On Error GoTo ErrSection:

    Dim FeedYard As cBrokerMessage      ' Feed yard object

    If g.Cattle.FeedYards.Exists(Str(SelectedFeedYard)) Then
        Set FeedYard = g.Cattle.FeedYards(Str(SelectedFeedYard))
        
        m.bCanEditLots = g.Cattle.CanEditLots And (FeedYard("CanEditLots") <> "0")
        tbToolbar.Tools("ID_Feed").Visible = m.bCanEditLots
    End If

    'GetVisibleCustomers
    GetAllLotColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.cboFeedYards_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboView_Click
'' Description: Change the view according to the user's selection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboView_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.cboView_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddCustomer_Click
'' Description: Allow the user to add a feedyard customer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddCustomer_Click()
On Error GoTo ErrSection:

    frmCattleManage.ShowMeFeedYardCustomers

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.AddCustomer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddFeedYard_Click
'' Description: Allow the user to add a feedyard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddFeedYard_Click()
On Error GoTo ErrSection:

    frmCattleManage.ShowMeFeedYards

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.AddFeedYard"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to lookup an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    Dim strCustomer As String           ' Customer that the user selected
    
    'strCustomer = g.AppBridge.AccountLookup(m.astrCustomers, , , True)
    strCustomer = frmOwnerLookup.ShowMe(m.astrCustomers, , , True)
    If Len(strCustomer) > 0 Then
        SelectCustomer strCustomer
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.cmdLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_AfterMoveColumn
'' Description: After a user moves a column, resave the display string
'' Inputs:      Column moved, Position moved to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.bLoadingColumns = False Then
        RebuildKeyValueToCol
        Dirty = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_AfterMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim strSymbol As String             ' Symbol for the mouse row
    Dim bCanEditLots As Boolean         ' Can the user edit lots?
    Dim lLotNumberCol As Long           ' Lot number column
    Dim bValidLotRow As Boolean         ' Is the current mouse row a valid lot row?
    Dim bOrder As Boolean               ' Is this row an order?

    m.iButtonPressed = Button
    If Button = vbRightButton Then
        lMouseRow = fgLots.MouseRow
        lMouseCol = fgLots.MouseCol
        
        fgLots.Row = lMouseRow
        strSymbol = SymbolForRow(lMouseRow)
        
        Enable mnuPriceLadder, (Len(strSymbol) > 0)
        mnuLots.Tag = strSymbol
        
        With fgLots
            lLotNumberCol = LotCol(eGDLotCol_LotNumber)
            If ValidGridRow(fgLots, lMouseRow) Then
                bValidLotRow = (.RowOutlineLevel(lMouseRow) <> 1) Or (UCase(.TextMatrix(lMouseRow, lLotNumberCol)) <> "TOTALS")
                mnuEditLot.Tag = Str(lMouseCol)
            Else
                bValidLotRow = False
            End If
        End With
        
        bCanEditLots = m.bCanEditLots 'g.Cattle.CanEditLots
        mnuAddLot.Visible = bCanEditLots
        mnuEditLot.Visible = bCanEditLots
        mnuEditLot.Enabled = bValidLotRow
        mnuSep2.Visible = bCanEditLots
        
        bOrder = (RowType(lMouseRow) = eGDRowType_Order)
        mnuCloseOrder.Visible = bOrder
        mnuSep3.Visible = bOrder
        
        PopupMenu mnuLots
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_BeforeMoveColumn
'' Description: Make sure certain columns stay where they are
'' Inputs:      Column to move, Position to move it to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:

    If m.bLoadingColumns = False Then
        Position = VerifyMoveColumn(Col, Position)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_BeforeMoveColumn"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_BeforeScroll
'' Description: Try to stop the FlexGrid vs. Streaming bug from happening
'' Inputs:      Old Top Row, Old Left Column, New Top Row, New Left Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    mCattle.GridScrollCheck fgLots, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_BeforeSort
'' Description: Handle the user sorting a column
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SortOnCol Col, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_BeforeSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_Click
'' Description: Handle the user clicking in a cell
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row of the cell in the grid where the user clicked
    Dim lMouseCol As Long               ' Column of the cell in the grid where the user clicked
    Dim lLotNumberCol As Long           ' Lot number column
    Dim Lot As cBrokerMessage           ' Lot object
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim Info As cBrokerMessage          ' Information object

    ' When the client clicks even in the non-client area of the grid, make sure that
    ' the focus moves to the grid (#6745)
    MoveFocus fgLots

    If m.iButtonPressed = vbLeftButton Then
        With fgLots
            lMouseRow = .MouseRow
            lMouseCol = .MouseCol
            lLotNumberCol = LotCol(eGDLotCol_LotNumber)
            Set LotColumn = LotColumnForCol(lMouseCol)
            
            If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
                If .RowOutlineLevel(lMouseRow) = 1 Then
                    If UCase(.TextMatrix(lMouseRow, lLotNumberCol)) <> "TOTALS" Then
                        If lMouseCol = lLotNumberCol Then
                            Set Lot = .RowData(lMouseRow)
                            g.AppBridge.CreateOrder Lot("FeedYardLotID")
                        ElseIf UCase(LotColumn.KeyValueField) = "ISLOTCLOSED" Then
                            EditLot LotColumn.KeyValueField
                        End If
                    End If
                ElseIf .RowOutlineLevel(lMouseRow) = 2 Then
                    If UCase(Left(.TextMatrix(lMouseRow, lMouseCol), 10)) = "CLICK HERE" Then
                        AssociateLotItems .GetNodeRow(lMouseRow, flexNTParent)
                    ElseIf RowType(lMouseRow) = eGDRowType_Symbol Then
                        Set Lot = .RowData(.GetNodeRow(lMouseRow, flexNTParent))
                        Set Info = .RowData(lMouseRow)
                        
                        g.AppBridge.CreateOrder Lot("FeedYardLotID"), Info("Symbol")
                    End If
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_Compare
'' Description: Perform a comparison for the two rows for sorting purposes
'' Inputs:      Row 1, Row 2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Dim strRow1 As String
    Dim strRow2 As String
    
    strRow1 = fgLots.TextMatrix(Row1, LotCol(eGDLotCol_SortKey))
    strRow2 = fgLots.TextMatrix(Row2, LotCol(eGDLotCol_SortKey))
    
    If strRow1 = strRow2 Then
        Cmp = 0
    ElseIf strRow1 < strRow2 Then
        Cmp = -1
    Else
        Cmp = 1
    End If
    
    If m.nSortedDir = flexSortStringDescending Then
        Cmp = Cmp * -1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_DblClick
'' Description: Handle the user double clicking in a cell
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row the mouse is over
    Dim lMouseCol As Long               ' Column the mouse is over
    Dim LotColumn As cLotColumn         ' Lot column information
    
    With fgLots
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If m.bCanEditLots Then
            If ValidGridRow(fgLots, lMouseRow) Then
                If .RowOutlineLevel(lMouseRow) = 1 Then
                    If UCase(.TextMatrix(lMouseRow, LotCol(eGDLotCol_LotNumber))) <> "TOTALS" Then
                        Set LotColumn = LotColumnForCol(lMouseCol)
                        
                        .Row = lMouseRow
                        
                        EditLot LotColumn.KeyValueField
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.fgLots_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgLots_MouseMove
'' Description: If the user hovers over a header, display a tooltip
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgLots_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseRow As Long               ' Row the mouse is over
    Dim lMouseCol As Long               ' Column the mouse is over
    Dim strTooltipText As String        ' Tool tip text to display
    Dim LotColumn As cLotColumn         ' Lot column information
    Dim lRootRow As Long                ' Root row for the mouse row
    
    strTooltipText = ""
    With fgLots
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow = 0 Then
            Set LotColumn = LotColumnForCol(lMouseCol)
            If Not LotColumn Is Nothing Then
                strTooltipText = LotColumn.TooltipText
            End If
        ElseIf (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            lRootRow = .GetNodeRow(lMouseRow, flexNTRoot)
            If (lRootRow >= .FixedRows) And (lRootRow < .Rows) Then
                strTooltipText = .TextMatrix(lRootRow, LotCol(eGDLotCol_LotNumber))
                If UCase(strTooltipText) <> "TOTALS" Then
                    strTooltipText = "Lot# " & strTooltipText
                End If
            End If
        End If
        
        If strTooltipText <> .TooltipText Then
            .TooltipText = strTooltipText
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Perform actions when the form is activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim lNumTemplates As Long           ' Number of saved templates
    Dim strOldDisplay As String         ' Old saved display
    Dim lTemplateNumber As Long         ' Template number
    
    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me

    Set m.LotColumns = New cGdTree
    Set m.KeyValueToIndex = New cGdTree
    Set m.KeyValueToCol = New cGdTree
    Set m.alLotCol = New cGdArray
    m.alLotCol.Create eGDARRAY_Longs, eGDLotCol_NumCols, kNullData
    
    Set m.astrCustomers = New cGdArray
    Set m.Accounts = New cGdTree
    Set m.AssociatedOrders = New cGdTree
    Set m.AssociatedFills = New cGdTree
    Set m.Fills = New cGdTree
    Set m.Trades = New cGdTree
    Set m.NotConnected = New cGdTree
    Set m.WaitAccounts = New cGdTree
    Set m.WaitAccountOrders = New cGdTree
    Set m.WaitAccountFills = New cGdTree
    Set m.WaitOrderOrders = New cGdTree
    Set m.WaitOrderFills = New cGdTree
    Set m.WaitFillFills = New cGdTree
    Set m.NewOrders = New cGdTree
    Set m.OrdersSent = New cGdTree
    Set m.AssociatedOrdersSent = New cGdTree
    Set m.AssociatedFillsSent = New cGdTree
    Set m.VisibleLots = New cGdTree
    Set m.Details = New cGdTree
    Set m.WaitDetails = New cGdTree
    
    mnuLots.Visible = False
    
    With tbToolbar
        .Tools("ID_Refresh").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kRefresh"))
        .Tools("ID_EditColumns").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_Settings"))
        .Tools("ID_ExpandAll").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kExpandAll"))
        .Tools("ID_CollapseAll").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kCollapseAll"))
        .Tools("ID_Print").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_Print"))
        .Tools("ID_TextIncrease").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_TextIncrease"))
        .Tools("ID_TextDecrease").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_TextDecrease"))
        .Tools("ID_Reports").Picture = g.AppBridge.Picture16("kBarChart")
        .Tools("ID_Ingredients").Picture = g.AppBridge.Picture16("kBlank")
        
        .Tools("ID_Feed").Visible = False
        
        .Tools("ID_ExpandAll").TooltipText = "Expand the tree one level"
        .Tools("ID_CollapseAll").TooltipText = "Collapse the tree one level"
    End With
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    tmrItemsSent.Interval = 1000
    tmrItemsSent.Enabled = False
    
    lNumTemplates = NumTemplates
    lTemplateNumber = GetIniFileProperty("TemplateNumber", 0&, "Turnkey", g.strIniFile)
    
    If lNumTemplates = 0& Then
        strOldDisplay = GetIniFileProperty("Display", "", "Turnkey", g.strIniFile)
        If Len(strOldDisplay) > 0 Then
            lNumTemplates = 1&
            lTemplateNumber = 1&
            
            NumTemplates = 1&
            TemplateName(1) = "SavedDisplay"
            TemplateDisplay(1) = strOldDisplay
        End If
    End If
    
    GridFontSize = GetIniFileProperty("FontSize", 8, "Turnkey", g.strIniFile)
    
    m.bDoneOnce = False
    Status = eGDConnectionStatus_Disconnected
    m.lSortedCol = -1&
    m.nSortedDir = -1&
    m.bLoadingColumns = False
    m.bFormUnloading = False
    TemplateNumber = lTemplateNumber
    
    cboView.AddItem "Show All Lots"
    cboView.ItemData(cboView.NewIndex) = LotView(eGDLotView_AllLots)
    cboView.AddItem "Lots With Hedge"
    cboView.ItemData(cboView.NewIndex) = LotView(eGDLotView_LotsWithHedge)
    cboView.AddItem "Lots Without Hedge"
    cboView.ItemData(cboView.NewIndex) = LotView(eGDLotView_LotsWithoutHedge)
    cboView.ListIndex = GetIniFileProperty("LotView", LotView(eGDLotView_AllLots), "Turnkey", g.strIniFile)
    If cboView.ListIndex = -1& Then
        cboView.ListIndex = LotView(eGDLotView_AllLots)
    End If
    
    cmdAddFeedYard.TooltipText = "Manage feed yards" ' "Add a new feed yard"
    '''RH cmdAddFeedYard.Picture = g.AppBridge.Picture16("kSystem")
    cmdAddFeedYard.Visible = False
    cmdAddCustomer.TooltipText = "Manage feed yard customers" ' "Add a new feed yard customer"
    '''RH cmdAddCustomer.Picture = g.AppBridge.Picture16("kSystem")
    cmdAddCustomer.Visible = False
    cmdLookup.TooltipText = "Lookup feed yard customer"
    '''RH cmdLookup.Picture = g.AppBridge.Picture16("kMagnify2")
    
    m.lExpandLevel = GetIniFileProperty("ExpandLevel", 4&, "Turnkey", g.strIniFile)
    
    Dirty = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Determine whether or not to let the form close
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = Not AskToSave
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Size and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lSpace As Long                  ' Space between controls
    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    Dim lTop As Long                    ' Top of a control
    
    lSpace = 60
    lMinScaleWidth = fraAccounts.Width + (lSpace * 2)
    lMinScaleHeight = fraAccounts.Height * 10 + (lSpace * 4)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With fraAccounts
            .Move lSpace, lSpace
        End With
        
        With fgLots
            lTop = fraAccounts.Height + (lSpace * 2)
            .Move lSpace, lTop, ScaleWidth - (lSpace * 2), ScaleHeight - lTop - lSpace
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    m.bFormUnloading = True

    If (Not g.Cattle Is Nothing) And (Status <> eGDConnectionStatus_Disconnected) Then
        g.Cattle.Disconnect
    End If
    
    SaveFormPlacement Me

    SetIniFileProperty "FontSize", GridFontSize, "Turnkey", g.strIniFile
    SetIniFileProperty "LotView", cboView.ListIndex, "Turnkey", g.strIniFile
    SetIniFileProperty "ExpandLevel", m.lExpandLevel, "Turnkey", g.strIniFile
    
    Set m.LotColumns = Nothing
    Set m.KeyValueToCol = Nothing
    Set m.KeyValueToIndex = Nothing
    
    Set m.astrCustomers = Nothing
    Set m.Accounts = Nothing
    Set m.AssociatedOrders = Nothing
    Set m.AssociatedFills = Nothing
    Set m.Fills = Nothing
    Set m.NotConnected = Nothing
    Set m.WaitAccounts = Nothing
    Set m.WaitAccountOrders = Nothing
    Set m.WaitAccountFills = Nothing
    Set m.WaitOrderOrders = Nothing
    Set m.WaitOrderFills = Nothing
    Set m.WaitFillFills = Nothing
    Set m.NewOrders = Nothing
    Set m.OrdersSent = Nothing
    Set m.AssociatedOrdersSent = Nothing
    Set m.AssociatedFillsSent = Nothing
    Set m.VisibleLots = Nothing
    
    tmrRealtime.Enabled = False
    tmrMenu.Enabled = False
    tmrItemsSent.Enabled = False
    
    m.bDoneOnce = False
    
    ' 02/01/2013 DAJ: Make sure to clear out any open InfBoxes here so that when people
    ' like Pete get impatient and close the dialog in the middle of a refresh, the notifications
    ' go away...
    InfBox ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    imgStatus_Click
'' Description: Allow the user to toggle the connection status
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub imgStatus_Click()
On Error GoTo ErrSection:

    ToggleConnection

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.imgStatus_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lblStatus_Click
'' Description: Allow the user to toggle the connection status
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblStatus_Click()
On Error GoTo ErrSection:

    ToggleConnection

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.lblStatus_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddLot_Click
'' Description: Allow the user to add a lot
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddLot_Click()
On Error GoTo ErrSection:

    Dim Lot As cBrokerMessage           ' Lot object
    Dim LotDetails As cGdTree           ' Lot details object
    Dim strRequestID As String          ' Request ID
    
    Set Lot = New cBrokerMessage
    Set LotDetails = New cGdTree
    
    If frmEditLot.ShowMe(SelectedFeedYard, Lot, LotDetails) Then
        Lot.Add "FeedYardID", Str(SelectedFeedYard)
        Lot.Add "FeedYardName", cboFeedYards.Text
        
        strRequestID = g.Cattle.NextRequestID
        Lot.Add "RequestID", strRequestID
        m.WaitDetails.Add LotDetails, strRequestID
    
        g.Cattle.AddLots Lot.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuAddLot_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCollapseAll_Click
'' Description: Allow the user to collapse all nodes of the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCollapseAll_Click()
On Error GoTo ErrSection:

    CollapseAll

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuCollapseAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditColumns_Click
'' Description: Allow the user to change fields in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditColumns_Click()
On Error GoTo ErrSection:

    StartMenuTimer "EditColumns", kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuEditColumns_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditLot_Click
'' Description: Allow the user to edit a lot
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditLot_Click()
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column object
    Dim strKeyValueField As String      ' Key value field

    strKeyValueField = ""
    If Len(mnuEditLot.Tag) > 0 Then
        Set LotColumn = LotColumnForCol(CLng(Val(mnuEditLot.Tag)))
        strKeyValueField = LotColumn.KeyValueField
    End If

    EditLot strKeyValueField

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuEditLot_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuExpandAll_Click
'' Description: Allow the user to expand all nodes of the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuExpandAll_Click()
On Error GoTo ErrSection:

    ExpandAll

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuExpandAll_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuExportToCsv_Click
'' Description: Allow the user to export the grid to a CSV file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuExportToCsv_Click()
On Error GoTo ErrSection:

    StartMenuTimer "ExportToCsv", kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuExportToCsv_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCloseOrder_Click
'' Description: Allow the user to mark the selected order as cancelled
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCloseOrder_Click()
On Error GoTo ErrSection:

    Dim Order As cBrokerMessage         ' Order object from the grid

    With fgLots
        If RowType(.Row) = eGDRowType_Order Then
            If TypeOf .RowData(.Row) Is cBrokerMessage Then
                If InfBox("This will close out your order on the Genesis Cattle servers so that it no longer shows up.||THIS WILL NOT AFFECT YOUR ORDER AT THE BROKER!||You should only do this if the order is no longer working at the broker.||Do you want to continue?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                    Set Order = .RowData(.Row)
                    
                    Order("Status") = Str(eTT_OrderStatus_Cancelled)
                    Order("IsWorking") = "0"
                    
                    UpdateOrder Order, "user marked as cancelled"
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuCloseOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPriceLadder_Click
'' Description: Allow the user to bring up a price ladder for the given row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPriceLadder_Click()
On Error GoTo ErrSection:

    g.AppBridge.ShowLadder mnuLots.Tag

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuPriceLadder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrint_Click
'' Description: Allow the user to print the report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    StartMenuTimer "Print", kNullData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuShowClosedLots_Click
'' Description: Determine whether the user wants to show closed lots
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuShowClosedLots_Click()
On Error GoTo ErrSection:

    mnuShowClosedLots.Checked = Not mnuShowClosedLots.Checked
    FilterGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.mnuShowClosedLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the user clicking on a tool in the toolbar
'' Inputs:      Tool
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim lTemplateNumber As Long         ' Template Number

    Select Case UCase(Tool.ID)
        Case "ID_REFRESH"
            RefreshTurnkey
        
        Case "ID_EDITCOLUMNS"
            EditColumns
            
        Case "ID_EXPANDALL"
            ExpandAll
            
        Case "ID_COLLAPSEALL"
            CollapseAll
            
        Case "ID_PRINT"
            PrintMe
            
        Case "ID_ASSOCIATE"
            
        Case "ID_ASSOCIATEACCOUNTS"
            AssociateAccounts
            
        Case "ID_ASSOCIATEORDERSFILLS"
            AssociateLotItems -1&
            
        Case "ID_IMPORTFILLS"
            ImportHistoricalFills
            
        Case "ID_NEWFILL"
            NewFill
            
        Case "ID_TEMPLATES"
            tbToolbar.Tools("ID_RenameTemplate").Enabled = (m.lTemplateNumber > 0&)
            LoadTemplatesMenu
            
        Case "ID_SAVETEMPLATE"
            SaveTemplate
        
        Case "ID_SAVEASTEMPLATE"
            SaveTemplateAs
        
        Case "ID_RENAMETEMPLATE"
            RenameTemplate
            
        Case "ID_REMOVETEMPLATES"
            RemoveTemplates
        
        Case "ID_DEFAULTTEMPLATE"
            TemplateNumber = 0&
            ShowGridColumns
            
        Case "ID_TEXTINCREASE"
            GridFontSize = GridFontSize + 1
        
        Case "ID_TEXTDECREASE"
            GridFontSize = GridFontSize - 1
            
        Case "ID_REPORTS"
            ShowReports
            
        Case "ID_INGREDIENTS"
            frmCattleManage.ShowMeIngredients
            
        Case "ID_RATIONS"
            frmCattleManage.ShowMeRations
            
        Case Else
            If Left(UCase(Tool.ID), 12) = "ID_TEMPLATE_" Then
                lTemplateNumber = CLng(Val(Mid(Tool.ID, 13)))
                If lTemplateNumber <> TemplateNumber Then
                    If AskToSave Then
                        TemplateNumber = lTemplateNumber
                        ShowGridColumns
                        Dirty = False
                    End If
                End If
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrItemsSent_Timer
'' Description: Check to see if too much time has elapsed since last item sent
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrItemsSent_Timer()
On Error GoTo ErrSection:

    g.AppBridge.TimerStart "frmLots.tmrItemsSent"
    If gdTickCount >= (m.dLastItemSent + 30000) Then
        m.dLastItemSent = kNullData
        tmrItemsSent.Enabled = False
        
        Refresh
    End If
    g.AppBridge.TimerEnd "frmLots.tmrItemsSent", tmrItemsSent.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.tmrItemsSent_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Perform the necessary menu command
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Tag of the timer control
    Static bInProgress As Boolean       ' Are we currently performing a command?

    g.AppBridge.TimerStart "frmLots.tmrMenu"
    If bInProgress = False Then
        bInProgress = True
        
        strTag = tmrMenu.Tag
        tmrMenu.Tag = ""
        
        Select Case UCase(Parse(strTag, vbTab, 1))
            Case "EDITCOLUMNS"
                EditColumns
                
            Case "EXPORTTOCSV"
                ExportToCsv
                
            Case "PRINT"
                PrintMe
                
        End Select
        
        bInProgress = False
    End If
    g.AppBridge.TimerEnd "frmLots.tmrMenu", tmrMenu.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRealtime_Timer
'' Description: Update the open equity and totals when streaming
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRealTime_Timer()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    g.AppBridge.TimerStart "frmLots.tmrRealTime"
    With fgLots
        For lIndex = .FixedRows To .Rows - 1
            If .RowOutlineLevel(lIndex) = 4 Then
                CalculateOpenEquityForRow lIndex
                CalcTotals True
            End If
        Next lIndex
    End With
    g.AppBridge.TimerEnd "frmLots.tmrRealTime", tmrRealtime.Interval
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.tmrRealtime_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitLotsGrid
'' Description: Initialize the lots grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitLotsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .WordWrap = True
        
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .RowHidden(0) = True
        
        m.nBackColor = .BackColor
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildFieldList
'' Description: Build the field list to send to frmQuoteBoardFields
'' Inputs:      None
'' Returns:     Field List
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildFieldList() As cGdArray
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrReturn As cGdArray          ' Array of field information
    Dim strActive As String             ' Is the column visible?
    Dim strShow As String               ' Do we want to show the column in the settings?
    Dim LotColumn As cLotColumn         ' Lot column object
    
    Set astrReturn = New cGdArray
    With fgLots
        For lIndex = 0 To .Cols - 1
            Set LotColumn = LotColumnForCol(lIndex)
            If Not LotColumn Is Nothing Then
                ' Active (vbChecked|vbUnchecked) \t Name \t Original Position \t Show (vbChecked|vbUnchecked)
                If LotColumn.UserHidden = True Then
                    strActive = Str(vbUnchecked)
                Else
                    strActive = Str(vbChecked)
                End If
                
                If (LotColumn.AllowUserMove = True) And (LotColumn.FeedyardHidden = False) Then
                    strShow = Str(vbChecked)
                ElseIf (LotColumn.KeyValueField = "OpenEquity") Or (LotColumn.KeyValueField = "ClosedProfit") Then
                    strShow = Str(vbChecked)
                Else
                    strShow = Str(vbUnchecked)
                End If
                
                astrReturn.Add strActive & vbTab & .TextMatrix(0, lIndex) & vbTab & Str(lIndex) & vbTab & strShow
            End If
        Next lIndex
    End With
    
    Set BuildFieldList = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.BuildFieldList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditColumns
'' Description: Allow the user to change fields in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditColumns()
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of field information
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Column number for the information
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim bUserHidden As Boolean          ' Did the user hide the column?
    
    Set astrFields = BuildFieldList
    
    If g.AppBridge.EditLotColumns(astrFields) Then
        With fgLots
            .Redraw = flexRDNone
            
            For lIndex = 0 To astrFields.Size - 1
                lCol = CLng(Val(Parse(astrFields(lIndex), vbTab, 3)))
                bUserHidden = (CLng(Val(Parse(astrFields(lIndex), vbTab, 1))) = flexUnchecked)
                
                Set LotColumn = LotColumnForCol(lCol)
                If Not LotColumn Is Nothing Then
                    LotColumn.UserHidden = bUserHidden
                    Set m.LotColumns(Str(LotColumn.ID)) = LotColumn
                
                    .ColHidden(lCol) = LotColumn.AlwaysHidden Or LotColumn.FeedyardHidden Or LotColumn.UserHidden
                End If
            Next lIndex
            Dirty = True
            
            .Redraw = flexRDBuffered
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.EditColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartMenuTimer
'' Description: Start the menu timer
'' Inputs:      Command, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartMenuTimer(ByVal strCommand As String, ByVal lRow As Long)
On Error GoTo ErrSection:

    tmrMenu.Interval = 100
    tmrMenu.Tag = strCommand & vbTab & Str(lRow)
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFeedyards
'' Description: Get the feedyards that the user can see
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetFeedyards()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        InfBox "Getting feedyard information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetFeedyards
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetFeedyards"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFeedyardCustomers
'' Description: Get the customers for the selected feed yard that the user can see
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetFeedyardCustomers()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID
    
    If (cboFeedYards.ListIndex >= 0) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting customer information.  Please wait...", , , g.Cattle.ProductName, True
        lFeedYardID = SelectedFeedYard
        g.Cattle.GetFeedyardCustomers lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetFeedyardCustomers"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleCustomers
'' Description: Get the customers for the selected feed yard that the user can see
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetVisibleCustomers()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID
    
    If (cboFeedYards.ListIndex >= 0) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting customer information.  Please wait...", , , g.Cattle.ProductName, True
        lFeedYardID = SelectedFeedYard
        g.Cattle.GetCustomers lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetVisibleCustomers"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectCustomer
'' Description: Select the given customer in the customers combo
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectCustomer(ByVal strCustomer As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To cboCustomers.ListCount - 1
        If Parse(cboCustomers.List(lIndex), "(", 1) = strCustomer Then
            cboCustomers.ListIndex = lIndex
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SelectCustomer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotColumnCategories
'' Description: Get a list of lot column categories
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetLotColumnCategories()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetLotColumnCategories
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetLotColumnCategories"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllLotColumns
'' Description: Get a list of lot columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllLotColumns()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If Not g.Cattle Is Nothing Then
        g.Cattle.GetAllLotColumns lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetAllLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleLotColumns
'' Description: Get a list of visible lot columns for the feedyard customer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetVisibleLotColumns()
On Error GoTo ErrSection:

    If Not g.Cattle Is Nothing Then
        g.Cattle.GetVisibleLotColumns
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetVisibleLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLots
'' Description: Get the lots for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetLots()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting lot information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetLots lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetLots"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotDetails
'' Description: Get the lot details for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetLotDetails()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        'InfBox "Getting lot detail information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetLotContentDetailsForYard Str(lFeedYardID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetLotDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLotDetailOptions
'' Description: Get the lot detail options for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetLotDetailOptions()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        g.Cattle.GetDetailOptions Str(lFeedYardID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetLotDetailOptions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetRations
'' Description: Get the rations for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetRations()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        g.Cattle.GetRations Str(lFeedYardID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetRations"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetIngredients
'' Description: Get the ingredients for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetIngredients()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        g.Cattle.GetIngredients Str(lFeedYardID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Get the accounts for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAccounts()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting account information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetAccounts lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTurnkeyOrders
'' Description: Get the orders for the selected feed yard from the Genesis
''              Turnkey server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetTurnkeyOrders()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting order information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetOrders lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetTurnkeyOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllTurnkeyFills
'' Description: Get all of the fills for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllTurnkeyFills()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting fill information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetAllFills lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetAllTurnkeyFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAssociatedTurnkeyFills
'' Description: Get all of the associated fills for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAssociatedTurnkeyFills()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting fill association information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetAssociatedFills lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetAssociatedTurnkeyFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTurnkeyTrades
'' Description: Get the trades for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetTurnkeyTrades()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting trade information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetTrades lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetTurnkeyTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTurnkeyPositions
'' Description: Get the positions for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetTurnkeyPositions()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting position information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetPositions lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetTurnkeyPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetVisibleLots
'' Description: Get the visible lot information for the selected feed yard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetVisibleLots()
On Error GoTo ErrSection:

    Dim lFeedYardID As Long             ' Feed Yard ID

    lFeedYardID = SelectedFeedYard
    If (lFeedYardID > -1&) And (Not g.Cattle Is Nothing) Then
        InfBox "Getting visible lot information.  Please wait...", , , g.Cattle.ProductName, True
        g.Cattle.GetVisibleLots lFeedYardID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetVisibleLots"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotToGrid
'' Description: Send the lot information to the grid
'' Inputs:      Lot Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LotToGrid(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:
    
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewRow As Boolean              ' Is this a new row?
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lPrevLastRow As Long            ' Previous last row in grid
    Dim Lot As cBrokerMessage           ' Lot information

    If (Len(turnkeyMessage("Deleted")) = 0) Or (turnkeyMessage("Deleted") = "0") Then
        With fgLots
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            bNewRow = False
            lRow = -1&
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) = 1 Then
                    If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                        Set Lot = .RowData(lIndex)
                        If Lot("FeedYardLotID") = turnkeyMessage("FeedYardLotID") Then
                            lRow = lIndex
                            Exit For
                        End If
                    End If
                End If
            Next lIndex
            
            lPrevLastRow = .Rows - 1
            
            If lRow = -1& Then
                .Rows = .Rows + 1
                lRow = .Rows - 1
                bNewRow = True
                
                .RowHidden(lRow) = True
            End If
            
            .IsSubtotal(lRow) = True
            .RowOutlineLevel(lRow) = 1
            .MergeRow(lRow) = False
            .RowData(lRow) = turnkeyMessage
            
            If lRow > .FixedRows Then
                If .Cell(flexcpBackColor, lRow - 1, 0) <> ALT_GRID_ROW_COLOR Then
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = ALT_GRID_ROW_COLOR
                End If
            End If
            
            For lIndex = 0 To NumberOfCols - 1
                Set LotColumn = LotColumnForCol(lIndex)
                If Not LotColumn Is Nothing Then
                    g.Cattle.GridValue(fgLots, lRow, lIndex, LotColumn) = turnkeyMessage(LotColumn.KeyValueField)
                
                    If UCase(LotColumn.Format) = "LINK" Then
                        .Cell(flexcpForeColor, lRow, lIndex) = vbBlue
                        .Cell(flexcpFontUnderline, lRow, lIndex) = True
                    End If
                End If
            Next lIndex
            
            .TextMatrix(lRow, LotCol(eGDLotCol_AscSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_" & Chr(31)
            .TextMatrix(lRow, LotCol(eGDLotCol_DescSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_}"
            .TextMatrix(lRow, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(lRow))
            
            RowType(lRow) = eGDRowType_Lot
            
            If bNewRow = True Then
                .Rows = .Rows + 1
                .IsSubtotal(.Rows - 1) = True
                .MergeRow(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 2
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_ClosedProfit) - 1) = "Click here to associate orders and fills to this lot"
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, .Rows - 2, 0)
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_ClosedProfit) - 1) = vbBlue
                .Cell(flexcpFontUnderline, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_ClosedProfit) - 1) = True
                
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_AscSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_}"
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_DescSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_" & Chr(31)
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(.Rows - 1))
                
                RowType(.Rows - 1) = eGDRowType_Associate
            End If
            
            .IsCollapsed(lRow) = flexOutlineCollapsed
            
            ' If this was an unsolicited lot that got added after the totals row, move the totals row
            ' back to the bottom...
            If turnkeyMessage("InRefresh") = "0" Then
                If lPrevLastRow <> .Rows - 1 Then
                    If .TextMatrix(lPrevLastRow, LotCol(eGDLotCol_LotNumber)) = "Totals" Then
                        .RowPosition(lPrevLastRow) = .Rows - 1
                    End If
                End If
            End If
            
            AutoSizeGrid
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.LotToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToGrid
'' Description: Send the order information to the grid
'' Inputs:      Order Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderToGrid(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid for the order
    Dim strSymbol As String             ' Symbol for the order
    Dim strAccountID As String          ' Account ID for the order
    Dim lGrandparent As Long            ' Grandparent to the row
    Dim bIsWorking As Boolean           ' Is the order working?
    Dim bRowAdded As Boolean            ' Was a new row added?

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        bIsWorking = (turnkeyMessage("IsWorking") = "1")
        
        If turnkeyMessage("NumberOfLegs") = 1 Then
            strSymbol = Parse(turnkeyMessage("Leg1"), ",", 3)
        Else
        End If
        
        strAccountID = turnkeyMessage("BrokerAccountID")
        
        lRow = GridRowForLotItem(turnkeyMessage("FeedyardLotID"), turnkeyMessage("ID"), eGDLotItemType_Order, strSymbol, strAccountID, bIsWorking, bRowAdded)
        If lRow <> -1& Then
            If bIsWorking Then
                .RowData(lRow) = turnkeyMessage
                
                .Cell(flexcpText, lRow, 0, lRow, .Cols - 1) = g.Cattle.OrderToString(turnkeyMessage)
                .MergeRow(lRow) = True
                .IsSubtotal(lRow) = True
                                
                .TextMatrix(lRow, LotCol(eGDLotCol_AscSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_" & Pad(turnkeyMessage("BrokerOrderID"), 40, "R")
                .TextMatrix(lRow, LotCol(eGDLotCol_DescSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_" & Pad(turnkeyMessage("BrokerOrderID"), 40, "R")
                .TextMatrix(lRow, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(lRow))
                
                RowType(lRow) = eGDRowType_Order
                
                lGrandparent = .GetNodeRow(.GetNodeRow(lRow, flexNTParent), flexNTParent)
                .IsCollapsed(lGrandparent) = .IsCollapsed(lGrandparent)
                
                If (bRowAdded = True) And (turnkeyMessage("InRefresh") = "0") Then
                    SortOnCol
                End If
            Else
                .RemoveItem lRow
            End If
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.OrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Send the fill information to the grid
'' Inputs:      Fill Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid for the order
    Dim lGrandparent As Long            ' Grandparent to the row
    Dim bRowAdded As Boolean            ' Was a new row added?

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = GridRowForLotItem(turnkeyMessage("FeedyardLotID"), turnkeyMessage("ID"), eGDLotItemType_Fill, turnkeyMessage("Symbol"), turnkeyMessage("BrokerAccountID"), True, bRowAdded)
        If lRow <> -1& Then
            .RowData(lRow) = turnkeyMessage
            
            .Cell(flexcpText, lRow, 0, lRow, LotCol(eGDLotCol_OpenEquity) - 1) = g.Cattle.FillToString(turnkeyMessage)
            .MergeRow(lRow) = True
            
            RowType(lRow) = eGDRowType_Fill
        
            lGrandparent = .GetNodeRow(.GetNodeRow(.GetNodeRow(lRow, flexNTParent), flexNTParent), flexNTParent)
            .IsCollapsed(lGrandparent) = .IsCollapsed(lGrandparent)
                
            If (bRowAdded = True) And (turnkeyMessage("InRefresh") = "0") Then
                SortOnCol
            End If
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeToGrid
'' Description: Send the trade information to the grid
'' Inputs:      Trade Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TradeToGrid(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid for the order
    Dim lGrandparent As Long            ' Grandparent to the row
    Dim Bars As cGdBars                 ' Bars object
    Dim strSymbol As String             ' Symbol
    Dim strAccountID As String          ' Account ID
    Dim bRowAdded As Boolean            ' Was a new row added?

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        strAccountID = turnkeyMessage("BrokerAccountID")
        
        lRow = GridRowForLotItem(turnkeyMessage("FeedyardLotID"), turnkeyMessage("ID"), eGDLotItemType_Trade, turnkeyMessage("Symbol"), strAccountID, True, bRowAdded)
        If lRow <> -1& Then
            .RowData(lRow) = turnkeyMessage
            
            .Cell(flexcpText, lRow, 0, lRow, LotCol(eGDLotCol_OpenEquity) - 1) = g.Cattle.TradeToString(turnkeyMessage, , False)
            CurrencyToGrid lRow, LotCol(eGDLotCol_ClosedProfit), turnkeyMessage("ClosedProfit")
            .MergeRow(lRow) = True
            .IsSubtotal(lRow) = True
 
            strSymbol = turnkeyMessage("Symbol")
            
            .TextMatrix(lRow, LotCol(eGDLotCol_AscSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_" & Chr(31) & "_" & Pad(turnkeyMessage("Sequence"), 5, "R")
            .TextMatrix(lRow, LotCol(eGDLotCol_DescSortKey)) = Pad(turnkeyMessage("FeedyardLotID"), 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_{_" & Pad(Str(99999 - CLng(Val(turnkeyMessage("Sequence")))), 5, "R")
            .TextMatrix(lRow, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(lRow))
            
            RowType(lRow) = eGDRowType_Trade
        
            CalculateOpenEquityForRow lRow
            
            lGrandparent = .GetNodeRow(.GetNodeRow(.GetNodeRow(lRow, flexNTParent), flexNTParent), flexNTParent)
            .IsCollapsed(lGrandparent) = .IsCollapsed(lGrandparent)
                
            If (bRowAdded = True) And (turnkeyMessage("InRefresh") = "0") Then
                SortOnCol
            End If
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.TradeToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionToGrid
'' Description: Send the position information to the grid
'' Inputs:      Position Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PositionToGrid(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid for the order
    Dim lPosition As Long               ' Current position for the symbol
    Dim lGrandparent As Long            ' Grandparent to the row
    Dim bRowAdded As Boolean            ' Was a new row added?

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lPosition = CLng(Val(turnkeyMessage("Quantity")))
        lRow = GridRowForLotItem(turnkeyMessage("FeedyardLotID"), turnkeyMessage("ID"), eGDLotItemType_Position, turnkeyMessage("Symbol"), turnkeyMessage("BrokerAccountID"), True, bRowAdded)
        If lRow <> -1& Then
            .RowData(lRow) = turnkeyMessage
            
            .Cell(flexcpText, lRow, 0, lRow, LotCol(eGDLotCol_OpenEquity) - 1) = g.Cattle.PositionToString(turnkeyMessage)
            .MergeRow(lRow) = True
            
            RowType(lRow) = eGDRowType_Position
        
            lGrandparent = .GetNodeRow(.GetNodeRow(lRow, flexNTParent), flexNTParent)
            .IsCollapsed(lGrandparent) = .IsCollapsed(lGrandparent)
                
            If (bRowAdded = True) And (turnkeyMessage("InRefresh") = "0") Then
                SortOnCol
            End If
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.PositionToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Filter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    Dim bHide As Boolean                ' Hide the rows?
    Dim lNextSibling As Long            ' Next sibling row in the grid
    Dim lLastDescendant As Long         ' Last descendant for this row
    Dim lFeedYardCustomerID As Long     ' Feed Yard customer ID
    Dim strVisibleLots As String        ' Visible lots
    Dim Lot As cBrokerMessage           ' Feed Yard Lot
    Dim nLotView As eGDLotView          ' Lot View
    Dim nCollapsed As CollapsedSettings ' Collapsed settings
    
    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lFeedYardCustomerID = SelectedCustomer
        If lFeedYardCustomerID > -1& Then
            If m.VisibleLots.Exists(Str(lFeedYardCustomerID)) Then
                strVisibleLots = "," & m.VisibleLots(Str(lFeedYardCustomerID)) & ","
            End If
        End If
        
        nLotView = cboView.ItemData(cboView.ListIndex)
        
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            Do While lRow <> -1&
                nCollapsed = .IsCollapsed(lRow)
                
                bHide = ((LotIsClosed(lRow) = True) And (mnuShowClosedLots.Checked = False))
                
                If bHide = False Then
                    If TypeOf .RowData(lRow) Is cBrokerMessage Then
                        Set Lot = .RowData(lRow)
                        
                        If lFeedYardCustomerID = -1& Then
                            bHide = False
                        Else
                            bHide = (InStr(strVisibleLots, "," & Lot("FeedYardLotID") & ",") = 0)
                        End If
                    End If
                End If
                
                lNextSibling = .GetNodeRow(lRow, flexNTNextSibling)
                If lNextSibling = -1& Then
                    lLastDescendant = .Rows - 1
                Else
                    lLastDescendant = lNextSibling - 1
                End If
                
                If (bHide = False) And (nLotView <> eGDLotView_AllLots) Then
                    If UCase(.TextMatrix(lRow, LotCol(eGDLotCol_LotNumber))) <> "TOTALS" Then
                        If lLastDescendant = lRow + 1 Then
                            bHide = (nLotView = eGDLotView_LotsWithHedge)
                        Else
                            bHide = (nLotView = eGDLotView_LotsWithoutHedge)
                        End If
                    End If
                End If
                
                For lIndex = lRow To lLastDescendant
                    .RowHidden(lIndex) = bHide
                Next lIndex
                
                If bHide = False Then
                    .IsCollapsed(lRow) = nCollapsed
                End If
                
                lRow = lNextSibling
            Loop
            
            ColorGridRows
            CalcTotals
        End If
                
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleConnection
'' Description: Toggle the Turnkey connection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleConnection()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout waiting for Turnkey to get disconnected

    If Not g.Cattle Is Nothing Then
        If Status = eGDConnectionStatus_Disconnected Then
            g.Cattle.Connect
        Else
            g.Cattle.Disconnect
            
            lTimeOut = 0&
            Do While (Status <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                Sleep 1
                lTimeOut = lTimeOut + 1&
            Loop
            
            If Status <> eGDConnectionStatus_Disconnected Then
                Status = eGDConnectionStatus_Disconnected
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ToggleConnection"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowForLotID
'' Description: Determine the grid row for the given lot ID
'' Inputs:      Lot ID
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowForLotID(ByVal strLotId As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    Dim turnkeyMessage As cBrokerMessage ' Broker message object
    
    lReturn = -1&
    With fgLots
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            
            Do While lRow <> -1&
                If TypeOf .RowData(lRow) Is cBrokerMessage Then
                    Set turnkeyMessage = .RowData(lRow)
                    If turnkeyMessage("FeedYardLotID") = strLotId Then
                        lReturn = lRow
                        Exit Do
                    End If
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
    End With
    
    GridRowForLotID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.GridRowForLotID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowForSymbol
'' Description: Determine the grid row for the given lot ID / Symbol
'' Inputs:      Lot ID, Symbol
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowForSymbol(ByVal strLotId As String, ByVal strSymbol As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    Dim lLotRow As Long                 ' Lot row
    
    lReturn = -1&
    With fgLots
        lLotRow = GridRowForLotID(strLotId)
        If lLotRow <> -1& Then
            lRow = .GetNodeRow(lLotRow, flexNTFirstChild)
            Do While lRow <> -1&
                If UCase(.TextMatrix(lRow, LotCol(eGDLotCol_LotNumber))) = UCase(strSymbol) Then
                    lReturn = lRow
                    Exit Do
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
    End With
    
    GridRowForSymbol = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.GridRowForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParentRowForLotItem
'' Description: Determine the parent row for the given information
'' Inputs:      Lot ID, Symbol, Account, Grid Row for Lot, Add Row if not Found?
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ParentRowForLotItem(ByVal strFeedYardLotID As String, ByVal strSymbol As String, ByVal strAccountID As String, Optional ByVal lGridRowForLotID As Long = -1&, Optional ByVal bAddRowIfNotFound As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid
    Dim lInsertAt As Long               ' Where to insert the parent row
    Dim strText As String               ' Text from the grid
    Dim strSymbolDisplay As String      ' Symbol display
    Dim strSymbolKey As String          ' Symbol key
    Dim Info As cBrokerMessage          ' Information object
    Dim strKey As String                ' Key from the grid
    
    lReturn = -1&
    lInsertAt = -1&
    
    If lGridRowForLotID = -1& Then
        lGridRowForLotID = GridRowForLotID(strFeedYardLotID)
    End If

    If lGridRowForLotID > -1& Then
        strSymbolDisplay = UCase(strSymbol) & " / " & UCase(g.Cattle.Accounts.DisplayAccountNumber(strAccountID)) & " (" & g.Cattle.Accounts.BrokerNameForAccountID(strAccountID) & ")"
        strSymbolKey = strSymbol & "|" & strAccountID
        
        With fgLots
            lRow = .GetNodeRow(lGridRowForLotID, flexNTFirstChild)
            Do While lRow <> -1&
                strText = UCase(.TextMatrix(lRow, 0))
                If TypeOf .RowData(lRow) Is cBrokerMessage Then
                    Set Info = .RowData(lRow)
                    strKey = Info("Key")
                Else
                    strKey = ""
                End If
                
                'If strText = UCase(strSymbolDisplay) Then
                If UCase(strKey) = UCase(strSymbolKey) Then
                    lReturn = lRow
                    Exit Do
                ElseIf (Left(strText, 10) = "CLICK HERE") And (lInsertAt = -1&) Then
                    lInsertAt = lRow
                    Exit Do
                ElseIf UCase(strSymbolDisplay) < strText Then
                    lInsertAt = lRow
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
            
            If (lReturn = -1&) And (bAddRowIfNotFound = True) And (lInsertAt > -1&) Then
                Set Info = New cBrokerMessage
                Info.Add "Symbol", strSymbol
                Info.Add "BrokerAccountID", strAccountID
                Info.Add "Key", strSymbolKey
                
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = Info
                .IsSubtotal(.Rows - 1) = True
                .MergeRow(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 2
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lGridRowForLotID, 0)
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_OpenEquity) - 1) = vbBlue
                .Cell(flexcpFontUnderline, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_OpenEquity) - 1) = True
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_OpenEquity) - 1) = strSymbolDisplay
                
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_AscSortKey)) = Pad(strFeedYardLotID, 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R")
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_DescSortKey)) = Pad(strFeedYardLotID, 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_}"
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(.Rows - 1))
                
                RowType(.Rows - 1) = eGDRowType_Symbol
                
                .RowPosition(.Rows - 1) = lInsertAt
                lReturn = lInsertAt
                
                .Rows = .Rows + 1
                .IsSubtotal(.Rows - 1) = True
                .MergeRow(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 3
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lGridRowForLotID, 0)
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, LotCol(eGDLotCol_OpenEquity) - 1) = "No Position"
                
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_AscSortKey)) = Pad(strFeedYardLotID, 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_" & Chr(31)
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_DescSortKey)) = Pad(strFeedYardLotID, 30, "R") & "_" & Pad(strSymbol, 40, "R") & "_" & Pad(strAccountID, 15, "R") & "_|"
                .TextMatrix(.Rows - 1, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(.Rows - 1))
                
                RowType(.Rows - 1) = eGDRowType_NoPosition
                
                .RowPosition(.Rows - 1) = lInsertAt + 1
            End If
        End With
    End If
    
    ParentRowForLotItem = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.ParentRowForLotItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridRowForLotItem
'' Description: Determine the grid row for the given information
'' Inputs:      Lot ID, Item ID, Type, Symbol, Account, Add Row if not Found?
'' Returns:     Grid Row (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GridRowForLotItem(ByVal strFeedYardLotID As String, ByVal strItemID As String, ByVal nType As eGDLotItemType, ByVal strSymbol As String, ByVal strAccountID As String, Optional ByVal bAddRowIfNotFound As Boolean = False, Optional bRowAdded As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lParentRow As Long              ' Parent row for the item
    Dim lRow As Long                    ' Row in the grid
    Dim turnkeyMessage As cBrokerMessage ' Broker message object
    Dim lLastChild As Long              ' Last child for the parent
    Dim lGridRowForLotID As Long        ' Grid row for the Lot
    Dim lPositionRow As Long            ' Grid row for the Position
    
    lReturn = -1&
    bRowAdded = False
    
    lGridRowForLotID = GridRowForLotID(strFeedYardLotID)
    If lGridRowForLotID > -1& Then
        With fgLots
            lParentRow = ParentRowForLotItem(strFeedYardLotID, strSymbol, strAccountID, lGridRowForLotID, True)
            If lParentRow > -1& Then
                Select Case nType
                    Case eGDLotItemType_Position
                        lReturn = .GetNodeRow(lParentRow, flexNTFirstChild)
                    
                    Case eGDLotItemType_Order
                        lRow = .GetNodeRow(.GetNodeRow(lParentRow, flexNTFirstChild), flexNTNextSibling)
                        Do While lRow <> -1&
                            If TypeOf .RowData(lRow) Is cBrokerMessage Then
                                Set turnkeyMessage = .RowData(lRow)
                                If turnkeyMessage("ID") = strItemID Then
                                    lReturn = lRow
                                    Exit Do
                                End If
                            End If
                            
                            lRow = .GetNodeRow(lRow, flexNTNextSibling)
                        Loop
                    
                    Case eGDLotItemType_Trade
                        lRow = .GetNodeRow(.GetNodeRow(lParentRow, flexNTFirstChild), flexNTFirstChild)
                        Do While lRow <> -1&
                            If TypeOf .RowData(lRow) Is cBrokerMessage Then
                                Set turnkeyMessage = .RowData(lRow)
                                If turnkeyMessage("ID") = strItemID Then
                                    lReturn = lRow
                                    Exit Do
                                End If
                            End If
                            
                            lRow = .GetNodeRow(lRow, flexNTNextSibling)
                        Loop
                
                End Select
                                
                If (lReturn = -1&) And (bAddRowIfNotFound = True) Then
                    .Rows = .Rows + 1
                    .IsSubtotal(.Rows - 1) = True
                    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lParentRow, 0)
                    
                    Select Case nType
                        Case eGDLotItemType_Position
                            .RowOutlineLevel(.Rows - 1) = 3
                            lReturn = lParentRow + 1&
                        
                        Case eGDLotItemType_Order
                            .RowOutlineLevel(.Rows - 1) = 3
                            lReturn = .GetNodeRow(lParentRow, flexNTLastChild) + 1&
                        
                        Case eGDLotItemType_Trade
                            .RowOutlineLevel(.Rows - 1) = 4
                            lPositionRow = .GetNodeRow(lParentRow, flexNTFirstChild)
                            lReturn = .GetNodeRow(lPositionRow, flexNTLastChild) + 1&
                            If lReturn = 0& Then
                                lReturn = lPositionRow + 1&
                            End If
                    End Select
                    
                    .RowPosition(.Rows - 1) = lReturn
                    bRowAdded = True
                End If
            End If
        End With
    End If
    
    GridRowForLotItem = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.GridRowForLotItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociateLotItems
'' Description: Associate items for a specific lot
'' Inputs:      Grid Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssociateLotItems(ByVal lLotRow As Long)
On Error GoTo ErrSection:

    Dim Lot As cBrokerMessage           ' Lot from the grid
    Dim Orders As cGdTree               ' Associated orders
    Dim Fills As cGdTree                ' Associated fills
    Dim lIndex As Long                  ' Index into a for loop
    Dim turnkeyMsg As cBrokerMessage    ' Turnkey message object
    Dim BrokerAccounts As cGdTree       ' Broker list of accounts
    Dim BrokerAccount As cBrokerMessage ' Broker account
    Dim astrAccountsAdded As cGdArray   ' List of accounts sent to the Genesis Turnkey server
    Dim lPos As Long                    ' Position in the array
    Dim strAccount As String            ' Account number
    Dim TurnkeyAcct As cBrokerMessage   ' Turnkey account to send
    Dim strAccountsKey As String        ' Key into the accounts collection
    Dim strBrokerOrderID As String      ' Broker order ID
    Dim strBrokerFillID As String       ' Broker fill ID
    
    If m.AssociatedOrdersSent.Count > 0 Then
        InfBox "You have order associations pending confirmation.  Please wait...", , , "Lot Item Association", True
    ElseIf m.AssociatedFillsSent.Count > 0 Then
        InfBox "You have fill associations pending confirmation.  Please wait...", , , "Lot Item Association", True
    Else
        With fgLots
            If lLotRow = -1& Then
                Set Lot = Nothing
            Else
                If TypeOf .RowData(lLotRow) Is cBrokerMessage Then
                    Set Lot = .RowData(lLotRow)
                Else
                    Set Lot = Nothing
                End If
            End If
            
            Set Orders = m.AssociatedOrders.MakeCopy
            Set Fills = m.AssociatedFills.MakeCopy
            
            If frmAssociateLotItems.ShowMe(Orders, Fills, m.Accounts, m.Fills, Lot) Then
                Set BrokerAccounts = g.AppBridge.GetBrokerAccounts
                Set astrAccountsAdded = New cGdArray
                
                For lIndex = 1 To Orders.Count
                    Set turnkeyMsg = Orders(lIndex)
                    If turnkeyMsg("NewFeedYardLotID") <> turnkeyMsg("FeedYardLotID") Then
                        turnkeyMsg.Add "FeedYardLotID", turnkeyMsg("NewFeedYardLotID")
                        
                        strAccountsKey = AccountKey(turnkeyMsg("Broker"), turnkeyMsg("BrokerAccountNumber"))
                        
                        If m.Accounts.Exists(turnkeyMsg("BrokerAccountID")) Then
                            UpdateOrder turnkeyMsg, "order being associated"
                        ElseIf BrokerAccounts.Exists(strAccountsKey) Then
                            strAccount = turnkeyMsg("BrokerAccountNumber")
                            
                            Set BrokerAccount = BrokerAccounts(strAccountsKey)
                            If astrAccountsAdded.BinarySearch(strAccountsKey, lPos) = False Then
                                astrAccountsAdded.Add strAccountsKey, lPos
                                
                                BrokerAccount.Add "FeedYardID", turnkeyMsg("FeedYardID")
                                g.Cattle.UpdateAccount BrokerAccount
                            End If
                            
                            strBrokerOrderID = OrderKey(turnkeyMsg("Broker"), turnkeyMsg("BrokerOrderID"))
                            
                            g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' added to list of orders waiting for the account to come back"
                            m.WaitAccountOrders.Add turnkeyMsg, strBrokerOrderID
                        End If
                    End If
                Next lIndex
                For lIndex = 1 To Fills.Count
                    Set turnkeyMsg = Fills(lIndex)
                    If turnkeyMsg("Changed") = "1" Then
                        strAccountsKey = AccountKey(turnkeyMsg("Broker"), turnkeyMsg("BrokerAccountNumber"))
                        
                        If m.Accounts.Exists(turnkeyMsg("BrokerAccountID")) Then
                            UpdateFillAssociation turnkeyMsg, "fill being associated"
                        ElseIf BrokerAccounts.Exists(strAccountsKey) Then
                            strAccount = turnkeyMsg("BrokerAccountNumber")
                            
                            Set BrokerAccount = BrokerAccounts(strAccountsKey)
                            If astrAccountsAdded.BinarySearch(strAccountsKey, lPos) = False Then
                                astrAccountsAdded.Add strAccountsKey, lPos
                                
                                BrokerAccount.Add "FeedYardID", turnkeyMsg("FeedYardID")
                                g.Cattle.UpdateAccount BrokerAccount
                            End If
                            
                            strBrokerFillID = AssociatedFillKey(turnkeyMsg("Broker"), turnkeyMsg("BrokerFillID"), turnkeyMsg("FeedYardLotID"))
                            
                            g.Cattle.DumpDebug "Fill '" & strBrokerFillID & "' added to list of fills waiting for the account to come back"
                            m.WaitAccountFills.Add turnkeyMsg, strBrokerFillID
                        End If
                    End If
                Next lIndex
                
                For lIndex = m.AssociatedOrders.Count To 1 Step -1
                    Set turnkeyMsg = m.AssociatedOrders(lIndex)
                    If Orders.Exists(OrderKey(turnkeyMsg("Broker"), turnkeyMsg("BrokerOrderID"))) = False Then
                        turnkeyMsg.Add "PreviousFeedYardLotID", turnkeyMsg("FeedyardLotID")
                        turnkeyMsg("FeedyardLotID") = "0"
                        UpdateOrder turnkeyMsg, "order being unassociated"
                    End If
                Next lIndex
                For lIndex = m.AssociatedFills.Count To 1 Step -1
                    Set turnkeyMsg = m.AssociatedFills(lIndex)
                    If Fills.Exists(AssociatedFillKey(turnkeyMsg("Broker"), turnkeyMsg("BrokerFillID"), turnkeyMsg("FeedYardLotID"))) = False Then
                        turnkeyMsg.Add "PreviousFeedYardLotID", turnkeyMsg("FeedyardLotID")
                        turnkeyMsg("FeedyardLotID") = "0"
                        UpdateFillAssociation turnkeyMsg, "fill being unassociated"
                    End If
                Next lIndex
                
                If (m.AssociatedOrdersSent.Count > 0) Or (m.AssociatedFillsSent.Count > 0) Then
                    InfBox "Sending order and fill assocations to server.  Please wait...", , , "Lot Item Association", True
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.AssociateLotItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckWaitAccountItems
'' Description: Check the pending items to see if they can be sent
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckWaitAccountItems()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim turnkeyMsg As cBrokerMessage    ' Turnkey message object
    Dim strBrokerAccountID As String    ' Broker Account ID
    
    For lIndex = m.WaitAccountOrders.Count To 1 Step -1
        Set turnkeyMsg = m.WaitAccountOrders(lIndex)
        strBrokerAccountID = g.Cattle.Accounts.AccountIdForBrokerNumber(turnkeyMsg("BrokerAccountNumber"), turnkeyMsg("Broker"))
        If Len(strBrokerAccountID) > 0 Then
            turnkeyMsg("BrokerAccountID") = strBrokerAccountID
            UpdateOrder turnkeyMsg, "account has been added"
            
            g.Cattle.DumpDebug "Order '" & turnkeyMsg("BrokerOrderID") & "' removed from list of orders waiting for the account to come back"
            m.WaitAccountOrders.Remove lIndex
        End If
    Next lIndex

    For lIndex = m.WaitAccountFills.Count To 1 Step -1
        Set turnkeyMsg = m.WaitAccountFills(lIndex)
        strBrokerAccountID = g.Cattle.Accounts.AccountIdForBrokerNumber(turnkeyMsg("BrokerAccountNumber"), turnkeyMsg("Broker"))
        If Len(strBrokerAccountID) > 0 Then
            turnkeyMsg("BrokerAccountID") = strBrokerAccountID
            UpdateFillAssociation turnkeyMsg, "account has been added"
            
            g.Cattle.DumpDebug "Fill '" & m.WaitAccountFills.Key(lIndex) & "' removed from list of fills waiting for the account to come back"
            m.WaitAccountFills.Remove lIndex
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CheckWaitAccountItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTotalsRow
'' Description: Add the totals row to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddTotalsRow()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .IsSubtotal(.Rows - 1) = True
        .MergeRow(.Rows - 1) = False
        .RowOutlineLevel(.Rows - 1) = 1
        
        .TextMatrix(.Rows - 1, LotCol(eGDLotCol_LotNumber)) = "Totals"
        .TextMatrix(.Rows - 1, LotCol(eGDLotCol_AscSortKey)) = "}"
        .TextMatrix(.Rows - 1, LotCol(eGDLotCol_DescSortKey)) = Chr(31)
        .TextMatrix(.Rows - 1, LotCol(eGDLotCol_OutlineLevel)) = Str(.RowOutlineLevel(.Rows - 1))
        
        RowType(.Rows - 1) = eGDRowType_Total
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.AddTotalsRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotals
'' Description: Calculate the totals
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcTotals(Optional ByVal bOnlyEquity As Boolean = False)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim dClosed As Double               ' Closed profit for the trade
    Dim dOpen As Double                 ' Open equity for the trade
    Dim dTotalPositionClosed As Double  ' Total closed profit for a position
    Dim dTotalPositionOpen As Double    ' Total open equity for a position
    Dim dTotalLotClosed As Double       ' Total closed profit for a lot
    Dim dTotalLotOpen As Double         ' Total open profit for a lot
    Dim dTotalLotProfit As Double       ' Open Equity + Closed Profit for the lot
    Dim lLotRow As Long                 ' Row for a lot in the grid
    Dim lSymbolRow As Long              ' Row for a symbol in the grid
    Dim lPositionRow As Long            ' Row for a position in the grid
    Dim lTradeRow As Long               ' Row for a trade in the grid
    Dim adTotals As cGdArray            ' Array of totals for appropriate columns
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFromCol As Long                ' From column for the for loop
    Dim lToCol As Long                  ' To column for the for loop
    Dim LotColumn As cLotColumn         ' Lot Column object

    Set adTotals = New cGdArray
    adTotals.Create eGDARRAY_Doubles, NumberOfCols, 0#

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If bOnlyEquity Then
            lFromCol = LotCol(eGDLotCol_OpenEquity)
            lToCol = LotCol(eGDLotCol_ClosedProfit)
        Else
            lFromCol = 1
            lToCol = .Cols - 1
        End If
        
        lLotRow = .FixedRows
        Do While lLotRow > -1&
            dTotalLotClosed = 0#
            dTotalLotOpen = 0#
            
            If UCase(.TextMatrix(lLotRow, LotCol(eGDLotCol_LotNumber))) <> "TOTALS" Then
                lSymbolRow = .GetNodeRow(lLotRow, flexNTFirstChild)
                Do While lSymbolRow > -1&
                    dTotalPositionClosed = 0#
                    dTotalPositionOpen = 0#
                    
                    lPositionRow = .GetNodeRow(lSymbolRow, flexNTFirstChild)
                    If lPositionRow > -1& Then
                        lTradeRow = .GetNodeRow(lPositionRow, flexNTFirstChild)
                        Do While lTradeRow > -1&
                            dClosed = ValOfText(.TextMatrix(lTradeRow, LotCol(eGDLotCol_ClosedProfit)))
                            dOpen = ValOfText(.TextMatrix(lTradeRow, LotCol(eGDLotCol_OpenEquity)))
                            
                            dTotalPositionOpen = dTotalPositionOpen + dOpen
                            dTotalPositionClosed = dTotalPositionClosed + dClosed
                            dTotalLotOpen = dTotalLotOpen + dOpen
                            dTotalLotClosed = dTotalLotClosed + dClosed
                            
                            lTradeRow = .GetNodeRow(lTradeRow, flexNTNextSibling)
                        Loop
                        
                        CurrencyToGrid lPositionRow, LotCol(eGDLotCol_OpenEquity), Str(dTotalPositionOpen)
                        CurrencyToGrid lPositionRow, LotCol(eGDLotCol_ClosedProfit), Str(dTotalPositionClosed)
                        CurrencyToGrid lSymbolRow, LotCol(eGDLotCol_OpenEquity), Str(dTotalPositionOpen)
                        CurrencyToGrid lSymbolRow, LotCol(eGDLotCol_ClosedProfit), Str(dTotalPositionClosed)
                    End If
                    
                    lSymbolRow = .GetNodeRow(lSymbolRow, flexNTNextSibling)
                Loop
            
                CurrencyToGrid lLotRow, LotCol(eGDLotCol_OpenEquity), Str(dTotalLotOpen)
                CurrencyToGrid lLotRow, LotCol(eGDLotCol_ClosedProfit), Str(dTotalLotClosed)
                
                dTotalLotProfit = dTotalLotOpen + dTotalLotClosed
                If dTotalLotProfit > 0 Then
                    .Cell(flexcpForeColor, lLotRow, LotCol(eGDLotCol_LotNumber)) = vbGreen
                ElseIf dTotalLotProfit < 0 Then
                    .Cell(flexcpForeColor, lLotRow, LotCol(eGDLotCol_LotNumber)) = vbRed
                Else
                    .Cell(flexcpForeColor, lLotRow, LotCol(eGDLotCol_LotNumber)) = vbBlue
                End If
                
                If .RowHidden(lLotRow) = False Then
                    For lIndex = lFromCol To lToCol
                        Set LotColumn = LotColumnForCol(lIndex)
                        If Not LotColumn Is Nothing Then
                            If LotColumn.Total = True Then
                                adTotals(lIndex) = adTotals(lIndex) + ValOfText(.TextMatrix(lLotRow, lIndex))
                            End If
                        End If
                    Next lIndex
                End If
            Else
                For lIndex = lFromCol To lToCol
                    Set LotColumn = LotColumnForCol(lIndex)
                    If LotColumn Is Nothing Then
                        .TextMatrix(lLotRow, lIndex) = ""
                    ElseIf LotColumn.Total = True Then
                        If UCase(LotColumn.Format) = "CURRENCY" Then
                            If (lIndex = LotCol(eGDLotCol_ClosedProfit)) Or (lIndex = LotCol(eGDLotCol_OpenEquity)) Then
                                CurrencyToGrid lLotRow, lIndex, Str(adTotals(lIndex))
                            Else
                                g.Cattle.GridValue(fgLots, lLotRow, lIndex, LotColumn) = Str(adTotals(lIndex))
                            End If
                        Else
                            g.Cattle.GridValue(fgLots, lLotRow, lIndex, LotColumn) = Str(adTotals(lIndex))
                        End If
                    Else
                        .TextMatrix(lLotRow, lIndex) = ""
                    End If
                Next lIndex
            End If
            
            lLotRow = .GetNodeRow(lLotRow, flexNTNextSibling)
        Loop
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CalcTotals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrencyToGrid
'' Description: Set the given cell to the given value
'' Inputs:      Row, Column, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CurrencyToGrid(ByVal lRow As Long, ByVal lCol As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value converted to a double
    Dim LotColumn As cLotColumn         ' Lot column object
    
    If Len(strValue) = 0 Then
        fgLots.TextMatrix(lRow, lCol) = ""
    Else
        Set LotColumn = LotColumnForCol(lCol)
        
        dValue = Val(strValue)
        g.Cattle.GridValue(fgLots, lRow, lCol, LotColumn) = strValue
        
        If dValue < 0 Then
            fgLots.Cell(flexcpForeColor, lRow, lCol) = vbRed
        ElseIf dValue = 0 Then
            fgLots.Cell(flexcpForeColor, lRow, lCol) = vbBlack
        Else
            fgLots.Cell(flexcpForeColor, lRow, lCol) = vbGreen
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CurrencyToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateOpenEquityForRow
'' Description: Calculate the open equity for the given row
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateOpenEquityForRow(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey message
    Dim dEntryPrice As Double           ' Entry price
    Dim dCurrentPrice As Double         ' Current price
    Dim dPriceDiff As Double            ' Price difference
    Dim dOpenEquity As Double           ' Open equity
    Dim dFees As Double                 ' Fees for the trade
    
    With fgLots
        Set turnkeyMessage = .RowData(lRow)
        
        If Len(turnkeyMessage("ExitBrokerFillID")) = 0 Then
            dEntryPrice = Val(turnkeyMessage("EntryFillPrice"))
            dCurrentPrice = g.AppBridge.LastKnownPrice(turnkeyMessage("Symbol"))
            
            If dCurrentPrice = kNullData Then
                CurrencyToGrid lRow, LotCol(eGDLotCol_OpenEquity), ""
            Else
                If turnkeyMessage("IsBuy") = "1" Then
                    dPriceDiff = dCurrentPrice - dEntryPrice
                Else
                    dPriceDiff = dEntryPrice - dCurrentPrice
                End If
                dFees = Val(turnkeyMessage("EntryCommission")) + Val(turnkeyMessage("ExitCommission"))
                dOpenEquity = g.AppBridge.Profit(turnkeyMessage("Symbol"), dPriceDiff, CLng(Val(turnkeyMessage("Quantity"))), , , , turnkeyMessage("AccountNumber")) - dFees
                
                CurrencyToGrid lRow, LotCol(eGDLotCol_OpenEquity), Str(dOpenEquity)
            End If
        Else
            CurrencyToGrid lRow, LotCol(eGDLotCol_OpenEquity), "0"
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CalculateOpenEquityForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteTrade
'' Description: Delete the given trade from the grid
'' Inputs:      Trade
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteTrade(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lTradeRow As Long               ' Row for a trade in the grid
    
    With fgLots
        lTradeRow = GridRowForLotItem(turnkeyMessage("FeedYardLotID"), turnkeyMessage("ID"), eGDLotItemType_Trade, turnkeyMessage("Symbol"), turnkeyMessage("BrokerAccountID"))
        If lTradeRow <> -1& Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            .RemoveItem lTradeRow
            
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.DeleteTrade"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrder
'' Description: Delete the given order from the grid
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteOrder(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lOrderRow As Long               ' Row for the order in the grid
    Dim orderMessage As cBrokerMessage  ' Order message
    Dim strSymbol As String             ' Symbol for the order
    Dim strBrokerOrderID As String      ' Broker order ID
    
    strBrokerOrderID = OrderKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerOrderID"))
    If m.AssociatedOrders.Exists(strBrokerOrderID) Then
        Set orderMessage = m.AssociatedOrders(strBrokerOrderID)
        g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' removed from list of associated orders"
        m.AssociatedOrders.Remove strBrokerOrderID
    
        With fgLots
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If orderMessage("NumberOfLegs") = 1 Then
                strSymbol = Parse(orderMessage("Leg1"), ",", 3)
            Else
            End If
            
            lOrderRow = GridRowForLotItem(turnkeyMessage("FeedYardLotID"), turnkeyMessage("ID"), eGDLotItemType_Order, strSymbol, turnkeyMessage("BrokerAccountID"))
            If lOrderRow <> -1& Then
                .RemoveItem lOrderRow
            End If
            
            .Redraw = nRedraw
        End With
    End If
    
    If m.AssociatedOrdersSent.Exists(strBrokerOrderID) Then
        g.Cattle.DumpDebug "Order '" & strBrokerOrderID & "' removed from list of associated orders sent"
        m.AssociatedOrdersSent.Remove strBrokerOrderID
        
        ItemReceived
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.DeleteOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteAssociatedFill
'' Description: Delete the given associated fill from the grid
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteAssociatedFill(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim strAssociatedFillKey As String  ' Associated fill key
    
    strAssociatedFillKey = AssociatedFillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"), turnkeyMessage("FeedYardLotID"))
    If m.AssociatedFills.Exists(strAssociatedFillKey) Then
        g.Cattle.DumpDebug "Fill '" & strAssociatedFillKey & "' removed from list of associated fills"
        m.AssociatedFills.Remove strAssociatedFillKey
    End If
    
    If m.AssociatedFillsSent.Exists(strAssociatedFillKey) Then
        g.Cattle.DumpDebug "Fill '" & strAssociatedFillKey & "' removed from list of associated fills sent"
        m.AssociatedFillsSent.Remove strAssociatedFillKey
        
        ItemReceived
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.DeleteAssociatedFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeletePosition
'' Description: Delete the given position from the grid
'' Inputs:      Trade
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeletePosition(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lPositionRow As Long            ' Row for a position in the grid
    Dim lSymbolRow As Long              ' Symbol row for the position row
    
    With fgLots
        lPositionRow = GridRowForLotItem(turnkeyMessage("FeedYardLotID"), turnkeyMessage("ID"), eGDLotItemType_Position, turnkeyMessage("Symbol"), turnkeyMessage("BrokerAccountID"))
        If lPositionRow <> -1& Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            lSymbolRow = .GetNodeRow(lPositionRow, flexNTParent)
            
            .RemoveItem lPositionRow
            
            ' If after deleting the position, there are no children for the symbol row,
            ' get rid of the symbol row as well...
            If .GetNodeRow(lSymbolRow, flexNTFirstChild) = -1& Then
                .RemoveItem lSymbolRow
            End If
            
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.DeletePosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateOrder
'' Description: Send an update for the given order to the Genesis Turnkey servers
'' Inputs:      Order, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateOrder(ByVal orderMessage As cBrokerMessage, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim strOrderKey As String           ' Order key

    strOrderKey = OrderKey(orderMessage("Broker"), orderMessage("BrokerOrderID"))

    g.Cattle.DumpDebug "Sending order '" & strOrderKey & "' because " & strReason
    If Len(orderMessage("ID")) = 0 Then
        g.Cattle.DumpDebug vbTab & "Order '" & strOrderKey & "' added to collection of orders sent"
        m.OrdersSent.Add orderMessage, strOrderKey
    End If

    g.Cattle.DumpDebug vbTab & "Order '" & strOrderKey & "' added to collection of associated orders sent"
    m.AssociatedOrdersSent.Add orderMessage, strOrderKey
    
    ItemSent

    g.Cattle.UpdateOrder orderMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.UpdateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFill
'' Description: Send an update for the given fill to the Genesis Turnkey servers
'' Inputs:      Fill, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateFill(ByVal FillMessage As cBrokerMessage, ByVal strReason As String)
On Error GoTo ErrSection:

    g.Cattle.DumpDebug "Sending fill '" & FillMessage("BrokerFillID") & "' because " & strReason
    g.Cattle.UpdateFill FillMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.UpdateFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateFillAssociation
'' Description: Send an update for the given fill association to the Genesis
''              Turnkey servers
'' Inputs:      Associated Fill, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateFillAssociation(ByVal associatedFillMessage As cBrokerMessage, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collections

    strKey = AssociatedFillKey(associatedFillMessage("Broker"), associatedFillMessage("BrokerFillID"), associatedFillMessage("FeedYardLotID"))
    
    g.Cattle.DumpDebug "Sending fill association '" & strKey & "' because " & strReason
    g.Cattle.AssociateFill associatedFillMessage
    
    If (associatedFillMessage("FeedYardLotID") = "0") And (Len(associatedFillMessage("PreviousFeedYardLotID")) > 0) Then
        strKey = AssociatedFillKey(associatedFillMessage("Broker"), associatedFillMessage("BrokerFillID"), associatedFillMessage("PreviousFeedYardLotID"))
    End If
    
    g.Cattle.DumpDebug vbTab & "Fill '" & strKey & "' added to collection of associated fills sent"
    m.AssociatedFillsSent.Add associatedFillMessage, strKey
    
    ItemSent

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.UpdateFillAssociation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckWaitOrderItems
'' Description: Check the items that may have been waiting for this order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckWaitOrderItems(ByVal orderMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Message to send to Genesis Turnkey servers
    Dim lIndex As Long                  ' Index into a for loop
    Dim strOrderKey As String           ' Order key
    
    strOrderKey = OrderKey(orderMessage("Broker"), orderMessage("BrokerOrderID"))
    
    If m.WaitOrderOrders.Exists(strOrderKey) Then
        Set turnkeyMessage = m.WaitOrderOrders(strOrderKey)
        
        turnkeyMessage.Add "ID", orderMessage("ID")
        turnkeyMessage.Add "FeedYardLotID", orderMessage("FeedYardLotID")
        
        UpdateOrder turnkeyMessage, "order is now associated"
        
        g.Cattle.DumpDebug "Order '" & strOrderKey & "' removed from list of orders waiting for order to come back"
        m.WaitOrderOrders.Remove strOrderKey
    End If
    
    For lIndex = m.WaitOrderFills.Count To 1 Step -1
        Set turnkeyMessage = m.WaitOrderFills(lIndex)
        If turnkeyMessage("BrokerOrderID") = orderMessage("BrokerOrderID") Then
            turnkeyMessage.Add "FeedYardLotID", orderMessage("FeedYardLotID")
            
            UpdateFillAssociation turnkeyMessage, "order is now associated"
            
            g.Cattle.DumpDebug "Fill '" & m.WaitOrderFills.Key(lIndex) & "' removed from list of fills waiting for order to come back"
            m.WaitOrderFills.Remove lIndex
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CheckWaitOrderItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckWaitFillItems
'' Description: Check the items that may have been waiting for this fill
'' Inputs:      Fill
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckWaitFillItems(ByVal FillMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim turnkeyMessage As cBrokerMessage ' Message to send to Genesis Turnkey servers
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = m.WaitFillFills.Count To 1 Step -1
        Set turnkeyMessage = m.WaitFillFills(lIndex)
        If turnkeyMessage("BrokerFillID") = FillMessage("BrokerFillID") Then
            turnkeyMessage.Add "FillID", FillMessage("ID")
            turnkeyMessage("ID") = ""
        
            UpdateFillAssociation turnkeyMessage, "fill confirmed"
            
            g.Cattle.DumpDebug "Associated Fill '" & m.WaitFillFills.Key(lIndex) & "' removed from list of associated fills waiting for fill to come back"
            m.WaitFillFills.Remove lIndex
            
            bReturn = True
        End If
    Next lIndex
    
    CheckWaitFillItems = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.CheckWaitFillItems"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleOrderUpdate
'' Description: Handle an order update from the broker
'' Inputs:      Turnkey Order, Key Value?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleOrderUpdate(ByVal turnkeyMessage As cBrokerMessage, ByVal bKeyValue As Boolean)
On Error GoTo ErrSection:

    Dim oldMessage As cBrokerMessage    ' Old Turnkey message
    Dim strOrderKey As String           ' Order key
    Dim strPreviousOrderKey As String   ' Previous order key
    Dim TurnkeyAcct As cBrokerMessage   ' Turnkey account message
    
    If Len(turnkeyMessage("BrokerOrderID")) > 0 Then
        strOrderKey = OrderKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerOrderID"))
        strPreviousOrderKey = OrderKey(turnkeyMessage("Broker"), turnkeyMessage("PreviousBrokerOrderID"))
        
        ' Order is already associated...
        If m.AssociatedOrders.Exists(strOrderKey) Then
            Set oldMessage = m.AssociatedOrders(strOrderKey)
            
            turnkeyMessage.Add "ID", oldMessage("ID")
            turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
            
            If oldMessage.ToString <> turnkeyMessage.ToString Then
                UpdateOrder turnkeyMessage, "order has changed"
            End If
            
        ' Previous version of this order was associated...
        ElseIf (m.AssociatedOrders.Exists(strPreviousOrderKey) = True) And (m.OrdersSent.Exists(strOrderKey) = False) Then
            Set oldMessage = m.AssociatedOrders(strPreviousOrderKey)
            
            oldMessage("Status") = Str(eTT_OrderStatus_Amended)
            oldMessage("StatusDate") = turnkeyMessage("StatusDate")
            oldMessage("IsWorking") = "0"
            
            UpdateOrder oldMessage, "order set to amended"
            
            turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
            
            UpdateOrder turnkeyMessage, "new version of order " & strPreviousOrderKey
            
        ' Order was waiting for a Broker ID before being associated...
        ElseIf m.NewOrders.Exists(turnkeyMessage("GenesisOrderID")) Then
            turnkeyMessage.Add "FeedYardLotID", m.NewOrders(turnkeyMessage("GenesisOrderID"))
            
            g.Cattle.DumpDebug "Order '" & turnkeyMessage("GenesisOrderID") & "' removed from list of new orders"
            m.NewOrders.Remove turnkeyMessage("GenesisOrderID")
            
            If m.Accounts.Exists(turnkeyMessage("BrokerAccountID")) Then
                UpdateOrder turnkeyMessage, "order now has a Broker Order ID"
            ElseIf m.WaitAccountOrders.Exists(strOrderKey) Then
                Set oldMessage = m.WaitAccountOrders(strOrderKey)
                
                turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
                turnkeyMessage.Add "FeedYardCustomerID", oldMessage("FeedYardCustomerID")
                
                g.Cattle.DumpDebug "Order '" & strOrderKey & "' updated in list of orders waiting for the account to come back"
                Set m.WaitAccountOrders(strOrderKey) = turnkeyMessage
            Else
                Set TurnkeyAcct = g.AppBridge.AccountForBrokerNumber(turnkeyMessage("BrokerAccountNumber"), CLng(Val(turnkeyMessage("Broker"))), bKeyValue)
                If Not TurnkeyAcct Is Nothing Then
                    TurnkeyAcct.Add "FeedYardID", Str(SelectedFeedYard)
                    g.Cattle.UpdateAccount TurnkeyAcct
                
                    g.Cattle.DumpDebug "Order '" & strOrderKey & "' added to list of orders waiting for the account to come back"
                    m.WaitAccountOrders.Add turnkeyMessage, strOrderKey
                End If
            End If
            
        ' Order is waiting for the account to come back from Turnkey...
        ElseIf m.WaitAccountOrders.Exists(strOrderKey) Then
            Set oldMessage = m.WaitAccountOrders(strOrderKey)
            
            turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
            turnkeyMessage.Add "FeedYardCustomerID", oldMessage("FeedYardCustomerID")
            
            g.Cattle.DumpDebug "Order '" & strOrderKey & "' updated in list of orders waiting for the account to come back"
            Set m.WaitAccountOrders(strOrderKey) = turnkeyMessage
            
        ' Order has been sent to Turnkey, but has not returned yet...
        ElseIf m.OrdersSent.Exists(strOrderKey) Then
            ' Order is waiting for association...
            If m.WaitOrderOrders.Exists(strOrderKey) Then
                g.Cattle.DumpDebug "Order '" & strOrderKey & "' overwritten in list waiting for order to come back"
                Set m.WaitOrderOrders(strOrderKey) = turnkeyMessage
                
            ' Order needs to wait for association...
            Else
                g.Cattle.DumpDebug "Order '" & strOrderKey & "' added to list waiting for order to come back"
                m.WaitOrderOrders.Add turnkeyMessage, strOrderKey
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.HandleOrderUpdate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleFillUpdate
'' Description: Handle a fill update from the broker
'' Inputs:      Turnkey Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleFillUpdate(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim oldMessage As cBrokerMessage    ' Old Turnkey message
    Dim strOrderKey As String           ' Order key
    Dim strFillKey As String            ' Fill key
    Dim lAssociatedQuantity As Long     ' Associated quantity for the fill
    
    strOrderKey = OrderKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerOrderID"))
    strFillKey = FillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"))
    lAssociatedQuantity = AssociatedQuantityForFill(turnkeyMessage)
    
    If m.Accounts.Exists(turnkeyMessage("BrokerAccountID")) Then
        If FillChanged(turnkeyMessage) Then
            turnkeyMessage.Add "FeedYardID", Str(SelectedFeedYard)
            UpdateFill turnkeyMessage, "fill has changed"
        End If
    End If
    
    ' Fill is already associated...
    If lAssociatedQuantity > 0 Then
        If lAssociatedQuantity > turnkeyMessage("Quantity") Then
            g.Cattle.DumpDebug "Fill '" & strFillKey & "' is now overbooked (Quantity = " & turnkeyMessage("Quantity") & "; Associated = " & Str(lAssociatedQuantity) & ")"
        End If
        
    ' Order for the fill is already associated...
    ElseIf m.AssociatedOrders.Exists(strOrderKey) Then
        Set oldMessage = m.AssociatedOrders(strOrderKey)
        
        turnkeyMessage.Add "AssociatedQuantity", turnkeyMessage("Quantity")
        turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
        
        If m.Fills.Exists(strFillKey) Then
            Set oldMessage = m.Fills(strFillKey)
            
            turnkeyMessage.Add "FillID", oldMessage("ID")
            turnkeyMessage("ID") = ""
            
            UpdateFillAssociation turnkeyMessage, "fill received for order already associated"
        Else
            strFillKey = AssociatedFillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"), turnkeyMessage("FeedYardLotID"))
            
            g.Cattle.DumpDebug "Associated Fill '" & strFillKey & "' added to list waiting for fill to come back"
            m.WaitFillFills.Add turnkeyMessage, strFillKey
        End If
        
    ' Order has been sent to Turnkey, but has not returned yet...
    ElseIf m.OrdersSent.Exists(strOrderKey) Then
        ' DAJ 09/13/2012: Don't wait for the order confirmation to come back from the Genesis Turnkey
        ' servers, just grab the FeedYardLotID out of the waiting order and go...
        
        Set oldMessage = m.OrdersSent(strOrderKey)
        
        turnkeyMessage.Add "AssociatedQuantity", turnkeyMessage("Quantity")
        turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
        
        If m.Fills.Exists(strFillKey) Then
            Set oldMessage = m.Fills(strFillKey)
            
            turnkeyMessage.Add "FillID", oldMessage("ID")
            turnkeyMessage("ID") = ""
            
            UpdateFillAssociation turnkeyMessage, "fill received for order already associated"
        Else
            strFillKey = AssociatedFillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"), turnkeyMessage("FeedYardLotID"))
            
            g.Cattle.DumpDebug "Associated Fill '" & strFillKey & "' added to list waiting for fill to come back"
            m.WaitFillFills.Add turnkeyMessage, strFillKey
        End If
        
    ' Order is waiting for account to come back from Turnkey...
    ElseIf m.WaitAccountOrders.Exists(strOrderKey) Then
        Set oldMessage = m.WaitAccountOrders(strOrderKey)
        
        turnkeyMessage.Add "FeedYardLotID", oldMessage("FeedYardLotID")
        strFillKey = AssociatedFillKey(turnkeyMessage("Broker"), turnkeyMessage("BrokerFillID"), turnkeyMessage("FeedYardLotID"))
        
        If m.WaitAccountFills.Exists(strFillKey) Then
            g.Cattle.DumpDebug "Fill '" & strFillKey & "' overwritten in list waiting for account to come back"
            Set m.WaitAccountFills(strFillKey) = turnkeyMessage
        Else
            g.Cattle.DumpDebug "Fill '" & strFillKey & "' added to list waiting for account to come back"
            m.WaitAccountFills.Add turnkeyMessage, strFillKey
        End If
    
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.HandleFillUpdate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortOnCol
'' Description: Sort the grid for the given column number and order
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortOnCol(Optional ByVal lCol As Long = kNullData, Optional ByVal nOrder As SortSettings = kNullData)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim bAscending As Boolean           ' Do we want to sort ascending?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strValue As String              ' Value for the column being sorted
    Dim LotColumn As cLotColumn         ' Lot Column object
    Dim strFormat As String             ' Format for the column

    If lCol = kNullData Then
        If m.lSortedCol = -1& Then
            lCol = LotCol(eGDLotCol_LotNumber)
        Else
            lCol = m.lSortedCol
        End If
    End If
    
    strFormat = ""
    Set LotColumn = LotColumnForCol(lCol)
    If Not LotColumn Is Nothing Then
        strFormat = LotColumn.Format
    End If
    
    If nOrder = kNullData Then
        If m.nSortedDir = -1& Then
            nOrder = flexSortStringAscending
        Else
            nOrder = m.nSortedDir
        End If
    End If

    If (nOrder = flexSortGenericAscending) Or (nOrder = flexSortNumericAscending) Or (nOrder = flexSortStringAscending) Or (nOrder = flexSortStringNoCaseAscending) Then
        bAscending = True
    Else
        bAscending = False
    End If

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        strValue = ""
        For lIndex = .FixedRows To .Rows - 2
            If .RowOutlineLevel(lIndex) = 1 Then
                strValue = g.Cattle.GridValue(fgLots, lIndex, lCol, LotColumn)
                
                Select Case UCase(strFormat)
                    Case "NUMBER", "DATE"
                        strValue = Pad(Format(Val(strValue), "#.0000000"), 50, "R")
                    Case "CURRENCY"
                        strValue = Pad(Format(Val(strValue), "#.000"), 50, "R")
                    Case Else
                        strValue = Pad(strValue, 50, "L")
                End Select
            End If
            
            If bAscending Then
                .TextMatrix(lIndex, LotCol(eGDLotCol_SortKey)) = strValue & "_" & .TextMatrix(lIndex, LotCol(eGDLotCol_AscSortKey))
            Else
                .TextMatrix(lIndex, LotCol(eGDLotCol_SortKey)) = strValue & "_" & .TextMatrix(lIndex, LotCol(eGDLotCol_DescSortKey))
            End If
            
            .RowOutlineLevel(lIndex) = 0
            .IsSubtotal(lIndex) = False
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Select .FixedRows, LotCol(eGDLotCol_SortKey), .Rows - 1, LotCol(eGDLotCol_SortKey)
        End If
        
        If bAscending Then
            '.Sort = flexSortStringAscending
            m.nSortedDir = flexSortStringAscending
        Else
            '.Sort = flexSortStringDescending
            m.nSortedDir = flexSortStringDescending
        End If
        .Sort = flexSortCustom
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0
        End If
        
        For lIndex = .FixedRows To .Rows - 2
            .RowOutlineLevel(lIndex) = CLng(Val(.TextMatrix(lIndex, LotCol(eGDLotCol_OutlineLevel))))
            .IsSubtotal(lIndex) = True
        Next lIndex
        
        If m.lSortedCol > -1& Then
            .Cell(flexcpPicture, 0, m.lSortedCol) = Nothing
            .Cell(flexcpBackColor, 0, m.lSortedCol) = .BackColorFixed
        End If
        'If bAscending Then
        '    .Cell(flexcpPicture, 0, lCol) = Picture16("kSortedUpRight")
        'Else
        '    .Cell(flexcpPicture, 0, lCol) = Picture16("kSortedDownRight")
        'End If
        .Cell(flexcpBackColor, 0, lCol) = &H80C0FF    ' 33023 ' Orange
        
        '.Cell(flexcpPictureAlignment, 0, lCol) = flexPicAlignRightBottom  '= flexPicAlignRightTop
        '.PicturesOver = True
        
        m.lSortedCol = lCol
        
        ColorGridRows
        
        AutoSizeGrid
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SortOnCol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorGridRows
'' Description: Set the background color of the grid rows as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorGridRows()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim bAlt As Boolean                 ' Alternate color?
    Dim lRow As Long                    ' Row in the grid
    Dim lNextSibling As Long            ' Next sibling row in the grid
    Dim lLastDescendant As Long         ' Last descendant for this row
    Dim bClosed As Boolean              ' Is the lot closed?
    
    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            bAlt = False
            Do While lRow <> -1&
                lNextSibling = .GetNodeRow(lRow, flexNTNextSibling)
                
                If .RowHidden(lRow) = False Then
                    If lNextSibling = -1& Then
                        lLastDescendant = .Rows - 1
                    Else
                        lLastDescendant = lNextSibling - 1
                    End If
                    
                    bClosed = LotIsClosed(lRow)
                    
                    For lIndex = lRow To lLastDescendant
                        If bClosed Then
                            .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = vbCyan
                        ElseIf bAlt Then
                            .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = ALT_GRID_ROW_COLOR
                        Else
                            .Cell(flexcpBackColor, lIndex, 0, lIndex, .Cols - 1) = m.nBackColor
                        End If
                    Next lIndex
                    
                    If bClosed = False Then
                        bAlt = Not bAlt
                    End If
                End If
                
                lRow = lNextSibling
            Loop
        End If
                
        .Redraw = nRedraw
    End With


ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ColorGridRows"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateLotColumns
'' Description: Create lot column objects for internal columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateLotColumns()
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column information
    
    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_OpenEquity + 10000
        .ColumnHeader = "Open Equity"
        .TooltipText = "Open Equity"
        .Format = "Currency"
        .DisplayFormat = "$#,##0.00"
        .KeyValueField = "OpenEquity"
        .Total = True
        
        .AlwaysHidden = False
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_ClosedProfit + 10000
        .ColumnHeader = "Closed Profit"
        .TooltipText = "Closed Profit"
        .Format = "Currency"
        .DisplayFormat = "$#,##0.00"
        .KeyValueField = "ClosedProfit"
        .Total = True
        
        .AlwaysHidden = False
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_AscSortKey + 10000
        .ColumnHeader = "Asc Sort Key"
        .TooltipText = "Ascending Sort Key"
        .Format = "Hidden"
        .KeyValueField = "AscSortKey"
        .Total = False
        
        .AlwaysHidden = True
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_DescSortKey + 10000
        .ColumnHeader = "Desc Sort Key"
        .TooltipText = "Descending Sort Key"
        .Format = "Hidden"
        .KeyValueField = "DescSortKey"
        .Total = False
        
        .AlwaysHidden = True
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_SortKey + 10000
        .ColumnHeader = "Sort Key"
        .TooltipText = "Sort Key"
        .Format = "Hidden"
        .KeyValueField = "SortKey"
        .Total = False
        
        .AlwaysHidden = True
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_OutlineLevel + 10000
        .ColumnHeader = "Outline Level"
        .TooltipText = "Outline Level"
        .Format = "Hidden"
        .KeyValueField = "OutlineLevel"
        .Total = False
        
        .AlwaysHidden = True
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

    Set LotColumn = New cLotColumn
    With LotColumn
        .ID = eGDLotCol_RowType + 10000
        .ColumnHeader = "Row Type"
        .TooltipText = "Row Type"
        .Format = "Hidden"
        .KeyValueField = "RowType"
        .Total = False
        
        .AlwaysHidden = True
        .FeedyardHidden = False
        .UserHidden = False
        .AllowUserMove = False
    End With
    m.KeyValueToIndex.Add m.LotColumns.Add(LotColumn, Str(LotColumn.ID)), LotColumn.KeyValueField

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CreateLotColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupGridColumns
'' Description: Setup the grid columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupGridColumns()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column information

    CreateLotColumns

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Cols = NumberOfCols
        .FixedCols = 0
        
        m.KeyValueToCol.Clear
        
        For lIndex = 1 To m.LotColumns.Count
            Set LotColumn = m.LotColumns(lIndex)
            
            LotColumn.Position = lIndex - 1
            
            m.KeyValueToCol.Add lIndex - 1, LotColumn.KeyValueField
            
            .ColData(lIndex - 1) = Str(LotColumn.ID)
            .TextMatrix(0, lIndex - 1) = LotColumn.ColumnHeader
        
            Select Case UCase(LotColumn.Format)
                Case "DATE"
                    .ColAlignment(lIndex - 1) = flexAlignCenterCenter
                    '.ColFormat(lIndex - 1) = LotColumn.DisplayFormat ' DateFormat("Format", MM_DD_YYYY)
                Case "CURRENCY", "NUMBER"
                    .ColAlignment(lIndex - 1) = flexAlignRightCenter
                    '.ColFormat(lIndex - 1) = LotColumn.DisplayFormat ' "$#,##0.00"
                Case "BOOLEAN"
                    .ColAlignment(lIndex - 1) = flexAlignCenterCenter
                    '.ColDataType(lIndex - 1) = flexDTBoolean
            End Select
            Select Case UCase(LotColumn.KeyValueField)
                Case "NUMBER"
                    m.alLotCol(eGDLotCol_LotNumber) = lIndex - 1
                    .ColAlignment(lIndex - 1) = flexAlignLeftCenter
                    LotColumn.AllowUserMove = False
                Case "STATUS"
                    m.alLotCol(eGDLotCol_LotStatus) = lIndex - 1
                Case "OPENEQUITY"
                    m.alLotCol(eGDLotCol_OpenEquity) = lIndex - 1
                Case "CLOSEDPROFIT"
                    m.alLotCol(eGDLotCol_ClosedProfit) = lIndex - 1
                Case "ASCSORTKEY"
                    m.alLotCol(eGDLotCol_AscSortKey) = lIndex - 1
                Case "DESCSORTKEY"
                    m.alLotCol(eGDLotCol_DescSortKey) = lIndex - 1
                Case "SORTKEY"
                    m.alLotCol(eGDLotCol_SortKey) = lIndex - 1
                Case "OUTLINELEVEL"
                    m.alLotCol(eGDLotCol_OutlineLevel) = lIndex - 1
                Case "ISLOTCLOSED"
                    m.alLotCol(eGDLotCol_IsClosed) = lIndex - 1
                Case "ROWTYPE"
                    m.alLotCol(eGDLotCol_RowType) = lIndex - 1
            End Select
                        
            Set m.LotColumns(lIndex) = LotColumn
            .ColHidden(lIndex - 1) = LotColumn.AlwaysHidden
        Next lIndex
        
        If LotCol(eGDLotCol_LotNumber) > 0 Then
            .ColPosition(LotCol(eGDLotCol_LotNumber)) = 0
            m.alLotCol(eGDLotCol_LotNumber) = 0
            RebuildKeyValueToCol
        End If
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        AutoSizeGrid
        .RowHidden(0) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SetupGridColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowGridColumns
'' Description: Show/Hide and move the grid columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowGridColumns()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column information
    Dim GridColumns As cGridColumns     ' Collection of grid column information
    Dim GridColumn As cGridColumn       ' Grid column information
    Dim strDisplay As String            ' Display string out of the INI file
    Dim lCol As Long                    ' Column in the grid

    Set GridColumns = New cGridColumns
    If TemplateNumber = 0& Then
        strDisplay = DefaultDisplayString
    Else
        strDisplay = TemplateDisplay(m.lTemplateNumber)
    End If

    SetFormCaption
    m.bLoadingColumns = True

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
               
        If Len(strDisplay) > 0 Then
            GridColumns.FromString strDisplay
            For lIndex = 1 To GridColumns.Count
                Set GridColumn = GridColumns.Item(lIndex)
                
                If m.KeyValueToCol.Exists(GridColumn.Name) Then
                    lCol = m.KeyValueToCol(GridColumn.Name)
                    
                    GridColumn.Position = VerifyMoveColumn(lCol, GridColumn.Position)
                    .ColPosition(lCol) = GridColumn.Position
                    RebuildKeyValueToCol
                    
                    lCol = GridColumn.Position
                    .ColWidth(lCol) = GridColumn.Width
                    
                    Set LotColumn = LotColumnForCol(lCol)
                    If Not LotColumn Is Nothing Then
                        LotColumn.Position = lCol
                        LotColumn.Width = GridColumn.Width
                        LotColumn.UserHidden = Not GridColumn.Visible
                        Set m.LotColumns(Str(LotColumn.ID)) = LotColumn
                        
                        m.KeyValueToCol(LotColumn.KeyValueField) = lCol
                        Select Case UCase(LotColumn.KeyValueField)
                            Case "NUMBER"
                                m.alLotCol(eGDLotCol_LotNumber) = lCol
                            Case "STATUS"
                                m.alLotCol(eGDLotCol_LotStatus) = lCol
                            Case "OPENEQUITY"
                                m.alLotCol(eGDLotCol_OpenEquity) = lCol
                            Case "CLOSEDPROFIT"
                                m.alLotCol(eGDLotCol_ClosedProfit) = lCol
                            Case "ASCSORTKEY"
                                m.alLotCol(eGDLotCol_AscSortKey) = lCol
                            Case "DESCSORTKEY"
                                m.alLotCol(eGDLotCol_DescSortKey) = lCol
                            Case "SORTKEY"
                                m.alLotCol(eGDLotCol_SortKey) = lCol
                            Case "OUTLINELEVEL"
                                m.alLotCol(eGDLotCol_OutlineLevel) = lCol
                            Case "ISLOTCLOSED"
                                m.alLotCol(eGDLotCol_IsClosed) = lCol
                            Case "ROWTYPE"
                                m.alLotCol(eGDLotCol_RowType) = lCol
                        End Select
                    End If
                End If
            Next lIndex
        End If
        
        ' Have to do this as a second pass because there may be new columns that weren't in the
        ' persisted information...
        For lIndex = 0 To .Cols - 1
            Set LotColumn = LotColumnForCol(lIndex)
            If Not LotColumn Is Nothing Then
                .ColHidden(lIndex) = LotColumn.AlwaysHidden Or LotColumn.FeedyardHidden Or LotColumn.UserHidden
            End If
        Next lIndex
        
        .RowHidden(0) = False
        AutoSizeGrid
        
        .Redraw = nRedraw
    End With
    
    m.bLoadingColumns = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ShowGridColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveGridColumns
'' Description: Save the grid column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveGridColumns()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column information
    Dim GridColumns As cGridColumns     ' Collection of grid column information
    Dim GridColumn As cGridColumn       ' Grid column information
    Dim strDisplay As String            ' Display string
    
    Set GridColumns = New cGridColumns
    With fgLots
        For lIndex = 0 To .Cols - 1
            Set LotColumn = LotColumnForCol(lIndex)
            If Not LotColumn Is Nothing Then
                If (LotColumn.AllowUserMove = True) Or (UCase(LotColumn.KeyValueField) = "OPENEQUITY") Or (UCase(LotColumn.KeyValueField) = "CLOSEDPROFIT") Then
                    Set GridColumn = New cGridColumn
                    GridColumn.Name = LotColumn.KeyValueField
                    GridColumn.Position = lIndex
                    GridColumn.Width = 0&
                    GridColumn.Visible = Not LotColumn.UserHidden
                    
                    GridColumns.Add GridColumn
                End If
            End If
        Next lIndex
    End With
    
    strDisplay = GridColumns.ToString(False)
    If Len(Parse(strDisplay, "|", 2)) > 0 Then
        TemplateDisplay(m.lTemplateNumber) = strDisplay
    End If

    Dirty = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SaveGridColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotColumnForCol
'' Description: Get the Lot Column object for the given grid column
'' Inputs:      Grid Column
'' Returns:     Lot Column object (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LotColumnForCol(ByVal lCol As Long) As cLotColumn
On Error GoTo ErrSection:

    Dim ReturnLotColumn As cLotColumn   ' Lot column to return
    Dim strLotColumnID As String        ' Lot column ID
        
    Set ReturnLotColumn = Nothing
    
    If ValidGridCol(fgLots, lCol) Then
        strLotColumnID = fgLots.ColData(lCol)
        If m.LotColumns.Exists(strLotColumnID) Then
            Set ReturnLotColumn = m.LotColumns(strLotColumnID)
        End If
    End If
    
    Set LotColumnForCol = ReturnLotColumn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.LotColumnForCol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RebuildKeyValueToCol
'' Description: Rebuild the key-value to column index
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RebuildKeyValueToCol()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    
    With fgLots
        For lIndex = 0 To .Cols - 1
            Set LotColumn = LotColumnForCol(lIndex)
            m.KeyValueToCol(LotColumn.KeyValueField) = lIndex
            
            LotColumn.Position = lIndex
            Set m.LotColumns(Str(LotColumn.ID)) = LotColumn
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.RebuildKeyValueToCol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTemplatesMenu
'' Description: Load the templates menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTemplatesMenu()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumTemplates As Long           ' Number of saved templates
    Dim strTemplateName As String       ' Name of the saved template
    Dim strToolID As String             ' Tool ID

    With tbToolbar.Tools("ID_Templates").Menu
        For lIndex = .Tools.Count To 7 Step -1
            .Tools.Remove lIndex
        Next lIndex
        
        If TemplateNumber = 0 Then
            tbToolbar.Tools("ID_DefaultTemplate").ForeColor = vbBlue
        Else
            tbToolbar.Tools("ID_DefaultTemplate").ForeColor = vbBlack
        End If
        
        lNumTemplates = NumTemplates
        For lIndex = 1 To lNumTemplates
            strTemplateName = TemplateName(lIndex)
            If Len(strTemplateName) > 0 Then
                strToolID = "ID_Template_" & Str(lIndex)
                .Tools.Add strToolID, ssTypeButton
                
                .Tools(lIndex + 6).Name = strTemplateName
                If lIndex = TemplateNumber Then
                    .Tools(lIndex + 6).ForeColor = vbBlue
                Else
                    .Tools(lIndex + 6).ForeColor = vbBlack
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.LoadTemplatesMenu"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetFormCaption
'' Description: Set the form caption
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormCaption()
On Error GoTo ErrSection:

    Dim strMode As String               ' Connection mode
    
    strMode = ""
    If g.Cattle.Mode = "D" Then
        strMode = "Demo "
    ElseIf g.Cattle.Mode = "L" Then
        strMode = "Live "
    ElseIf g.Cattle.Mode = "T" Then
        strMode = "Test "
    End If

    If TemplateNumber = 0& Then
        Caption = strMode & g.Cattle.ProductName & " Feed Lot Information"
    Else
        Caption = strMode & g.Cattle.ProductName & " Feed Lot Information - [" & TemplateName(m.lTemplateNumber) & "]"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SetFormCaption"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameTemplate
'' Description: Rename the current display template
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RenameTemplate()
On Error GoTo ErrSection:

    Dim strOldName As String            ' Old name of the display template
    Dim strNewName As String            ' New name of the display template
    
    If m.lTemplateNumber > 0& Then
        strOldName = TemplateName(m.lTemplateNumber)
        
        strNewName = AskForNewTemplateName("Rename Template...", strOldName)
        If (Len(strNewName) > 0) And (strNewName <> strOldName) Then
            TemplateName(m.lTemplateNumber) = strNewName
            SetFormCaption
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.RenameTemplate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveTemplate
'' Description: Save the current display template
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveTemplate()
On Error GoTo ErrSection:

    If TemplateNumber = 0& Then
        SaveTemplateAs
    Else
        SaveGridColumns
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SaveTemplate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveTemplateAs
'' Description: Save the current display template as a new name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveTemplateAs()
On Error GoTo ErrSection:

    Dim strOldName As String            ' Old name of the display template
    Dim strNewName As String            ' New name of the display template
    Dim lNumTemplates As Long           ' Number of saved templates
    
    lNumTemplates = NumTemplates
    strOldName = TemplateName(m.lTemplateNumber)
    
    strNewName = AskForNewTemplateName("Save Template As...", strOldName)
    If (Len(strNewName) > 0) And (strNewName <> strOldName) Then
        lNumTemplates = lNumTemplates + 1&
        NumTemplates = lNumTemplates
        TemplateNumber = lNumTemplates
        TemplateName(m.lTemplateNumber) = strNewName
        
        SaveGridColumns
        
        SetFormCaption
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SaveTemplateAs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskToSave
'' Description: Ask the user if they want to save display changes
'' Inputs:      None
'' Returns:     True if not cancel, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim strResponse As String           ' Response from the question
    
    bReturn = True
    If Dirty Then
        strResponse = InfBox("Do you want to save your display changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                bReturn = False
            Case "Y"
                SaveTemplate
        End Select
    End If
    
    AskToSave = bReturn
        
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmLots.AskToSave"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultDisplayString
'' Description: Build the default display string
'' Inputs:      None
'' Returns:     Default display string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DefaultDisplayString() As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column information
    Dim GridColumns As cGridColumns     ' Collection of grid column information
    Dim GridColumn As cGridColumn       ' Grid column information
    
    Set GridColumns = New cGridColumns
    For lIndex = 1 To m.LotColumns.Count
        Set LotColumn = m.LotColumns(lIndex)
        If Not LotColumn Is Nothing Then
            If LotColumn.AllowUserMove Then
                Set GridColumn = New cGridColumn
                GridColumn.Name = LotColumn.KeyValueField
                GridColumn.Position = lIndex - 1
                GridColumn.Width = 0&
                GridColumn.Visible = True
                
                GridColumns.Add GridColumn
            End If
        End If
    Next lIndex
    
    DefaultDisplayString = GridColumns.ToString(False)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.DefaultDisplayString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveTemplates
'' Description: Allow the user to remove saved templates
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveTemplates()
On Error GoTo ErrSection:

    Dim astrTemplates As cGdArray       ' Array of saved templates to send to delete form
    Dim lIndex As Long                  ' Index into a for loop
    Dim lNumTemplates As Long           ' Number of saved templates
    Dim astrToDelete As cGdArray        ' Array of indexes to delete
    Dim lTemplateNumber As Long         ' Template number
    Dim lToDelete As Long               ' Index to delete
    Dim lTemplateIndex As Long          ' Template index
    
    Set astrTemplates = New cGdArray
    lNumTemplates = NumTemplates
    For lIndex = 1 To lNumTemplates
        astrTemplates.Add TemplateName(lIndex) & vbTab & Str(lIndex) & vbTab & TemplateDisplay(lIndex)
    Next lIndex
    
    Set astrToDelete = frmDelete.ShowMe(astrTemplates)
    If Not astrToDelete Is Nothing Then
        For lIndex = astrToDelete.Size - 1 To 0 Step -1
            lToDelete = CLng(Val(astrToDelete(lIndex)))
            astrTemplates.Remove lToDelete - 1
            If lToDelete = TemplateNumber Then
                lTemplateNumber = 0&
            End If
        Next lIndex
        
        For lIndex = 0 To astrTemplates.Size - 1
            lTemplateIndex = lIndex + 1
            TemplateName(lTemplateIndex) = Parse(astrTemplates(lIndex), vbTab, 1)
            TemplateDisplay(lTemplateIndex) = Parse(astrTemplates(lIndex), vbTab, 3)
            
            If CLng(Val(Parse(astrTemplates(lIndex), vbTab, 2))) = TemplateNumber Then
                lTemplateNumber = lTemplateIndex
            End If
        Next lIndex
        For lIndex = astrTemplates.Size + 1 To lNumTemplates
            TemplateName(lIndex) = ""
            TemplateDisplay(lIndex) = ""
        Next lIndex
        NumTemplates = astrTemplates.Size
        
        If lTemplateNumber <> TemplateNumber Then
            TemplateNumber = lTemplateNumber
            ShowGridColumns
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.RemoveTemplates"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TemplateNameIsUnique
'' Description: Determine if the given name is unique
'' Inputs:      New Name, Show Message?
'' Returns:     True if unique, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TemplateNameIsUnique(ByVal strNewName As String, Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = True
    For lIndex = 1 To NumTemplates
        If TemplateName(lIndex) = Trim(strNewName) Then
            bReturn = False
            
            If bShowMessage Then
                InfBox "Display template name must be unique", "!", , "Display Template"
            End If
            
            Exit For
        End If
    Next lIndex
    
    TemplateNameIsUnique = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.TemplateNameIsUnique"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TemplateNameIsValid
'' Description: Determine if the given name is valid
'' Inputs:      New Name, Show Message?
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TemplateNameIsValid(ByVal strNewName As String, Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If InStr(strNewName, "=") <> 0 Then
        bReturn = False
        
        If bShowMessage Then
            InfBox "Display template name cannot contain an equals sign (=)", "!", , "Display Template"
        End If
    End If
    
    TemplateNameIsValid = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.TemplateNameIsValid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskForNewTemplateName
'' Description: Ask for a new template name
'' Inputs:      Caption, Default
'' Returns:     New Name ( Blank if cancelled )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AskForNewTemplateName(ByVal strCaption As String, ByVal strDefault As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    Do
        strReturn = InfBox("Please enter a new name for this display template", , "+OK|-Cancel", strCaption, , , , , , "string", strDefault)
    Loop While (Len(strReturn) > 0) And ((Not TemplateNameIsUnique(strReturn)) Or (Not TemplateNameIsValid(strReturn)))

    AskForNewTemplateName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.AskForNewTemplateName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociateAccounts
'' Description: Allow the user to associate accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssociateAccounts()
On Error GoTo ErrSection:

    Dim BrokerAccounts As cGdTree       ' Broker list of accounts
    Dim AssociatedAccounts As cGdTree   ' Associated accounts
    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cBrokerMessage       ' Account
    Dim strKey As String                ' Key into the collection
    Dim lSelectedFeedYard As Long       ' Selected feedyard ID
    
    lSelectedFeedYard = SelectedFeedYard
    If lSelectedFeedYard > -1& Then
        Set BrokerAccounts = g.AppBridge.GetBrokerAccounts.MakeCopy
        Set AssociatedAccounts = m.Accounts.MakeCopy
            
        If frmCattleSelect.ShowMeAccounts(BrokerAccounts, AssociatedAccounts, cboFeedYards.Text) Then
            For lIndex = 1 To AssociatedAccounts.Count
                Set Account = AssociatedAccounts(lIndex)
                
                If Len(Account("ID")) = 0 Then
                    strKey = AccountKey(Account("Broker"), Account("Number"))
                    Account.Add "FeedYardID", Str(lSelectedFeedYard)
                    g.Cattle.UpdateAccount Account
                    
                    m.WaitAccounts.Add Account, strKey
                End If
            Next lIndex
            
            For lIndex = 1 To m.Accounts.Count
                Set Account = m.Accounts(lIndex)
                strKey = AccountKey(Account("Broker"), Account("Number"))
                
                If AssociatedAccounts.Exists(strKey) = False Then
                    g.Cattle.RemoveAccount Account
                End If
            Next lIndex
        End If
    Else
        InfBox "Please select a feedyard before associating accounts", "!", , g.Cattle.ProductName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.AssociateAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllBrokerOrders
'' Description: Get the broker orders for the accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllBrokerOrders()
On Error GoTo ErrSection:

    Dim lAccount As Long                ' Index into a for loop
    Dim lOrder As Long                  ' Index into a for loop
    Dim Account As cBrokerMessage       ' Account
    Dim Orders As cGdTree               ' Collection of orders
    Dim Order As cBrokerMessage         ' Order message
    Dim strKey As String                ' Key into the orders collection
    Dim AssocOrder As cBrokerMessage    ' Associated order
    
    For lAccount = 1 To m.Accounts.Count
        Set Account = m.Accounts(lAccount)
        
        Set Orders = g.AppBridge.GetBrokerOrdersForAccount(Account)
        
        For lOrder = 1 To Orders.Count
            Set Order = Orders(lOrder)
            
            strKey = OrderKey(Order("Broker"), Order("BrokerOrderID"))
            If m.AssociatedOrders.Exists(strKey) Then
                Set AssocOrder = m.AssociatedOrders(strKey)
                If OrderChanged(Order) Then
                    Order.Add "ID", AssocOrder("ID")
                    Order.Add "FeedYardID", Str(SelectedFeedYard)
                    Order.Add "FeedYardLotID", AssocOrder("FeedYardLotID")
                    
                    g.Cattle.UpdateOrder Order
                End If
            End If
        Next lOrder
    Next lAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetAllBrokerOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderChanged
'' Description: Determine if any of the order information changed
'' Inputs:      Order Message
'' Returns:     True if changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderChanged(ByVal orderMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strKey As String                ' Key into the orders collection
    Dim astrFields As cGdArray          ' List of fields to check
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cBrokerMessage         ' Order from the collection
    
    strKey = OrderKey(orderMessage("Broker"), orderMessage("BrokerOrderID"))
    
    If m.AssociatedOrders.Exists(strKey) Then
        Set Order = m.AssociatedOrders(strKey)
        
        Set astrFields = New cGdArray
        astrFields.Create eGDARRAY_Strings, 14
        
        astrFields(0) = "BrokerOrderID"
        astrFields(1) = "GenesisOrderID"
        astrFields(2) = "Quantity"
        astrFields(3) = "Type"
        astrFields(4) = "LimitPrice"
        astrFields(5) = "StopPrice"
        astrFields(6) = "TimeInForce"
        astrFields(7) = "ExpirationDate"
        astrFields(8) = "Status"
        astrFields(9) = "StatusDate"
        astrFields(10) = "IsWorking"
        astrFields(11) = "Side"
        astrFields(12) = "Leg1"
        astrFields(13) = "Symbol"
        
        bReturn = False
        For lIndex = 0 To astrFields.Size - 1
            If Order(astrFields(lIndex)) <> orderMessage(astrFields(lIndex)) Then
                bReturn = True
                Exit For
            End If
        Next lIndex
    Else
        bReturn = True
    End If
    
    OrderChanged = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.OrderChanged"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAllBrokerFills
'' Description: Get the broker fills for the accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAllBrokerFills()
On Error GoTo ErrSection:

    Dim lAccount As Long                ' Index into a for loop
    Dim lFill As Long                   ' Index into a for loop
    Dim Account As cBrokerMessage       ' Account
    Dim Fills As cGdTree                ' Collection of fills
    Dim Fill As cBrokerMessage          ' Fill message
    
    For lAccount = 1 To m.Accounts.Count
        Set Account = m.Accounts(lAccount)
        
        Set Fills = g.AppBridge.GetBrokerFillsForAccount(Account)
        
        For lFill = 1 To Fills.Count
            Set Fill = Fills(lFill)
            
            If FillChanged(Fill) Then
                Fill.Add "FeedYardID", Str(SelectedFeedYard)
                g.Cattle.UpdateFill Fill
            End If
        Next lFill
    Next lAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.GetAllBrokerFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillChanged
'' Description: Determine if any of the fill information changed
'' Inputs:      Fill Message
'' Returns:     True if changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillChanged(ByVal FillMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strKey As String                ' Key into the fills collection
    Dim astrFields As cGdArray          ' List of fields to check
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cBrokerMessage          ' Fill from the collection
    
    strKey = FillKey(FillMessage("Broker"), FillMessage("BrokerFillID"))
    
    If m.Fills.Exists(strKey) Then
        Set Fill = m.Fills(strKey)
        
        Set astrFields = New cGdArray
        astrFields.Create eGDARRAY_Strings, 8
        
        astrFields(0) = "BrokerOrderID"
        astrFields(1) = "BrokerFillID"
        astrFields(2) = "FillTime"
        astrFields(3) = "Symbol"
        astrFields(4) = "IsBuy"
        astrFields(5) = "Quantity"
        astrFields(6) = "Price"
        astrFields(7) = "Commission"
        
        bReturn = False
        For lIndex = 0 To astrFields.Size - 1
            If Fill(astrFields(lIndex)) <> FillMessage(astrFields(lIndex)) Then
                bReturn = True
                Exit For
            End If
        Next lIndex
    Else
        bReturn = True
    End If
    
    FillChanged = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.FillChanged"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociatedQuantityForFill
'' Description: Determine the associated quantity for the given fill
'' Inputs:      Fill
'' Returns:     Associated Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AssociatedQuantityForFill(ByVal Fill As cBrokerMessage) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim AssociatedFill As cBrokerMessage ' Associated fill
    
    lReturn = 0&
    For lIndex = 1 To m.AssociatedFills.Count
        Set AssociatedFill = m.AssociatedFills(lIndex)
        If AssociatedFill("BrokerFillID") = Fill("BrokerFillID") Then
            lReturn = lReturn + CLng(Val(AssociatedFill("AssociatedQuantity")))
        End If
    Next lIndex
    
    AssociatedQuantityForFill = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.AssociatedQuantityForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTurnkey
'' Description: Call for a refresh from the Turnkey servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshTurnkey()
On Error GoTo ErrSection:

    m.bHasDetails = False
    m.bHasIngredients = False

    'GetAllLotColumns
    GetFeedyards

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.RefreshTurnkey"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ItemSent
'' Description: An item has been sent to the Genesis Turnkey servers for
''              association
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ItemSent()
On Error GoTo ErrSection:

    m.dLastItemSent = gdTickCount
    tmrItemsSent.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ItemSent"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ItemReceived
'' Description: An item has been received from the Genesis Turnkey servers to
''              confirm association
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ItemReceived()
On Error GoTo ErrSection:

    If (m.AssociatedOrdersSent.Count = 0) And (m.AssociatedFillsSent.Count = 0) Then
        InfBox ""
        m.dLastItemSent = kNullData
        tmrItemsSent.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ItemReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpandAll
'' Description: Expand all of the nodes in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExpandAll(Optional lLevel As Long = -1)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgLots
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If lLevel = -1& Then
                m.lExpandLevel = SetGridLevel(fgLots, m.lExpandLevel + 1)
            Else
                m.lExpandLevel = SetGridLevel(fgLots, lLevel)
            End If
            
            FilterGrid
            
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ExpandAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CollapseAll
'' Description: Collapse all of the nodes in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CollapseAll(Optional lLevel As Long = -1)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgLots
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If lLevel = -1& Then
                m.lExpandLevel = SetGridLevel(fgLots, m.lExpandLevel - 1)
            Else
                m.lExpandLevel = SetGridLevel(fgLots, lLevel)
            End If
            
            FilterGrid
            
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.CollapseAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportToCsv
'' Description: Export the grid to a CSV file
'' Inputs:      Include Hidden Rows?, Include Hidden Columns?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportToCsv(Optional ByVal bIncludeHiddenRows As Boolean = True, Optional ByVal bIncludeHiddenColumns As Boolean = False)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Array to dump to the file
    Dim astrLine As cGdArray            ' Line to add to the array
    Dim lRow As Long                    ' Row in the grid
    Dim lCol As Long                    ' Column in the grid
    Dim strText As String               ' Text from the cell
    Dim strDefaultPath As String        ' Default path for the CSV file
    Dim strFileName As String           ' Filename for the CSV file
    Dim strLastText As String           ' Last text printed
    
    strDefaultPath = GetIniFileProperty("CsvPath", AddSlash(g.strAppPath), "Turnkey", g.strIniFile)
    strFileName = CommonDialogFile(g.frmMain.CommonDialog1, True, "CSV Files (*.csv)|*.csv", strDefaultPath, g.Cattle.ProductName)
    If Len(strFileName) > 0 Then
        SetIniFileProperty "CsvPath", FilePath(strFileName), "Turnkey", g.strIniFile
        
        Set astrFile = New cGdArray
        astrFile.Create eGDARRAY_Strings
        
        With fgLots
            For lRow = 0 To .Rows - 1
                If Not ((.RowOutlineLevel(lRow) = 2) And (.MergeRow(lRow) = True)) Then
                    Set astrLine = New cGdArray
                    astrLine.Create eGDARRAY_Strings
                    
                    If (bIncludeHiddenRows = True) Or (.RowHidden(lRow) = False) Then
                        strText = ""
                        strLastText = ""
                        
                        For lCol = .FixedCols To .Cols - 1
                            If (bIncludeHiddenColumns = True) Or (.ColHidden(lCol) = False) Then
                                strText = .Cell(flexcpTextDisplay, lRow, lCol)
                                
                                If (.MergeRow(lRow) = False) Or (strText <> strLastText) Then
                                    If InStr(strText, ",") <> 0 Then
                                        astrLine.Add Chr(34) & strText & Chr(34)
                                    Else
                                        astrLine.Add strText
                                    End If
                                    
                                    strLastText = strText
                                Else
                                    astrLine.Add ""
                                End If
                            End If
                        Next lCol
                        
                        astrFile.Add astrLine.JoinFields(",")
                    End If
                End If
            Next lRow
        End With
        
        astrFile.ToFile strFileName
        
        InfBox "Lot Information saved to file:|" & strFileName & "||", "i", , g.Cattle.ProductName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ExportToCsv"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolForRow
'' Description: Determine the symbol for the given row
'' Inputs:      Row
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolForRow(ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lRowOutlineLevel As Long        ' Row outline level for the given row
    Dim lSymbolRow As Long              ' Symbol row for the child
    Dim Info As cBrokerMessage          ' Information object
    
    strReturn = ""
    With fgLots
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            lRowOutlineLevel = .RowOutlineLevel(lRow)
            
            If (lRowOutlineLevel = 2) And (RowType(lRow) = eGDRowType_Symbol) Then
                Set Info = .RowData(lRow)
                strReturn = Info("Symbol")
            
            ElseIf lRowOutlineLevel > 2 Then
                lSymbolRow = .GetNodeRow(lRow, flexNTParent)
                Do While .RowOutlineLevel(lSymbolRow) > 2
                    lSymbolRow = .GetNodeRow(lSymbolRow, flexNTParent)
                Loop
                
                Set Info = .RowData(lSymbolRow)
                strReturn = Info("Symbol")
            End If
        End If
    End With
    
    SymbolForRow = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.SymbolForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoSizeGrid
'' Description: Auto size the grid columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AutoSizeGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim astrHeader As cGdArray          ' Header fields
    Dim lIndex As Long                  ' Index into a for loop
    Dim iSpace As Integer               ' Index of a space
    Dim strHeader As String             ' Header text

    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        Set astrHeader = New cGdArray
        astrHeader.Create eGDARRAY_Strings, .Cols
        
        For lIndex = 0 To .Cols - 1
            strHeader = .TextMatrix(0, lIndex)

            astrHeader(lIndex) = strHeader
            If InStr(strHeader, " ") <> 0 Then
                iSpace = InStr(Int(Len(strHeader) / 2), strHeader, " ")
                If iSpace = 0 Then
                    iSpace = InStrRev(strHeader, " ", Int(Len(strHeader) / 2))
                    If iSpace < 3 Then
                        .TextMatrix(0, lIndex) = strHeader ' & "00"
                    Else
                        .TextMatrix(0, lIndex) = Right(strHeader, Len(strHeader) - iSpace) ' & "00"
                    End If
                Else
                    If iSpace > Len(strHeader) - 3 Then
                        .TextMatrix(0, lIndex) = strHeader ' & "00"
                    Else
                        .TextMatrix(0, lIndex) = Left(strHeader, iSpace) ' & "00"
                    End If
                End If
            Else
                .TextMatrix(0, lIndex) = strHeader ' & "00"
            End If
        Next lIndex
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1, False ', 75
        
        For lIndex = 0 To .Cols - 1
            .TextMatrix(0, lIndex) = astrHeader(lIndex)
        Next lIndex
        
        .RowHeight(0) = -1&
        .RowHeight(0) = .RowHeight(0) * 2
        '.Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightBottom
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.AutoSizeGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveDetails
'' Description: Remove the appropriate details out of the collection
'' Inputs:      Feed Yard ID, Feed Yard Lot ID, Lot Column ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveDetails(ByVal strFeedyardID As String, Optional ByVal strFeedYardLotID As String = "", Optional ByVal strLotColumnID As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Detail As cBrokerMessage        ' Detail object
    
    For lIndex = m.Details.Count To 1 Step -1
        Set Detail = m.Details(lIndex)
        
        If Detail("FeedYardID") = strFeedyardID Then
            If (Len(strFeedYardLotID) = 0) Or (Detail("FeedYardLotID") = strFeedYardLotID) Then
                If (Len(strLotColumnID) = 0) Or (Detail("LotColumnID") = strLotColumnID) Then
                    m.Details.Remove lIndex
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLots.RemoveDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDetails
'' Description: Get the appropriate details out of the collection
'' Inputs:      Feed Yard ID, Feed Yard Lot ID, Lot Column ID
'' Returns:     Details collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDetails(ByVal strFeedyardID As String, Optional ByVal strFeedYardLotID As String = "", Optional ByVal strLotColumnID As String = "") As cGdTree
On Error GoTo ErrSection:

    Dim Details As cGdTree              ' Collection of details
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Details = New cGdTree
    For lIndex = 1 To m.Details.Count
        Set Detail = m.Details(lIndex)
        
        If Detail("FeedYardID") = strFeedyardID Then
            If (Len(strFeedYardLotID) = 0) Or (Detail("FeedYardLotID") = strFeedYardLotID) Then
                If (Len(strLotColumnID) = 0) Or (Detail("LotColumnID") = strLotColumnID) Then
                    Details.Add Detail, Detail("ID")
                End If
            End If
        End If
    Next lIndex
    
    Set GetDetails = Details

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.GetDetails"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDetailString
'' Description: Build a string out of the appropriate details
'' Inputs:      Feed Yard ID, Feed Yard Lot ID, Lot Column ID
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDetailString(ByVal strFeedyardID As String, ByVal strFeedYardLotID As String, ByVal strLotColumnID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrValues As cGdArray          ' Values for the details
    Dim Details As cGdTree              ' Collection of details
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lArrayIndex As Long             ' Index into the array
    
    strReturn = ""
    Set Details = GetDetails(strFeedyardID, strFeedYardLotID, strLotColumnID)
    If Details.Count = 1 Then
        Set Detail = Details(1)
        strReturn = Detail("Value")
    ElseIf Details.Count > 0 Then
        Set astrValues = New cGdArray
        astrValues.Create eGDARRAY_Strings, Details.Count
        
        For lIndex = 1 To Details.Count
            Set Detail = Details(lIndex)
            
            lArrayIndex = Minute(Val(Detail("Date"))) - 2
            astrValues(lArrayIndex) = Detail("Value")
        Next lIndex
        
        strReturn = astrValues.JoinFields(";")
    End If
    
    GetDetailString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.GetDetailString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowReports
'' Description: Show the reports
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowReports()
On Error GoTo ErrSection:

    Dim Lots As cGdTree                 ' Collection of lots
    Dim lRow As Long                    ' Row in the grid
    
    If SelectedFeedYard = -1& Then
        InfBox "Please select a feedyard before viewing reports", "!", , "Error"
    Else
        Set Lots = New cGdTree
        
        With fgLots
            lRow = .FixedRows
            Do While lRow <> -1&
                If UCase(.TextMatrix(lRow, LotCol(eGDLotCol_LotNumber))) <> "TOTALS" Then
                    If TypeOf .RowData(lRow) Is cBrokerMessage Then
                        Lots.Add .RowData(lRow)
                    End If
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End With
        
        If Lots.Count > 0 Then
            frmCattleReport.ShowMe Lots, m.Details, m.LotColumns, m.Trades
        Else
            InfBox "There are no lots for this feedyard", "!", , "Error"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ShowReports"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportHistoricalFills
'' Description: Allow the user to import historical fills
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ImportHistoricalFills()
On Error GoTo ErrSection:

    Dim Fills As cGdTree                ' Collection of fills
    Dim AccountFills As cGdTree         ' Collection of fills for an account
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Account As cBrokerMessage       ' Account object
    Dim brokerFill As cBrokerMessage    ' Fill object
    Dim Fill As cBrokerMessage          ' Fill object
    Dim strKey As String                ' Key into the fills collection
    Dim lQuantityKnown As Long          ' Quantity already known for the fill
    Dim lFillQuantity As Long           ' Fill quantity
    Dim Consolidated As cFillQuantities ' Consolidated carried fills

    Set Fills = New cGdTree

    If g.AppBridge.BrokerViewLoaded Then
        Set AccountFills = g.AppBridge.GetCarriedFillsForBroker
        
        If AccountFills.Count > 0 Then
            Set Consolidated = New cFillQuantities
            For lIndex = 1 To AccountFills.Count
                Consolidated.AddFromFill AccountFills(lIndex)
            Next lIndex
            
            For lIndex = 1 To Consolidated.Count
                Set brokerFill = Consolidated.FillForItem(lIndex).MakeCopy
                
                lFillQuantity = CLng(Val(brokerFill("Quantity")))
                lQuantityKnown = g.Cattle.FillQuantities.QuantityForFill(brokerFill)
                
                If lFillQuantity > lQuantityKnown Then
                    brokerFill("Quantity") = Str(lFillQuantity - lQuantityKnown)
                    
                    strKey = FillKey(brokerFill("Broker"), brokerFill("BrokerFillID"))
                    If m.Fills.Exists(strKey) = False Then
                        If Fills.Exists(brokerFill("BrokerFillID")) = False Then
                            Fills.Add brokerFill, strKey
                        End If
                    End If
                End If
            Next lIndex
        End If
    Else
        For lIndex = 1 To m.Accounts.Count
            Set Account = m.Accounts(lIndex)
            
            Set AccountFills = g.AppBridge.GetHistoricalFillsForAccount(Account("Number"))
            For lIndex2 = 1 To AccountFills.Count
                Set brokerFill = AccountFills(lIndex2)
                
                strKey = FillKey(brokerFill("Broker"), brokerFill("BrokerFillID"))
                If m.Fills.Exists(strKey) = False Then
                    If Fills.Exists(brokerFill("BrokerFillID")) = False Then
                        Fills.Add brokerFill, strKey
                    End If
                End If
            Next lIndex2
        Next lIndex
    End If
    
    If Fills.Count = 0 Then
        InfBox "There are no fills to import", "i", , "Import Fills"
    ElseIf frmCattleSelect.ShowMeFills(Fills) Then
        Set Fills = frmCattleSelect.Fills
        For lIndex = 1 To Fills.Count
            Fill.Add "FeedYardID", Str(SelectedFeedYard)
            g.Cattle.UpdateFill Fill
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.ImportHistoricalFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotDetailsToGrid
'' Description: Send the lot details to the grid for certain columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LotDetailsToGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim lCol As Long                    ' Column in the grid
    Dim lRow As Long                    ' Row in the grid
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim Lot As cBrokerMessage           ' Lot object
    Dim strDetail As String             ' Detail string
    Dim bChanged As Boolean             ' Did we change anything?
    
    With fgLots
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        bChanged = False
        For lCol = 0 To .Cols - 1
            Set LotColumn = LotColumnForCol(lCol)
            If Not LotColumn Is Nothing Then
                If (UCase(LotColumn.KeyValueField) = "OWNERNAME") Or (UCase(LotColumn.KeyValueField) = "OWNERNUMBER") Or (UCase(LotColumn.KeyValueField) = "PERCENTOWNED") Then
                    For lRow = .FixedRows To .Rows - 1
                        If .RowOutlineLevel(lRow) = 1 Then
                            If TypeOf .RowData(lRow) Is cBrokerMessage Then
                                Set Lot = .RowData(lRow)
                                
                                strDetail = GetDetailString(Lot("FeedYardID"), Lot("FeedYardLotID"), LotColumn.ID)
                                If Len(strDetail) > 0 Then
                                    .TextMatrix(lRow, lCol) = strDetail
                                    bChanged = True
                                End If
                            End If
                        End If
                    Next lRow
                End If
            End If
        Next lCol
        
        If bChanged Then
            AutoSizeGrid
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.LotDetailsToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditLot
'' Description: Allow the user to edit a lot
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditLot(Optional ByVal strDefaultKeyValueField As String = "")
On Error GoTo ErrSection:

    Dim Lot As cBrokerMessage           ' Lot object
    Dim LotDetails As cGdTree           ' Lot details object
    
    Set Lot = SelectedLot
    If Not Lot Is Nothing Then
        Set LotDetails = GetDetails(Str(SelectedFeedYard), Lot("FeedYardLotID"))
    
        If frmEditLot.ShowMe(SelectedFeedYard, Lot, LotDetails, strDefaultKeyValueField) Then
            g.Cattle.AddLots Lot.ToString
            g.Cattle.UpdateLotContentDetails LotDetails, Str(SelectedFeedYard), Lot("FeedYardLotID")
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.EditLot"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyMoveColumn
'' Description: Verify a column move
'' Inputs:      Column, Position
'' Returns:     Actual new position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyMoveColumn(ByVal Col As Long, ByVal Position As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lLotNumberCol As Long           ' Lot number column
    Dim lOpenEquityCol As Long          ' Open equity column

    lReturn = Position
    lLotNumberCol = LotCol(eGDLotCol_LotNumber)
    lOpenEquityCol = LotCol(eGDLotCol_OpenEquity)
    
    With fgLots
        If Col = lLotNumberCol Then
            lReturn = Col
        ElseIf Position <= lLotNumberCol Then
            lReturn = lLotNumberCol + 1
        ElseIf Col >= lOpenEquityCol Then
            lReturn = Col
        ElseIf Position >= lOpenEquityCol Then
            lReturn = lOpenEquityCol - 1
        End If
    End With

    VerifyMoveColumn = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.VerifyMoveColumn"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotIsClosed
'' Description: Determine if the lot for the given row is closed
'' Inputs:      Row
'' Returns:     True if closed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LotIsClosed(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    With fgLots
        If LotCol(eGDLotCol_IsClosed) = kNullData Then
            bReturn = (UCase(Left(.TextMatrix(lRow, LotCol(eGDLotCol_LotStatus)), 1)) = "C")
        Else
            bReturn = (CheckedCell(fgLots, lRow, LotCol(eGDLotCol_IsClosed)) = True)
        End If
    End With
    
    LotIsClosed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLots.LotIsClosed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SynchronizeIngredients
'' Description: Synchronize the ingredients between the lot content details and
''              the ingredient list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SynchronizeIngredients()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim strDetailOptions As String      ' Detail options
    Dim astrIngredients As cGdArray     ' Ingredients from the details
    Dim Ingredients As cGdTree          ' Collection of ingredient objects
    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim strID As String                 ' ID for the ingredient
    Static bInProgress As Boolean       ' Are we currently in progress?

    If bInProgress = False Then
        bInProgress = True
        
        If m.bHasDetails And m.bHasIngredients Then
            If m.KeyValueToIndex.Exists("Ingredient") Then
                Set Ingredients = g.Cattle.Ingredients
                
                lIndex = m.KeyValueToIndex("Ingredient")
                Set LotColumn = m.LotColumns(lIndex)
                strDetailOptions = g.Cattle.DetailOptions(Str(LotColumn.ID))
                
                Set astrIngredients = New cGdArray
                astrIngredients.SplitFields strDetailOptions, "|"
                
                For lIndex = 0 To astrIngredients.Size - 1
                    strID = g.Cattle.IngredientIDForName(astrIngredients(lIndex))
                    If Len(strID) = 0 Then
                        Set Ingredient = New cBrokerMessage
                        
                        Ingredient.Add "FeedYardID", Str(SelectedFeedYard)
                        Ingredient.Add "Ingredient", astrIngredients(lIndex)
                        Ingredient.Add "DryFeedPct", ""
                        Ingredient.Add "CostPerPound", ""
                        
                        g.Cattle.UpdateIngredient Ingredient
                    End If
                Next lIndex
            End If
        End If
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.SynchronizeIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewFill
'' Description: Allow the user to create a new fill and associate it with a lot
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewFill()
On Error GoTo ErrSection:

    Dim CattleFill As cBrokerMessage    ' Fill object
    Dim Account As cAccount             ' Account object
    Dim strFillKey As String            ' Key into the wait fills collection
    
    Set CattleFill = New cBrokerMessage
    
    If g.AppBridge.ShowCattleFill(CattleFill) = True Then
        Set Account = g.Cattle.Accounts(CattleFill("BrokerAccountID"))
        If Not Account Is Nothing Then
            CattleFill.Add "Broker", Str(Account.Broker)
            CattleFill.Add "BrokerAccountNumber", Account.AccountNumber
            CattleFill.Add "FcmAccount", Account.FcmNumber
            CattleFill.Add "AssociatedQuantity", CattleFill("Quantity")
            
            strFillKey = AssociatedFillKey(CattleFill("Broker"), CattleFill("BrokerFillID"), CattleFill("FeedYardLotID"))
            
            g.Cattle.DumpDebug "Associated Fill '" & strFillKey & "' added to list waiting for fill to come back"
            m.WaitFillFills.Add CattleFill, strFillKey
            
            UpdateFill CattleFill, "manual fill created"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLots.NewFill"
    
End Sub

