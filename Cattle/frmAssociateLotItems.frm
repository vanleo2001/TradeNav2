VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAssociateLotItems 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraFilters 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7635
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
      Caption         =   "frmAssociateLotItems.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAssociateLotItems.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssociateLotItems.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraSubFilters 
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   0
         Width           =   4755
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
         Caption         =   "frmAssociateLotItems.frx":0068
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAssociateLotItems.frx":0094
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAssociateLotItems.frx":00B4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboAccounts 
            Height          =   315
            Left            =   3240
            TabIndex        =   7
            Top             =   0
            Width           =   1515
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
            Tip             =   "frmAssociateLotItems.frx":00D0
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmAssociateLotItems.frx":00F0
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboSymbols 
            Height          =   315
            Left            =   720
            TabIndex        =   5
            Top             =   0
            Width           =   1515
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
            Tip             =   "frmAssociateLotItems.frx":010C
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmAssociateLotItems.frx":012C
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAccounts 
            Height          =   195
            Left            =   2400
            Top             =   60
            Width           =   735
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
            Caption         =   "frmAssociateLotItems.frx":0148
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAssociateLotItems.frx":017C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAssociateLotItems.frx":019C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSymbols 
            Height          =   195
            Left            =   0
            Top             =   60
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
            Caption         =   "frmAssociateLotItems.frx":01B8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAssociateLotItems.frx":01EA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAssociateLotItems.frx":020A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniComboImageXP cboLots 
         Height          =   315
         Left            =   420
         TabIndex        =   2
         Top             =   0
         Width           =   2355
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
         Tip             =   "frmAssociateLotItems.frx":0226
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAssociateLotItems.frx":0246
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblLots 
         Height          =   195
         Left            =   0
         Top             =   60
         Width           =   435
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
         Caption         =   "frmAssociateLotItems.frx":0262
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAssociateLotItems.frx":028E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAssociateLotItems.frx":02AE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
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
      Caption         =   "frmAssociateLotItems.frx":02CA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAssociateLotItems.frx":02F6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssociateLotItems.frx":0316
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmAssociateLotItems.frx":0332
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAssociateLotItems.frx":0360
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAssociateLotItems.frx":0380
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1215
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
         Caption         =   "frmAssociateLotItems.frx":039C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAssociateLotItems.frx":03C2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAssociateLotItems.frx":03E2
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrders 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   780
      Width           =   2895
      _cx             =   5106
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
   Begin VSFlex7LCtl.VSFlexGrid fgFills 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2340
      Width           =   2895
      _cx             =   5106
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
   Begin HexUniControls.ctlUniLabelXP lblFills 
      Height          =   195
      Left            =   120
      Top             =   2100
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
      Caption         =   "frmAssociateLotItems.frx":03FE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAssociateLotItems.frx":042A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssociateLotItems.frx":044A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOrders 
      Height          =   195
      Left            =   120
      Top             =   540
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
      Caption         =   "frmAssociateLotItems.frx":0466
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAssociateLotItems.frx":0494
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssociateLotItems.frx":04B4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmAssociateLotItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAssociateLotItems.frm
'' Description: Form for allowing user to associate orders/fills for lots
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/14/2012   DAJ         Set key correctly in returning collections
'' 06/14/2012   DAJ         Fix for getting orders and fills off of the broker view form
'' 06/25/2012   DAJ         Changed key value field for ID on Lots
'' 09/11/2012   DAJ         Filters, Associate parts of fills
'' 09/14/2012   DAJ         Pass selected lot into the fill assign form for new association
'' 09/18/2012   DAJ         Send associated fills on new association as well
'' 10/22/2012   DAJ         Rename Turnkey to HedgeLinc
'' 11/15/2013   DAJ         Changed how to get Turnkey Icon for the form
'' 11/22/2013   DAJ         Renamed frmTurnkeySelectLot to frmTurnkeySelect
'' 11/22/2013   DAJ         Import historical fills for Turnkey
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 05/08/2014   DAJ         Made remove association confirmation more verbose
'' 05/22/2014   DAJ         Renamed frmTurnkeyFillAssign to frmAssignFillToLot; Renamed
''                          frmTurnkeySelect to frmCattleSelect
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Utilized new accounts object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kstrModify As String = "Modify"
Private Const kstrRemove As String = "Remove"

Private Enum eGDOrderCols
    eGDOrderCols_On = 0
    eGDOrderCols_OrderId
    eGDOrderCols_Description
    eGDOrderCols_Account
    eGDOrderCols_Symbol
    eGDOrderCols_Lot
    eGDOrderCols_FeedYardLotID
    eGDOrderCols_NumCols
End Enum

Private Enum eGDFillCols
    eGDFillCols_BrokerFillId = 0
    eGDFillCols_Side
    eGDFillCols_Quantity
    eGDFillCols_AssignedQuantity
    eGDFillCols_Symbol
    eGDFillCols_Price
    eGDFillCols_Account
    eGDFillCols_Time
    eGDFillCols_TurnkeyFillId
    eGDFillCols_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog?
    Lot As cBrokerMessage               ' Lot passed in
    AssociatedAccounts As cGdTree       ' Collection of associated accounts
    
    astrSymbols As cGdArray             ' Unique list of symbols from the orders and fills
    astrAccounts As cGdArray            ' Unique list of accounts from the orders and fills
End Type
Private m As mPrivate

Private Function OrderCol(ByVal nCol As eGDOrderCols) As Long
    OrderCol = nCol
End Function

Private Function FillCol(ByVal nCol As eGDFillCols) As Long
    FillCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Associated Orders, Associated Fills, Associated Accounts, Fills, Lot
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(AssociatedOrders As cGdTree, AssociatedFills As cGdTree, ByVal AssociatedAccounts As cGdTree, ByVal BrokerFills As cGdTree, Optional ByVal Lot As cBrokerMessage = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim turnkeyMessage As cBrokerMessage ' Turnkey message
    Dim strFeedYardLotID As String      ' Feed Yard Lot ID

    Set m.Lot = Lot
    Set m.AssociatedAccounts = AssociatedAccounts
        
    If Lot Is Nothing Then
        strFeedYardLotID = ""
        Caption = "Associated Items for Lots"
    Else
        strFeedYardLotID = Lot("FeedYardLotID")
        Caption = "Associated Items for Lot #" & g.Cattle.LotDisplay(Lot)
    End If

    InitOrdersGrid
    LoadOrdersGrid AssociatedOrders
    
    InitFillsGrid
    LoadFillsGrid AssociatedFills, BrokerFills
    
    LoadLotsCombo strFeedYardLotID
    LoadSymbolsCombo
    LoadAccountsCombo
    
    FilterGrids

    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK Then
        AssociatedOrders.Clear
        With fgOrders
            For lIndex = .FixedRows To .Rows - 1
                If CheckedCell(fgOrders, lIndex, OrderCol(eGDOrderCols_On)) Then
                    Set turnkeyMessage = .RowData(lIndex)
                    If m.Lot Is Nothing Then
                        turnkeyMessage.Add "NewFeedYardLotID", .TextMatrix(lIndex, OrderCol(eGDOrderCols_FeedYardLotID))
                    Else
                        turnkeyMessage.Add "NewFeedYardLotID", m.Lot("FeedYardLotID")
                    End If
                    
                    AssociatedOrders.Add turnkeyMessage, turnkeyMessage("Broker") & "|" & turnkeyMessage("BrokerOrderID")
                End If
            Next lIndex
        End With
        
        AssociatedFills.Clear
        With fgFills
            For lIndex = .FixedRows To .Rows - 1
                If (.RowOutlineLevel(lIndex) = 2) And (IsClickHereRow(lIndex) = False) Then
                    Set turnkeyMessage = .RowData(lIndex)
                    AssociatedFills.Add turnkeyMessage, turnkeyMessage("Broker") & "|" & turnkeyMessage("BrokerFillID") & "|" & turnkeyMessage("FeedYardLotID")
                End If
            Next lIndex
        End With
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmAssociateLotItems.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: User has chosen to change the account filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrids
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.cboAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboLots_Click
'' Description: User has chosen to change the lot filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboLots_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrids
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.cboLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbols_Click
'' Description: User has chosen to change the symbol filter
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbols_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrids
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.cboSymbols_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: User has chosen to cancel the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: User has chosen to OK the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_Click
'' Description: Toggle check box when user clicks in column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row of the cell clicked on in the grid
    Dim lMouseCol As Long               ' Column of the cell clicked on in the grid

    With fgFills
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If IsClickHereRow(lMouseRow) Then
                NewAssociation .GetNodeRow(lMouseRow, flexNTParent)
            ElseIf .RowOutlineLevel(lMouseRow) = 2 Then
                If .TextMatrix(lMouseRow, lMouseCol) = kstrModify Then
                    EditAssociation lMouseRow
                ElseIf .TextMatrix(lMouseRow, lMouseCol) = kstrRemove Then
                    RemoveAssociation lMouseRow
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.fgFills_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_Click
'' Description: Toggle check box when user clicks in column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row of the cell clicked on in the grid
    Dim lMouseCol As Long               ' Column of the cell clicked on in the grid

    With fgOrders
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If lMouseCol = OrderCol(eGDOrderCols_On) Then
                If CheckedCell(fgOrders, lMouseRow, lMouseCol) = True Then
                    CheckedCell(fgOrders, lMouseRow, lMouseCol) = False
                    
                    .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_Lot)) = ""
                    .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_FeedYardLotID)) = ""
                Else
                    CheckedCell(fgOrders, lMouseRow, lMouseCol) = True
                    
                    If m.Lot Is Nothing Then
                        If frmCattleSelect.ShowMeLot Then
                            .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_Lot)) = g.Cattle.LotDisplayForID(frmCattleSelect.FeedYardLotID)
                            .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_FeedYardLotID)) = frmCattleSelect.FeedYardLotID
                        Else
                            CheckedCell(fgOrders, lMouseRow, lMouseCol) = True
                            .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_Lot)) = ""
                            .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_FeedYardLotID)) = ""
                        End If
                    Else
                        .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_Lot)) = g.Cattle.LotDisplay(m.Lot)
                        .TextMatrix(lMouseRow, OrderCol(eGDOrderCols_FeedYardLotID)) = m.Lot("FeedYardLotID")
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.fgOrders_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    
    Set m.astrSymbols = New cGdArray
    m.astrSymbols.Create eGDARRAY_Strings
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: User has chosen to close the dialog
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = True
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.Form_QueryUnload"

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

    Dim lGridHeight As Long             ' Height of each of the grids
    Dim lSpace As Long                  ' Space between controls

    lSpace = 120

    If LimitFormSize(Me, 7920, 4365) = False Then
        With fraFilters
            .Move lSpace, lSpace, ScaleWidth - (lSpace * 2)
        End With
        
        With fraSubFilters
            .Move fraFilters.Width - .Width
        End With
        
        With cboLots
            .Move .Left, .Top, fraSubFilters.Left - .Left - lSpace
        End With
        
        With fraButtons
            .Move (ScaleWidth / 2) - (.Width / 2), ScaleHeight - .Height - lSpace
        End With
        
        lGridHeight = (ScaleHeight - fraFilters.Height - lblOrders.Height - lblFills.Height - fraButtons.Height - (lSpace * 4)) / 2
        With fgFills
            .Move lSpace, fraButtons.Top - lGridHeight - 120, ScaleWidth - (lSpace * 2), lGridHeight
        End With
        With lblFills
            .Move lSpace, fgFills.Top - .Height
        End With
        With fgOrders
            .Move lSpace, lblFills.Top - lGridHeight, ScaleWidth - (lSpace * 2), lGridHeight
        End With
        With lblOrders
            .Move lSpace, fgOrders.Top - .Height
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

    SaveFormPlacement Me
    
    Set m.astrAccounts = Nothing
    Set m.astrSymbols = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitOrdersGrid
'' Description: Initialize the orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitOrdersGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .MergeCells = flexMergeNever
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = OrderCol(eGDOrderCols_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, OrderCol(eGDOrderCols_On)) = "Use"
        .TextMatrix(0, OrderCol(eGDOrderCols_OrderId)) = "ID"
        .TextMatrix(0, OrderCol(eGDOrderCols_Description)) = "Description"
        .TextMatrix(0, OrderCol(eGDOrderCols_Account)) = "Account"
        .TextMatrix(0, OrderCol(eGDOrderCols_Symbol)) = "Symbol"
        .TextMatrix(0, OrderCol(eGDOrderCols_Lot)) = "Lot"
        .TextMatrix(0, OrderCol(eGDOrderCols_FeedYardLotID)) = "Lot ID"
        
        .ColHidden(OrderCol(eGDOrderCols_Account)) = True
        .ColHidden(OrderCol(eGDOrderCols_Symbol)) = True
        .ColHidden(OrderCol(eGDOrderCols_Lot)) = (Not m.Lot Is Nothing)
        .ColHidden(OrderCol(eGDOrderCols_FeedYardLotID)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.InitOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrdersGrid
'' Description: Load the orders grid
'' Inputs:      Associated Orders
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOrdersGrid(ByVal AssociatedOrders As cGdTree)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Order As cBrokerMessage         ' Order object
    Dim BrokerOrders As cGdTree         ' Broker orders
    Dim lPos As Long                    ' Position in the array

    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To AssociatedOrders.Count
            Set Order = AssociatedOrders(lIndex)
            
            .Rows = .Rows + 1
            
            .RowData(.Rows - 1) = Order
            
            CheckedCell(fgOrders, .Rows - 1, OrderCol(eGDOrderCols_On)) = True
            .Cell(flexcpPictureAlignment, .Rows - 1, OrderCol(eGDOrderCols_On)) = flexAlignCenterCenter
            
            OrderToGrid .Rows - 1, Order
        Next lIndex
        
        Set BrokerOrders = g.AppBridge.GetWorkingOrders
        For lIndex = 1 To BrokerOrders.Count
            Set Order = BrokerOrders(lIndex)
            
            If AccountIsAssociated(Order("Broker"), Order("BrokerAccountNumber")) Then
                If AssociatedOrders.Exists(Order("Broker") & "|" & Order("BrokerOrderID")) = False Then
                    .Rows = .Rows + 1
                    
                    .RowData(.Rows - 1) = Order
                    CheckedCell(fgOrders, .Rows - 1, OrderCol(eGDOrderCols_On)) = False
                    .Cell(flexcpPictureAlignment, .Rows - 1, OrderCol(eGDOrderCols_On)) = flexAlignCenterCenter
                    
                    OrderToGrid .Rows - 1, Order
                End If
            End If
        Next lIndex
        
        SetBackColors fgOrders
        .AutoSize 0, .Cols - 1, False, 75
    
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.LoadOrdersGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToGrid
'' Description: Send the order to the grid
'' Inputs:      Row, Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderToGrid(ByVal lRow As Long, ByVal turnkeyOrder As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim Bars As cGdBars                 ' Bars object
    Dim strAccount As String            ' Account to display
    Dim lPos As Long                    ' Position in an array
    Dim strFeedYardLotID As String      ' Feed Yard Lot ID for the order
    
    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        strAccount = g.Cattle.Accounts.DisplayAccountNumber(turnkeyOrder("BrokerAccountID"))
        
        .TextMatrix(lRow, OrderCol(eGDOrderCols_OrderId)) = turnkeyOrder("BrokerOrderID")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Description)) = g.Cattle.OrderToString(turnkeyOrder, turnkeyOrder("FcmAccount"))
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Account)) = strAccount
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Symbol)) = turnkeyOrder("Symbol")
        
        strFeedYardLotID = turnkeyOrder("FeedYardLotID")
        If (Len(strFeedYardLotID) = 0) Or (strFeedYardLotID = "0") Then
            .TextMatrix(lRow, OrderCol(eGDOrderCols_Lot)) = ""
            .TextMatrix(lRow, OrderCol(eGDOrderCols_FeedYardLotID)) = ""
        Else
            .TextMatrix(lRow, OrderCol(eGDOrderCols_Lot)) = g.Cattle.LotDisplayForID(strFeedYardLotID)
            .TextMatrix(lRow, OrderCol(eGDOrderCols_FeedYardLotID)) = strFeedYardLotID
        End If
        
        If Len(strAccount) > 0 Then
            If m.astrAccounts.BinarySearch(strAccount, lPos) = False Then
                m.astrAccounts.Add strAccount, lPos
            End If
        End If
        If Len(turnkeyOrder("Symbol")) > 0 Then
            If m.astrSymbols.BinarySearch(turnkeyOrder("Symbol"), lPos) = False Then
                m.astrSymbols.Add turnkeyOrder("Symbol"), lPos
            End If
        Else
            lPos = lPos
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.OrderToGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFillsGrid
'' Description: Initialize the fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFillsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = FillCol(eGDFillCols_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, FillCol(eGDFillCols_BrokerFillId)) = "ID"
        .TextMatrix(0, FillCol(eGDFillCols_Side)) = "Side"
        .TextMatrix(0, FillCol(eGDFillCols_Quantity)) = "Qty"
        .TextMatrix(0, FillCol(eGDFillCols_AssignedQuantity)) = "Assigned"
        .TextMatrix(0, FillCol(eGDFillCols_Symbol)) = "Symbol"
        .TextMatrix(0, FillCol(eGDFillCols_Price)) = "Price"
        .TextMatrix(0, FillCol(eGDFillCols_Account)) = "Account"
        .TextMatrix(0, FillCol(eGDFillCols_Time)) = "Time"
        .TextMatrix(0, FillCol(eGDFillCols_TurnkeyFillId)) = "ID"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .ColFormat(FillCol(eGDFillCols_Time)) = DateFormat("Format", MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
        
        .ColAlignment(FillCol(eGDFillCols_BrokerFillId)) = flexAlignLeftCenter
        .ColAlignment(FillCol(eGDFillCols_Price)) = flexAlignRightCenter
        .ColAlignment(FillCol(eGDFillCols_Account)) = flexAlignLeftCenter
        
        .ColHidden(FillCol(eGDFillCols_TurnkeyFillId)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.InitFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFillsGrid
'' Description: Load the fills grid
'' Inputs:      Associated Fills
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFillsGrid(ByVal AssociatedFills As cGdTree, ByVal BrokerFills As cGdTree)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cBrokerMessage          ' Fill object
    Dim lFillRow As Long                ' Row for the fill
    Dim lNewRow As Long                 ' New row number for the associated fill

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To BrokerFills.Count
            .Rows = .Rows + 1
            .RowOutlineLevel(.Rows - 1) = 1
            .IsSubtotal(.Rows - 1) = True
            .MergeRow(.Rows - 1) = False
            
            FillToGrid .Rows - 1, BrokerFills(lIndex)
        Next lIndex
        
        
        For lIndex = 1 To AssociatedFills.Count
            Set Fill = AssociatedFills(lIndex)
            AddAssociatedFill Fill
        Next lIndex
        
        If .Rows > .FixedRows Then
            lFillRow = .FixedRows
            Do While lFillRow > -1&
                CalculateFillInfo lFillRow
                lFillRow = .GetNodeRow(lFillRow, flexNTNextSibling)
            Loop
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
    
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.LoadFillsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Send a fill to the grid
'' Inputs:      Row, Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(ByVal lRow As Long, ByVal turnkeyFill As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strAccount As String            ' Account to display
    Dim lPos As Long                    ' Position in an array
    Dim strFeedYardLotID As String      ' Feed Yard Lot ID for the fill

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        strAccount = g.Cattle.Accounts.DisplayAccountNumber(turnkeyFill("BrokerAccountID"))
        
        .RowData(lRow) = turnkeyFill
        
        .TextMatrix(lRow, FillCol(eGDFillCols_BrokerFillId)) = turnkeyFill("BrokerFillID")
        If turnkeyFill("IsBuy") = "1" Then
            .TextMatrix(lRow, FillCol(eGDFillCols_Side)) = "Bought"
        Else
            .TextMatrix(lRow, FillCol(eGDFillCols_Side)) = "Sold"
        End If
        .TextMatrix(lRow, FillCol(eGDFillCols_Quantity)) = turnkeyFill("Quantity")
        .TextMatrix(lRow, FillCol(eGDFillCols_Symbol)) = turnkeyFill("Symbol")
        .TextMatrix(lRow, FillCol(eGDFillCols_Price)) = g.AppBridge.PriceDisplay(Val(turnkeyFill("Price")), turnkeyFill("Symbol"))
        .TextMatrix(lRow, FillCol(eGDFillCols_Account)) = strAccount
        .TextMatrix(lRow, FillCol(eGDFillCols_Time)) = Val(turnkeyFill("FillTime"))
        .TextMatrix(lRow, FillCol(eGDFillCols_TurnkeyFillId)) = turnkeyFill("ID")
        
        If Len(strAccount) > 0 Then
            If m.astrAccounts.BinarySearch(strAccount, lPos) = False Then
                m.astrAccounts.Add strAccount, lPos
            End If
        End If
        If Len(turnkeyFill("Symbol")) > 0 Then
            If m.astrSymbols.BinarySearch(turnkeyFill("Symbol"), lPos) = False Then
                m.astrSymbols.Add turnkeyFill("Symbol"), lPos
            End If
        Else
            lPos = lPos
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddAssociatedFill
'' Description: Add the given associated fill to the grid
'' Inputs:      Associated Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddAssociatedFill(ByVal turnkeyAssociatedFill As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lFillRow As Long                ' Row for the fill
    Dim lNewRow As Long                 ' New row number for the associated fill

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
            
        lFillRow = RowForFill(turnkeyAssociatedFill("FillID"))
        If lFillRow > -1& Then
            .Rows = .Rows + 1
            .RowOutlineLevel(.Rows - 1) = 1
            .IsSubtotal(.Rows - 1) = True
            
            lNewRow = NewAssociatedFillRow(lFillRow)
            If lNewRow <> .Rows - 1 Then
                .RowPosition(.Rows - 1) = lNewRow
            End If
            
            .RowOutlineLevel(lNewRow) = 2
            .IsSubtotal(lNewRow) = True
            .MergeRow(lNewRow) = True
            
            AssociatedFillToGrid lNewRow, turnkeyAssociatedFill
        End If
            
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.AddAssociatedFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociatedFillToGrid
'' Description: Send an associated fill to the grid
'' Inputs:      Row, Associated Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AssociatedFillToGrid(ByVal lRow As Long, ByVal turnkeyAssociatedFill As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strLotDisplay As String         ' Lot display that the fill is associated with
    Dim strText As String               ' Text to display in the grid

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .RowData(lRow) = turnkeyAssociatedFill
        strLotDisplay = g.Cattle.LotDisplayForID(turnkeyAssociatedFill("FeedYardLotID"))
        strText = turnkeyAssociatedFill("AssociatedQuantity") & " assigned to Lot# " & strLotDisplay
        .Cell(flexcpText, lRow, FillCol(eGDFillCols_BrokerFillId), lRow, FillCol(eGDFillCols_Price)) = strText
        
        .Cell(flexcpForeColor, lRow, FillCol(eGDFillCols_Account), lRow, FillCol(eGDFillCols_Time)) = vbBlue
        .Cell(flexcpFontUnderline, lRow, FillCol(eGDFillCols_Account), lRow, FillCol(eGDFillCols_Time)) = True
        .TextMatrix(lRow, FillCol(eGDFillCols_Account)) = kstrModify
        .TextMatrix(lRow, FillCol(eGDFillCols_Time)) = kstrRemove
                
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.AssociatedFillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLotsCombo
'' Description: Load the lots combo from the array
'' Inputs:      Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLotsCombo(ByVal strFeedYardLotID As String)
On Error GoTo ErrSection:

    g.Cattle.LoadLotsCombo cboLots, strFeedYardLotID, "All"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.LoadLotsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbolsCombo
'' Description: Load the symbols combo from the array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSymbolsCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With cboSymbols
        .Clear
        
        .AddItem "All"
        For lIndex = 0 To m.astrSymbols.Size - 1
            .AddItem m.astrSymbols(lIndex)
        Next lIndex
        
        .ListIndex = 0&
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.LoadSymbolsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsCombo
'' Description: Load the accounts combo from the array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAccountsCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With cboAccounts
        .Clear
        
        .AddItem "All"
        For lIndex = 0 To m.astrAccounts.Size - 1
            .AddItem m.astrAccounts(lIndex)
        Next lIndex
        
        .ListIndex = 0&
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.LoadAccountsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrids
'' Description: Filter the grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrids()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim bHide As Boolean                ' Hide the row?
    Dim Order As cBrokerMessage         ' Order object

    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            Set Order = .RowData(lIndex)
            
            bHide = (Order("IsWorking") <> "1")
            If bHide = False Then
                If (cboLots.ListIndex > 0) And (Len(.TextMatrix(lIndex, OrderCol(eGDOrderCols_FeedYardLotID))) > 0) Then
                    bHide = (.TextMatrix(lIndex, OrderCol(eGDOrderCols_FeedYardLotID)) <> Str(cboLots.ItemData(cboLots.ListIndex)))
                End If
            End If
            If bHide = False Then
                If (cboAccounts.ListIndex > 0) And (.TextMatrix(lIndex, OrderCol(eGDOrderCols_Account)) <> cboAccounts.Text) Then
                    bHide = True
                End If
            End If
            If bHide = False Then
                If (cboSymbols.ListIndex > 0) And (.TextMatrix(lIndex, OrderCol(eGDOrderCols_Symbol)) <> cboSymbols.Text) Then
                    bHide = True
                End If
            End If
            
            .RowHidden(lIndex) = bHide
        Next lIndex
        
        .Redraw = nRedraw
    End With

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            bHide = False
            
            If bHide = False Then
                If (cboAccounts.ListIndex > 0) And (.TextMatrix(lIndex, FillCol(eGDFillCols_Account)) <> cboAccounts.Text) Then
                    bHide = True
                End If
            End If
            If bHide = False Then
                If (cboSymbols.ListIndex > 0) And (.TextMatrix(lIndex, FillCol(eGDFillCols_Symbol)) <> cboSymbols.Text) Then
                    bHide = True
                End If
            End If
            
            .RowHidden(lIndex) = bHide
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.FilterGrids"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountIsAssociated
'' Description: Determine if the given account information is associated
'' Inputs:      Broker, Account Number
'' Returns:     True if Associated, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AccountIsAssociated(ByVal strBroker As String, ByVal strBrokerAccountNumber As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cBrokerMessage       ' Associated account
    
    bReturn = False
    For lIndex = 1 To m.AssociatedAccounts.Count
        Set Account = m.AssociatedAccounts(lIndex)
        If Account("Broker") = strBroker Then
            If Account("Number") = strBrokerAccountNumber Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    AccountIsAssociated = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.AccountIsAssociated"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsClickHereRow
'' Description: Determine if the given row is a click here row in the fills grid
'' Inputs:      Row
'' Returns:     True if Click Here row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsClickHereRow(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    With fgFills
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            If .RowOutlineLevel(lRow) = 2 Then
                bReturn = (UCase(.TextMatrix(lRow, FillCol(eGDFillCols_TurnkeyFillId))) = "CLICK")
            End If
        End If
    End With
    
    IsClickHereRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.IsClickHereRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForFill
'' Description: Determine the row in the fills grid for the given Turnkey Fill ID
'' Inputs:      Turnkey Fill ID
'' Returns:     Row for the given ID ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForFill(ByVal strTurnkeyFillID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Row in the grid

    lReturn = -1&
    With fgFills
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            Do While lRow <> -1&
                If .TextMatrix(lRow, FillCol(eGDFillCols_TurnkeyFillId)) = strTurnkeyFillID Then
                    lReturn = lRow
                    Exit Do
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
    End With
    
    RowForFill = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.RowForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForAssociatedFill
'' Description: Determine the row in the fills grid for the given Turnke
''              Associated Fill ID
'' Inputs:      Turnkey Fill ID, Turnkey Associated Fill ID
'' Returns:     Row for the given ID ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForAssociatedFill(ByVal strTurnkeyFillID As String, ByVal strTurnkeyAssociatedFillID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lFillRow As Long                ' Row in the grid for the fill
    Dim lRow As Long                    ' Row in the grid

    lReturn = -1&
    With fgFills
        If .Rows > .FixedRows Then
            lFillRow = RowForFill(strTurnkeyFillID)
            
            lRow = .GetNodeRow(lFillRow, flexNTFirstChild)
            Do While lRow <> -1&
                If .TextMatrix(lRow, FillCol(eGDFillCols_TurnkeyFillId)) = strTurnkeyAssociatedFillID Then
                    lReturn = lRow
                    Exit Do
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
    End With
    
    RowForAssociatedFill = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.RowForAssociatedFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociatedQuantityForFillRow
'' Description: Determine the associated quantity for the given fill row
'' Inputs:      Fill Row
'' Returns:     Associated Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AssociatedQuantityForFillRow(ByVal lFillRow As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lRow As Long                    ' Child row for the given fill row
    Dim AssociatedFill As cBrokerMessage ' Associated fill from the grid
    
    lReturn = 0&
    With fgFills
        lRow = .GetNodeRow(lFillRow, flexNTFirstChild)
        Do While lRow > -1&
            If IsClickHereRow(lRow) = False Then
                Set AssociatedFill = .RowData(lRow)
                lReturn = lReturn + CLng(Val(AssociatedFill("AssociatedQuantity")))
            End If
            
            lRow = .GetNodeRow(lRow, flexNTNextSibling)
        Loop
    End With
    
    AssociatedQuantityForFillRow = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociatedLotItems.AssociatedQuantityForFillRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewAssociatedFillRow
'' Description: Determine the row to put a new associated fill
'' Inputs:      Fill Row
'' Returns:     Associated Fill Row
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewAssociatedFillRow(ByVal lFillRow As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lLastChild As Long              ' Last child row for the fill row

    With fgFills
        lLastChild = .GetNodeRow(lFillRow, flexNTLastChild)
        If lLastChild = -1& Then
            lReturn = lFillRow + 1&
        ElseIf IsClickHereRow(lLastChild) Then
            lReturn = lLastChild
        Else
            lReturn = lLastChild + 1&
        End If
    End With
    
    NewAssociatedFillRow = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.NewAssociatedFillRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereRowForFill
'' Description: Determine the "Click Here" row in the grid for the given fill
'' Inputs:      Fill Row
'' Returns:     Click Here Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereRowForFill(ByVal lFillRow As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lLastChild As Long              ' Last child for the fill row
    
    lReturn = -1&
    With fgFills
        lLastChild = .GetNodeRow(lFillRow, flexNTLastChild)
        If lLastChild > -1& Then
            If IsClickHereRow(lLastChild) Then
                lReturn = lLastChild
            End If
        End If
    End With
    
    ClickHereRowForFill = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.ClickHereRowForFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateFillInfo
'' Description: Take care of a fill
'' Inputs:      Fill Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateFillInfo(ByVal lFillRow As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw settings
    Dim lFillQuantity As Long           ' Fill quantity
    Dim lAssociatedQuantity As Long     ' Associated fill quantity
    Dim lClickHereRow As Long           ' "Click Here" row for the fill

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lFillQuantity = CLng(Val(.TextMatrix(lFillRow, FillCol(eGDFillCols_Quantity))))
        lAssociatedQuantity = AssociatedQuantityForFillRow(lFillRow)
        lClickHereRow = ClickHereRowForFill(lFillRow)
        
        .TextMatrix(lFillRow, FillCol(eGDFillCols_AssignedQuantity)) = Str(lAssociatedQuantity)
        
        If lAssociatedQuantity = lFillQuantity Then
            If lClickHereRow > -1& Then
                .RemoveItem lClickHereRow
            End If
        Else
            If lClickHereRow = -1& Then
                lClickHereRow = NewAssociatedFillRow(lFillRow)
                
                .Rows = .Rows + 1
                If lClickHereRow <> .Rows - 1 Then
                    .RowPosition(.Rows - 1) = lClickHereRow
                End If
                
                .MergeRow(lClickHereRow) = True
                .RowOutlineLevel(lClickHereRow) = 2
                .IsSubtotal(lClickHereRow) = True
                .Cell(flexcpText, lClickHereRow, 0, lClickHereRow, FillCol(eGDFillCols_Time)) = "Click here to associate fill"
                .Cell(flexcpForeColor, lClickHereRow, 0, lClickHereRow, FillCol(eGDFillCols_Time)) = vbBlue
                .Cell(flexcpFontUnderline, lClickHereRow, 0, lClickHereRow, FillCol(eGDFillCols_Time)) = True
                
                .TextMatrix(lClickHereRow, FillCol(eGDFillCols_TurnkeyFillId)) = "CLICK"
            End If
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.CalculateFillInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewAssociation
'' Description: Allow the user to create a new fill association
'' Inputs:      Fill Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewAssociation(ByVal lFillRow As Long)
On Error GoTo ErrSection:

    Dim Fill As cBrokerMessage          ' Fill object
    Dim AssociatedFill As cBrokerMessage ' Associated fill
    Dim strFeedYardLotID As String      ' Feedyard lot ID
    Dim AssociatedFills As cGdTree      ' Associated fills

    If cboLots.ItemData(cboLots.ListIndex) <> -1& Then
        strFeedYardLotID = Str(cboLots.ItemData(cboLots.ListIndex))
    Else
        strFeedYardLotID = ""
    End If

    Set Fill = fgFills.RowData(lFillRow)
    Set AssociatedFills = AssociatedFillsForFillRow(lFillRow)
    
    If frmAssignFillToLot.ShowMe(Fill, Nothing, AssociatedFills, strFeedYardLotID) Then
        Set AssociatedFill = Fill.MakeCopy
        
        AssociatedFill("ID") = ""
        AssociatedFill.Add "FillID", Fill("ID")
        AssociatedFill.Add "AssociatedQuantity", Str(frmAssignFillToLot.AssignedQuantity)
        AssociatedFill.Add "FeedYardLotID", frmAssignFillToLot.FeedYardLotID
        AssociatedFill.Add "Changed", "1"
        
        AddAssociatedFill AssociatedFill
        CalculateFillInfo lFillRow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.NewAssociation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditAssociation
'' Description: Allow the user to edit an existing fill association
'' Inputs:      Associated Fill Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditAssociation(ByVal lAssociatedFillRow As Long)
On Error GoTo ErrSection:

    Dim lFillRow As Long                ' Fill row
    Dim Fill As cBrokerMessage          ' Fill object
    Dim AssociatedFill As cBrokerMessage ' Associated fill
    Dim AssociatedFills As cGdTree      ' Associated fills
    Dim lRow As Long                    ' Row in the grid
    Dim bChanged As Boolean             ' Did the information change?
    
    With fgFills
        lFillRow = .GetNodeRow(lAssociatedFillRow, flexNTParent)
        If lFillRow > -1& Then
            Set Fill = .RowData(lFillRow)
            Set AssociatedFill = .RowData(lAssociatedFillRow)
            Set AssociatedFills = AssociatedFillsForFillRow(lFillRow)
            
            If frmAssignFillToLot.ShowMe(Fill, AssociatedFill, AssociatedFills) Then
                bChanged = False
                If AssociatedFill("AssociatedQuantity") <> Str(frmAssignFillToLot.AssignedQuantity) Then
                    AssociatedFill("AssociatedQuantity") = Str(frmAssignFillToLot.AssignedQuantity)
                    bChanged = True
                End If
                If AssociatedFill("FeedYardLotID") <> frmAssignFillToLot.FeedYardLotID Then
                    AssociatedFill("FeedYardLotID") = frmAssignFillToLot.FeedYardLotID
                    bChanged = True
                End If
                
                If bChanged Then
                    AssociatedFill.Add "Changed", "1"
                    AssociatedFillToGrid lAssociatedFillRow, AssociatedFill
                    CalculateFillInfo lFillRow
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.EditAssociation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAssociation
'' Description: Allow the user to remove an existing fill association
'' Inputs:      Associated Fill Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveAssociation(ByVal lAssociatedFillRow As Long)
On Error GoTo ErrSection:

    Dim lFillRow As Long                ' Fill row
    Dim Fill As cBrokerMessage          ' Fill object
    Dim AssociatedFill As cBrokerMessage ' Associated fill object
    Dim strLotDisplay As String         ' Lot display that the fill is associated with
    Dim strMessage As String            ' Message to display

    With fgFills
        lFillRow = .GetNodeRow(lAssociatedFillRow, flexNTParent)
        If lFillRow > -1& Then
            Set Fill = .RowData(lFillRow)
            Set AssociatedFill = .RowData(lAssociatedFillRow)
            strLotDisplay = g.Cattle.LotDisplayForID(AssociatedFill("FeedYardLotID"))
            strMessage = "Fill '" & Fill("BrokerFillID") & "' will no longer be associated|with Lot #" & strLotDisplay & "||Do you want to continue?|"
            
            If InfBox(strMessage, "?", "+Yes|-No", "Confirmation") = "Y" Then
                .RemoveItem lAssociatedFillRow
                CalculateFillInfo lFillRow
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssociateLotItems.RemoveAssociation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociatedFillsForFillRow
'' Description: Get associated fills for the given fill row
'' Inputs:      Fill Row
'' Returns:     Associated Fills
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AssociatedFillsForFillRow(ByVal lFillRow As Long) As cGdTree
On Error GoTo ErrSection:

    Dim AssociatedFills As cGdTree      ' Collection of associated fills for the fill row
    Dim lRow As Long                    ' Row in the grid
    
    Set AssociatedFills = New cGdTree
    lRow = fgFills.GetNodeRow(lFillRow, flexNTFirstChild)
    Do While lRow > -1&
        If IsClickHereRow(lRow) = False Then
            AssociatedFills.Add fgFills.RowData(lRow)
        End If
        
        lRow = fgFills.GetNodeRow(lRow, flexNTNextSibling)
    Loop
    
    Set AssociatedFillsForFillRow = AssociatedFills

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssociateLotItems.AssociatedFillsForFillRow"
    
End Function

