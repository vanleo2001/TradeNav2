VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmBrokerView 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRefresh 
      Left            =   7200
      Top             =   5940
   End
   Begin VB.Timer tmrRealTime 
      Left            =   6720
      Top             =   5940
   End
   Begin VB.Timer tmrMenu 
      Left            =   6240
      Top             =   5940
   End
   Begin HexUniControls.ctlUniFrameWL fraConnection 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
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
      Caption         =   "frmBrokerView.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerView.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerView.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDisconnect 
         Height          =   315
         Left            =   2820
         TabIndex        =   3
         Top             =   0
         Width           =   1155
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
         Caption         =   "frmBrokerView.frx":005C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerView.frx":0092
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":00B2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdConnect 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   0
         Width           =   1155
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
         Caption         =   "frmBrokerView.frx":00CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerView.frx":00FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":011E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStatus 
         Height          =   195
         Left            =   300
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
         Caption         =   "frmBrokerView.frx":013A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerView.frx":0174
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":0194
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Image imgStatus 
         Height          =   195
         Left            =   60
         Picture         =   "frmBrokerView.frx":01B0
         Top             =   60
         Width           =   195
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrders 
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   2580
      Width           =   6015
      _cx             =   10610
      _cy             =   1720
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
      TabIndex        =   4
      Top             =   600
      Width           =   8415
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
      Caption         =   "frmBrokerView.frx":0436
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBrokerView.frx":0456
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerView.frx":0476
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPriceLadder 
         Height          =   315
         Left            =   6960
         TabIndex        =   9
         Top             =   0
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
         Caption         =   "frmBrokerView.frx":0492
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerView.frx":04CC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":04EC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookupAccount 
         Height          =   315
         Left            =   5460
         TabIndex        =   8
         Top             =   0
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
         Caption         =   "frmBrokerView.frx":0508
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerView.frx":0546
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":0566
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRefreshAccount 
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   0
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
         Caption         =   "frmBrokerView.frx":0582
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmBrokerView.frx":05C2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":05E2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         Top             =   0
         Width           =   3015
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
         Tip             =   "frmBrokerView.frx":05FE
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":061E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAccounts 
         Height          =   195
         Left            =   0
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
         Caption         =   "frmBrokerView.frx":063A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBrokerView.frx":066C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBrokerView.frx":068C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgFills 
      Height          =   975
      Left            =   60
      TabIndex        =   5
      Top             =   3960
      Width           =   6015
      _cx             =   10610
      _cy             =   1720
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
      Height          =   975
      Left            =   60
      TabIndex        =   10
      Top             =   5340
      Width           =   6015
      _cx             =   10610
      _cy             =   1720
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
   Begin VSFlex7LCtl.VSFlexGrid fgAccountDetails 
      Height          =   975
      Left            =   60
      TabIndex        =   11
      Top             =   1305
      Width           =   6015
      _cx             =   10610
      _cy             =   1720
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
   Begin HexUniControls.ctlUniLabelXP lblAccountDetails 
      Height          =   195
      Left            =   60
      Top             =   1080
      Width           =   1335
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
      Caption         =   "frmBrokerView.frx":06A8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerView.frx":06E8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerView.frx":0708
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPositions 
      Height          =   195
      Left            =   60
      Top             =   5100
      Width           =   3555
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
      Caption         =   "frmBrokerView.frx":0724
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerView.frx":078C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerView.frx":07AC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblFills 
      Height          =   195
      Left            =   60
      Top             =   3720
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
      Caption         =   "frmBrokerView.frx":07C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerView.frx":07F4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerView.frx":0814
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblOrders 
      Height          =   195
      Left            =   60
      Top             =   2355
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
      Caption         =   "frmBrokerView.frx":0830
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerView.frx":085E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerView.frx":087E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "Orders"
      Begin VB.Menu mnuOrdersCreate 
         Caption         =   "Create New Order"
      End
      Begin VB.Menu mnuOrdersModify 
         Caption         =   "Modify Order"
      End
      Begin VB.Menu mnuOrdersCancel 
         Caption         =   "Cancel Order"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdersPriceLadder 
         Caption         =   "Price Ladder"
      End
   End
   Begin VB.Menu mnuFills 
      Caption         =   "Fills"
      Begin VB.Menu mnuFillsPriceLadder 
         Caption         =   "Price Ladder"
      End
   End
   Begin VB.Menu mnuPositions 
      Caption         =   "Positions"
      Begin VB.Menu mnuPositionsFlatten 
         Caption         =   "Flatten Position"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPositionsPriceLadder 
         Caption         =   "Price Ladder"
      End
   End
End
Attribute VB_Name = "frmBrokerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBrokerView.frm
'' Description: User interface for a broker view on the data coming from the broker
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 04/11/2012   DAJ         Added the FCM Account Number
'' 04/12/2012   DAJ         Added Genesis Symbols
'' 04/20/2012   DAJ         Mods for communicating with price ladder
'' 04/25/2012   DAJ         Removed broker from account lookup call, raise error on "events"
'' 05/01/2012   DAJ         Realtime equity updates, account details, price ladder
'' 05/03/2012   DAJ         Remove Flatten from grid, Account sync with ladder,
''                          Price Ladder button, Lookup Account bug
'' 05/10/2012   DAJ         Additional logging, disable controls during refresh
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/11/2012   DAJ         Make Turnkey work with all brokers, Show Account in Simple Order dialog
'' 06/12/2012   DAJ         Make sure to update trades on unsolicited fill
'' 09/11/2012   DAJ         Collections of refresh items, Refresh of non-selected account
'' 09/27/2012   DAJ         Renamed Ladder_ChangeAccount to ChangeAccount, Exposed LookupAccount
'' 10/23/2012   DAJ         Before bringing up Simple Order form, verify an account is selected
'' 01/09/2013   DAJ         Optionally take side in the CreateOrderForLot call
'' 01/09/2013   DAJ         Exposed the ShowLadder call
'' 01/31/2013   DAJ         Expose m.Accounts collection
'' 06/24/2013   DAJ         Timer Logging
'' 10/24/2013   DAJ         Pass account number to g.Profit.Profit
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDOrderCols
    eGDOrderCols_BrokerID = 0
    eGDOrderCols_Side
    eGDOrderCols_Quantity
    eGDOrderCols_Symbol
    eGDOrderCols_GenesisSymbol
    eGDOrderCols_LimitPrice
    eGDOrderCols_StopPrice
    eGDOrderCols_Type
    eGDOrderCols_TIF
    eGDOrderCols_Status
    eGDOrderCols_Modify
    eGDOrderCols_Cancel
    eGDOrderCols_Blank
    eGDOrderCols_NumCols
End Enum

Private Enum eGDFillCols
    eGDFillCols_BrokerSymbol = 0
    eGDFillCols_GenesisSymbol
    eGDFillCols_EntryTime
    eGDFillCols_EntryOrderID
    eGDFillCols_EntryFillID
    eGDFillCols_EntrySide
    eGDFillCols_EntryQuantity
    eGDFillCols_EntryPrice
    eGDFillCols_ExitTime
    eGDFillCols_ExitOrderID
    eGDFillCols_ExitFillID
    eGDFillCols_ExitSide
    eGDFillCols_ExitQuantity
    eGDFillCols_ExitPrice
    eGDFillCols_OpenEquity
    eGDFillCols_Reserved
    eGDFillCols_ClosedProfit
    eGDFillCols_Blank
    eGDFillCols_NumCols
End Enum

Private Enum eGDPositionCols
    eGDPositionCols_Symbol = 0
    eGDPositionCols_GenesisSymbol
    eGDPositionCols_Carried
    eGDPositionCols_Buys
    eGDPositionCols_Sells
    eGDPositionCols_Current
    eGDPositionCols_Flatten
    eGDPositionCols_EntryPrice
    eGDPositionCols_CurrentPrice
    eGDPositionCols_Profit
    eGDPositionCols_Blank
    eGDPositionCols_NumCols
End Enum

Private Type mPrivate
    nBroker As eTT_AccountType          ' Broker that we are working with
    BrokerObj As cBroker                ' Broker object for the given broker
    nStatus As eGDConnectionStatus      ' Connection status to the broker
    strCaption As String                ' Form caption
    strAccount As String                ' Currently selected account
    strFcmAccount As String             ' FCM Account number for the selected account
    bSetAccountFromCode As Boolean      ' Have we set the account combo from code?
    dRefreshStart As Double             ' Time the refresh was started
    
    strRefreshAccount As String         ' Account refresh is being called for ( if not selected account )
    strRefreshFcmAccount As String      ' FCM Account number for the refresh account
    bRefreshDone As Boolean             ' Is the requested refresh done?
    bRefreshTimedOut As Boolean         ' Did the requested refresh time out?
    RefreshOrders As cGdTree            ' Collection of orders for the requested refresh
    RefreshFills As cGdTree             ' Collection of fills for the requested refresh
    RefreshCarriedFills As cGdTree      ' Collection of carried fills for the requested refresh
    
    Accounts As cGdTree                 ' Collection of accounts
    astrAccounts As cGdArray            ' List of accounts
    NumBuysToday As cGdTree             ' Number of buys per symbol
    NumSellsToday As cGdTree            ' Number of sells per symbol
    
    CarriedPositions As cGdTree         ' Collection of carried positions
    CarriedFills As cGdTree             ' Collection of carried fills
    TodaysFills As cGdTree              ' Collection of fills done today
    PositionFills As cGdTree            ' Collection of fills that make up the position
    Positions As cGdTree                ' Collection of positions
    
    Trades As cGdTree                   ' Collection of trades
    
    BarProps As cGdTree                 ' Collection of Bars properties
    BarsColl As cGdTree                 ' Collection of Bars for Real Time
    frmLadder As frmTickDistribution    ' Price ladder form
    
    dPrevBalance As Double              ' Previous balance
    dTotalClosedProfit As Double        ' Total closed profit for the account
    dAccountBalance As Double           ' Account Balance (Previous Balance + Closed Profit)
    dTotalOpenEquity As Double          ' Total open equity for the account
    dNetLiquidity As Double             ' Net Liquidity Value (Account Balance + Open Equity)
End Type
Private m As mPrivate

Public Property Get Account() As String
    Account = m.strFcmAccount
End Property

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccount
End Property

Public Property Get Accounts() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnAccounts As cGdTree       ' Collection of accounts to return

    Set ReturnAccounts = Nothing
    Select Case m.nBroker
        Case eTT_AccountType_RjoCqg
            Set ReturnAccounts = g.RjoCqg.Accounts
            
    End Select
    
    Set Accounts = ReturnAccounts

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmBrokerView.Accounts.Get"

End Property

Public Property Get AllAccounts() As cGdTree
    Set AllAccounts = m.Accounts
End Property

Public Property Get Orders() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnOrders As cGdTree         ' Collection of orders to return

    Set ReturnOrders = Nothing
    Select Case m.nBroker
        Case eTT_AccountType_RjoCqg
            Set ReturnOrders = g.RjoCqg.Orders
            
    End Select
    
    Set Orders = ReturnOrders

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmBrokerView.Orders.Get"

End Property

Public Property Get Fills() As cGdTree
On Error GoTo ErrSection:

    Dim ReturnFills As cGdTree          ' Collection of fills to return

    Set ReturnFills = Nothing
    Select Case m.nBroker
        Case eTT_AccountType_RjoCqg
            Set ReturnFills = g.RjoCqg.Fills
            
    End Select
    
    Set Fills = ReturnFills

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmBrokerView.Fills.Get"

End Property

Public Property Get CarriedFills() As cGdTree
    Set CarriedFills = m.CarriedFills
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get ConnectionStatus() As eGDConnectionStatus
    ConnectionStatus = m.nStatus
End Property

Public Property Get RefreshDone() As Boolean
    RefreshDone = m.bRefreshDone
End Property

Public Property Get RefreshTimedOut() As Boolean
    RefreshTimedOut = m.bRefreshTimedOut
End Property

Public Property Get RefreshOrders() As cGdTree
    Set RefreshOrders = m.RefreshOrders
End Property

Public Property Get RefreshFills() As cGdTree
    Set RefreshFills = m.RefreshFills
End Property

Public Property Get RefreshCarriedFills() As cGdTree
    Set RefreshCarriedFills = m.RefreshCarriedFills
End Property

Private Property Get OrderCol(ByVal nCol As eGDOrderCols) As Long
    OrderCol = nCol
End Property

Private Property Get FillCol(ByVal nCol As eGDFillCols) As Long
    FillCol = nCol
End Property

Private Property Get PositionCol(ByVal nCol As eGDPositionCols) As Long
    PositionCol = nCol
End Property

Private Property Get RefreshStart() As Double
    RefreshStart = m.dRefreshStart
End Property
Private Property Let RefreshStart(ByVal dRefreshStart As Double)
On Error GoTo ErrSection:

    m.dRefreshStart = dRefreshStart
    If dRefreshStart = 0& Then
        tmrRefresh.Enabled = False
        EnableControls
        InfBox ""
        m.bRefreshDone = True
    Else
        tmrRefresh.Interval = 500
        tmrRefresh.Enabled = True
        EnableControls
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmBrokerView.RefreshStart.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    m.nBroker = nBroker
    Set m.BrokerObj = g.Broker.Broker(nBroker)
    If m.BrokerObj Is Nothing Then
        m.strCaption = "Broker View"
        'SetConnectionStatus eGDConnectionStatus_Disconnected
    Else
        m.strCaption = "Broker View for " & m.BrokerObj.BrokerName
        'SetConnectionStatus m.BrokerObj.ConnectionStatus
    End If
    Caption = m.strCaption
    
    InitAccountDetailsGrid
    InitOrdersGrid
    InitFillsGrid
    InitPositionsGrid
    
    EnableControls
    
    tmrRealTime.Interval = frmQuotes.tmrRealTime.Interval
    tmrRealTime.Enabled = g.RealTime.Active
    
    ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_ConnectionStatus
'' Description: The connection status to the broker has changed
'' Inputs:      New Connection Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_ConnectionStatus(ByVal nNewStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    SetConnectionStatus nNewStatus

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_ConnectionStatus", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_Account
'' Description: An account has been received from the broker
'' Inputs:      Account Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_Account(ByVal strAccountInfo As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim strAccount As String            ' Account information
    Dim lPos As Long                    ' Position of account in the array
    
    strFirstField = Parse(strAccountInfo, vbTab, 1)
    
    Select Case UCase(strFirstField)
        Case "BEGIN"
            cboAccounts.Clear
            m.astrAccounts.Clear
            
        Case "END"
            InfBox ""
        
        Case Else
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strAccountInfo
            If Len(brokerMessage("FcmAccount")) = 0 Then
                brokerMessage.Add "FcmAccount", brokerMessage("Account")
            End If
            strAccount = brokerMessage("FcmAccount")
            
            If m.Accounts.Exists(strAccount) Then
                lPos = m.Accounts.Index(strAccount)
            Else
                lPos = m.Accounts.Add(brokerMessage, strAccount)
            End If
            
            If Len(brokerMessage("AccountName")) = 0 Then
                cboAccounts.AddItem strAccount
            Else
                cboAccounts.AddItem strAccount & " (" & brokerMessage("AccountName") & ")"
            End If
            cboAccounts.ItemData(cboAccounts.NewIndex) = lPos
            
            strAccount = strAccount & vbTab & vbTab & brokerMessage("AccountName")
            If m.astrAccounts.BinarySearch(strAccount, lPos) = False Then
                m.astrAccounts.Add strAccount, lPos
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_Account", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_AccountDetails
'' Description: An account details record has been received from the broker
'' Inputs:      Account Details
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_AccountDetails(ByVal strAccountDetails As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim Acct As cPtAccount              ' Account object
    
    strFirstField = Parse(strAccountDetails, vbTab, 1)
    
    Select Case UCase(strFirstField)
        Case "BEGIN"
            
        Case "END"
        
        Case Else
            g.RjoCqg.AccountFromInfo strAccountDetails, Acct
            
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strAccountDetails
            If brokerMessage("Account") = m.strAccount Then
                AccountDetailsToGrid brokerMessage
            End If
            
            RefreshStart = 0
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_AccountDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_Order
'' Description: An order has been received from the broker
'' Inputs:      Order Information, Coming in a refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_Order(ByVal strOrderInfo As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim Order As cPtOrder               ' Order object
    
    strFirstField = Parse(strOrderInfo, vbTab, 1)
    
    Select Case UCase(strFirstField)
        Case "BEGIN"
            m.RefreshOrders.Clear
            If Len(m.strRefreshAccount) = 0 Then
                fgOrders.Redraw = flexRDNone
                fgOrders.Rows = fgOrders.FixedRows
            End If
            
        Case "END"
            If Len(m.strRefreshAccount) = 0 Then
                AddClickHereRow
                fgOrders.AutoSize 0, fgOrders.Cols - 1, False, 75
                fgOrders.Redraw = flexRDBuffered
                
                InfBox "Requesting fills from the " & m.BrokerObj.BrokerName & " servers for account " & m.strFcmAccount & ".  Please wait...", , , "Fills", True
                GetFills m.strAccount
            Else
                InfBox "Requesting fills from the " & m.BrokerObj.BrokerName & " servers for account " & m.strRefreshFcmAccount & ".  Please wait...", , , "Fills", True
                GetFills m.strRefreshAccount
            End If
        
        Case Else
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strOrderInfo
            AddGenesisSymbolToOrder brokerMessage
            If bRefresh = True Then
                m.RefreshOrders.Add brokerMessage
            Else
                g.RjoCqg.OrderFromInfo strOrderInfo, Order
            End If
            
            If (Len(m.strRefreshAccount) = 0) Or (bRefresh = False) Then
                If brokerMessage("Account") = m.strAccount Then
                    OrderToGrid brokerMessage, bRefresh
                End If
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_Order", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_Fill
'' Description: A fill has been received from the broker
'' Inputs:      Fill Information, Coming in a refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_Fill(ByVal strFillInfo As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim lPos As Long                    ' Position in the array
    Dim Fill As cPtFill                 ' Fill object
    
    strFirstField = Parse(strFillInfo, vbTab, 1)
    
    Select Case UCase(strFirstField)
        Case "BEGIN"
            m.RefreshFills.Clear
            If Len(m.strRefreshAccount) = 0 Then
                m.NumBuysToday.Clear
                m.NumSellsToday.Clear
                m.TodaysFills.Clear
            End If
            
        Case "END"
            If Len(m.strRefreshAccount) = 0 Then
                fgFills.AutoSize 0, fgFills.Cols - 1, False, 75
                fgFills.Redraw = flexRDBuffered
                
                InfBox "Requesting positions from the " & m.BrokerObj.BrokerName & " servers for account " & m.strFcmAccount & ".  Please wait...", , , "Positions", True
                GetPositions m.strAccount
            Else
                InfBox "Requesting positions from the " & m.BrokerObj.BrokerName & " servers for account " & m.strRefreshFcmAccount & ".  Please wait...", , , "Positions", True
                GetPositions m.strRefreshAccount
            End If
        
        Case Else
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strFillInfo
            AddGenesisSymbolToFill brokerMessage
            
            If bRefresh = True Then
                m.RefreshFills.Add brokerMessage
            Else
                g.RjoCqg.FillFromInfo strFillInfo, Fill
            End If
        
            If (Len(m.strRefreshAccount) = 0) Or (bRefresh = False) Then
                If brokerMessage("Account") = m.strAccount Then
                    UpdateTodayNumbers brokerMessage
                    
                    m.TodaysFills.Add brokerMessage
                    If bRefresh = False Then
                        AddTodayFillToPositionFills brokerMessage, False
                        
                        AddTodayFillToTrades brokerMessage
                        TradesToGrid
                    End If
                End If
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_Fill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_CarriedFill
'' Description: A carried fill has been received from the broker
'' Inputs:      Fill Information, Coming in a refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_CarriedFill(ByVal strFillInfo As String, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim lPos As Long                    ' Position in the array
    Dim FillsForSymbol As cGdTree       ' Fills for the symbol
    Dim strSymbol As String             ' Symbol for the fill
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Trades As cGdTree               ' Collection of trades
    
    strFirstField = Parse(strFillInfo, vbTab, 1)
    
    Select Case UCase(strFirstField)
        Case "BEGIN"
            m.RefreshCarriedFills.Clear
            If Len(m.strRefreshAccount) = 0 Then
                m.CarriedFills.Clear
                m.Trades.Clear
            End If
            
        Case "END"
            If Len(m.strRefreshAccount) = 0 Then
                For lIndex = 1 To m.TodaysFills.Count
                    AddTodayFillToTrades m.TodaysFills(lIndex)
                Next lIndex
                
                TradesToGrid
            End If
        
        Case Else
            Set brokerMessage = New cBrokerMessage
            brokerMessage.FromString strFillInfo
            AddGenesisSymbolToFill brokerMessage
            
            If bRefresh = True Then
                m.RefreshCarriedFills.Add brokerMessage
            End If
        
            If (Len(m.strRefreshAccount) = 0) Or (bRefresh = False) Then
                If brokerMessage("Account") = m.strAccount Then
                    m.CarriedFills.Add brokerMessage
                    AddTodayFillToTrades brokerMessage
                End If
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_CarriedFill", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_Position
'' Description: A position has been received from the broker
'' Inputs:      Position Information, Coming in a refresh?, Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_Position(ByVal strPositionInfo As String, ByVal bRefresh As Boolean, Optional BrokerPosition As cPtPosition = Nothing)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim lPos As Long                    ' Position in the array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol for a position
    
    strFirstField = Parse(strPositionInfo, vbTab, 1)
    
    Select Case UCase(strFirstField)
        Case "BEGIN"
            If Len(m.strRefreshAccount) = 0 Then
                m.CarriedPositions.Clear
                m.Positions.Clear
                fgPositions.Redraw = flexRDNone
                fgPositions.Rows = fgPositions.FixedRows
                AddPositionsTotalRow
            End If
            
        Case "END"
            If Len(m.strRefreshAccount) = 0 Then
                BuildPositionFills
                
                For lIndex = 1 To m.PositionFills.Count
                    strSymbol = m.PositionFills.Key(lIndex)
                    CalcPositionFromPositionFills strSymbol
                    PositionToGrid m.Positions(strSymbol), bRefresh
                Next lIndex
                
                fgPositions.AutoSize 0, fgPositions.Cols - 1, False, 75
                fgPositions.Redraw = flexRDBuffered
                            
                If Not g.RjoCqg Is Nothing Then
                    g.RjoCqg.GetAccountDetails m.strAccount
                End If
            Else
                RefreshStart = 0
            End If
        
        Case Else
            If (Len(m.strRefreshAccount) = 0) Or (bRefresh = False) Then
                Set brokerMessage = New cBrokerMessage
                brokerMessage.FromString strPositionInfo
                If brokerMessage("Account") = m.strAccount Then
                    AddGenesisSymbolToPosition brokerMessage
                    m.CarriedPositions.Add brokerMessage, brokerMessage("Symbol")
                    
                    If bRefresh = False Then
                        CalcPositionFromPositionFills brokerMessage("Symbol")
                        PositionToGrid brokerMessage, bRefresh
                    End If
                End If
            End If
        
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Broker_Position", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Ladder_Unloaded
'' Description: Notification that the associated price ladder has been unloaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Ladder_Unloaded()
On Error GoTo ErrSection:

    Set m.frmLadder = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Ladder_Unloaded", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Ladder_CreateOrder
'' Description: Notification that the user wants to create an order from the ladder
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Ladder_CreateOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    CreateOrder Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Ladder_CreateOrder", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Ladder_ModifyOrder
'' Description: Notification that the user wants to modify an order from the ladder
'' Inputs:      Order, New Price, New Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Ladder_ModifyOrder(ByVal Order As cPtOrder, Optional ByVal dNewPrice As Double = 0#, Optional ByVal lNewQuantity As Long = 0&)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid
    
    lRow = RowForOrder(Order)
    If lRow > -1& Then
        ModifyOrder lRow, dNewPrice, lNewQuantity
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Ladder_ModifyOrder", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Ladder_CancelOrder
'' Description: Notification that the user wants to cancel an order from the ladder
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Ladder_CancelOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid
    
    lRow = RowForOrder(Order)
    If lRow > -1& Then
        CancelOrder lRow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Ladder_CancelOrder", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Ladder_FlattenPosition
'' Description: Notification that the user wants to flatten a position from the ladder
'' Inputs:      Genesis Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Ladder_FlattenPosition(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid
    
    With fgPositions
        For lRow = .FixedRows To .Rows - 1
            If .TextMatrix(lRow, PositionCol(eGDPositionCols_GenesisSymbol)) = strGenesisSymbol Then
                FlattenPosition lRow
                Exit For
            End If
        Next lRow
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Ladder_FlattenPosition", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeAccount
'' Description: Notification that the user changed the account outside of the form
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeAccount(ByVal lAccountID As Long, ByVal strSource As String)
On Error GoTo ErrSection:

    If SelectAccountID(lAccountID) Then
        AddToLog "User changed to account '" & cboAccounts.Text & "' from '" & strSource & "'"
        RefreshSelectedAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ChangeAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionStringForSymbol
'' Description: Build a position string for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     Position String
''
'' Fields:      Direction|Quantity|Open Equity|Average Entry Display|Session
''              Quantity|Session Profit|Average Entry Actual
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionStringForSymbol(ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim Position As cBrokerMessage      ' Position for the given symbol
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 7
    astrReturn(0) = "Flat"
    
    With fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                Set Position = .RowData(lIndex)
                If Position("GenesisSymbol") = strGenesisSymbol Then
                    astrReturn(0) = Position("Direction")
                    astrReturn(1) = Position("Quantity")
                    astrReturn(2) = .TextMatrix(lIndex, PositionCol(eGDPositionCols_Profit))
                    astrReturn(3) = .TextMatrix(lIndex, PositionCol(eGDPositionCols_EntryPrice))
                    astrReturn(4) = ""
                    astrReturn(5) = ""
                    astrReturn(6) = Position("AverageEntry")
                    
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    PositionStringForSymbol = astrReturn.JoinFields("|")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.PositionStringForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WorkingOrdersForSymbol
'' Description: Build a collection of working orders for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     Collection of Working Orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WorkingOrdersForSymbol(ByVal strGenesisSymbol As String) As cGdTree
On Error GoTo ErrSection:

    Dim ReturnOrders As cGdTree         ' Collection of working orders to return
    Dim lIndex As Long                  ' Index into a for loop
    Dim orderMessage As cBrokerMessage  ' Order from the grid
    Dim Order As cPtOrder               ' Order object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim BrokerKv As cBrokerKeyValue     ' Routines for key-value broker messages
    
    Set ReturnOrders = New cGdTree
    Set BrokerKv = New cBrokerKeyValue
    
    strBrokerSymbol = g.RjoCqg.BrokerSymbol(strGenesisSymbol, BrokerSym)
    If Len(strBrokerSymbol) > 0 Then
        For lIndex = fgOrders.FixedRows To fgOrders.Rows - 1
            If TypeOf fgOrders.RowData(lIndex) Is cBrokerMessage Then
                Set orderMessage = fgOrders.RowData(lIndex)
                If orderMessage("Symbol") = strBrokerSymbol Then
                    If IsOpenOrder(orderMessage("Status")) Then
                        Set Order = BrokerKv.OrderFromMessage(orderMessage, m.BrokerObj, strGenesisSymbol, BrokerSym, 0)
                        Order.OrderID = lIndex
                        ReturnOrders.Add Order, Str(Order.OrderID)
                    End If
                End If
            End If
        Next lIndex
    End If
    
    Set WorkingOrdersForSymbol = ReturnOrders

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.WorkingOrdersForSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrderForLot
'' Description: Allow the user to create an order for the given lot
'' Inputs:      Feed Yard Lot ID, Symbol, Side
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateOrderForLot(ByVal strFeedYardLotID As String, Optional ByVal strSymbol As String = "", Optional ByVal strSide As String = "")
On Error GoTo ErrSection:

    CreateOrder Nothing, strSymbol, strFeedYardLotID, strSide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CreateOrderForLot"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasAccountNumber
'' Description: Is the given account number in the accounts collection?
'' Inputs:      Account Number, Account
'' Returns:     True if in accounts collection, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasAccountNumber(ByVal strAccountNumber As String, Optional Account As cBrokerMessage = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.Accounts.Count
        Set Account = m.Accounts(lIndex)
        If Account("Account") = strAccountNumber Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasAccountNumber = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.HasAccountNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccount
'' Description: Refresh the given account numer
'' Inputs:      Account Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshAccount(ByVal strAccountNumber As String)
On Error GoTo ErrSection:

    Dim strSelectedAccount As String    ' Selected account number
    Dim Account As cBrokerMessage       ' Account object
    
    strSelectedAccount = SelectedAccountNumber
    If strAccountNumber = strSelectedAccount Then
        RefreshSelectedAccount
    ElseIf HasAccountNumber(strAccountNumber, Account) Then
        m.strRefreshAccount = strAccountNumber
        m.strRefreshFcmAccount = Account("FcmAccount")
        
        m.bRefreshDone = False
        m.bRefreshTimedOut = False
        RefreshStart = gdTickCount
        
        InfBox "Requesting orders from the " & m.BrokerObj.BrokerName & " servers for account " & m.strRefreshFcmAccount & ".  Please wait...", , , "Orders", True
        GetOrders m.strRefreshAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.RefreshAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupAccount
'' Description: Lookup an account
'' Inputs:      None
'' Returns:     Account ID chosen ( -1 if Cancelled )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LookupAccount() As Long
On Error GoTo ErrSection:

    Dim strFcmAccount As String         ' FCM Account chosen
    Dim lReturn As Long                 ' Return value for the function

    lReturn = -1&
    
    strFcmAccount = frmAccountLookup.ShowMe(m.astrAccounts)
    If Len(strFcmAccount) > 0 Then
        SelectFcmAccount strFcmAccount
        AddToLog "User changed to account '" & cboAccounts.Text & "' with the lookup button on broker view"
        RefreshSelectedAccount
        
        lReturn = cboAccounts.ItemData(cboAccounts.ListIndex)
    End If
    
    LookupAccount = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.LookupAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLadder
'' Description: Show the price ladder for the given Genesis symbol
'' Inputs:      Genesis Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowLadder(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID for the given Genesis symbol
    Dim astrSymbols As cGdArray         ' Symbols array back from the symbol selector
    
    If Len(strGenesisSymbol) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", False, , "Select a symbol for the price ladder")
        If astrSymbols.Size > 0 Then
            strGenesisSymbol = astrSymbols(0)
        End If
    End If
    
    If Len(strGenesisSymbol) > 0 Then
        lSymbolID = GetSymbolID(strGenesisSymbol)
    
        If m.frmLadder Is Nothing Then
            Set m.frmLadder = New frmTickDistribution
            m.frmLadder.ShowMe lSymbolID, 0, , Me
        Else
            m.frmLadder.ChangeSymbol lSymbolID
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ShowLadder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: Handle the user changing accounts in the combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If Visible Then
        If m.bSetAccountFromCode = False Then
            AddToLog "User changed to account '" & cboAccounts.Text & "' by changing combo on broker view"
        End If
        
        RefreshSelectedAccount
    End If
    
    m.bSetAccountFromCode = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.cboAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdConnect_Click
'' Description: Allow the user to connect to the broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdConnect_Click()
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        m.BrokerObj.Connect
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.cmdConnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDisconnect_Click
'' Description: Allow the user to disconnect to the broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDisconnect_Click()
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        m.BrokerObj.Disconnect
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.cmdDisconnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPriceLadder_Click
'' Description: Allow the user to bring up the price ladder
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPriceLadder_Click()
On Error GoTo ErrSection:

    ShowLadder ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.cmdPriceLadder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRefreshAccount_Click
'' Description: Load the information for the selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRefreshAccount_Click()
On Error GoTo ErrSection:

    RefreshSelectedAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.cmdRefreshAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookupAccount_Click
'' Description: Allow the user to lookup an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookupAccount_Click()
On Error GoTo ErrSection:

    LookupAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.cmdLookupAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Column that the mouse is on
    Dim lMouseRow As Long               ' Row that the mouse is on
    
    With fgFills
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        .Row = lMouseRow
        
        If Button = vbRightButton Then
            PopupMenu mnuFills
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.fgFills_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFills_DblClick
'' Description: Handle the user double clicking on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFills_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid the user double clicked
    
    lMouseRow = fgFills.MouseRow
    ShowLadderForFillRow lMouseRow, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.fgFills_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Column that the mouse is on
    Dim lMouseRow As Long               ' Row that the mouse is on
    Dim bEnable As Boolean              ' Enable the menu items?
    
    With fgOrders
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        .Row = lMouseRow
        
        If Button = vbRightButton Then
            bEnable = False
            If IsValidRow(fgOrders, lMouseRow) Then
                bEnable = IsOpenOrder(.TextMatrix(lMouseRow, OrderCol(eGDOrderCols_Status)))
            End If
            
            Enable mnuOrdersModify, bEnable
            Enable mnuOrdersCancel, bEnable
            
            PopupMenu mnuOrders
        ElseIf Button = vbLeftButton Then
            If IsValidRow(fgOrders, lMouseRow) Then
                If .Cell(flexcpForeColor, lMouseRow, lMouseCol) = vbBlue Then
                    If .MergeRow(lMouseRow) = True Then
                        CreateOrder
                    ElseIf lMouseCol = OrderCol(eGDOrderCols_Modify) Then
                        ModifyOrder lMouseRow
                    ElseIf lMouseCol = OrderCol(eGDOrderCols_Cancel) Then
                        CancelOrder lMouseRow
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.fgOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_DblClick
'' Description: Handle the user double clicking on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid the user double clicked
    
    lMouseRow = fgOrders.MouseRow
    ShowLadderForOrderRow lMouseRow, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.fgOrders_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in a cell
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel the click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Column that the mouse is on
    Dim lMouseRow As Long               ' Row that the mouse is on
    Dim bEnable As Boolean              ' Enable the menu items?
    
    With fgPositions
        lMouseCol = .MouseCol
        lMouseRow = .MouseRow
        
        .Row = lMouseRow
        
        If Button = vbRightButton Then
            bEnable = False
            If IsValidRow(fgPositions, lMouseRow) Then
                bEnable = (UCase(.TextMatrix(lMouseRow, PositionCol(eGDPositionCols_Current))) <> "FLAT")
            End If
            
            Enable mnuPositionsFlatten, bEnable
            
            PopupMenu mnuPositions
        ElseIf Button = vbLeftButton Then
            If IsValidRow(fgPositions, lMouseRow) Then
                If lMouseCol = PositionCol(eGDPositionCols_Flatten) Then
                    If .Cell(flexcpForeColor, lMouseRow, lMouseCol) = vbBlue Then
                        FlattenPosition lMouseRow
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.fgPositions_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_DblClick
'' Description: Handle the user double clicking on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid the user double clicked
    
    lMouseRow = fgPositions.MouseRow
    ShowLadderForPositionsRow lMouseRow, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.fgPositions_DblClick"
    
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

    Static bAlreadyDone As Boolean      ' Have we already done something?

    If cboAccounts.ListCount = 0 Then
        If m.BrokerObj Is Nothing Then
            SetConnectionStatus eGDConnectionStatus_Disconnected
        ElseIf (bAlreadyDone = False) And (m.BrokerObj.ConnectionStatus = eGDConnectionStatus_Disconnected) Then
            bAlreadyDone = True
            m.BrokerObj.Connect
        Else
            SetConnectionStatus m.BrokerObj.ConnectionStatus
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement for the form
    
    strPlacement = GetIniFileProperty("frmBrokerView", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    g.Styler.StyleForm Me
    
    Icon = Picture16("kBlank")
    
    m.bSetAccountFromCode = False
    
    Set m.Accounts = New cGdTree
    Set m.astrAccounts = New cGdArray
    m.astrAccounts.Create eGDARRAY_Strings
    Set m.NumBuysToday = New cGdTree
    Set m.NumSellsToday = New cGdTree
    
    Set m.CarriedPositions = New cGdTree
    Set m.CarriedFills = New cGdTree
    Set m.TodaysFills = New cGdTree
    Set m.PositionFills = New cGdTree
    Set m.Positions = New cGdTree
    Set m.Trades = New cGdTree
    Set m.BarProps = New cGdTree
    Set m.BarsColl = New cGdTree
    
    m.strRefreshAccount = ""
    m.strRefreshFcmAccount = ""
    Set m.RefreshOrders = New cGdTree
    Set m.RefreshFills = New cGdTree
    Set m.RefreshCarriedFills = New cGdTree
    
    Set m.frmLadder = Nothing
    
    mnuOrders.Visible = False
    mnuFills.Visible = False
    mnuPositions.Visible = False
    
    tmrMenu.Enabled = False
    tmrRefresh.Enabled = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Form_Load"
    
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
    Dim lMinScaleHeight As Long         ' Minimum scale height for the form
    Dim lTotalGridHeight As Long        ' Total height available for grids
    Dim lSeparation As Long             ' Separation between controls
    
    lSeparation = 120
    lMinScaleWidth = fraAccounts.Width + (lSeparation * 2)
    lMinScaleHeight = 6480 ' 5160
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        lTotalGridHeight = ScaleHeight - fraConnection.Height - fraAccounts.Height - lblAccountDetails.Height - lblOrders.Height - lblFills.Height - lblPositions.Height - (lSeparation * 7)
        
        With fgAccountDetails
            .Move lSeparation, .Top, ScaleWidth - (lSeparation * 2), lTotalGridHeight / 4
        End With
        
        With lblOrders
            .Move lSeparation, fgAccountDetails.Top + fgAccountDetails.Height + lSeparation
        End With
        With fgOrders
            .Move lSeparation, lblOrders.Top + lblOrders.Height, ScaleWidth - (lSeparation * 2), lTotalGridHeight / 4
        End With
        
        With lblFills
            .Move lSeparation, fgOrders.Top + fgOrders.Height + lSeparation
        End With
        With fgFills
            .Move lSeparation, lblFills.Top + lblFills.Height, ScaleWidth - (lSeparation * 2), lTotalGridHeight / 4
        End With
        
        With lblPositions
            .Move lSeparation, fgFills.Top + fgFills.Height + lSeparation
        End With
        With fgPositions
            .Move lSeparation, lblPositions.Top + lblPositions.Height, ScaleWidth - (lSeparation * 2), lTotalGridHeight / 4
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    SetIniFileProperty "frmBrokerView", GetFormPlacement(Me), "Placement", g.strIniFile
    
    tmrMenu.Enabled = False
    tmrRealTime.Enabled = False
    tmrRefresh.Enabled = False
    
    For lIndex = 1 To m.BarsColl.Count
        g.RealTime.RemoveTickBuffer m.BarsColl(lIndex)
    Next lIndex
    Set m.BarsColl = Nothing
    
    Set m.BrokerObj = Nothing
    
    Set m.RefreshOrders = Nothing
    Set m.RefreshFills = Nothing
    Set m.RefreshCarriedFills = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuFillsPriceLadder_Click
'' Description: Allow the user to show the price ladder
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuFillsPriceLadder_Click()
On Error GoTo ErrSection:

    ShowLadderForFillRow fgFills.Row

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.mnuFillsPriceLadder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersCancel_Click
'' Description: Allow the user to cancel an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersCancel_Click()
On Error GoTo ErrSection:

    If IsValidRow(fgOrders, fgOrders.Row) Then
        StartMenuTimer "CancelOrder", fgOrders.Row
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.mnuOrdersCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersCreate_Click
'' Description: Allow the user to create a new order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersCreate_Click()
On Error GoTo ErrSection:

    StartMenuTimer "CreateOrder", -1&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.mnuOrdersCreate_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersModify_Click
'' Description: Allow the user to modify an order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersModify_Click()
On Error GoTo ErrSection:

    If IsValidRow(fgOrders, fgOrders.Row) Then
        StartMenuTimer "ModifyOrder", fgOrders.Row
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.mnuOrdersModify_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuOrdersPriceLadder_Click
'' Description: Allow the user to show the price ladder
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuOrdersPriceLadder_Click()
On Error GoTo ErrSection:

    ShowLadderForOrderRow fgOrders.Row

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.mnuOrdersPriceLadder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPositionsFlatten_Click
'' Description: Allow the user to flatten a position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPositionsFlatten_Click()
On Error GoTo ErrSection:

    If IsValidRow(fgPositions, fgPositions.Row) Then
        StartMenuTimer "FlattenPosition", fgPositions.Row
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.mnuPositionsFlatten_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPositionsPriceLadder_Click
'' Description: Allow the user to show the price ladder
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPositionsPriceLadder_Click()
On Error GoTo ErrSection:

    ShowLadderForPositionsRow fgPositions.Row

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmBrokerView.mnuPositionsPriceLadder_Click"

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

    TimerStart "frmBrokerView.tmrMenu"
    If bInProgress = False Then
        bInProgress = True
        
        strTag = tmrMenu.Tag
        tmrMenu.Tag = ""
        tmrMenu.Enabled = False
        
        Select Case UCase(Parse(strTag, vbTab, 1))
            Case "CANCELORDER"
                CancelOrder CLng(Val(Parse(strTag, vbTab, 2)))
            Case "CREATEORDER"
                CreateOrder
            Case "FLATTENPOSITION"
                FlattenPosition CLng(Val(Parse(strTag, vbTab, 2)))
            Case "MODIFYORDER"
                ModifyOrder CLng(Val(Parse(strTag, vbTab, 2)))
        End Select
        
        bInProgress = False
    End If
    TimerEnd "frmBrokerView.tmrMenu", tmrMenu.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.tmrMenu_Timer"
    
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

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewBar As Boolean              ' Do we have a new bar?
    Dim Bars As cGdBars                 ' Temporary bars object
    
    TimerStart "frmBrokerView.tmrRealTime"
    For lIndex = 1 To m.BarsColl.Count
        If g.RealTime.UpdateBars(m.BarsColl(lIndex), bNewBar) Then
            If bNewBar = True Then
                Set Bars = m.BarsColl(lIndex)
                LoadBars Bars, Bars.Prop(eBARS_Symbol)
                If Bars.BarsHandle <> m.BarsColl(lIndex).BarsHandle Then
                    Set m.BarsColl(lIndex) = Bars
                End If
            End If
            
            RefreshPositionPrices m.BarsColl(lIndex)
            RefreshTradePrices m.BarsColl(lIndex)
        End If
    Next lIndex
    TimerEnd "frmBrokerView.tmrRealTime", tmrRealTime.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.tmrRealTime_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRefresh_Timer
'' Description: Timer to handle a time-out situation on a refresh
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRefresh_Timer()
On Error GoTo ErrSection:

    TimerStart "frmBrokerView.tmrRefresh"
    If gdTickCount > RefreshStart + 2000 Then
        RefreshStart = 0
        m.bRefreshTimedOut = True
    End If
    TimerEnd "frmBrokerView.tmrRefresh", tmrRefresh.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.tmrRefresh_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable cmdConnect, (m.nStatus = eGDConnectionStatus_Disconnected)
    Enable cmdDisconnect, (m.nStatus = eGDConnectionStatus_Connected)
    Enable lblAccounts, (RefreshStart = 0)
    Enable cboAccounts, (RefreshStart = 0)
    Enable cmdRefreshAccount, ((cboAccounts.ListIndex > -1) And (RefreshStart = 0))
    Enable cmdLookupAccount, (RefreshStart = 0)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitAccountDetailsGrid
'' Description: Initialize the account details grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitAccountDetailsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw state for the grid
    
    With fgAccountDetails
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = 0
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 10
        .FixedRows = 0
        .Cols = 7
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Account"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 4) = "Currency"
        
        .TextMatrix(1, 0) = "Initial Margin"
        .TextMatrix(1, 2) = "Maintenance Margin"
        
        .TextMatrix(2, 0) = "Cash Excess"
        .TextMatrix(2, 2) = "Collateral"
        .TextMatrix(2, 4) = "Market Value of Options"
        
        .TextMatrix(4, 0) = "Previous Balance"
        .TextMatrix(5, 0) = "Closed Profit"
        .TextMatrix(6, 0) = "Account Balance"
        .TextMatrix(7, 0) = "Open Equity"
        .TextMatrix(8, 0) = "OTE+P/L"
        .TextMatrix(9, 0) = "Net Liquidity Value"
        
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignRightTop
                
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.InitAccountDetailsGrid"
    
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

    Dim nRedraw As RedrawSettings       ' Redraw state for the grid
    
    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = OrderCol(eGDOrderCols_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, OrderCol(eGDOrderCols_BrokerID)) = "Broker ID"
        .TextMatrix(0, OrderCol(eGDOrderCols_Side)) = "Side"
        .TextMatrix(0, OrderCol(eGDOrderCols_Quantity)) = "Quantity"
        .TextMatrix(0, OrderCol(eGDOrderCols_Symbol)) = "Broker Symbol"
        .TextMatrix(0, OrderCol(eGDOrderCols_GenesisSymbol)) = "Genesis Symbol"
        .TextMatrix(0, OrderCol(eGDOrderCols_LimitPrice)) = "Limit Price"
        .TextMatrix(0, OrderCol(eGDOrderCols_StopPrice)) = "Stop Price"
        .TextMatrix(0, OrderCol(eGDOrderCols_Type)) = "Type"
        .TextMatrix(0, OrderCol(eGDOrderCols_TIF)) = "TIF"
        .TextMatrix(0, OrderCol(eGDOrderCols_Status)) = "Status"
        .TextMatrix(0, OrderCol(eGDOrderCols_Modify)) = "Modify"
        .TextMatrix(0, OrderCol(eGDOrderCols_Cancel)) = "Cancel"
        
        .ColAlignment(OrderCol(eGDOrderCols_LimitPrice)) = flexAlignRightTop
        .ColAlignment(OrderCol(eGDOrderCols_StopPrice)) = flexAlignRightTop
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.InitOrdersGrid"
    
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

    Dim nRedraw As RedrawSettings       ' Redraw state for the grid
    
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
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 2
        .FixedRows = 2
        .Cols = FillCol(eGDFillCols_NumCols)
        .FixedCols = 0
        
        .ColAlignment(FillCol(eGDFillCols_EntryTime)) = flexAlignCenterTop
        .ColAlignment(FillCol(eGDFillCols_EntryPrice)) = flexAlignRightTop
        .ColAlignment(FillCol(eGDFillCols_ExitTime)) = flexAlignCenterTop
        .ColAlignment(FillCol(eGDFillCols_ExitPrice)) = flexAlignRightTop
        
        .Cell(flexcpText, 0, FillCol(eGDFillCols_BrokerSymbol), 0, FillCol(eGDFillCols_GenesisSymbol)) = "Symbol"
        .Cell(flexcpText, 0, FillCol(eGDFillCols_EntryTime), 0, FillCol(eGDFillCols_EntryPrice)) = "Entry"
        .Cell(flexcpText, 0, FillCol(eGDFillCols_ExitTime), 0, FillCol(eGDFillCols_ExitPrice)) = "Exit"
        .Cell(flexcpText, 0, FillCol(eGDFillCols_OpenEquity), 0, FillCol(eGDFillCols_ClosedProfit)) = "Profit"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .MergeRow(0) = True
        
        .TextMatrix(1, FillCol(eGDFillCols_BrokerSymbol)) = "Broker"
        .TextMatrix(1, FillCol(eGDFillCols_GenesisSymbol)) = "Genesis"
        .TextMatrix(1, FillCol(eGDFillCols_EntryTime)) = "Time"
        .TextMatrix(1, FillCol(eGDFillCols_EntryOrderID)) = "Order ID"
        .TextMatrix(1, FillCol(eGDFillCols_EntryFillID)) = "Fill ID"
        .TextMatrix(1, FillCol(eGDFillCols_EntrySide)) = "Side"
        .TextMatrix(1, FillCol(eGDFillCols_EntryQuantity)) = "Quantity"
        .TextMatrix(1, FillCol(eGDFillCols_EntryPrice)) = "Price"
        
        .TextMatrix(1, FillCol(eGDFillCols_BrokerSymbol)) = "Broker"
        .TextMatrix(1, FillCol(eGDFillCols_GenesisSymbol)) = "Genesis"
        .TextMatrix(1, FillCol(eGDFillCols_ExitTime)) = "Time"
        .TextMatrix(1, FillCol(eGDFillCols_ExitOrderID)) = "Order ID"
        .TextMatrix(1, FillCol(eGDFillCols_ExitFillID)) = "Fill ID"
        .TextMatrix(1, FillCol(eGDFillCols_ExitSide)) = "Side"
        .TextMatrix(1, FillCol(eGDFillCols_ExitQuantity)) = "Quantity"
        .TextMatrix(1, FillCol(eGDFillCols_ExitPrice)) = "Price"
        
        .TextMatrix(1, FillCol(eGDFillCols_OpenEquity)) = "Open"
        .TextMatrix(1, FillCol(eGDFillCols_ClosedProfit)) = "Closed"
        
        ' Add the hidden reserved row so that the Open Equity and Closed Profit columns will not
        ' be merged if they happen to be the same value...
        .ColHidden(FillCol(eGDFillCols_Reserved)) = True
                
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.InitFillsGrid"
    
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

    Dim nRedraw As RedrawSettings       ' Redraw state for the grid
    
    With fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = PositionCol(eGDPositionCols_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, PositionCol(eGDPositionCols_Symbol)) = "Broker Symbol"
        .TextMatrix(0, PositionCol(eGDPositionCols_GenesisSymbol)) = "Genesis Symbol"
        .TextMatrix(0, PositionCol(eGDPositionCols_Carried)) = "Carried"
        .TextMatrix(0, PositionCol(eGDPositionCols_Buys)) = "Bought Today"
        .TextMatrix(0, PositionCol(eGDPositionCols_Sells)) = "Sold Today"
        .TextMatrix(0, PositionCol(eGDPositionCols_Current)) = "Current"
        .TextMatrix(0, PositionCol(eGDPositionCols_Flatten)) = "Flatten"
        .TextMatrix(0, PositionCol(eGDPositionCols_EntryPrice)) = "Average Entry"
        .TextMatrix(0, PositionCol(eGDPositionCols_CurrentPrice)) = "Last Price"
        .TextMatrix(0, PositionCol(eGDPositionCols_Profit)) = "Profit"
        
        .ColAlignment(PositionCol(eGDPositionCols_EntryPrice)) = flexAlignRightTop
        .ColAlignment(PositionCol(eGDPositionCols_CurrentPrice)) = flexAlignRightTop
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        ' 05/02/2012 DAJ - Rance doesn't want the Flatten option here because it
        ' is too dangerous.  For now, I am just going to hide it instead of removing
        ' the code...
        .ColHidden(PositionCol(eGDPositionCols_Flatten)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.InitPositionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConnectionStatus
'' Description: Set the UI controls based on the connection status
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetConnectionStatus(ByVal nStatus As eGDConnectionStatus)
On Error GoTo ErrSection:

    Dim nOldStatus As eGDConnectionStatus ' Previous connectionstatus

    nOldStatus = m.nStatus
    m.nStatus = nStatus
    
    Select Case nStatus
        Case eGDConnectionStatus_Disconnected
            imgStatus.Picture = frmOnlineBroker.imgRed.Picture
            lblStatus.Caption = "Disconnected"
            
        Case eGDConnectionStatus_Disconnecting
            imgStatus.Picture = frmOnlineBroker.imgYellow.Picture
            lblStatus.Caption = "Disconnecting"
        
        Case eGDConnectionStatus_Connecting
            imgStatus.Picture = frmOnlineBroker.imgYellow.Picture
            lblStatus.Caption = "Connecting"
        
        Case eGDConnectionStatus_Connected
            imgStatus.Picture = frmOnlineBroker.imgGreen.Picture
            lblStatus.Caption = "Connected"
    
    End Select
    
    EnableControls
    
    If (m.nStatus = eGDConnectionStatus_Connected) And ((nOldStatus <> eGDConnectionStatus_Connected) Or (cboAccounts.ListCount = 0)) Then
        InfBox "Requesting accounts from the " & m.BrokerObj.BrokerName & " servers.  Please wait...", , , "Accounts", True
        GetAccounts
    ElseIf (nStatus = eGDConnectionStatus_Disconnected) And (nOldStatus <> eGDConnectionStatus_Disconnected) Then
        ClearAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.SetConnectionStatus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request accounts from the broker
'' Inputs:      Show Message?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAccounts()
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        m.BrokerObj.GetAccounts
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Request positions from the broker for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetOrders(ByVal strAccount As String)
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        m.BrokerObj.GetOrders strAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Request positions from the broker for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetFills(ByVal strAccount As String)
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        m.BrokerObj.GetFills strAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Request positions from the broker for the given account
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetPositions(ByVal strAccount As String)
On Error GoTo ErrSection:

    If Not m.BrokerObj Is Nothing Then
        m.BrokerObj.GetPositions strAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountDetailsToGrid
'' Description: Update the account details in the grid
'' Inputs:      Broker Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AccountDetailsToGrid(ByVal brokerMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgAccountDetails
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .TextMatrix(0, 1) = brokerMessage("FcmAccount")
        .TextMatrix(0, 3) = brokerMessage("AccountName")
        .TextMatrix(0, 5) = brokerMessage("Currency")
        
        CurrencyToGrid fgAccountDetails, 1, 1, brokerMessage("InitialMargin")
        CurrencyToGrid fgAccountDetails, 1, 3, brokerMessage("MaintenanceMargin")
        
        CurrencyToGrid fgAccountDetails, 2, 1, brokerMessage("CashExcess")
        CurrencyToGrid fgAccountDetails, 2, 3, brokerMessage("Collateral")
        CurrencyToGrid fgAccountDetails, 2, 5, brokerMessage("MarketValue")
        
        m.dPrevBalance = Val(brokerMessage("EndingBalance"))
        UpdateAccountDetails
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AccountDetailsToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToGrid
'' Description: Update or add the given order in the grid
'' Inputs:      Broker Message, Coming in a refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OrderToGrid(ByVal brokerMessage As cBrokerMessage, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim bHasGenesisSymbol As Boolean    ' Does the message have the Genesis symbol?
    Dim bNewlyAdded As Boolean          ' Has the order just been added to the grid?
    
    bNewlyAdded = False
    
    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, OrderCol(eGDOrderCols_BrokerID)) = brokerMessage("BrokerID") Then
                lRow = lIndex
                Exit For
            End If
        Next lIndex
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
            bNewlyAdded = True
        End If
        
        bHasGenesisSymbol = (Len(brokerMessage("GenesisSymbol")) > 0)
        
        .RowData(lRow) = brokerMessage
        
        .TextMatrix(lRow, OrderCol(eGDOrderCols_BrokerID)) = brokerMessage("BrokerID")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Side)) = brokerMessage("Side")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Quantity)) = brokerMessage("Quantity")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Symbol)) = brokerMessage("Symbol")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_GenesisSymbol)) = brokerMessage("GenesisSymbol")
        If bHasGenesisSymbol = False Then
            .TextMatrix(lRow, OrderCol(eGDOrderCols_LimitPrice)) = brokerMessage("LimitPrice")
            .TextMatrix(lRow, OrderCol(eGDOrderCols_StopPrice)) = brokerMessage("StopPrice")
        Else
            .TextMatrix(lRow, OrderCol(eGDOrderCols_LimitPrice)) = FormatGenesisPrice(brokerMessage("GenesisLimitPrice"), brokerMessage("GenesisSymbol"))
            .TextMatrix(lRow, OrderCol(eGDOrderCols_StopPrice)) = FormatGenesisPrice(brokerMessage("GenesisStopPrice"), brokerMessage("GenesisSymbol"))
        End If
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Type)) = brokerMessage("Type")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_TIF)) = brokerMessage("TIF")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Status)) = brokerMessage("Status")
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Modify)) = "Modify"
        .TextMatrix(lRow, OrderCol(eGDOrderCols_Cancel)) = "Cancel"
        
        If bHasGenesisSymbol = False Then
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = RGB(128, 128, 128)
        Else
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = vbBlack
        End If
        
        .Cell(flexcpFontUnderline, lRow, OrderCol(eGDOrderCols_Modify), lRow, OrderCol(eGDOrderCols_Cancel)) = True
        If IsOpenOrder(brokerMessage("Status")) Then
            .Cell(flexcpForeColor, lRow, OrderCol(eGDOrderCols_Modify), lRow, OrderCol(eGDOrderCols_Cancel)) = vbBlue
        Else
            .Cell(flexcpForeColor, lRow, OrderCol(eGDOrderCols_Modify), lRow, OrderCol(eGDOrderCols_Cancel)) = RGB(128, 128, 128)
        End If
        
        If bRefresh = False Then
            If bNewlyAdded = True Then
                .RowPosition(lRow) = .Rows - 2
            End If
            .AutoSize 0, .Cols - 1, False, 75
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.OrderToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToGrid
'' Description: Update or add the given fill in the grid
'' Inputs:      Broker Message, Coming in a refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillToGrid(ByVal brokerMessage As cBrokerMessage, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFillQuantity As Long           ' Fill quantity
    Dim bIsBuy As Boolean               ' Is this fill a buy?
    Dim strSymbol As String             ' Symbol for the record
    Dim bHasGenesisSymbol As Boolean    ' Does the message have the Genesis symbol?
    
    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, FillCol(eGDFillCols_EntryFillID)) = brokerMessage("FillID") Then
                lRow = lIndex
                Exit For
            End If
        Next lIndex
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            .RowPosition(.Rows - 1) = .Rows - 2
            lRow = .Rows - 2
        End If
        
        bHasGenesisSymbol = (Len(brokerMessage("GenesisSymbol")) > 0)
        
        strSymbol = brokerMessage("Symbol")
        lFillQuantity = CLng(Val(brokerMessage("Quantity")))
        If Len(brokerMessage("Side")) > 0 Then
            bIsBuy = (Left(UCase(brokerMessage("Side")), 3) = "BUY")
        Else
            bIsBuy = (lFillQuantity >= 0)
        End If
        
        .RowData(lRow) = brokerMessage
        .TextMatrix(lRow, FillCol(eGDFillCols_BrokerSymbol)) = strSymbol
        .TextMatrix(lRow, FillCol(eGDFillCols_GenesisSymbol)) = brokerMessage("GenesisSymbol")
        
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryOrderID)) = brokerMessage("BrokerID")
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryFillID)) = brokerMessage("FillID")
        If bIsBuy Then
            .TextMatrix(lRow, FillCol(eGDFillCols_EntrySide)) = "Bought"
        Else
            .TextMatrix(lRow, FillCol(eGDFillCols_EntrySide)) = "Sold"
        End If
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryQuantity)) = Str(Abs(lFillQuantity))
        If bHasGenesisSymbol = False Then
            .TextMatrix(lRow, FillCol(eGDFillCols_EntryPrice)) = brokerMessage("FillPrice")
        Else
            .TextMatrix(lRow, FillCol(eGDFillCols_EntryPrice)) = FormatGenesisPrice(brokerMessage("GenesisFillPrice"), brokerMessage("GenesisSymbol"))
        End If
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryTime)) = brokerMessage("FillDate")
        
        If bIsBuy Then
            If m.NumBuysToday.Exists(strSymbol) Then
                m.NumBuysToday(strSymbol) = m.NumBuysToday(strSymbol) + lFillQuantity
            Else
                m.NumBuysToday(strSymbol) = lFillQuantity
            End If
        Else
            If m.NumSellsToday.Exists(strSymbol) Then
                m.NumSellsToday(strSymbol) = m.NumSellsToday(strSymbol) + Abs(lFillQuantity)
            Else
                m.NumSellsToday(strSymbol) = Abs(lFillQuantity)
            End If
        End If
        
        If bHasGenesisSymbol = False Then
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = RGB(128, 128, 128)
        Else
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = vbBlack
        End If
        
        If bRefresh = False Then
            .AutoSize 0, .Cols - 1, False, 75
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.FillToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionToGrid
'' Description: Update or add the given position in the grid
'' Inputs:      Broker Message, Coming in a refresh?, Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PositionToGrid(ByVal brokerMessage As cBrokerMessage, ByVal bRefresh As Boolean, Optional BrokerPosition As cPtPosition = Nothing)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRow As Long                    ' Row in the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCarriedPosition As Long        ' Carried position
    Dim lCurrentPosition As Long        ' Current position
    Dim strSymbol As String             ' Symbol for the record
    Dim lNumBuy As Long                 ' Number bought today
    Dim lNumSell As Long                ' Number sold today
    Dim bHasGenesisSymbol As Boolean    ' Does the message have the Genesis symbol?
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim Bars As cGdBars                 ' Bars object
    
    With fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, PositionCol(eGDPositionCols_Symbol)) = brokerMessage("Symbol") Then
                lRow = lIndex
                Exit For
            End If
        Next lIndex
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            .RowPosition(.Rows - 1) = .Rows - 2
            lRow = .Rows - 2
        End If
        
        strSymbol = brokerMessage("Symbol")
        strGenesisSymbol = brokerMessage("GenesisSymbol")
        bHasGenesisSymbol = (Len(strGenesisSymbol) > 0)
        
        .RowData(lRow) = brokerMessage
        
        .TextMatrix(lRow, PositionCol(eGDPositionCols_Symbol)) = strSymbol
        .TextMatrix(lRow, PositionCol(eGDPositionCols_GenesisSymbol)) = strGenesisSymbol
        
        lCarriedPosition = CLng(Val(brokerMessage("CarriedQuantity")))
        If lCarriedPosition = 0 Then
            .TextMatrix(lRow, PositionCol(eGDPositionCols_Carried)) = "Flat"
        ElseIf UCase(brokerMessage("CarriedDirection")) = "LONG" Then
            .TextMatrix(lRow, PositionCol(eGDPositionCols_Carried)) = "Long " & brokerMessage("CarriedQuantity")
        Else
            .TextMatrix(lRow, PositionCol(eGDPositionCols_Carried)) = "Short " & brokerMessage("CarriedQuantity")
        End If
        
        If m.NumBuysToday.Exists(strSymbol) Then
            lNumBuy = m.NumBuysToday(strSymbol)
        Else
            lNumBuy = 0
        End If
        .TextMatrix(lRow, PositionCol(eGDPositionCols_Buys)) = Str(lNumBuy)
        
        If m.NumSellsToday.Exists(strSymbol) Then
            lNumSell = m.NumSellsToday(strSymbol)
        Else
            lNumSell = 0
        End If
        .TextMatrix(lRow, PositionCol(eGDPositionCols_Sells)) = Str(lNumSell)
        
        lCurrentPosition = CLng(Val(brokerMessage("Quantity")))
        If lCurrentPosition = 0 Then
            .TextMatrix(lRow, PositionCol(eGDPositionCols_Current)) = "Flat*"
        Else
            .TextMatrix(lRow, PositionCol(eGDPositionCols_Current)) = brokerMessage("Direction") & " " & brokerMessage("Quantity") & "*"
        End If
        
        .TextMatrix(lRow, PositionCol(eGDPositionCols_Flatten)) = "Flatten"
        
        If bHasGenesisSymbol = False Then
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = RGB(128, 128, 128)
            .TextMatrix(lRow, PositionCol(eGDPositionCols_EntryPrice)) = brokerMessage("AverageEntry")
            .TextMatrix(lRow, PositionCol(eGDPositionCols_CurrentPrice)) = ""
            .TextMatrix(lRow, PositionCol(eGDPositionCols_Profit)) = ""
        Else
            .Cell(flexcpForeColor, lRow, 0, lRow, .Cols - 1) = vbBlack
            
            Set Bars = GetBars(strGenesisSymbol)
            RefreshPositionPrices Bars, lRow
        End If
        
        ' Make sure to do this down here to override the forecolor setting for the whole row...
        .Cell(flexcpFontUnderline, lRow, PositionCol(eGDPositionCols_Flatten)) = True
        If lCurrentPosition <> 0 Then
            .Cell(flexcpForeColor, lRow, PositionCol(eGDPositionCols_Flatten)) = vbBlue
        Else
            .Cell(flexcpForeColor, lRow, PositionCol(eGDPositionCols_Flatten)) = RGB(128, 128, 128)
        End If
                
        If bRefresh = False Then
            .AutoSize 0, .Cols - 1, False, 75
        End If
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.PositionToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOpenOrder
'' Description: Determine if the given order status is "open"
'' Inputs:      Status
'' Returns:     True if open, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsOpenOrder(ByVal strStatus As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    Select Case UCase(strStatus)
        Case "NEW", "PARTIAL", "REPLACED", "PARKED"
            bReturn = True
        
        Case Else
            bReturn = False
    End Select
    
    IsOpenOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.IsOpenOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrder
'' Description: Allow the user to create an order
'' Inputs:      Order, Symbol, Feed Yard Lot ID, Side
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateOrder(Optional ByVal Order As cPtOrder = Nothing, Optional ByVal strGenesisSymbol As String = "", Optional ByVal strFeedYardLotID As String = "", Optional ByVal strSide As String = "")
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim BrokerKv As cBrokerKeyValue     ' Routines for key-value broker messages
    Dim strBrokerMessage As String      ' Broker message
    Dim strGenesisID As String          ' Genesis order ID for the new order

    If cboAccounts.ListIndex >= 0& Then
        Set brokerMessage = New cBrokerMessage
        If Not Order Is Nothing Then
            Set BrokerKv = New cBrokerKeyValue
            strBrokerSymbol = g.RjoCqg.BrokerSymbol(Order.Symbol, BrokerSym)
            If Len(strBrokerSymbol) > 0 Then
                strBrokerMessage = BrokerKv.OrderToMessage(Order, True, strBrokerSymbol, BrokerSym, m.BrokerObj)
                If Len(strBrokerMessage) > 0 Then
                    brokerMessage.FromString strBrokerMessage
                End If
            End If
        Else
            brokerMessage.Add "GenesisSymbol", strGenesisSymbol
        End If
        If Len(strSide) > 0 Then
            brokerMessage.Add "Side", strSide
        End If
        
        If frmSimpleOrder.ShowMe(brokerMessage, cboAccounts.ItemData(cboAccounts.ListIndex), , , strFeedYardLotID) = True Then
            If AddBrokerSymbolToOrder(brokerMessage) Then
                strGenesisID = NextGenesisOrderID(m.strAccount, m.nBroker)
                brokerMessage.Add "GenesisID", strGenesisID
                brokerMessage.Add "Account", m.strAccount
                
                If Len(strFeedYardLotID) > 0 Then
                    g.CattleBridge.SetUpNewOrder strGenesisID, strFeedYardLotID
                End If
                
                g.RjoCqg.AddOrderFromString brokerMessage.ToString(False)
            Else
                InfBox "Could not convert " & brokerMessage("GenesisSymbol") & " to a " & g.Broker.BrokerName(m.nBroker) & " symbol", "!", , "Order Error"
            End If
        End If
    Else
        MoveFocus cboAccounts
        InfBox "No account selected", "!", , "Create Order Error"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CreateOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ModifyOrder
'' Description: Allow the user to modify an order
'' Inputs:      Row of the Order, New Price, New Quantity
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ModifyOrder(ByVal lRow As Long, Optional ByVal dNewPrice As Double = 0#, Optional ByVal lNewQuantity As Long = 0&)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message object

    With fgOrders
        If IsValidRow(fgOrders, lRow) Then
            If TypeOf .RowData(lRow) Is cBrokerMessage Then
                Set brokerMessage = .RowData(lRow)
                
                If IsOpenOrder(brokerMessage("Status")) Then
                    If frmSimpleOrder.ShowMe(brokerMessage, cboAccounts.ItemData(cboAccounts.ListIndex), dNewPrice, lNewQuantity) = True Then
                        If AddBrokerSymbolToOrder(brokerMessage) Then
                            brokerMessage("PreviousGenesisID") = brokerMessage("GenesisID")
                            brokerMessage("GenesisID") = NextGenesisOrderID(m.strAccount, m.nBroker)
                        
                            g.RjoCqg.AmendOrderFromString brokerMessage.ToString(False)
                        Else
                            InfBox "Could not convert " & brokerMessage("GenesisSymbol") & " to a " & g.Broker.BrokerName(m.nBroker) & " symbol", "!", , "Order Error"
                        End If
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ModifyOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Allow the user to cancel an order
'' Inputs:      Row of the Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CancelOrder(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message object

    With fgOrders
        If IsValidRow(fgOrders, lRow) Then
            If TypeOf .RowData(lRow) Is cBrokerMessage Then
                Set brokerMessage = .RowData(lRow)
                
                If IsOpenOrder(brokerMessage("Status")) Then
                    If InfBox("Are you sure you want to cancel order '" & brokerMessage("BrokerID") & "'?", "?", "+Cancel|-Don't Cancel", "Cancel Order Confirmation") = "C" Then
                        brokerMessage("PreviousGenesisID") = brokerMessage("GenesisID")
                        brokerMessage("GenesisID") = NextGenesisOrderID(m.strAccount, m.nBroker)
                        
                        g.RjoCqg.CancelOrderFromString brokerMessage.ToString(False)
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CancelOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FlattenPosition
'' Description: Allow the user to flatten a position
'' Inputs:      Row of the Position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FlattenPosition(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol from the grid
    Dim strPosition As String           ' Position from the grid
    Dim strDirection As String          ' Direction of the position
    Dim lPosition As Long               ' Quantity of the position
    Dim strOrderString As String        ' Order string to send to the broker
    Dim strAction As String             ' Action to display to the user

    If IsValidRow(fgPositions, lRow) Then
        strSymbol = fgPositions.TextMatrix(lRow, PositionCol(eGDPositionCols_Symbol))
        strPosition = fgPositions.TextMatrix(lRow, PositionCol(eGDPositionCols_Current))
        
        strDirection = UCase(Parse(strPosition, " ", 1))
        lPosition = CLng(Val(Parse(strPosition, " ", 2)))
        
        If strDirection <> "FLAT" Then
            If strDirection = "LONG" Then
                strAction = "Sell " & Str(lPosition) & " " & strSymbol & " at Market"
            Else
                strAction = "Buy " & Str(lPosition) & " " & strSymbol & " at Market"
            End If
            
            If InfBox("You are about to submit an order to|" & strAction & "| in account " & m.strFcmAccount & "||Do you want to continue?", "?", "+Submit|-Don't Submit", "Flatten Position Confirmation") = "S" Then
                strOrderString = CreateMarketOrder(strSymbol, strDirection = "SHORT", lPosition)
                g.RjoCqg.AddOrderFromString strOrderString
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.FlattenPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsValidRow
'' Description: Is the given row valid in the given grid?
'' Inputs:      Grid, Row
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsValidRow(fgGrid As VSFlexGrid, ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    IsValidRow = (lRow >= fgGrid.FixedRows) And (lRow < fgGrid.Rows)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.IsValidRow"
    
End Function

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
    RaiseError "frmBrokerView.StartMenuTimer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectFcmAccount
'' Description: Select the given account in the accounts combo
'' Inputs:      Account
'' Returns:     True if selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectFcmAccount(ByVal strFcmAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the given Fcm account
    
    bReturn = False
    If m.Accounts.Exists(strFcmAccount) Then
        lAccountID = m.Accounts.Index(strFcmAccount)
        
        For lIndex = 0 To cboAccounts.ListCount - 1
            If cboAccounts.ItemData(lIndex) = lAccountID Then
                If cboAccounts.ListIndex <> lIndex Then
                    m.bSetAccountFromCode = True
                    cboAccounts.ListIndex = lIndex
                    If Not m.frmLadder Is Nothing Then
                        m.frmLadder.TradeAccountID = cboAccounts.ItemData(lIndex)
                    End If
                    
                    bReturn = True
                End If
                
                Exit For
            End If
        Next lIndex
    End If
    
    SelectFcmAccount = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.SelectFcmAccount"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectAccountID
'' Description: Select the given account in the accounts combo
'' Inputs:      Account ID
'' Returns:     True if selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectAccountID(ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 0 To cboAccounts.ListCount - 1
        If cboAccounts.ItemData(lIndex) = lAccountID Then
            If cboAccounts.ListIndex <> lIndex Then
                m.bSetAccountFromCode = True
                cboAccounts.ListIndex = lIndex
                If Not m.frmLadder Is Nothing Then
                    m.frmLadder.TradeAccountID = cboAccounts.ItemData(lIndex)
                End If
            
                bReturn = True
            End If
            
            Exit For
        End If
    Next lIndex
    
    SelectAccountID = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.SelectAccountID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateMarketOrder
'' Description: Create a market order string from the given information
'' Inputs:      Symbol, Is Buy?, Quantity
'' Returns:     Order String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateMarketOrder(ByVal strSymbol As String, ByVal bIsBuy As Boolean, ByVal lQuantity As Long) As String
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker Message object

    Set brokerMessage = New cBrokerMessage
    brokerMessage.Add "GenesisID", NextGenesisOrderID(m.strAccount, m.nBroker)
    brokerMessage.Add "Account", m.strAccount
    brokerMessage.Symbol = strSymbol
    brokerMessage.Add "Type", "Market"
    If bIsBuy Then
        brokerMessage.Add "Side", "Buy"
    Else
        brokerMessage.Add "Side", "Sell"
    End If
    brokerMessage.Add "Quantity", Str(lQuantity)
    brokerMessage.Add "TIF", "Day"
    brokerMessage.Add "Manual", "Y"
    brokerMessage.Add "Direction", "C"
    
    CreateMarketOrder = brokerMessage.ToString(False)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.CreateMarketOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereRow
'' Description: Add the "Click Here" row to the orders grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereRow()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With fgOrders
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to create new order"
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        .Cell(flexcpFontUnderline, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
        .MergeRow(.Rows - 1) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AddClickHereRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddFillsTotalRow
'' Description: Add the Total row to the fills grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddFillsTotalRow()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, FillCol(eGDFillCols_OpenEquity) - 1) = "Total"
        CurrencyToGrid fgFills, .Rows - 1, FillCol(eGDFillCols_OpenEquity), "0"
        CurrencyToGrid fgFills, .Rows - 1, FillCol(eGDFillCols_ClosedProfit), "0"
        .MergeRow(.Rows - 1) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AddFillsTotalRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddPositionsTotalRow
'' Description: Add the Total row to the positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddPositionsTotalRow()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    
    With fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, PositionCol(eGDPositionCols_Profit) - 1) = "Total"
        CurrencyToGrid fgPositions, .Rows - 1, PositionCol(eGDPositionCols_Profit), "0"
        .MergeRow(.Rows - 1) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AddPositionsTotalRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshSelectedAccount
'' Description: Refresh the currently selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshSelectedAccount()
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message from the accounts collection

    Set brokerMessage = SelectedAccount
    If Not brokerMessage Is Nothing Then
        If m.strAccount <> brokerMessage("Account") Then
            RefreshStart = gdTickCount
            
            m.strAccount = brokerMessage("Account")
            m.strFcmAccount = brokerMessage("FcmAccount")
            
            m.strRefreshAccount = ""
            m.strRefreshFcmAccount = ""
            m.bRefreshDone = False
            m.bRefreshTimedOut = False
        
            Caption = m.strCaption & " - [" & m.strFcmAccount & "]"
            
            m.dPrevBalance = 0#
            m.dAccountBalance = 0#
            m.dNetLiquidity = 0#
            m.dTotalClosedProfit = 0#
            m.dTotalOpenEquity = 0#
            
            UpdateAccountDetails
            
            If Not m.frmLadder Is Nothing Then
                m.frmLadder.TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
            End If
        End If
        
        InfBox "Requesting orders from the " & m.BrokerObj.BrokerName & " servers for account " & m.strFcmAccount & ".  Please wait...", , , "Orders", True
        GetOrders m.strAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.RefreshSelectedAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddGenesisSymbolToOrder
'' Description: Add the Genesis symbol and prices to an order message
'' Inputs:      Order Message
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddGenesisSymbolToOrder(brokerOrderMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(brokerOrderMessage("Symbol"), brokerOrderMessage("Exchange"), BrokerSym)
    If Len(strGenesisSymbol) > 0 Then
        brokerOrderMessage.Add "GenesisSymbol", strGenesisSymbol
        
        If Len(brokerOrderMessage("StopPrice")) > 0 Then
            brokerOrderMessage.Add "GenesisStopPrice", Str(m.BrokerObj.GenesisPrice(brokerOrderMessage("StopPrice"), BrokerSym.PriceMult))
        End If
        If Len(brokerOrderMessage("LimitPrice")) > 0 Then
            brokerOrderMessage.Add "GenesisLimitPrice", Str(m.BrokerObj.GenesisPrice(brokerOrderMessage("LimitPrice"), BrokerSym.PriceMult))
        End If
        
        bReturn = True
    End If
    
    AddGenesisSymbolToOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.AddGenesisSymbolToOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerSymbolToOrder
'' Description: Add the broker symbol and prices to an order message
'' Inputs:      Order Message
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddBrokerSymbolToOrder(brokerOrderMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    strBrokerSymbol = g.RjoCqg.BrokerSymbol(brokerOrderMessage("GenesisSymbol"), BrokerSym)
    If Len(strBrokerSymbol) > 0 Then
        brokerOrderMessage.Add "Symbol", strBrokerSymbol
        
        If Len(brokerOrderMessage("GenesisStopPrice")) > 0 Then
            brokerOrderMessage.Add "StopPrice", Str(m.BrokerObj.BrokerPrice(Val(brokerOrderMessage("GenesisStopPrice")), BrokerSym.PriceMult, BrokerSym.PriceFormat))
        End If
        If Len(brokerOrderMessage("GenesisLimitPrice")) > 0 Then
            brokerOrderMessage.Add "LimitPrice", Str(m.BrokerObj.BrokerPrice(Val(brokerOrderMessage("GenesisLimitPrice")), BrokerSym.PriceMult, BrokerSym.PriceFormat))
        End If
        
        bReturn = True
    End If
    
    AddBrokerSymbolToOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.AddBrokerSymbolToOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddGenesisSymbolToFill
'' Description: Add the Genesis symbol and prices to a fill message
'' Inputs:      Fill Message
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddGenesisSymbolToFill(brokerFillMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(brokerFillMessage("Symbol"), brokerFillMessage("Exchange"), BrokerSym)
    If Len(strGenesisSymbol) > 0 Then
        brokerFillMessage.Add "GenesisSymbol", strGenesisSymbol
        
        If Len(brokerFillMessage("FillPrice")) > 0 Then
            brokerFillMessage.Add "GenesisFillPrice", Str(m.BrokerObj.GenesisPrice(brokerFillMessage("FillPrice"), BrokerSym.PriceMult))
        End If
    End If
    
    AddGenesisSymbolToFill = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.AddGenesisSymbolToFill"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddGenesisSymbolToPosition
'' Description: Add the Genesis symbol to a position message
'' Inputs:      Position Message
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddGenesisSymbolToPosition(brokerPositionMessage As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(brokerPositionMessage("Symbol"), brokerPositionMessage("Exchange"), BrokerSym)
    If Len(strGenesisSymbol) > 0 Then
        brokerPositionMessage.Add "GenesisSymbol", strGenesisSymbol
    End If
    
    AddGenesisSymbolToPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.AddGenesisSymbolToPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormatPrice
'' Description: Format a Genesis price for the given Genesis symbol
'' Inputs:      Genesis Price, Genesis Symbol
'' Returns:     True if added, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FormatGenesisPrice(ByVal strGenesisPrice As String, ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim Bars As cGdBars                 ' Bars object
    Dim strReturn As String             ' Return value for the function
    
    If Len(strGenesisPrice) = 0 Then
        strReturn = ""
    Else
        If m.BarProps.Exists(strGenesisSymbol) Then
            Set Bars = m.BarProps(strGenesisSymbol)
        Else
            Set Bars = New cGdBars
            If SetBarProperties(Bars, strGenesisSymbol) Then
                m.BarProps.Add Bars, strGenesisSymbol
            Else
                m.BarProps.Add Nothing, strGenesisSymbol
            End If
        End If
        
        If Bars Is Nothing Then
            strReturn = strGenesisPrice
        Else
            strReturn = Bars.PriceDisplay(Val(strGenesisPrice))
        End If
    End If
    
    FormatGenesisPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.FormatPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildPositionFills
'' Description: Build the position fills collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildPositionFills()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim FillsForSymbol As cGdTree       ' Collection of fills for the symbol
    Dim strSymbol As String             ' Symbol for the fill
    Dim carriedFill As cBrokerMessage   ' Carried fill from the collection
    
    m.PositionFills.Clear
    For lIndex = 1 To m.CarriedFills.Count
        AddTodayFillToPositionFills m.CarriedFills(lIndex), True
#If 0 Then
        Set carriedFill = m.CarriedFills(lIndex)
        strSymbol = carriedFill("Symbol")
        If m.PositionFills.Exists(strSymbol) Then
            Set FillsForSymbol = m.PositionFills(strSymbol)
            FillsForSymbol.Add carriedFill
            Set m.PositionFills(strSymbol) = FillsForSymbol
        Else
            Set FillsForSymbol = New cGdTree
            FillsForSymbol.Add carriedFill
            m.PositionFills.Add FillsForSymbol, strSymbol
        End If
#End If
    Next lIndex
    
    For lIndex = 1 To m.TodaysFills.Count
        AddTodayFillToPositionFills m.TodaysFills(lIndex), True
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.BuildPositionFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTodayFillToPositionFills
'' Description: Add the "Today Fill" to the position fills collection
'' Inputs:      Today Fill, Coming from a Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddTodayFillToPositionFills(todayFill As cBrokerMessage, ByVal bRefresh As Boolean)
On Error GoTo ErrSection:

    Dim FillsForSymbol As cGdTree       ' Collection of fills for the symbol
    Dim strSymbol As String             ' Symbol for the fill
    Dim carriedFill As cBrokerMessage   ' Carried fill from the collection
    Dim lOldQuantity As Long            ' Quantity in existing fill
    Dim lNewQuantity As Long            ' Quantity in new fill

    strSymbol = todayFill("Symbol")
    If m.PositionFills.Exists(strSymbol) Then
        Set FillsForSymbol = m.PositionFills(strSymbol)
        If FillsForSymbol.Count = 0 Then
            FillsForSymbol.Add todayFill
        Else
            Set carriedFill = FillsForSymbol(1)
            If todayFill("Side") = carriedFill("Side") Then
                FillsForSymbol.Add todayFill
            Else
                lNewQuantity = todayFill("Quantity")
                
                Do While (lNewQuantity > 0) And (FillsForSymbol.Count > 0)
                    Set carriedFill = FillsForSymbol(1)
                    lOldQuantity = carriedFill("Quantity")
                    
                    If lOldQuantity = lNewQuantity Then
                        FillsForSymbol.Remove 1
                        lNewQuantity = 0
                    ElseIf lOldQuantity > lNewQuantity Then
                        carriedFill("Quantity") = Str(lOldQuantity - lNewQuantity)
                        Set FillsForSymbol(1) = carriedFill
                        lNewQuantity = 0
                    Else
                        lNewQuantity = lNewQuantity - lOldQuantity
                        FillsForSymbol.Remove 1
                    End If
                Loop
                
                If lNewQuantity > 0 Then
                    todayFill("Quantity") = Str(lNewQuantity)
                    FillsForSymbol.Add todayFill
                End If
            End If
        End If
        Set m.PositionFills(strSymbol) = FillsForSymbol
    Else
        Set FillsForSymbol = New cGdTree
        FillsForSymbol.Add todayFill
        m.PositionFills.Add FillsForSymbol, strSymbol
    End If
    
    If bRefresh = False Then
        CalcPositionFromPositionFills strSymbol
        PositionToGrid m.Positions(strSymbol), bRefresh
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AddTodayFillToPositionFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcPositionFromPositionFills
'' Description: Calculate the position from the position fills
'' Inputs:      Broker Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcPositionFromPositionFills(ByVal strBrokerSymbol As String)
On Error GoTo ErrSection:

    Dim Position As cBrokerMessage      ' Position message to send to the grid
    Dim FillsForSymbol As cGdTree       ' Collection of fills for the symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim brokerFill As cBrokerMessage    ' Fill from the collection
    Dim lQuantity As Long               ' Quantity for the fill
    Dim dFillPriceBroker As Double      ' Price for the fill
    Dim dFillPriceGenesis As Double     ' Price for the fill
    Dim lPosition As Long               ' Position
    Dim dPriceSumBroker As Double       ' Sum of the fill prices
    Dim dPriceSumGenesis As Double      ' Sum of the fill prices
    Dim dAverageEntryBroker As Double   ' Average Entry price
    Dim dAverageEntryGenesis As Double  ' Average Entry price
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim CarriedPosition As cBrokerMessage 'Carried position
    
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(strBrokerSymbol, "", BrokerSym)
    lPosition = 0&
    dAverageEntryBroker = 0#
    dAverageEntryGenesis = 0#
    dPriceSumBroker = 0#
    dPriceSumGenesis = 0#
    
    Set Position = New cBrokerMessage
    Position.Add "Symbol", strBrokerSymbol
    Position.Add "GenesisSymbol", strGenesisSymbol
    Position.Add "Direction", "Flat"
    
    If m.CarriedPositions.Exists(strBrokerSymbol) Then
        Set CarriedPosition = m.CarriedPositions(strBrokerSymbol)
        Position.Add "CarriedQuantity", CarriedPosition("Carried")
        If UCase(CarriedPosition("CarriedSide")) = "BUY" Then
            Position.Add "CarriedDirection", "Long"
        Else
            Position.Add "CarriedDirection", "Short"
        End If
    Else
        Position.Add "CarriedQuantity", "0"
        Position.Add "CarriedDirection", "Flat"
    End If
                
    If m.PositionFills.Exists(strBrokerSymbol) Then
        Set FillsForSymbol = m.PositionFills(strBrokerSymbol)
        For lIndex = 1 To FillsForSymbol.Count
            Set brokerFill = FillsForSymbol(lIndex)
            
            If lIndex = 1 Then
                If UCase(brokerFill("Side")) = "BUY" Then
                    Position("Direction") = "Long"
                Else
                    Position("Direction") = "Short"
                End If
            End If
            
            lQuantity = CLng(Val(brokerFill("Quantity")))
            
            dFillPriceBroker = Val(brokerFill("FillPrice"))
            dPriceSumBroker = dPriceSumBroker + (dFillPriceBroker * CDbl(lQuantity))
            
            If Len(strGenesisSymbol) > 0 Then
                dFillPriceGenesis = Val(brokerFill("GenesisFillPrice"))
                dPriceSumGenesis = dPriceSumGenesis + (dFillPriceGenesis * CDbl(lQuantity))
            End If
            
            lPosition = lPosition + lQuantity
        Next lIndex
        
        If lPosition <> 0& Then
            dAverageEntryBroker = dPriceSumBroker / CDbl(lPosition)
            If Len(strGenesisSymbol) > 0 Then
                dAverageEntryGenesis = dPriceSumGenesis / CDbl(lPosition)
            End If
        Else
            dAverageEntryBroker = 0#
            If Len(strGenesisSymbol) > 0 Then
                dAverageEntryGenesis = 0
            End If
        End If
    End If
    
    Position.Add "Quantity", Str(lPosition)
    Position.Add "AverageEntry", Str(dAverageEntryBroker)
    If Len(strGenesisSymbol) > 0 Then
        Position.Add "GenesisAverageEntry", Str(dAverageEntryGenesis)
    End If
    
    If m.Positions.Exists(strBrokerSymbol) Then
        Set m.Positions(strBrokerSymbol) = Position
    Else
        m.Positions.Add Position, strBrokerSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CalcPositionFromPositionFills"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLadderForOrderRow
'' Description: Show the price ladder for the given row in the orders grid
'' Inputs:      Row in Orders Grid, Only if Valid Row?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLadderForOrderRow(ByVal lOrderRow As Long, Optional ByVal bOnlyIfValidRow As Boolean = False)
On Error GoTo ErrSection:

    Dim orderMessage As cBrokerMessage  ' Order message from the grid
    Dim strGenesisSymbol As String      ' Genesis symbol
    
    strGenesisSymbol = ""
    With fgOrders
        If (lOrderRow >= .FixedRows) And (lOrderRow < .Rows) Then
            If TypeOf .RowData(lOrderRow) Is cBrokerMessage Then
                Set orderMessage = .RowData(lOrderRow)
                strGenesisSymbol = orderMessage("GenesisSymbol")
            End If
        End If
    End With
    
    If (Len(strGenesisSymbol) > 0) Or (bOnlyIfValidRow = False) Then
        ShowLadder strGenesisSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ShowLadderForOrderRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLadderForFillRow
'' Description: Show the price ladder for the given row in the fills grid
'' Inputs:      Row in Fills Grid, Only if Valid Row?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLadderForFillRow(ByVal lFillRow As Long, Optional ByVal bOnlyIfValidRow As Boolean = False)
On Error GoTo ErrSection:

    Dim Trade As cTrade                 ' Trade object
    Dim strGenesisSymbol As String      ' Genesis symbol
    
    strGenesisSymbol = ""
    With fgFills
        If (lFillRow >= .FixedRows) And (lFillRow < .Rows) Then
            If TypeOf .RowData(lFillRow) Is cTrade Then
                Set Trade = .RowData(lFillRow)
                strGenesisSymbol = Trade.GenesisSymbol
            End If
        End If
    End With
    
    If (Len(strGenesisSymbol) > 0) Or (bOnlyIfValidRow = False) Then
        ShowLadder strGenesisSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ShowLadderForFillRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLadderForPositionRow
'' Description: Show the price ladder for the given row in the positions grid
'' Inputs:      Row in Positions Grid, Only if Valid Row?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLadderForPositionsRow(ByVal lPositionRow As Long, Optional ByVal bOnlyIfValidRow As Boolean = False)
On Error GoTo ErrSection:

    Dim positionMessage As cBrokerMessage ' Position message from the grid
    Dim strGenesisSymbol As String      ' Genesis symbol
    
    strGenesisSymbol = ""
    With fgPositions
        If (lPositionRow >= .FixedRows) And (lPositionRow < .Rows) Then
            If TypeOf .RowData(lPositionRow) Is cBrokerMessage Then
                Set positionMessage = .RowData(lPositionRow)
                strGenesisSymbol = positionMessage("GenesisSymbol")
            End If
        End If
    End With
        
    If (Len(strGenesisSymbol) > 0) Or (bOnlyIfValidRow = False) Then
        ShowLadder strGenesisSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ShowLadderForOrderRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForOrder
'' Description: Determine the row for the given order
'' Inputs:      Order
'' Returns:     Row (-1 if not valid or found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForOrder(ByVal Order As cPtOrder) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim orderMessage As cBrokerMessage  ' Order Message from the grid
    Dim lRow As Long                    ' Row in the grid
        
    lReturn = -1&
    lRow = Order.OrderID
    
    With fgOrders
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            If TypeOf .RowData(lRow) Is cBrokerMessage Then
                Set orderMessage = .RowData(lRow)
                If Order.BrokerID = orderMessage("BrokerID") Then
                    lReturn = lRow
                End If
            End If
        End If
    End With
    
    RowForOrder = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.RowForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrencyToGrid
'' Description: Set the given cell to the given value
'' Inputs:      Grid, Row, Column, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CurrencyToGrid(fgGrid As VSFlexGrid, ByVal lRow As Long, ByVal lCol As Long, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value converted to a double
    
    If Len(strValue) = 0 Then
        fgGrid.TextMatrix(lRow, lCol) = ""
    Else
        dValue = Val(strValue)
        fgGrid.TextMatrix(lRow, lCol) = Format(dValue, "$#,##0.00")
        If dValue < 0 Then
            fgGrid.Cell(flexcpForeColor, lRow, lCol) = vbRed
        ElseIf dValue = 0 Then
            fgGrid.Cell(flexcpForeColor, lRow, lCol) = vbBlack
        Else
            fgGrid.Cell(flexcpForeColor, lRow, lCol) = vbGreen
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CurrencyToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBars
'' Description: Get bars for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBars(ByVal strGenesisSymbol As String) As cGdBars
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    If m.BarsColl.Exists(strGenesisSymbol) = False Then
        Set Bars = New cGdBars
        LoadBars Bars, strGenesisSymbol
        m.BarsColl.Add Bars, strGenesisSymbol
    End If
    
    Set GetBars = m.BarsColl(strGenesisSymbol)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.GetBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load and splice the bars for the given Genesis symbol
'' Inputs:      Bars to Load, Genesis Symbol
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadBars(Bars As cGdBars, ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value
    
    Bars.ArrayMask = eBARS_EodBidAsk
    bReturn = DM_GetBars(Bars, strGenesisSymbol, , LastDailyDownload - 5, , , False)
    
    g.RealTime.AddTickBuffer Bars
    g.RealTime.SpliceBars Bars, , True
    
    LoadBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.LoadBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPositionPrices
'' Description: Refresh the prices on the positions grid
'' Inputs:      Bars, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshPositionPrices(Bars As cGdBars, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim brokerMessage As cBrokerMessage ' Broker message
    Dim dCurrentPrice As Double         ' Current price
    Dim dEntryPrice As Double           ' Entry price
    Dim dProfit As Double               ' Profit
    Dim lCurrentPosition As Long        ' Current position
    
    strGenesisSymbol = Bars.Prop(eBARS_Symbol)
    With fgPositions
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        Set brokerMessage = Nothing
        If lRow = -1& Then
            For lIndex = .FixedRows To .Rows - 2
                If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                    Set brokerMessage = .RowData(lIndex)
                    If brokerMessage("GenesisSymbol") = strGenesisSymbol Then
                        lRow = lIndex
                        Exit For
                    Else
                        Set brokerMessage = Nothing
                    End If
                End If
            Next lIndex
        ElseIf TypeOf .RowData(lRow) Is cBrokerMessage Then
            Set brokerMessage = .RowData(lRow)
        End If
        
        If Not brokerMessage Is Nothing Then
            dCurrentPrice = Bars(eBARS_Close, Bars.Size - 1)
            dEntryPrice = Val(brokerMessage("GenesisAverageEntry"))
            
            lCurrentPosition = CLng(Val(brokerMessage("Quantity")))
            If lCurrentPosition = 0 Then
                dProfit = 0#
            
                .TextMatrix(lRow, PositionCol(eGDPositionCols_EntryPrice)) = ""
                .TextMatrix(lRow, PositionCol(eGDPositionCols_CurrentPrice)) = ""
                .TextMatrix(lRow, PositionCol(eGDPositionCols_Profit)) = ""
            Else
                If brokerMessage("Direction") = "Long" Then
                    dProfit = g.Profit.Profit(strGenesisSymbol, dCurrentPrice - dEntryPrice, lCurrentPosition, , , , AccountNumber)
                ElseIf brokerMessage("Direction") = "Short" Then
                    dProfit = g.Profit.Profit(strGenesisSymbol, dEntryPrice - dCurrentPrice, Abs(lCurrentPosition), , , , AccountNumber)
                End If
                
                .TextMatrix(lRow, PositionCol(eGDPositionCols_EntryPrice)) = FormatGenesisPrice(brokerMessage("GenesisAverageEntry"), strGenesisSymbol)
                
                If dCurrentPrice = kNullData Then
                    .TextMatrix(lRow, PositionCol(eGDPositionCols_CurrentPrice)) = ""
                    CurrencyToGrid fgPositions, lRow, PositionCol(eGDPositionCols_Profit), ""
                Else
                    .TextMatrix(lRow, PositionCol(eGDPositionCols_CurrentPrice)) = FormatGenesisPrice(Str(dCurrentPrice), strGenesisSymbol)
                    CurrencyToGrid fgPositions, lRow, PositionCol(eGDPositionCols_Profit), Str(dProfit)
                End If
            End If
        End If
        
        CalcTotalOpenEquityPositions
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.RefreshPositionPrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotalOpenEquityPositions
'' Description: Calculate the total open equity for the positions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcTotalOpenEquityPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dTotalOpenEquity As Double      ' Total Open equity

    dTotalOpenEquity = 0#
    With fgPositions
        For lIndex = .FixedRows To .Rows - 2
            If Len(.TextMatrix(lIndex, PositionCol(eGDPositionCols_Profit))) > 0 Then
                dTotalOpenEquity = dTotalOpenEquity + ValOfText(.TextMatrix(lIndex, PositionCol(eGDPositionCols_Profit)))
            End If
        Next lIndex
        
        CurrencyToGrid fgPositions, .Rows - 1, PositionCol(eGDPositionCols_Profit), Str(dTotalOpenEquity)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CalcTotalOpenEquityPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotalsTrades
'' Description: Calculate the total open equity / closed profit for the trades grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcTotalsTrades()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dTotalOpenEquity As Double      ' Total open equity
    Dim dTotalClosedProfit As Double    ' Total closed profit

    dTotalOpenEquity = 0#
    dTotalClosedProfit = 0#
    
    With fgFills
        For lIndex = .FixedRows To .Rows - 2
            If Len(.TextMatrix(lIndex, FillCol(eGDFillCols_OpenEquity))) > 0 Then
                dTotalOpenEquity = dTotalOpenEquity + ValOfText(.TextMatrix(lIndex, FillCol(eGDFillCols_OpenEquity)))
            End If
            If Len(.TextMatrix(lIndex, FillCol(eGDFillCols_ClosedProfit))) > 0 Then
                dTotalClosedProfit = dTotalClosedProfit + ValOfText(.TextMatrix(lIndex, FillCol(eGDFillCols_ClosedProfit)))
            End If
        Next lIndex
        
        CurrencyToGrid fgFills, .Rows - 1, FillCol(eGDFillCols_OpenEquity), Str(dTotalOpenEquity)
        CurrencyToGrid fgFills, .Rows - 1, FillCol(eGDFillCols_ClosedProfit), Str(dTotalClosedProfit)
    End With
    
    m.dTotalOpenEquity = dTotalOpenEquity
    m.dTotalClosedProfit = dTotalClosedProfit
    UpdateAccountDetails

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.CalcTotalsTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateToDouble
'' Description: Convert a broker date to a double
'' Inputs:      Date to Convert
'' Returns:     "Julian" date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DateToDouble(ByVal strDateToConvert As String) As Double
On Error GoTo ErrSection:

    Dim BrokerKv As New cBrokerKeyValue ' Broker Key/Value object
    
    DateToDouble = BrokerKv.DateToDouble(strDateToConvert)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillToTrade
'' Description: Convert a fill into a trade
'' Inputs:      Fill
'' Returns:     Trade
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillToTrade(ByVal brokerMessage As cBrokerMessage) As cTrade
On Error GoTo ErrSection:

    Dim returnTrade As cTrade           ' Trade object to return
    
    Set returnTrade = New cTrade
    returnTrade.AccountNumber = brokerMessage("Account")
    returnTrade.BrokerSymbol = brokerMessage("Symbol")
    returnTrade.GenesisSymbol = brokerMessage("GenesisSymbol")
    returnTrade.EntryTime = DateToDouble(brokerMessage("FillDate"))
    returnTrade.EntryOrderID = brokerMessage("BrokerID")
    returnTrade.EntryFillID = brokerMessage("FillID")
    returnTrade.EntryIsBuy = (UCase(brokerMessage("Side")) = "BUY")
    returnTrade.EntryQuantity = CLng(Val(brokerMessage("Quantity")))
    returnTrade.EntryPriceBroker = Val(brokerMessage("FillPrice"))
    returnTrade.EntryPriceGenesis = Val(brokerMessage("GenesisFillPrice"))
    
    Set FillToTrade = returnTrade
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.FillToTrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddCarriedFillToTrades
'' Description: Add the given carried fill to the trades collection
'' Inputs:      Carried Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddCarriedFillToTrades(ByVal brokerMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol to use for lookup
    Dim TradesForSymbol As cGdTree      ' Trades for the given symbol
    
    If Len(brokerMessage("GenesisSymbol")) > 0 Then
        strSymbol = brokerMessage("GenesisSymbol")
    Else
        strSymbol = brokerMessage("Symbol")
    End If
    
    If m.Trades.Exists(strSymbol) Then
        Set TradesForSymbol = m.Trades(strSymbol)
        TradesForSymbol.Add FillToTrade(brokerMessage)
        Set m.Trades(strSymbol) = TradesForSymbol
    Else
        Set TradesForSymbol = New cGdTree
        TradesForSymbol.Add FillToTrade(brokerMessage)
        m.Trades.Add TradesForSymbol, strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AddCarriedFillToTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTodayFillToTrades
'' Description: Add the given today fill to the trades collection
'' Inputs:      Today Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddTodayFillToTrades(ByVal brokerMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol to use for lookup
    Dim TradesForSymbol As cGdTree      ' Trades for the given symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRemainingQuantity As Long      ' Remaining quantity on the fill
    Dim Trade As cTrade                 ' Trade from the collection
    Dim NewTrade As cTrade              ' New Trade to add to the collection
    
    If Len(brokerMessage("GenesisSymbol")) > 0 Then
        strSymbol = brokerMessage("GenesisSymbol")
    Else
        strSymbol = brokerMessage("Symbol")
    End If
    
    If m.Trades.Exists(strSymbol) Then
        Set TradesForSymbol = m.Trades(strSymbol)
        
        lRemainingQuantity = CLng(Val(brokerMessage("Quantity")))
        For lIndex = 1 To TradesForSymbol.Count
            Set Trade = TradesForSymbol(lIndex)
            If Trade.ExitQuantity = 0 Then
                If ((UCase(brokerMessage("Side")) = "BUY") And (Trade.EntryIsBuy = False)) Or ((UCase(brokerMessage("Side")) = "SELL") And (Trade.EntryIsBuy = True)) Then
                    If lRemainingQuantity = Trade.EntryQuantity Then
                        Trade.ExitTime = DateToDouble(brokerMessage("FillDate"))
                        Trade.ExitOrderID = brokerMessage("BrokerID")
                        Trade.ExitFillID = brokerMessage("FillID")
                        Trade.ExitIsBuy = (UCase(brokerMessage("Side")) = "BUY")
                        Trade.ExitQuantity = lRemainingQuantity
                        Trade.ExitPriceBroker = Val(brokerMessage("FillPrice"))
                        Trade.ExitPriceGenesis = Val(brokerMessage("GenesisFillPrice"))
                        
                        Set TradesForSymbol(lIndex) = Trade
                        lRemainingQuantity = 0
                        Exit For
                        
                    ElseIf lRemainingQuantity > Trade.EntryQuantity Then
                        Trade.ExitTime = DateToDouble(brokerMessage("FillDate"))
                        Trade.ExitOrderID = brokerMessage("BrokerID")
                        Trade.ExitFillID = brokerMessage("FillID")
                        Trade.ExitIsBuy = (UCase(brokerMessage("Side")) = "BUY")
                        Trade.ExitQuantity = Trade.EntryQuantity
                        Trade.ExitPriceBroker = Val(brokerMessage("FillPrice"))
                        Trade.ExitPriceGenesis = Val(brokerMessage("GenesisFillPrice"))
                        
                        Set TradesForSymbol(lIndex) = Trade
                        lRemainingQuantity = lRemainingQuantity - Trade.EntryQuantity
                    Else
                        Set NewTrade = Trade.MakeCopy
                        NewTrade.EntryQuantity = lRemainingQuantity
                        NewTrade.ExitTime = DateToDouble(brokerMessage("FillDate"))
                        NewTrade.ExitOrderID = brokerMessage("BrokerID")
                        NewTrade.ExitFillID = brokerMessage("FillID")
                        NewTrade.ExitIsBuy = (UCase(brokerMessage("Side")) = "BUY")
                        NewTrade.ExitQuantity = lRemainingQuantity
                        NewTrade.ExitPriceBroker = Val(brokerMessage("FillPrice"))
                        NewTrade.ExitPriceGenesis = Val(brokerMessage("GenesisFillPrice"))
                        
                        Set TradesForSymbol(lIndex) = NewTrade
                        
                        Trade.EntryQuantity = Trade.EntryQuantity - lRemainingQuantity
                        TradesForSymbol.Add Trade, , lIndex
                        
                        lRemainingQuantity = 0
                        Exit For
                    End If
                End If
            End If
        Next lIndex
        
        If lRemainingQuantity > 0 Then
            Set NewTrade = FillToTrade(brokerMessage)
            NewTrade.EntryQuantity = lRemainingQuantity
            TradesForSymbol.Add NewTrade
        End If
    Else
        Set TradesForSymbol = New cGdTree
        TradesForSymbol.Add FillToTrade(brokerMessage)
        m.Trades.Add TradesForSymbol, strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.AddTodayFillToTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradesToGrid
'' Description: Fill the trades grid from the trades collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TradesToGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Trades As cGdTree               ' Collection of trades
    
    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        AddFillsTotalRow
        
        For lIndex = 1 To m.Trades.Count
            Set Trades = m.Trades(lIndex)
            For lIndex2 = 1 To Trades.Count
                TradeToGrid Trades(lIndex2)
            Next lIndex2
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.TradesToGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeToGrid
'' Description: Send the given trade to the grid
'' Inputs:      Trade
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TradeToGrid(ByVal Trade As cTrade, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim Bars As cGdBars                 ' Bars object

    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            .RowPosition(.Rows - 1) = .Rows - 2
            lRow = .Rows - 2
        End If
        
        If Len(Trade.GenesisSymbol) > 0 Then
            Set Bars = GetBars(Trade.GenesisSymbol)
        Else
            Set Bars = New cGdBars
        End If
        
        .RowData(lRow) = Trade
        
        .TextMatrix(lRow, FillCol(eGDFillCols_BrokerSymbol)) = Trade.BrokerSymbol
        .TextMatrix(lRow, FillCol(eGDFillCols_GenesisSymbol)) = Trade.GenesisSymbol
        
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryTime)) = DateFormat(Trade.EntryTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryOrderID)) = Trade.EntryOrderID
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryFillID)) = Trade.EntryFillID
        If Trade.EntryIsBuy Then
            .TextMatrix(lRow, FillCol(eGDFillCols_EntrySide)) = "Bought"
        Else
            .TextMatrix(lRow, FillCol(eGDFillCols_EntrySide)) = "Sold"
        End If
        .TextMatrix(lRow, FillCol(eGDFillCols_EntryQuantity)) = Str(Trade.EntryQuantity)
        If Len(Trade.GenesisSymbol) = 0 Then
            .TextMatrix(lRow, FillCol(eGDFillCols_EntryPrice)) = Str(Trade.EntryPriceBroker)
        Else
            .TextMatrix(lRow, FillCol(eGDFillCols_EntryPrice)) = Bars.PriceDisplay(Trade.EntryPriceGenesis)
        End If
        
        If Trade.ExitQuantity = 0 Then
            .Cell(flexcpText, lRow, FillCol(eGDFillCols_ExitTime), lRow, FillCol(eGDFillCols_ExitPrice)) = ""
        Else
            .TextMatrix(lRow, FillCol(eGDFillCols_ExitTime)) = DateFormat(Trade.ExitTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER, True)
            .TextMatrix(lRow, FillCol(eGDFillCols_ExitOrderID)) = Trade.ExitOrderID
            .TextMatrix(lRow, FillCol(eGDFillCols_ExitFillID)) = Trade.ExitFillID
            If Trade.ExitIsBuy Then
                .TextMatrix(lRow, FillCol(eGDFillCols_ExitSide)) = "Bought"
            Else
                .TextMatrix(lRow, FillCol(eGDFillCols_ExitSide)) = "Sold"
            End If
            .TextMatrix(lRow, FillCol(eGDFillCols_ExitQuantity)) = Str(Trade.ExitQuantity)
            If Len(Trade.GenesisSymbol) = 0 Then
                .TextMatrix(lRow, FillCol(eGDFillCols_ExitPrice)) = Str(Trade.ExitPriceBroker)
            Else
                .TextMatrix(lRow, FillCol(eGDFillCols_ExitPrice)) = Bars.PriceDisplay(Trade.ExitPriceGenesis)
            End If
        End If
        
        RefreshTradePrices Bars, lRow
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.TradeToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradePrices
'' Description: Refresh the prices on the trades grid
'' Inputs:      Bars, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshTradePrices(Bars As cGdBars, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow > -1& Then
            RefreshTradePricesForRow Bars, lRow
        Else
            For lIndex = .FixedRows To .Rows - 2
                RefreshTradePricesForRow Bars, lIndex
            Next lIndex
        End If
        
        CalcTotalsTrades
                
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.RefreshTradePrices"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshTradePricesForRow
'' Description: Refresh the prices on the trades grid
'' Inputs:      Bars, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshTradePricesForRow(Bars As cGdBars, ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim Trade As cTrade                 ' Trade object
    
    strGenesisSymbol = Bars.Prop(eBARS_Symbol)
    With fgFills
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If TypeOf .RowData(lRow) Is cTrade Then
            Set Trade = .RowData(lRow)
            If Trade.GenesisSymbol = strGenesisSymbol Then
                If Bars(eBARS_Close, Bars.Size - 1) = kNullData Then
                    CurrencyToGrid fgFills, lRow, FillCol(eGDFillCols_OpenEquity), ""
                Else
                    CurrencyToGrid fgFills, lRow, FillCol(eGDFillCols_OpenEquity), Str(Trade.OpenProfit(Bars(eBARS_Close, Bars.Size - 1)))
                End If
                CurrencyToGrid fgFills, lRow, FillCol(eGDFillCols_ClosedProfit), Str(Trade.ClosedProfit)
            End If
        End If
                
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.RefreshTradePricesForRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateAccountDetails
'' Description: Update the account details in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateAccountDetails()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgAccountDetails
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        m.dAccountBalance = m.dPrevBalance + m.dTotalClosedProfit
        m.dNetLiquidity = m.dAccountBalance + m.dTotalOpenEquity
        
        CurrencyToGrid fgAccountDetails, 4, 1, Str(m.dPrevBalance)
        CurrencyToGrid fgAccountDetails, 5, 1, Str(m.dTotalClosedProfit)
        CurrencyToGrid fgAccountDetails, 6, 1, Str(m.dAccountBalance)
        CurrencyToGrid fgAccountDetails, 7, 1, Str(m.dTotalOpenEquity)
        CurrencyToGrid fgAccountDetails, 8, 1, Str(m.dTotalClosedProfit + m.dTotalOpenEquity)
        CurrencyToGrid fgAccountDetails, 9, 1, Str(m.dNetLiquidity)
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.UpdateAccountDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearAccount
'' Description: Clear the grids and the accounts combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearAccount()
On Error GoTo ErrSection:

    m.bSetAccountFromCode = True
    cboAccounts.ListIndex = -1&
    cboAccounts.Clear
    
    InitAccountDetailsGrid
    m.dPrevBalance = 0#
    m.dAccountBalance = 0#
    m.dNetLiquidity = 0#
    m.dTotalClosedProfit = 0#
    m.dTotalOpenEquity = 0#
    
    fgOrders.Rows = fgOrders.FixedRows
    
    fgFills.Rows = fgFills.FixedRows
    AddFillsTotalRow
    
    fgPositions.Rows = fgPositions.FixedRows
    AddPositionsTotalRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.ClearAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToLog
'' Description: Add the given message to the appropriate log file
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddToLog(ByVal strMessage As String)
On Error GoTo ErrSection:

    g.Broker.BrokerDebug m.nBroker, strMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateTodayNumbers
'' Description: Update the "TodayBuys" and "TodaySells" numbers
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateTodayNumbers(ByVal brokerMessage As cBrokerMessage)
On Error GoTo ErrSection:

    Dim lFillQuantity As Long           ' Fill quantity
    Dim bIsBuy As Boolean               ' Is this fill a buy?
    Dim strSymbol As String             ' Symbol for the record
    
    strSymbol = brokerMessage("Symbol")
    lFillQuantity = CLng(Val(brokerMessage("Quantity")))
    If Len(brokerMessage("Side")) > 0 Then
        bIsBuy = (Left(UCase(brokerMessage("Side")), 3) = "BUY")
    Else
        bIsBuy = (lFillQuantity >= 0)
    End If
        
    If bIsBuy Then
        If m.NumBuysToday.Exists(strSymbol) Then
            m.NumBuysToday(strSymbol) = m.NumBuysToday(strSymbol) + lFillQuantity
        Else
            m.NumBuysToday(strSymbol) = lFillQuantity
        End If
    Else
        If m.NumSellsToday.Exists(strSymbol) Then
            m.NumSellsToday(strSymbol) = m.NumSellsToday(strSymbol) + Abs(lFillQuantity)
        Else
            m.NumSellsToday(strSymbol) = Abs(lFillQuantity)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerView.UpdateTodayNumbers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedAccountNumber
'' Description: Determine the selected account number
'' Inputs:      None
'' Returns:     Account Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedAccountNumber() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim brokerMessage As cBrokerMessage ' Broker message from the accounts collection

    strReturn = ""
    Set brokerMessage = SelectedAccount
    If Not brokerMessage Is Nothing Then
        strReturn = brokerMessage("Account")
    End If
    
    SelectedAccountNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.SelectedAccountNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedAccount
'' Description: Determine the selected account
'' Inputs:      None
'' Returns:     Account
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedAccount() As cBrokerMessage
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message from the accounts collection

    Set brokerMessage = Nothing
    If cboAccounts.ListIndex > -1& Then
        If m.Accounts.Exists(cboAccounts.ItemData(cboAccounts.ListIndex)) Then
            Set brokerMessage = m.Accounts(cboAccounts.ItemData(cboAccounts.ListIndex))
        End If
    End If
    
    Set SelectedAccount = brokerMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerView.SelectedAccount"
    
End Function

