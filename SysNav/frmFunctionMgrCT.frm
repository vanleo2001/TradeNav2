VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFunctionMgrCT 
   Caption         =   "Function Manager"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   9060
   Begin HexUniControls.ctlUniFrameWL fraFgButtons 
      Height          =   375
      Left            =   3360
      TabIndex        =   28
      Top             =   5160
      Width           =   5295
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
      Caption         =   "frmFunctionMgrCT.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFunctionMgrCT.frx":0038
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFunctionMgrCT.frx":0058
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdFgDown 
         Height          =   375
         Left            =   4020
         TabIndex        =   0
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
         Caption         =   "frmFunctionMgrCT.frx":0074
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":00A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":00C6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFgUp 
         Height          =   375
         Left            =   2700
         TabIndex        =   2
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
         Caption         =   "frmFunctionMgrCT.frx":00E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":0110
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0130
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFgDelete 
         Height          =   375
         Left            =   1380
         TabIndex        =   7
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
         Caption         =   "frmFunctionMgrCT.frx":014C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":0178
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0198
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFgAdd 
         Height          =   375
         Left            =   60
         TabIndex        =   29
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
         Caption         =   "frmFunctionMgrCT.frx":01B4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":01E2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0202
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgSpread 
      Height          =   585
      Left            =   2700
      TabIndex        =   9
      Top             =   5340
      Visible         =   0   'False
      Width           =   1905
      _cx             =   3360
      _cy             =   1032
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7LCtl.VSFlexGrid fgScoring 
      Height          =   585
      Left            =   60
      TabIndex        =   21
      Top             =   5340
      Visible         =   0   'False
      Width           =   1905
      _cx             =   3360
      _cy             =   1032
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
   Begin VSFlex7LCtl.VSFlexGrid fgPlanet 
      Height          =   585
      Left            =   1740
      TabIndex        =   25
      Top             =   5340
      Visible         =   0   'False
      Width           =   1905
      _cx             =   3360
      _cy             =   1032
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   8400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   7
      DisplayContextMenu=   0   'False
      Tools           =   "frmFunctionMgrCT.frx":021E
      ToolBars        =   "frmFunctionMgrCT.frx":0473
   End
   Begin HexUniControls.ctlUniFrameWL fraFunctionInfo 
      Height          =   1815
      Left            =   100
      TabIndex        =   16
      Top             =   100
      Width           =   8835
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
      Caption         =   "frmFunctionMgrCT.frx":0617
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFunctionMgrCT.frx":0643
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFunctionMgrCT.frx":0663
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtCodedText 
         DataField       =   "Description"
         Height          =   720
         Left            =   1380
         TabIndex        =   4
         Top             =   420
         Width           =   7890
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmFunctionMgrCT.frx":067F
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
         MultiLine       =   -1  'True
         Alignment       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmFunctionMgrCT.frx":069F
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":06BF
      End
      Begin HexUniControls.ctlUniCheckXP chkAutoMultiplier 
         Height          =   195
         Left            =   5760
         TabIndex        =   26
         Top             =   60
         Visible         =   0   'False
         Width           =   2055
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
         Caption         =   "frmFunctionMgrCT.frx":06DB
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":0719
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0739
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkAdvanced 
         Height          =   195
         Left            =   3720
         TabIndex        =   15
         Top             =   45
         Width           =   2055
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
         Caption         =   "frmFunctionMgrCT.frx":0755
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":079F
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":07BF
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtDesc 
         DataField       =   "Description"
         Height          =   720
         Left            =   900
         TabIndex        =   3
         Top             =   435
         Width           =   7890
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmFunctionMgrCT.frx":07DB
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
         MultiLine       =   -1  'True
         Alignment       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmFunctionMgrCT.frx":07FB
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":081B
      End
      Begin HexUniControls.ctlUniComboImageXP cboCategory 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   -15
         Width           =   2550
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
         Tip             =   "frmFunctionMgrCT.frx":0837
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0857
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRichTextBoxXP TradeSense 
         Height          =   570
         Left            =   900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1215
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   1005
         BackColor       =   12632256
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmFunctionMgrCT.frx":0873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   -1
         MultiLine       =   -1  'True
         Alignment       =   0
         ScrollBars      =   3
         PasswordChar    =   ""
         TrapTab         =   0   'False
         RaiseChangeEvent=   -1  'True
         RaiseUpdateEvent=   0   'False
         RaiseSelChangeEvent=   -1  'True
         Tip             =   "frmFunctionMgrCT.frx":0893
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":08B3
         ViewMode        =   0
         TextModeText    =   2
         TextModeUndoLevel=   8
         TextModeCodePage=   32
         AutoURLDetect   =   0   'False
         FileName        =   ""
         VerticalLayout  =   0   'False
         OnlyNumbers     =   0   'False
         NoIME           =   0   'False
         SelfIME         =   0   'False
         LanguageOptions =   150
         RaiseRequestResizeEvent=   0   'False
         RaiseMsgFilterEvent=   0   'False
         SubClassPaintMessage=   0   'False
         TabSize         =   4
         TypographyOptions=   0
         BlockAutoCopy   =   0   'False
         BlockAutoCut    =   0   'False
         BlockAutoPaste  =   0   'False
         BlockAutoUndo   =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   210
         Index           =   0
         Left            =   0
         Top             =   1215
         Width           =   900
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
         Caption         =   "frmFunctionMgrCT.frx":08CF
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":08F9
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0919
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Index           =   5
         Left            =   0
         Top             =   30
         Width           =   720
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
         Caption         =   "frmFunctionMgrCT.frx":0935
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":0965
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0985
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   210
         Index           =   6
         Left            =   0
         Top             =   435
         Width           =   900
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
         Caption         =   "frmFunctionMgrCT.frx":09A1
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFunctionMgrCT.frx":09D7
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":09F7
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   2760
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4868
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
      FrontTabForeColor=   -2147483635
      Caption         =   "&Function|&Inputs|&Advanced"
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
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin HexUniControls.ctlUniFrameWL fraAdvanced 
         Height          =   2385
         Left            =   9720
         TabIndex        =   19
         Top             =   330
         Width           =   8685
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
         Caption         =   "frmFunctionMgrCT.frx":0A13
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFunctionMgrCT.frx":0A3F
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0A5F
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtRequiredMod 
            Height          =   315
            Left            =   1740
            TabIndex        =   27
            Top             =   1020
            Width           =   5715
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmFunctionMgrCT.frx":0A7B
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
            Tip             =   "frmFunctionMgrCT.frx":0A9B
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgrCT.frx":0ABB
         End
         Begin HexUniControls.ctlUniFrameWL fraUsage 
            Height          =   675
            Left            =   180
            TabIndex        =   20
            Top             =   180
            Width           =   4515
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
            Caption         =   "frmFunctionMgrCT.frx":0AD7
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmFunctionMgrCT.frx":0B01
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgrCT.frx":0B21
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkMM 
               Height          =   195
               Left            =   2040
               TabIndex        =   14
               Top             =   60
               Visible         =   0   'False
               Width           =   1395
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
               Caption         =   "frmFunctionMgrCT.frx":0B3D
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgrCT.frx":0B71
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgrCT.frx":0B91
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkCriteria 
               Height          =   195
               Left            =   2880
               TabIndex        =   13
               Top             =   300
               Width           =   1395
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
               Caption         =   "frmFunctionMgrCT.frx":0BAD
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgrCT.frx":0BE9
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgrCT.frx":0C09
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkCharting 
               Height          =   195
               Left            =   1800
               TabIndex        =   12
               Top             =   300
               Width           =   1095
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
               Caption         =   "frmFunctionMgrCT.frx":0C25
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgrCT.frx":0C55
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgrCT.frx":0C75
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkSystemTesting 
               Height          =   195
               Left            =   240
               TabIndex        =   11
               Top             =   300
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
               Caption         =   "frmFunctionMgrCT.frx":0C91
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgrCT.frx":0CD1
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgrCT.frx":0CF1
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblRequiredMod 
            Height          =   255
            Left            =   180
            Top             =   1050
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
            Caption         =   "frmFunctionMgrCT.frx":0D0D
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgrCT.frx":0D4F
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgrCT.frx":0D6F
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraInputs 
         Height          =   2385
         Left            =   9420
         TabIndex        =   18
         Top             =   330
         Width           =   8685
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
         Caption         =   "frmFunctionMgrCT.frx":0D8B
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFunctionMgrCT.frx":0DB7
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":0DD7
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraMoveButtons 
            Height          =   375
            Left            =   3240
            TabIndex        =   22
            Top             =   1980
            Width           =   2655
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
            Caption         =   "frmFunctionMgrCT.frx":0DF3
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmFunctionMgrCT.frx":0E1F
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgrCT.frx":0E3F
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdMoveDown 
               Height          =   375
               Left            =   1440
               TabIndex        =   24
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
               Caption         =   "frmFunctionMgrCT.frx":0E5B
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgrCT.frx":0E8F
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgrCT.frx":0EAF
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdMoveUp 
               Height          =   375
               Left            =   0
               TabIndex        =   23
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
               Caption         =   "frmFunctionMgrCT.frx":0ECB
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmFunctionMgrCT.frx":0EFB
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmFunctionMgrCT.frx":0F1B
               RightToLeft     =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid vsInputs 
            Height          =   1560
            Left            =   105
            TabIndex        =   10
            Top             =   345
            Width           =   8445
            _cx             =   14896
            _cy             =   2752
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
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   225
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   9225
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
            Caption         =   "frmFunctionMgrCT.frx":0F37
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgrCT.frx":1045
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgrCT.frx":1065
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraFunction 
         Height          =   2385
         Left            =   45
         TabIndex        =   17
         Top             =   330
         Width           =   8685
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
         Caption         =   "frmFunctionMgrCT.frx":1081
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmFunctionMgrCT.frx":10AD
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFunctionMgrCT.frx":10CD
         RightToLeft     =   0   'False
         Begin NavTradeSenseV3.Editor Editor1 
            Height          =   1905
            Left            =   100
            TabIndex        =   8
            Top             =   360
            Width           =   8445
            _ExtentX        =   14896
            _ExtentY        =   3360
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   240
            Left            =   60
            Top             =   60
            Width           =   9225
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
            Caption         =   "frmFunctionMgrCT.frx":10E9
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmFunctionMgrCT.frx":1195
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmFunctionMgrCT.frx":11B5
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmFunctionMgrCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmFunctionMgrCT.frm
'' Description: Allows the user to edit a Coded Text Function
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/03/2009   DAJ         Trim editor text when passing off to validate
'' 04/29/2009   DAJ         Fixed overflow error in ConvertParmTokens
'' 08/11/2010   DAJ         Added in a backdoor way to show coded text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

'constants for planet combo grid
Private Const kBodyList = "#-999;(none)|#0;Sun|#2;Mercury|#3;Venus|#14;Earth|#1;Moon|#4;Mars|#5;Jupiter|#6;Saturn|#7;Uranus|#8;Neptune|#9;Pluto|#15;Chiron|#16;Pholus|#17;Ceres|#18;Pallas|#19;Juno|#20;Vesta|#10;Mean_Node|#11;True_Node|#12;Mean_Apog|#13;Oscu_Apog|#40;Cupido|#41;Hades|#42;Zeus|#43;Kronos|#44;Apollon|#45;Admetos|#46;Vulkanus|#47;Poseidon|#48;Isis|#49;Nibiru|#50;Harrington|#51;Neptune_Leverrier|#52;Neptune_Adams|#53;Pluto_Lowell|#54;Pluto_Pickering|#59;Carrington"
Private Const kSystemList = "#0;Geocentric|#1;Heliocentric|#2;Right Ascension|#3;Barycentric"
Private Const kValueList1 = "Longitude|Latitude|Distance|Speed|Acceleration|Aspect"
Private Const kValueList2 = "Rectascension|Declination|Distance|Speed|Acceleration|Aspect"
Private Const kPlanetCategory = 23
Private Const kPlanetCols = 9
Private Const kZeroColWidth = 500

'constants for function categories
Private Const kScoringCols = 3
Private Const kScoringCategory = 25
Private Const kSpreadCategory = 27

'constants for spread operators
Private Const kPlus = "plus"
Private Const kMinus = "minus"
Private Const kDivide = "divide"
Private Const kOpAll = "plus|minus|divide"
Private Const kOpPlusMinus = "plus|minus|"

Private Type mPrivate
    Function As cFunction
    FunctionCategories As cFunctionCategories
    Inputs As cInputs
    ListLoading As cListLoading
    ReturnValue As Variant
    Usage As Byte
    strCalledFrom As String
    bResavingAll As Boolean
    bSkipAutoIf As Boolean
        
    strName As String
    
    strConditions As String    'available boolean conditions for scoring grid
    strSaveCondtion As String
    
    nPrevColWidth As Long   ' used for custom extended column
    frmChartEditor As frmChartCfg
End Type
Private m As mPrivate

'Columns in planet grid
Private Enum ePLCols
    ePLCol_Use = 0
    ePLCol_Body1
    ePLCol_Body2
    ePLCol_PlanetSystem
    ePLCol_Value
    ePLCol_Weight
    ePLCol_Offset
    ePLCol_Harmonic
    ePLCol_Orb
End Enum

'columns in scoring grid
Private Enum eSCRCols
    eSCRCol_Use = 0
    eSCRCol_Points
    eSCRCol_Condition
End Enum

'Columns in the inputs grid
Private Enum eGDCols
    eGDCol_ParmID = 0
    eGDCol_InputName
    eGDCol_DefaultValue
    eGDCol_FromVal
    eGDCol_ToVal
    eGDCol_ParmTypeID
    eGDCol_ParmDesc
    eGDCol_Sort
    eGDCol_Req
    eGDCol_Expression
    eGDCol_NumCols
End Enum

'Usage masks
Private Enum eUsageMask
    eUsageMask_MM = 1
    eUsageMask_SystemTesting = 2
    eUsageMask_Charting = 4
    eUsageMask_Criteria = 8
End Enum

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function
Private Function UsageMask(ByVal lUsageMask As eUsageMask) As Long
    UsageMask = lUsageMask
End Function

Public Property Let CalledFrom(pData As String)
    m.strCalledFrom = pData
End Property

'1=Money Management functions, 2=System functions
Property Let Usage(pData As Byte)
    m.Usage = pData
End Property

Public Property Get ID() As Long
    ID = m.Function.FunctionID
End Property

Private Property Get Inputs() As cInputs
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim vDefault As Variant
    Dim tmpInputs As cInputs
    
    With vsInputs
        Set tmpInputs = New cInputs
        For lIndex = 1 To .Rows - 1
            If Not IsNumeric(.TextMatrix(lIndex, GDCol(eGDCol_DefaultValue))) Then
                vDefault = .TextMatrix(lIndex, GDCol(eGDCol_DefaultValue))
            Else
                vDefault = ConvertInputValue(.TextMatrix(lIndex, GDCol(eGDCol_DefaultValue)), .Cell(flexcpValue, lIndex, GDCol(eGDCol_ParmTypeID)))
            End If
            
            tmpInputs.Add "", lIndex, .TextMatrix(lIndex, GDCol(eGDCol_InputName)), _
                .TextMatrix(lIndex, GDCol(eGDCol_ParmDesc)), _
                .Cell(flexcpValue, lIndex, GDCol(eGDCol_ParmID)), _
                "", 0, 0, 0, 0, 0, 0, _
                .Cell(flexcpValue, lIndex, GDCol(eGDCol_ParmTypeID)), _
                vDefault, _
                .TextMatrix(lIndex, GDCol(eGDCol_Req)), _
                .Cell(flexcpValue, lIndex, GDCol(eGDCol_FromVal)), _
                .Cell(flexcpValue, lIndex, GDCol(eGDCol_ToVal)), 0, "", ""
        Next lIndex
    End With
        
    Set Inputs = tmpInputs
    
ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmFunctionMgrCT.Inputs.Get"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Intialize and Show the form
'' Inputs:      Function ID to load, Text for Editor, Function Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal lFunctionID As Long, _
    Optional ByVal strText As String = "", _
    Optional ByVal strName As String = "", _
    Optional ByVal lCategoryID As Long = -1&, _
    Optional ByVal bScoring As Boolean = False, _
    Optional ByRef frm As frmChartCfg = Nothing)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    Set m.frmChartEditor = frm
    
    If lFunctionID = 0 Then
        Add bScoring
        If strName <> "" Then m.strName = Trim(strName)
        If strText <> "" Then
            Editor1.Text = Trim(strText)
            Editor1.ExprIsFormatted = False
            If bScoring Then
                cboCategory.Text = "Scoring"
                LoadScoringGrid
            End If
            Editor1_Change
        End If
        If lCategoryID <> -1& Then
            For lIndex = 0 To cboCategory.ListCount - 1
                If cboCategory.ItemData(lIndex) = lCategoryID Then
                    cboCategory.ListIndex = lIndex
                    Exit For
                End If
            Next lIndex
        End If
        MoveFocus Editor1
    Else
        If Not LoadRec(lFunctionID) Then
            Unload Me
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    If Not frm Is Nothing Then
        If Not frm.Chart Is Nothing Then CenterFormOnChart Me, frm.Chart            '6499
    End If
    ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    ''Unload Me
    Exit Sub

ErrSection:
    Unload Me
    RaiseError "frmFunctionMgrCT.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResaveFunctions
'' Description: Resave (and reverify) the given functions
'' Inputs:      List of Function IDs
'' Returns:     True if all Successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ResaveFunctions(alFunctionIDs As cGdArray) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bContinue As Boolean            ' Do we want to continue through the loop?
    Dim strOldCoded As String           ' Coded text before the verify/save
    Dim strOldFormatted As String       ' Formatted text before the verify/save
    Dim strOldEnglish As String         ' English text before the verify/save
    
    ResaveLog "BEGIN Resaving Selected Functions"
    
    For lIndex = 0 To alFunctionIDs.Size - 1
        If LoadRec(alFunctionIDs(lIndex), False) Then
            bContinue = False
            
            InfBox "Resaving " & m.Function.FunctionName, , , "Resaving Functions", True
            
            strOldCoded = m.Function.CodedText
            strOldFormatted = m.Function.Formatted
            strOldEnglish = Trim(Editor1.Text)
            
            If Verify = True Then
                If Save("ID_Save", False) = True Then
                    bContinue = True
                    
                    If strOldCoded <> m.Function.CodedText Then
                        ResaveLog m.Function.FunctionName & ": Coded Text Changed" & vbCrLf & "Before: " & strOldCoded & vbCrLf & "After:  " & m.Function.CodedText
                    End If
                    
                    If strOldFormatted <> m.Function.Formatted Then
                        ResaveLog m.Function.FunctionName & ": Formatted Text Changed" & vbCrLf & "Before: " & strOldFormatted & vbCrLf & "After:  " & m.Function.Formatted
                    End If
                    
                    If Trim(UCase(strOldEnglish)) <> Trim(UCase(Editor1.Text)) Then
                        ResaveLog m.Function.FunctionName & ": English Text Changed" & vbCrLf & "Before: " & Trim(strOldEnglish) & vbCrLf & "After:  " & Trim(Editor1.Text)
                    End If
                Else
                    ResaveLog m.Function.FunctionName & ": Could not save function"
                End If
            Else
                ResaveLog m.Function.FunctionName & ": Could not verify function"
            End If
            
            If bContinue = False Then
                ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR
                Exit For
            End If
        End If
    Next lIndex
    
    InfBox ""
    ResaveLog "END Resaving Selected Functions"
    
    ResaveFunctions = bContinue

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmFunctionMgrCT.ResaveFunctions"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Gets the function manager ready to add a new function
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(Optional ByVal bScoring As Boolean = False)
On Error GoTo ErrSection:
    
    m.bSkipAutoIf = False ' True
    
    vsIndexTab1.CurrTab = 0
    tbToolbar.Tools("ID_Verify").Enabled = False
    ClearFunction
    
    Set m.Function = New cFunction
    With m.Function
        .FunctionID = 0
        .Load
        If Not bScoring Then cboCategory.Text = "Indicator"
        chkCharting.Value = vbChecked
        chkCriteria.Value = vbChecked
        chkSystemTesting.Value = vbChecked
    End With
    
    Set m.Inputs = New cInputs
    If Not bScoring Then
        vsInputs.Redraw = flexRDNone
        InitGrid
        LoadGrid
        vsInputs.Redraw = flexRDBuffered
    End If
    
    SetEditorCaption Me, "Function", ""
    
    TradeSense.Text = ""
    txtRequiredMod.Text = ""
    EnableToolbar False

    If Not bScoring Then AdvancedDisplay

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Add"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadRec
'' Description: Loads the function manager with the given function
'' Inputs:      Function ID to load
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadRec(pFunctionID As Long, Optional ByVal bAskPassword As Boolean = True) As Boolean
On Error GoTo ErrSection:
    
    Dim lUsage As Long
    Dim nCategory As Long
    
    Dim strSpreadExpr$, iDollarMultiplier&
    Dim Inputs As cInputs
        
    ClearFunction
    
    'Load the function...
    Set m.Function = New cFunction
    With m.Function
        .FunctionID = pFunctionID
        .Load

        If (bAskPassword = True) Or (IsIDE = False) Then
            If Not g.Security.CanEdit(.SecurityLevel, .Password) Then
                GoTo ErrExit:
            End If
        End If
        
        m.strName = .FunctionName
        txtDesc.Text = .Description
        If m.FunctionCategories.Found(.FunctionCategoryID) Then
            cboCategory.Text = _
                m.FunctionCategories.Item(CStr(.FunctionCategoryID)).FunctionCategory
        End If
        Editor1.TextRTF = m.Function.GetRTF(.Formatted)
        nCategory = .FunctionCategoryID
        
        lUsage = .Usage
        SetCtl chkMM, lUsage And UsageMask(eUsageMask_MM)
        SetCtl chkSystemTesting, lUsage And UsageMask(eUsageMask_SystemTesting)
        SetCtl chkCharting, lUsage And UsageMask(eUsageMask_Charting)
        SetCtl chkCriteria, lUsage And UsageMask(eUsageMask_Criteria)
        
        txtRequiredMod.Text = .RequiredMod
    End With
    
    SetEditorCaption Me, "Function", m.strName
          
    'Load inputs grid...
    If nCategory = kPlanetCategory Then
        Me.Icon = Picture16(ToolbarIcon("ID_PlanetData"), , True)
        LoadPlanetGrid
    ElseIf nCategory = kScoringCategory Then
        LoadScoringGrid
    ElseIf nCategory = kSpreadCategory Then
        strSpreadExpr = LoadSpreadGrid
        'set auto multiplier check box
        If Len(strSpreadExpr) > 0 Then
            Set Inputs = m.Function.Inputs
            If Not Inputs Is Nothing Then
                If Inputs.Item(Inputs.Count).ParmName = "Auto Multiplier" Then
                    If Inputs.Item(Inputs.Count).DefaultValue = 1 Then
                        'make sure that multipliers are really dollar multiplier before checking the box
                        iDollarMultiplier = IsDollarMultiplier(strSpreadExpr)
                    End If
                End If
            End If
            chkAutoMultiplier.Value = Abs(iDollarMultiplier)
        End If
    Else
        vsInputs.Redraw = flexRDNone
        InitGrid
        LoadGrid
        vsInputs.Redraw = flexRDBuffered
        RemoveCategory "Planet Combo"
    End If

    ShowParmLine TradeSense
    
    EnableToolbar False
    tbToolbar.Tools("ID_Verify").Enabled = False
    m.ReturnValue = LockWindowUpdate(0)
    
    AdvancedDisplay
    LoadRec = True
            
ErrExit:
    m.ReturnValue = LockWindowUpdate(0)
    Exit Function

ErrSection:
    RaiseError "frmFunctionMgrCT.LoadRec"
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the function to the database
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Save(ByVal strButton As String, Optional ByVal bAskPassword As Boolean = True) As Boolean
On Error GoTo ErrSection:
    
    Dim lUsage As Long
    Dim bSaveAs As Boolean
    Dim strNewName As String
    Dim strText As String
    Dim strError As String
    Dim lOldID As Long
    Dim lSpreadFuncId As Long
    
    ' Saving or editing a scoring function requires Platinum
    ' (let them see and work with it so they know what it is, but they can't save it)
    If UCase(Trim(cboCategory.Text)) = "SCORING" Then
        If Not HasPlatinum(True, "Saving or editing a Scoring Function") Then Exit Function
    End If
    
    ' Verify if we need to
    If tbToolbar.Tools("ID_Verify").Enabled Or cboCategory.Text = "Spread" Then         'aardvark 3938
        Verify
        If tbToolbar.Tools("ID_Verify").Enabled Then Exit Function
    End If
       
    ' Handle Rename/Save As
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        strText = "Save the current Function as..."
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    ElseIf strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Function as..."
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & m.strName & " #02 ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        strText = "Rename the current Function as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    Else
        strNewName = m.strName
    End If
    
    ' Verify that it is a good name
    Do While Len(Trim(strNewName)) > 0
        ' Strip out a colon if it exists in the name...
        If InStr(strNewName, ":") Then
            strNewName = Replace(strNewName, ":", "")
        End If
        ' TLB 11/23/2010: don't allow multiple consecutive spaces
        Do While InStr(strNewName, "  ") > 0
            strNewName = Replace(strNewName, "  ", " ")
        Loop
        strError = m.Function.ValidName(strNewName)
        If strError <> "" Then
            InfBox strError, "e", , "Error"
        ElseIf FunctionExists(strNewName) Then
            InfBox "'" & strNewName & "' already exists.", "e", , "Error"
        Else
            ' Name is OK so we can exit loop
            Exit Do
        End If
        strText = "Rename the Function as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
    Loop
    
    If Len(Trim(strNewName)) = 0 Then
        Exit Function 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    m.strName = Trim(strNewName)
    SetEditorCaption Me, "Function", m.strName

    If bSaveAs Then
        Set m.Function = New cFunction
        m.Function.FunctionID = 0
        Verify
    End If
    'If bSaveAs Then
    '    lOldID = m.Function.FunctionID
    '    Set m.Function = New cFunction
    '    m.Function.FunctionID = lOldID
    '    m.Function.Load
    '    m.Function.FunctionID = 0&
    'End If

    Screen.MousePointer = vbHourglass
        
    ' Validate function fields
    With m.Function
        
        ' Pass Inputs collection into Function class for validation.
        .Inputs = Inputs
    
        ' Check each input to see if any functions are specified in
        ' for the default value.  If yes, verify this string and updated
        ' the input token type to Series of numbers, or boolean
        ConvertParmTokens
        
        ' User must be authorized to save (don't prompt for new functions or if
        ' copying an existing function
        If m.Function.FunctionID = 0 Then
            .SecurityLevel = 0
            .CannotDelete = False
            .LibraryID = kSN_UserLibrary
            .Password = ""
        Else
            If (bAskPassword = True) Or (IsIDE = False) Then
                If Not g.Security.CanSave(.SecurityLevel, .Password) Then
                    GoTo ErrExit:
                End If
            End If
        End If
        
        ' Set values specific to BOTH builtin and user functions
        .FunctionName = m.strName
        If Len(Trim(txtDesc.Text)) = 0 Then
            .Description = " "
        Else
            .Description = Trim(txtDesc.Text)
        End If
        .Description = txtDesc.Text
        .FunctionCategoryID = GetCatID(cboCategory.Text)
        .CodedName = StripStr(.FunctionName, " ")
        .ImplementationTypeID = kSN_Custom

        ' Usage
        lUsage = 0
        If chkMM Then lUsage = lUsage Or UsageMask(eUsageMask_MM)
        If chkSystemTesting Then lUsage = lUsage Or UsageMask(eUsageMask_SystemTesting)
        If chkCharting Then lUsage = lUsage Or UsageMask(eUsageMask_Charting)
        If chkCriteria Then lUsage = lUsage Or UsageMask(eUsageMask_Criteria)
        .Usage = lUsage
        
        .RequiredMod = txtRequiredMod.Text
        
        ' Function rule...
        ShowParmLine TradeSense
        .TradeSenseUsage = TradeSense.Tag
        .Reverify = False
        ' update LastModified (unless skip file exists -- e.g. when updating master mdb)
        If .LastModified <= 0 Or Not FileExist(App.Path & "\lastmod.skp") Then
            .LastModified = Now()
        Else
            StatusMsg "LastModified not changed"
        End If
        .Save
        g.bDirtyLibrariesMDB = True
    End With
             
    'get spread function ID
    If m.Function.FunctionCategoryID = kSpreadCategory Then lSpreadFuncId = m.Function.FunctionID
            
    RefreshFunction m.Function
    RefreshReverify lSpreadFuncId
    Screen.MousePointer = vbDefault
    
    tbToolbar.Tools("ID_Verify").Enabled = False
    EnableToolbar False
    Save = True
    
    If Not m.frmChartEditor Is Nothing Then
        Me.Hide
        m.frmChartEditor.NewFunctionAdded Editor1.Text, m.Function.FunctionName, m.Function.ReturnTypeID
    End If
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Function

ErrSection:
    Screen.MousePointer = vbDefault
    Select Case m.Function.ErrNbr
        Case 4: Editor1.SetFocus
        Case 5: cboCategory.SetFocus
    End Select
    RaiseError "frmFunctionMgrCT.Save"
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConvertParmTokens
'' Description: Convert each of the parameter tokens
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ConvertParmTokens()
On Error GoTo ErrSection:

    Dim strFirstChar    As String
    Dim strDefaultVal   As String
    Dim lPos            As Long
    Dim strCodedTextNew As String
    Dim lIndex          As Long
    Dim strSearch       As String
    Dim strToken        As String
    
    With m.Function
        For lIndex = 1 To .Inputs.Count
            strDefaultVal = Trim(.Inputs.Item(lIndex).DefaultValue)
            If Len(strDefaultVal) > 0 And _
               UCase(strDefaultVal) <> "MARKET1" And _
               UCase(strDefaultVal) <> "TRADES" Then
                        
                    With .Inputs.Item(lIndex)
                        'Fix parm type
#If 0 Then
                        strFirstChar = UCase(Left(strDefaultVal, 1))
                        If strFirstChar >= "A" And strFirstChar <= "Z" Then
                            'parm is a series (array, expression)
                            If .ParmTypeID = 1 Then
                                .ParmTypeID = 4
                            ElseIf .ParmTypeID = 6 Then
                                .ParmTypeID = 3
                            End If
                        Else
                            'parm is a constant
                            m.Function.Inputs.Item(lIndex).DefaultValue = Str(Val(strDefaultVal))
                            If .ParmTypeID = 4 Then
                                .ParmTypeID = 1
                            ElseIf .ParmTypeID = 3 Then
                                .ParmTypeID = 6
                            End If
                        End If
#End If
                        'Get new token
                        Select Case .ParmTypeID
                            Case 1 'numeric constant
                                strToken = "05"
                            Case 4 'numeric series (array)
                                strToken = "27"
                            Case 6 'boolean constant
                                strToken = "06"
                            Case 3 'boolean series (array)
                                strToken = "28"
                        End Select
                    End With
                    
                    'Search for the Token for this input and set to
                    'correct token type based on (new) parm type.
                    '(need to search for Length+Name so "Var" won't match with "VarA")
                    strSearch = Format(Len(.Inputs.Item(lIndex).ParmName), "000") _
                                & .Inputs.Item(lIndex).ParmName
                    lPos = 1
                    Do Until lPos > Len(.CodedText)
                        lPos = InStr(lPos, .CodedText, strSearch)
                        If lPos = 0 Then Exit Do
                        If Mid(.CodedText, lPos - 2, 2) <> strToken Then
                            strCodedTextNew = .CodedText
                            Mid(strCodedTextNew, lPos - 2, 2) = strToken
                            .CodedText = strCodedTextNew
                        End If
                        lPos = lPos + 1
                    Loop
                    
                    'Attempt to verify default value.
                    'cExpression.mode = 3       'Autodetect type
                    'cExpression.ValidateRule strDefaultVal
                    'If cExpression.ProcessedOK Then
                        '1=true/false, 2=numeric
                    'End If
                    'MsgBox "OK:" & cExpression.ProcessedOK & _
                          " Mode:" & cExpression.mode
                    
                End If
        Next lIndex
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.ConvertParmTokens"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCatID
'' Description: Get the category ID of a function
'' Inputs:      Function to get the ID for
'' Returns:     Category ID for the function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCatID(ByVal strName As String) As Long
On Error GoTo ErrSection:

    Dim lIndex       As Integer
    
    GetCatID = 0
    For lIndex = 1 To m.FunctionCategories.Count
        With m.FunctionCategories.Item(lIndex)
            If .FunctionCategory = strName Then
                GetCatID = .FunctionCategoryID
                Exit For
            End If
        End With
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionMgrCT.GetCatID"
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowParmLine
'' Description: Show the parameter line
'' Inputs:      Rich Text Box to show the line in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowParmLine(rtbTradeSense As RichTextBox)
On Error GoTo ErrSection:
    
    Dim strTradeSenseText As String
    Dim lLeftParen As Long
    Dim lTextLen As Long
    Dim lIndex As Long
    Dim nInputs As Byte
    
    'Default the function name to beginning of TradeSense text
    If Len(Trim(m.strName)) = 0 Then
        strTradeSenseText = "New Function ("
    Else
        strTradeSenseText = m.strName & " ("
    End If
    lLeftParen = 0
    nInputs = 0
    
    If vsInputs.Rows > 1 Then
        For lIndex = 1 To vsInputs.Rows - 1
            If vsInputs.TextMatrix(lIndex, GDCol(eGDCol_InputName)) <> "Trades" And _
                vsInputs.TextMatrix(lIndex, GDCol(eGDCol_ParmTypeID)) <> "5" Then
               'UCase(vsInputs.TextMatrix(lIndex, GDCol(eGDCol_InputName))) <> "MARKET1" Then
                strTradeSenseText = strTradeSenseText + _
                    Trim(vsInputs.TextMatrix(lIndex, GDCol(eGDCol_InputName))) + ", "
                nInputs = nInputs + 1
            End If
        Next lIndex
    End If
    
    'Add parm right paren to end of string
    If nInputs > 0 Then
        strTradeSenseText = Left(strTradeSenseText, Len(strTradeSenseText) - 2) + ")"
        lLeftParen = InStr(1, strTradeSenseText, "(")
        lTextLen = Len(strTradeSenseText) - Len(Trim(m.strName))
    Else
        strTradeSenseText = Left(strTradeSenseText, Len(strTradeSenseText) - 2)
    End If
    
    'Simulate text entered into RTF box...
    With rtbTradeSense
        .Tag = strTradeSenseText
        .Text = strTradeSenseText
        If lLeftParen > 0 Then
            .SelStart = lLeftParen
            .SelLength = lTextLen
            .SelItalic = True
            .SelLength = 0
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.ShowParmLine"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCategory_Click
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCategory_Click()
On Error GoTo ErrSection:
    
    Dim bGrid As Boolean
        
    If cboCategory.Text = "Planet Combo" Then
        Me.Icon = Picture16(ToolbarIcon("ID_PlanetData"), , True)
        EnableScoringGrid False
        EnablePlanetGrid True
        EnableSpreadGrid False
        bGrid = True
    ElseIf cboCategory.Text = "Scoring" Then
        EnableScoringGrid True
        EnablePlanetGrid False
        EnableSpreadGrid False
        bGrid = True
    ElseIf cboCategory.Text = "Spread" Then
        EnableScoringGrid False
        EnablePlanetGrid False
        EnableSpreadGrid True
        AddBlankRow True
        bGrid = True
    Else
        Me.Icon = Picture16(ToolbarIcon("ID_Functions"), , True)
        EnableScoringGrid False
        EnablePlanetGrid False
        EnableSpreadGrid False
        
        Label1(0).Visible = True   'Usage
        TradeSense.Visible = True
        chkAdvanced.Visible = True
        chkAutoMultiplier.Visible = False
        vsIndexTab1.Visible = True
        fraFgButtons.Visible = False
        
        EnableToolbar True
    End If
    
    If bGrid Then FormResize Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.cboCategory_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAdvanced_Click
'' Description: Show/Hide the advanced stuff as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAdvanced_Click()
On Error GoTo ErrSection:

    AdvancedDisplay
    SetIniFileProperty "Advanced", chkAdvanced.Value, "FuncMgrCT", g.strIniFile
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.chkAdvanced_Click"
    Resume ErrExit
    
End Sub

Private Sub chkAutoMultiplier_Click()
On Error GoTo ErrSection:

    Dim bError As Boolean
    
    bError = ToggleAutoMultiplier(fgSpread, chkAutoMultiplier.Value)
    
    If Not bError Then
        tbToolbar.Tools("ID_Verify").Enabled = True
        tbToolbar.Tools("ID_Save").Enabled = True
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.chkAutoMultiplier_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCharting_Click
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCharting_Click()
On Error GoTo ErrSection:

    If Me.Visible Then Verify True
    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.chkCharting_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCriteria_Click
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCriteria_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then Verify True
    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.chkCriteria_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkSystemTesting_Click
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkSystemTesting_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then Verify True
    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.chkSystemTesting_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to Print the function
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "SNV FunctionCT", Me, 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.PrintMe"
    
End Sub

Private Sub cmdFgAdd_Click()
On Error GoTo ErrSection:

    Dim i&

    If fgPlanet.Visible Then
        With fgPlanet
            .AddItem "", .Row
            .Row = .Row
            .Cell(flexcpChecked, .Row, ePLCol_Use) = flexChecked
            .TextMatrix(.Row, ePLCol_Body1) = -999
            .EditCell
            SendKeys "{F4}" '(to dropdown the combo list)
        End With
    ElseIf fgScoring.Visible Then
        With fgScoring
            .AddItem "", .Row
            .Row = .Row
            .Cell(flexcpChecked, .Row, eSCRCol_Use) = flexChecked
            .TextMatrix(.Row, eSCRCol_Points) = "1"
            .TextMatrix(.Row, eSCRCol_Condition) = ""
            .Col = eSCRCol_Use
            fgScoring_BeforeEdit .Row, .Col, False
        End With
    ElseIf fgSpread.Visible Then
        i = AddBlankRow
        If i > 0 Then
            If fgSpread.TextMatrix(i, 0) = kDivide Then
                InfBox "A ratio spread can only have 2 symbols.", "I", , Me.Caption
            ElseIf Not fgSpread.MergeRow(i) Then
                InfBox "One or more items in row " & Str(i) & " is blank." & vbCrLf & _
                       "Please complete this row before " & vbCrLf & "adding a new one.", _
                       "I", , Me.Caption
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.cmdFgAdd_Click"
    
End Sub

Private Sub cmdFgDelete_Click()
On Error GoTo ErrSection:
    
    If fgPlanet.Visible Then
        With fgPlanet
            .RemoveItem .Row
            If .Rows = 2 Then cmdFgDelete.Enabled = False
        End With
    ElseIf fgScoring.Visible Then
        With fgScoring
            .RemoveItem .Row
            If .Rows = 2 Then cmdFgDelete.Enabled = False
        End With
    ElseIf fgSpread.Visible Then
        With fgSpread
            If .Rows > .FixedRows + 1 Then
                If .Row >= .FixedRows And .Row <= .Rows - 1 Then
                    .RemoveItem .Row
                    .TextMatrix(.FixedRows, 0) = ""
                    .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                End If
            End If
            If .Rows = 2 Then
                cmdFgDelete.Enabled = False
                cmdFgAdd.Enabled = True
            End If
        End With
    Else
        Exit Sub
    End If
    
    tbToolbar.Tools("ID_Verify").Enabled = True
    tbToolbar.Tools("ID_Save").Enabled = True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.cmdFgDelete_Click"

End Sub

Private Sub cmdFgDown_Click()
On Error GoTo ErrSection:

    Dim fg As VSFlexGrid
    
    If fgPlanet.Visible Then
        Set fg = fgPlanet
    ElseIf fgScoring.Visible Then
        Set fg = fgScoring
    Else
        Exit Sub
    End If
    
    If fg Is Nothing Then
        Exit Sub
    ElseIf fg.Row >= fg.Rows - fg.FixedRows Then
        Exit Sub  'precautionary, should never happen
    End If
        
    With fg
        .RowPosition(.Row) = .Row + 1
        .Row = .Row + 1
    End With
    
    EnableFgButtons fg.Row

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.cmdFgDown_Click"

End Sub

Private Sub cmdFgUp_Click()
On Error GoTo ErrSection:

    Dim fg As VSFlexGrid
    
    If fgPlanet.Visible Then
        Set fg = fgPlanet
    ElseIf fgScoring.Visible Then
        Set fg = fgScoring
    Else
        Exit Sub
    End If
    
    If fg Is Nothing Then
        Exit Sub
    ElseIf fg.Row <= 1 Then
        Exit Sub
    End If
    
    With fg
        .RowPosition(.Row) = .Row - 1
        .Row = .Row - 1
    End With
    
    EnableFgButtons fg.Row

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.cmdFgUp_Click"

End Sub

Private Sub cmdMoveDown_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim bOptional As Boolean

    With vsInputs
        .RowPosition(.Row) = .Row + 1
        .Row = .Row + 1
        .RowSel = .Row
        
        bOptional = False
        For lIndex = .FixedRows To .Rows - 1
            If Not CheckedCell(vsInputs, lIndex, GDCol(eGDCol_Req)) Then
                bOptional = True
            ElseIf bOptional Then
                CheckedCell(vsInputs, lIndex, GDCol(eGDCol_Req)) = False
            End If
        Next lIndex
        
        MoveFocus vsInputs
    End With
    
    EnableButtons
    ShowParmLine TradeSense
    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.cmdMoveDown_Click"
    Resume ErrExit
    
End Sub

Private Sub cmdMoveUp_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim bOptional As Boolean

    With vsInputs
        .RowPosition(.Row) = .Row - 1
        .Row = .Row - 1
        .RowSel = .Row
        
        bOptional = False
        For lIndex = .FixedRows To .Rows - 1
            If Not CheckedCell(vsInputs, lIndex, GDCol(eGDCol_Req)) Then
                bOptional = True
            ElseIf bOptional Then
                CheckedCell(vsInputs, lIndex, GDCol(eGDCol_Req)) = False
            End If
        Next lIndex
        
        MoveFocus vsInputs
    End With
    
    EnableButtons
    ShowParmLine TradeSense
    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.cmdMoveUp_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_Change
'' Description: Set the function dirty when the editor changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_Change()
On Error GoTo ErrSection:

    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = (Len(Trim(Editor1.Text)) > 0)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_Change"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_EditFunction
'' Description: Allow the user to edit a function from the coded text
'' Inputs:      ID and Name of Function, Whether it was Found
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    Err.Raise vbObjectError + 1000, , "Sub-Functions cannot be edited or added here"

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_EditFunction"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_KeyDown
'' Description: Make sure that a Tab key inside the editor changes focus
'' Inputs:      KeyCode of key pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    If KeyCode = 112 Then
        InfBox "F1 Function help not availble in this version", "i", , "Message"
        Exit Sub
    End If
    
    'Tab key pressed (or F6), Move focus to inputs grid
    If KeyCode = 9 Or KeyCode = 117 Then
        If Shift <> 0 Then
            'Shift-Tab, Move focus to functions grid
            txtDesc.SetFocus
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_KeyDown"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_GotFocus
'' Description: Intialize when the editor gets focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_GotFocus()
On Error GoTo ErrSection:
    
    Set g.ActiveEditor = Editor1
    InitEditor
    
    If Len(Trim(Editor1.Text)) = 0 And Not m.bSkipAutoIf Then
        Editor1.Text = ""
        SendKeys " "
    End If
    
    m.bSkipAutoIf = False
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_GotFocus"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_KeyUp
'' Description: Show the tree if the the correct keys are pressed
'' Inputs:      KeyCode of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    'Alt-1 to show rule tree
    If KeyCode = 49 And Shift = 4 Then
        VerifyFunctionDebug
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_KeyUp"
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_LostFocus
'' Description: Get rid of the TradeSense upon losing focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_LostFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Nothing
    Editor1.RemoveTradeSense

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_LostFocus"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_NewFunction
'' Description: Allow the user to create a new function
'' Inputs:      Category ID the Function List form was currently on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim frm As frmFunctionMgrCT         ' New Function Manager form
    
    Set frm = New frmFunctionMgrCT
    frm.ShowMe 0&, , , lCategoryID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.Editor1_NewFunction"
    Resume ErrExit
    
End Sub

Private Sub fgPlanet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    EnablePlanetGrid True
    With fgPlanet
        If .Col + 1 < .Cols Then
            .Col = .Col + 1
        Else
            .Col = 0
        End If
    End With
    
    tbToolbar.Tools("ID_Verify").Enabled = True
    tbToolbar.Tools("ID_Save").Enabled = True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.fgPlanet_AfterEdit"

End Sub

Private Sub fgPlanet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim i&
    
    'user must select a valid body1 first
    If fgPlanet.TextMatrix(fgPlanet.Row, ePLCol_Body1) = "-999" And fgPlanet.Col <> ePLCol_Body1 Then
        Cancel = True
        If Col = ePLCol_Use Then
            CheckedCell(fgPlanet, Row, Col) = True
            fgPlanet.Col = ePLCol_Body1
            fgPlanet.EditCell
            SendKeys "{F4}" '(to dropdown the combo list)
        End If
        Exit Sub
    End If
    
    'check orb edit
    If fgPlanet.Col = ePLCol_Orb And fgPlanet.TextMatrix(fgPlanet.Row, ePLCol_Value) <> "Aspect" Then
        MsgBox "Orb is used only if the 'Value' field is 'Aspect'."
        fgPlanet.SetFocus
        fgPlanet.Col = ePLCol_Orb - 1
        Cancel = True
        Exit Sub
    End If
    
    With fgPlanet
        If Col = ePLCol_Value Then
            If .TextMatrix(Row, ePLCol_PlanetSystem) = "2" Then
                .ComboList = kValueList2
            Else
                .ComboList = kValueList1
            End If
        ElseIf Col = ePLCol_Body1 Or Col = ePLCol_Body2 Then
            .ColComboList(ePLCol_Body1) = kBodyList
        ElseIf Col = ePLCol_PlanetSystem Then
            .ColComboList(ePLCol_PlanetSystem) = kSystemList
        Else
            .ComboList = ""
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.fgPlanet_BeforeEdit"

End Sub

Private Sub fgPlanet_Click()
    EnableFgButtons fgPlanet.Row
End Sub

Private Sub fgPlanet_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    With fgPlanet
        If .Col = ePLCol_PlanetSystem Then
            If .ComboIndex >= 0 Then
                If Val(.ComboData(.ComboIndex)) = 2 Then
                    If .TextMatrix(.Row, ePLCol_Value) = "Longitude" Then
                        .TextMatrix(.Row, ePLCol_Value) = "Rectascension"
                    ElseIf .TextMatrix(.Row, ePLCol_Value) = "Latitude" Then
                        .TextMatrix(.Row, ePLCol_Value) = "Declination"
                    End If
                Else
                    If .TextMatrix(.Row, ePLCol_Value) = "Rectascension" Then
                        .TextMatrix(.Row, ePLCol_Value) = "Longitude"
                    ElseIf .TextMatrix(.Row, ePLCol_Value) = "Declination" Then
                        .TextMatrix(.Row, ePLCol_Value) = "Latitude"
                    End If
                End If
            End If
        End If
    End With
    
    FinishEdit = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.fgPlanet_ComboCloseUp"

End Sub

Private Sub fgScoring_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strCond$
                
    If Col = eSCRCol_Condition Then
        With fgScoring
            If .MouseCol = eSCRCol_Condition Then
                If .ComboIndex = 0 Or (.ComboIndex = 1 And Len(m.strSaveCondtion) > 0) Then
                        If .MouseCol = eSCRCol_Condition Then
                            EditScoringCondition Row, Col
                        End If
                ElseIf .ComboIndex > 0 Then
                    strCond = .ComboItem(.ComboIndex)
                    strCond = BooleanConditionText(strCond)
                    If Len(strCond) > 0 Then
                        .TextMatrix(.Row, eSCRCol_Condition) = strCond
                        fgScoring.Cell(flexcpChecked, Row, eSCRCol_Use) = flexChecked
                        If Len(.TextMatrix(.Row, eSCRCol_Points)) < 1 Then .TextMatrix(.Row, eSCRCol_Points) = "1"
                        EnableScoringGrid True, True
                    End If
                End If
            End If
        End With
    End If
    
    m.strSaveCondtion = ""
    tbToolbar.Tools("ID_Verify").Enabled = True
    tbToolbar.Tools("ID_Save").Enabled = True
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.fgScoring_AfterEdit"

End Sub

Private Sub fgScoring_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strCond$
    
    If fgScoring.TextMatrix(Row, eSCRCol_Condition) <> "<Custom Condition>" Then
        m.strSaveCondtion = fgScoring.TextMatrix(Row, eSCRCol_Condition)
    End If
    
    strCond = "<Custom Condition>|" & m.strSaveCondtion & "|" & m.strConditions
    
    With fgScoring
        .ColComboList(eSCRCol_Condition) = strCond

        If Col = ePLCol_Use Then
            If Len(.TextMatrix(Row, eSCRCol_Condition)) < 1 Then
                CheckedCell(fgScoring, Row, Col) = True
                .Col = eSCRCol_Condition
                .EditCell
                SendKeys "{F4}" '(to dropdown the combo list)
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.fgScoring_BeforeEdit"

End Sub

Private Sub fgScoring_Click()
On Error GoTo ErrSection:

    With fgScoring
        EnableFgButtons .Row
        If .MouseCol = eSCRCol_Condition Then
            .EditCell
            SendKeys "{F4}" '(to dropdown the combo list)
        End If
    End With
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.fgScoring_Click"

End Sub

Private Sub fgScoring_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    FinishEdit = True
End Sub


Private Sub fgSpread_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim i&

    i = AddBlankRow
    If i > 0 Then
        With fgSpread
            If Col = 0 Then
                If Len(.TextMatrix(.Row, 0)) > 0 And Len(.TextMatrix(.Row, 1)) = 0 Then
                    fgSpread_CellButtonClick .Row, 1
                End If
            ElseIf Col = 1 Then
                If Len(.TextMatrix(.Row, 2)) = 0 Then
                    'new row just got added: default multiplier & contracts to 1
                    .TextMatrix(.Row, 2) = "1"
                    .TextMatrix(.Row, 3) = "1"
                    .TextMatrix(.Row, 4) = "1"      'save initial multiplier to hidden column
                    AddBlankRow
                    .Col = 2
                    .EditCell
                End If
            ElseIf Col = 2 And chkAutoMultiplier.Value = 0 Then
                'save user-entered multiplier to hidden column
                fgSpread.TextMatrix(fgSpread.Row, 4) = fgSpread.TextMatrix(fgSpread.Row, 2)
            End If
        End With
    End If
    
    'disallow changing to another category
    If fgSpread.Rows > 2 Then
        cboCategory.Enabled = False
        tbToolbar.Tools("ID_Verify").Enabled = True
        tbToolbar.Tools("ID_Save").Enabled = True
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".fgSpread_AfterEdit"

End Sub

Private Sub fgSpread_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    With fgSpread
        Select Case .Rows
            Case 1, 2, 3
                .ColComboList(0) = kOpAll
            Case 4
                If .MergeRow(3) Then
                    .ColComboList(0) = kOpAll
                Else
                    .ColComboList(0) = kOpPlusMinus
                End If
            Case Else
                .ColComboList(0) = kOpPlusMinus
        End Select
        If Row = .FixedRows And Col = 0 Then Cancel = True
        If Col = 2 And chkAutoMultiplier.Value = 1 Then
            Cancel = True
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_BeforeEdit"

End Sub

Private Sub fgSpread_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) back from the symbol selector
    Dim lSymbolID As Long               ' Symbol ID for the symbol selected
    Dim strSym$, strNext$
    
    Dim nID&, dTickVM#
    Dim Bars As cGdBars
    
    If frmSymbolSelector.Visible Then Exit Sub
    
    With fgSpread
        If .Row = .FixedRows Then
            strSym = .TextMatrix(.Row, 1)
        ElseIf .Row - 1 >= .FixedRows Then
            strSym = .TextMatrix(.Row - 1, 1)
            ' look for next contract forward
            strNext = GetNextContract(strSym)
            If Len(strNext) > 0 Then
                strSym = strNext
            End If
        End If
    End With
    
    Set astrSymbols = frmSymbolSelector.ShowMe(strSym, False)
    
    If astrSymbols.Size > 0 Then
        lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
    End If
    If lSymbolID = 0 Then
        Beep
    Else
        fgSpread.TextMatrix(Row, 1) = astrSymbols(0)
        fgSpread_AfterEdit Row, 1
        If chkAutoMultiplier.Value = 1 Then
            Set Bars = New cGdBars
            strSym = fgSpread.TextMatrix(fgSpread.Row, 1)
            nID = GetMarketInfo(strSym, Bars)
            If nID > 0 And Bars.Prop(eBARS_TickValue) > 0 And Bars.Prop(eBARS_TickMove) > 0 Then
                dTickVM = Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
            End If
            If nID > 0 And dTickVM > 0 Then
                fgSpread.TextMatrix(fgSpread.Row, 2) = Str(dTickVM)
            End If
        End If
    End If
    
    Set astrSymbols = Nothing
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_CellButtonClick"

End Sub

Private Sub fgSpread_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error Resume Next
    
    FinishEdit = True

End Sub

Private Sub fgSpread_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim strChar$
    Dim astrSymbols As New cGdArray     ' Symbol(s) back from the symbol selector
    Dim lSymbolID As Long               ' Symbol ID for the symbol selected
    
    With fgSpread
        If .MouseCol = 1 Then
            strChar = UCase(Chr(KeyCode))
            Set astrSymbols = frmSymbolSelector.ShowMe(strChar, False, , , , False, True)
            
            If astrSymbols.Size > 0 Then
                lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
            End If
            If lSymbolID = 0 Then
                Beep
            Else
                .TextMatrix(.Row, 0) = astrSymbols(0)
                fgSpread_AfterEdit .Row, 0
            End If
        End If
    End With
    
    Set astrSymbols = Nothing
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_KeyDown"
    
End Sub

Private Sub fgSpread_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Static nPrevMouseRow&

    With fgSpread
    
        If .Row >= .FixedRows And .Row < .Rows And .Row = .MouseRow Then
            If .MergeRow(.Row) Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                .MergeRow(.Row) = False
                If .Row = .FixedRows Then
                    fgSpread_CellButtonClick .FixedRows, 2
                    .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                Else
                    .Col = 0
                    .EditCell
                    SendKeys "{F4}"         'dropdown the combo list for (+, -, /) column
                End If
            ElseIf .MouseCol = 0 Then
                .EditCell
                SendKeys "{F4}"
            ElseIf .MouseCol = 1 Then
                'prevent symbol selector from coming up a second time from a double-click
                If .MouseRow <> nPrevMouseRow Then
                    fgSpread_CellButtonClick .Row, 2
                End If
            ElseIf .MouseCol = 2 Then
                .EditCell
            End If
        End If
    
        nPrevMouseRow = .MouseRow
    End With
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_MouseUp"

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Reload the function list upon activating the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim bUnload As Boolean
    
    If g.Functions Is Nothing Then
        InitFunctions
    End If
    
    ' Load internally generated TradeSense lists (Symbols, etc.)
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    ' Quickly check the Reverify flag.  If on then force a reverify...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "WHERE [FunctionID]=" & m.Function.FunctionID & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        
        If rs!CheckSum = 0.5 Then
            EnableToolbar False
            bUnload = True
            Err.Raise vbObjectError + 1000, , "This Function is no longer Valid"
        End If
        
        If rs!Reverify Then
            EnableToolbar True
            tbToolbar.Tools("ID_Verify").Enabled = True
        End If
    End If
    rs.Close
        
    If GetActiveWindow = Me.hWnd Then
        MoveFocus Editor1
        If Len(Editor1.Text) = 0 And Not m.bSkipAutoIf Then SendKeys " "
    End If
    
    ''vsIndexTab1.CurrTab = 0
    ''Editor1.SetFocus
    
ErrExit:
    Set rs = Nothing
    If bUnload Then Unload Me
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Form_Activate"
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    Else
        frmMain.DockPro_ShortcutKeyDown KeyCode, Shift, Me.Name
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.Form_KeyDown"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Intialize the controls and form upon loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText As String
    Dim lIndex As Long
    Dim strFont As String
    Dim bAdd As Boolean

    Me.Icon = Picture16(ToolbarIcon("ID_Functions"), , True)
    
    g.Styler.StyleForm Me
    
    With tbToolbar
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Verify").Picture = Picture16(ToolbarIcon("kVerify"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
    
    With chkAdvanced
        chkAutoMultiplier.Move .Left, .Top
    End With
        
    CenterTheForm Me

    strText = GetIniFileProperty("FuncMgrCT", "", "Placement", g.strIniFile)
    SetFormPlacement Me, strText, "LHT"
    
    m.strName = ""
    vsIndexTab1.CurrTab = 0
    
    ' Only show the required module controls if running from IDE...
    lblRequiredMod.Visible = IsIDE
    txtRequiredMod.Visible = IsIDE
    
    InitPlanetGrid
    InitScoringGrid
    InitSpreadGrid
  
    Set m.FunctionCategories = New cFunctionCategories
    m.FunctionCategories.Load
    For lIndex = 1 To m.FunctionCategories.Count
        With m.FunctionCategories.Item(lIndex)
            Select Case .FunctionCategory
                Case "Reserved", "Actions"
                    bAdd = False
                Case "Planet Combo"
                    bAdd = HasModule("ASTR")
                Case Else
                    bAdd = True
            End Select
            If bAdd Then
                cboCategory.AddItem m.FunctionCategories.Item(lIndex).FunctionCategory
                cboCategory.ItemData(cboCategory.NewIndex) = m.FunctionCategories.Item(lIndex).FunctionCategoryID
            End If
        End With
    Next lIndex
    
    chkAdvanced.Value = GetIniFileProperty("Advanced", vbUnchecked, "FuncMgrCT", g.strIniFile)
    AdvancedDisplay
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("FunctionMgrCT", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsInputs.Font, strFont
    
    If Not DirExist(AddSlash(App.Path) & "Resave") Then
        MkDir AddSlash(App.Path) & "Resave"
    End If
    
    txtCodedText.Visible = False
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Form_Load"
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim nColWidth&, Mid&
    
    If WindowState = vbMinimized Then
        If TypeOf ActiveControl Is Editor Then
            Set g.ActiveEditor = Nothing
            ActiveControl.RemoveTradeSense
        End If
    End If

    If LimitFormSize(Me, chkAdvanced.Left + chkAdvanced.Width + 600, vsIndexTab1.Top + 1800) Then Exit Sub
    
    With fraFunctionInfo
        .Move .Left, .Top, ScaleWidth - (.Left * 2)
    End With
    
    With txtDesc
        .Move .Left, .Top, fraFunctionInfo.Width - .Left
    End With
    
    With txtCodedText
        .Move txtDesc.Left, txtDesc.Top, txtDesc.Width, txtDesc.Height
    End With
    
    With TradeSense
        .Move .Left, .Top, fraFunctionInfo.Width - .Left
    End With

    With vsIndexTab1
        .Move fraFunctionInfo.Left, fraFunctionInfo.Height + (fraFunctionInfo.Top * 2), _
                ScaleWidth - (fraFunctionInfo.Left * 2), _
                ScaleHeight - fraFunctionInfo.Height - (fraFunctionInfo.Top * 3)
        .Refresh
    End With
        
    With fgPlanet
        .Move fraFunctionInfo.Left, TradeSense.Top + 130, _
                ScaleWidth - (fraFunctionInfo.Left * 2) + 15, _
                ScaleHeight - fraFgButtons.Height - txtDesc.Height * 2
        Mid = .Left + .Width / 2
        Mid = Mid - fraFgButtons.Width / 2
        fraFgButtons.Move Mid, .Top + .Height + 50
       'resize individual columns except the first one
        nColWidth = (.Width - kZeroColWidth) / kPlanetCols
        .ColWidth(ePLCol_Body1) = nColWidth * 1.3
        .ColWidth(ePLCol_Body2) = nColWidth * 1.3
        .ColWidth(ePLCol_PlanetSystem) = nColWidth * 1.3
        .ColWidth(ePLCol_Value) = nColWidth * 1.3
        .ColWidth(ePLCol_Weight) = nColWidth * 0.8
        .ColWidth(ePLCol_Offset) = nColWidth
        .ColWidth(ePLCol_Harmonic) = nColWidth
        .ColWidth(ePLCol_Orb) = nColWidth * 0.5
        .Refresh
    End With
        
    With fgScoring
        .Move fgPlanet.Left, fgPlanet.Top, fgPlanet.Width, fgPlanet.Height
        .Refresh
    End With
    
    With fgSpread
        .Move fgPlanet.Left, fgPlanet.Top, fgPlanet.Width, fgPlanet.Height
        If cboCategory.Text = "Spread" Then
            fraFgButtons.Left = fraFgButtons.Left + fraFgButtons.Width / 4
            cmdFgUp.Visible = False
            cmdFgDown.Visible = False
        Else
            cmdFgUp.Visible = True
            cmdFgDown.Visible = True
        End If
        .ColWidth(1) = .ClientWidth / 2
        .ColWidth(0) = .ClientWidth / 6
        .ColWidth(2) = .ClientWidth / 6
        .ColWidth(3) = .ClientWidth / 6
        .Refresh
    End With
    
    With Editor1
        .Move .Left, .Top, vsIndexTab1.ClientWidth - (.Left * 2), _
                vsIndexTab1.ClientHeight - Label5.Height - (Label5.Top * 3)
    End With
    
    With fraMoveButtons
        .Move (vsIndexTab1.ClientWidth - .Width) / 2, vsIndexTab1.ClientHeight - .Height - Label1(1).Top
    End With
    
    With vsInputs
        .Move .Left, .Top, vsIndexTab1.ClientWidth - (.Left * 2), _
                vsIndexTab1.ClientHeight - Label1(1).Height - (Label1(1).Top * 4) - fraMoveButtons.Height
    End With
    
    ExtendCustomColumn
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form gets unloaded, save off some settings
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "FuncMgrCT", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "FunctionCategoryID", cboCategory.Text, "Misc", g.strIniFile
    SetIniFileProperty "FunctionMgrCT", FontToString(vsInputs.Font), "Fonts", g.strIniFile
    
    If Not m.frmChartEditor Is Nothing Then m.frmChartEditor.tmrChartCfg.Enabled = True
    
    Set m.Function = Nothing
    Set m.FunctionCategories = Nothing
    Set m.frmChartEditor = Nothing
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Form_Unload"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearFunction
'' Description: Clear the form to get ready for loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearFunction()
On Error GoTo ErrSection:

    Dim CatText     As String
    
    m.strName = ""
    txtDesc.Text = ""
    
    'Get INI Defaults
    CatText = GetIniFileProperty("FunctionCategory", "", "Misc", g.strIniFile)
    If Len(CatText) > 0 Then
        cboCategory.Text = CatText
    End If
    TradeSense.Text = ""
    Editor1.Text = ""
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.ClearFunction"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unload the form if the user clicks on the 'lIndex'
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from a user question

    If UnloadMode <> vbFormCode Then
        Cancel = AskToSave
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.Form_QueryUnload"
    Resume ErrExit
    
End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo ErrSection:

    If Index = 6 Then
        If FileExist(AddSlash(App.Path) & "ShowCT.FLG") Then
            If txtCodedText.Visible Then
                txtDesc.Visible = True
                txtCodedText.Visible = False
            Else
                txtDesc.Visible = False
                txtCodedText.Visible = True
                txtCodedText.Text = m.Function.CodedText
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.Label1_Click"
    
End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsInputs, True
    RefreshGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.mnuChangeFont_Click"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the action that the user chose on the toolbar
'' Inputs:      Tool Clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strID$
    Dim bAutoIf As Boolean

    bAutoIf = m.bSkipAutoIf
    m.bSkipAutoIf = True
    ToggleFocus Me, Me.vsIndexTab1

    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            Save Tool.ID
        
        Case "ID_Print"
            PrintMe
        
        Case "ID_Verify"
            Verify
        
        Case "ID_Toolbox"
            If Not AskToSave Then
                strID = CStr(m.Function.FunctionID)
                Unload Me
                frmToolbox.ShowMe eTab_Functions, strID
            End If
        
        Case "ID_Close"
            If Not AskToSave Then
                Unload Me
            End If
    
    End Select
    
    m.bSkipAutoIf = bAutoIf

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.tbToolbar_ToolClick"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDesc_Change
'' Description: If the description changes, dirty the function
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDesc_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.txtDesc_Change"
    Resume ErrExit
    
End Sub

Private Sub txtRequiredMod_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.txtRequiredMod_Change"
    Resume ErrExit
    
End Sub

Private Sub vsIndexTab1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    Select Case NewTab
        Case 0
            MoveFocus Editor1
            
        Case 1
            MoveFocus vsInputs
            
        Case 2
            MoveFocus chkSystemTesting
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsIndexTab1_Switch"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterEdit
'' Description: Do some cleanup after the user edits
'' Inputs:      Row and Column of Cell being edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    Select Case Col
        Case GDCol(eGDCol_Expression)
            If CheckedCell(vsInputs, Row, Col) = False Then
                If ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_ParmTypeID))) = kSN_RetText Then
                    vsInputs.TextMatrix(Row, GDCol(eGDCol_DefaultValue)) = Chr(34) & StripStr(vsInputs.TextMatrix(Row, GDCol(eGDCol_DefaultValue)), Chr(34)) & Chr(34)
                ElseIf IsNumeric(vsInputs.TextMatrix(Row, GDCol(eGDCol_DefaultValue))) = False Then
                    vsInputs.TextMatrix(Row, GDCol(eGDCol_DefaultValue)) = FormatNum(0)
                    InfBox "Default value for a numeric constant field" & _
                        "|must be numeric.  The Default value has been|set back to zero", _
                        "!", , "Warning"
                End If
            Else
                If ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_ParmTypeID))) = kSN_RetText Then
                    InfBox "Text parameters cannot be set to expressions", "!", , "Function Input Error"
                    CheckedCell(vsInputs, Row, GDCol(eGDCol_Expression)) = False
                End If
            End If
            EnableToolbar True
    
        Case GDCol(eGDCol_DefaultValue)
            With vsInputs
                If ValOfText(.TextMatrix(Row, GDCol(eGDCol_ParmTypeID))) = kSN_RetText Then
                    If CheckedCell(vsInputs, Row, GDCol(eGDCol_Expression)) = False Then
                        .TextMatrix(Row, GDCol(eGDCol_DefaultValue)) = Chr(34) & StripStr(.TextMatrix(Row, GDCol(eGDCol_DefaultValue)), Chr(34)) & Chr(34)
                    End If
                End If
            End With
            'If Not IsNumeric(vsInputs.TextMatrix(Row, Col)) Then
            '    If CheckedCell(vsInputs, Row, GDCol(eGDCol_Expression)) = False Then
            '        CheckedCell(vsInputs, Row, GDCol(eGDCol_Expression)) = True
            '    End If
            'End If
    
        Case GDCol(eGDCol_Req)
            With vsInputs
                .Redraw = flexRDNone
                If CheckedCell(vsInputs, Row, Col) = True Then
                    For lIndex = .FixedRows To Row
                        CheckedCell(vsInputs, lIndex, Col) = True
                    Next lIndex
                Else
                    For lIndex = Row To .Rows - 1
                        CheckedCell(vsInputs, lIndex, Col) = False
                    Next lIndex
                End If
                .Redraw = flexRDBuffered
            End With
            EnableToolbar True
    End Select
    
    ' Color the cell apppropriately
    ColorCell Row, Col
    
    ' Format the value if numeric
    If IsNumeric(vsInputs.TextMatrix(Row, Col)) Then
        vsInputs.TextMatrix(Row, Col) = FormatNum(ValOfText(vsInputs.TextMatrix(Row, Col)))
    End If
    
    ' Make sure that the parm type id is set correctly
    With vsInputs
        If CheckedCell(vsInputs, Row, GDCol(eGDCol_Expression)) = True Then
            Select Case ValOfText(.TextMatrix(Row, GDCol(eGDCol_ParmTypeID)))
                Case kSN_RetTrueFalseConstant
                    .TextMatrix(Row, GDCol(eGDCol_ParmTypeID)) = FormatNum(kSN_RetTrueFalse)
                Case kSN_RetNumericConstant
                    .TextMatrix(Row, GDCol(eGDCol_ParmTypeID)) = FormatNum(kSN_RetNumeric)
            End Select
        Else
            Select Case ValOfText(.TextMatrix(Row, GDCol(eGDCol_ParmTypeID)))
                Case kSN_RetTrueFalse
                    .TextMatrix(Row, GDCol(eGDCol_ParmTypeID)) = FormatNum(kSN_RetTrueFalseConstant)
                Case kSN_RetNumeric
                    .TextMatrix(Row, GDCol(eGDCol_ParmTypeID)) = FormatNum(kSN_RetNumericConstant)
            End Select
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_AfterEdit"

End Sub

Private Sub vsInputs_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    vsInputs.Row = Position
    vsInputs.RowSel = Position
    
    RefreshGrid
    ShowParmLine TradeSense
    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_AfterMoveRow"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterRowColChange
'' Description: Edit the cell upon the user changing the row or column
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With vsInputs
        If .Visible And NewCol = GDCol(eGDCol_DefaultValue) Then
            .EditCell
        End If
    End With
    
    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_AfterRowColChange"
    Resume ErrExit
    
End Sub

Private Sub vsInputs_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    ExtendCustomColumn Col

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_AfterUserResize"
    Resume ErrExit
End Sub

Private Sub vsInputs_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lPos As Long

    With vsInputs
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            If Button = vbLeftButton Then
                .Row = lMouseRow
                .RowSel = lMouseRow
            
                .Refresh
                lPos = .DragRow(lMouseRow)
                If lPos <> lMouseRow Then
                    Cancel = True
                End If
            ElseIf Button = vbRightButton Then
                .Row = lMouseRow
                PopupMenu mnuPopUp
            Else
                Cancel = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_BeforeMouseDown"
    Resume ErrExit
    
End Sub

Private Sub vsInputs_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' save current size in case after custom extended column
    m.nPrevColWidth = vsInputs.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_BeforeUserResize"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_ChangeEidt
'' Description: If the user changes a cell, dirty the function
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_ChangeEdit()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_ChangeEdit"
    Resume ErrExit
    
End Sub

Private Sub vsInputs_GotFocus()
On Error GoTo ErrSection:

    Dim lCol As Long
    
    With vsInputs
        lCol = .Col
        If lCol >= .FixedCols And .Row > .FixedRows Then
            If .ColDataType(lCol) <> flexDTBoolean Then
                .EditCell
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_GotFocus"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_ValidateEdit
'' Description: Validate what the user has entered
'' Inputs:      Row and Column of Cell being Edited, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim dDefaultValue As Double         ' Default value for the input
    Dim dFromVal As Double              ' From value for the input
    Dim dToVal As Double                ' To value for the input
    
    ' Get input values
    dDefaultValue = ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_DefaultValue)))
    dFromVal = ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_FromVal)))
    dToVal = ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_ToVal)))
    
    Select Case Col
        Case GDCol(eGDCol_FromVal)
            If Not IsNumeric(vsInputs.EditText) Then
                Cancel = True
                Exit Sub
            End If
            dFromVal = ValOfText(vsInputs.EditText)
        
        Case GDCol(eGDCol_ToVal)
            If Not IsNumeric(vsInputs.EditText) Then
                Cancel = True
                Exit Sub
            End If
            dToVal = ValOfText(vsInputs.EditText)
        
        Case GDCol(eGDCol_DefaultValue)
            If IsNumeric(vsInputs.EditText) Then
                dDefaultValue = vsInputs.EditText
                'If dFromVal <> 0 Or dToVal <> 0 Then
                '    If dDefaultValue < dFromVal Or dDefaultValue > dToVal Then
                '        Cancel = True
                '        Err.Raise vbObjectError + 1000, , "Please enter a value between " & Format(dFromVal, "#,##0") & " and " & Format(dToVal, "#,##0")
                '    End If
                'Else
                    If (dDefaultValue < -100000000000# Or dDefaultValue > 100000000000#) Then
                        Cancel = True
                        Err.Raise vbObjectError + 1000, , "Please enter a value between -100,000,000,000 and 100,000,000,000"
                    End If
                'End If
            End If
            
        Case GDCol(eGDCol_Expression)
        
        Case GDCol(eGDCol_Req)
        
        Case GDCol(eGDCol_ParmDesc)
        
        Case Else
            Cancel = True
            Exit Sub
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_ValidateEdit"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeEdit
'' Description: Only allow the user to edit certain cells
'' Inputs:      Row and Column to Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    ' No changes allowed for bar structure inputs
    If ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_ParmTypeID))) = kSN_RetBars Then
        Cancel = True
    End If
    
    ' Do not allow changes to the Input Name
    If Col = GDCol(eGDCol_InputName) Then
        Cancel = True
    End If
    
    ' Do not allow changes to bars or trades type expressions
    If Col = GDCol(eGDCol_Expression) Then
        Select Case ValOfText(vsInputs.TextMatrix(Row, GDCol(eGDCol_ParmTypeID)))
            Case kSN_RetBars, kSN_RetTrades
                Cancel = True
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.vsInputs_BeforeEdit"
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Verify
'' Description: Verify the coded text
'' Inputs:      None
'' Returns:     True if Verified, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Verify(Optional ByVal bShowMsg As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    Dim svErr As Long
    Dim svErrDesc As String
    Dim svSource As String
    Dim strMsg As String
    Dim strFormatted As String          ' Text returned from the verify
    Dim Expr As cExpression
    Dim tmpInputs As New cInputs
    Dim lIndex As Long                  ' Index into a for loop
    
    ' TLB 5/23/2005: need to make sure we don't recurse into here (e.g. when check boxes are set from within here)
    Static bInHere As Boolean
    If bInHere Then Exit Function
    bInHere = True

    'Save current input values from grid...
    'SaveGridValues tmpInputs
    
    FixPeriodInMarkets
    
    'Shut things off, get ready for verifying rule
    Screen.MousePointer = vbHourglass
    m.ReturnValue = LockWindowUpdate(Me.hWnd)
    
    If fgPlanet.Visible Then
        BuildPlanetText
    ElseIf fgScoring.Visible Then
        BuildScoringText
        'turn all usage flags on
        With m.Function
            .Usage = .Usage Or UsageMask(eUsageMask_MM)
            .Usage = .Usage Or UsageMask(eUsageMask_SystemTesting)
            .Usage = .Usage Or UsageMask(eUsageMask_Charting)
            .Usage = .Usage Or UsageMask(eUsageMask_Criteria)
        End With
        'turn all checkboxes on
        chkMM.Value = 1
        chkSystemTesting.Value = 1
        chkCharting.Value = 1
        chkCriteria.Value = 1
    ElseIf fgSpread.Visible Then
        If Not BuildSpreadText Then
            bInHere = False
            Exit Function
        End If
    End If
    
    'Verify...
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        ' TLB 3/24/2015: replace a Tab with a space (e.g. in case was cut-and-paste from a text file)
        .ValidateFunctionRule Replace(Editor1.Text, vbTab, " ")
        strFormatted = .EditText
        
        ' Verify any "Symbol,Period" market types...
        If Not .Inputs Is Nothing Then
            For lIndex = 1 To .Inputs.Count
                If ValidMarket(.Inputs.Item(lIndex).ParmName) = False Then
                    Err.Raise vbObjectError + 1000, , "No data can be loaded for " & .Inputs.Item(lIndex).ParmName
                End If
            Next lIndex
        End If
        
        ' Convert to rich text...
        Editor1.TurnOffEditing
        Editor1.TextRTF = m.Function.GetRTF(strFormatted)
        Editor1.ExprIsFormatted = True
    
        ' Save verify settings...
        m.Function.FunctionIDs = .GetFIDs
        m.Function.Formatted = .EditText
        m.Function.FormattedWithFillWords = .Preview
        m.Function.CodedText = .CodedText
        m.Function.DataTypeID = .FunctionReturnType
        m.Function.ReturnTypeID = .FunctionReturnType
        m.Function.LateCalculating = .LateCondition
    
        'set auto mutliplier input for spread functions
        If cboCategory.Text = "Spread" Then
            Dim NewParm As New cInput
            
            NewParm.ParmName = "Auto Multiplier"
            NewParm.ParmTypeID = 6          'boolean constant
            NewParm.DefaultValue = chkAutoMultiplier.Value
            
            With NewParm
                Expr.Inputs.Add .RuleName, .ParmSeq, .ParmName, _
                    .ParmDesc, .ParmID, .Value, .IfOptimize, _
                    .OptFromValue, .OptToValue, .OptStepValue, _
                    .OptListID, .RuleID, .ParmTypeID, .DefaultValue, _
                    .Required, .FromValue, .ToValue, .ListID, _
                    .FillPre, .FillPost
            End With
            
            Set NewParm = Nothing
        End If
    End With
            
    If VerifyUsage(Expr.GetFIDs, bShowMsg) = False Then
        chkSystemTesting.Value = vbUnchecked
    End If
        
    'If Not Expr.Inputs Is Nothing Then
    '    ShowParmLine TradeSense
    '    m.Function.TradeSenseUsage = TradeSense.Tag
    'End If
    
    'Restore saved values over the current values
    'Set m.Inputs = Expr.Inputs
    'LoadGrid
    'strMsg = RestoreGridValues(tmpInputs)
    'If Len(strMsg) > 0 Then
    '    'ask about new inputs
    '    strMsg = "Unrecognized as existing functions or inputs:|" _
    '            & strMsg & "|Add the above as new INPUTS to this function?|"
    '    If AskBox("i=? ; b=+Add|-Cancel ; h=Add Inputs ; " & strMsg) = "C" Then
    '        'consider it "unverified"
    '        Err.Raise vbObjectError + 1000, , "Need to fix unrecognized functions."
    '    Else
    '        vsIndexTab1.CurrTab = 1
    '    End If
    'End If
    RefreshInputs Expr.Inputs
    If Not Expr.Inputs Is Nothing Then
        ShowParmLine TradeSense
        m.Function.TradeSenseUsage = TradeSense.Tag
    End If
       
    tbToolbar.Tools("ID_Verify").Enabled = False
    EnableToolbar True
    
    Screen.MousePointer = vbDefault
    m.ReturnValue = LockWindowUpdate(0)
    
    Verify = True
    
ErrExit:
    Editor1.TurnOnEditing
    Set Expr = Nothing
    bInHere = False
    Exit Function
    
ErrSection:
    bInHere = False
    Screen.MousePointer = vbDefault
    m.ReturnValue = LockWindowUpdate(0)
    
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        svErr = Err.Number
        svSource = Err.Source
        svErrDesc = Err.Description
        
        'Highlight error in advanced editor...
        If Expr.EditText <> "" Then
            With Editor1
                .AppPath = App.Path
                .TurnOffEditing
                strFormatted = Expr.EditText
                .ExprIsFormatted = False
                .TextRTF = m.Function.GetRTF(strFormatted)
                .ExprIsFormatted = True
            End With
            Editor1.TurnOnEditing
        End If
        
        Set Expr = Nothing
        Err.Raise svErr, svSource, svErrDesc
    Else
        Set Expr = Nothing
        RaiseError "frmFunctionMgrCT.Verify"
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyFunctionDebug
'' Description: Verify the coded text and show the tree
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyFunctionDebug()
On Error GoTo ErrSection:
    
    Dim svErr           As Long
    Dim svErrDesc       As String
    Dim svSource        As String
    Dim strMsg          As String
    Dim wrkText         As String
    Dim Expr            As cExpression
    Dim tmpInputs As New cInputs
    
    'Save current input values from grid...
    'SaveGridValues tmpInputs
    
    'Shut things off, get ready for verifying rule
    Screen.MousePointer = vbHourglass
    m.ReturnValue = LockWindowUpdate(Me.hWnd)
    
    'Verify...
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule Trim(Editor1.Text)
        
        'Convert to rich text
        Editor1.TurnOffEditing
        wrkText = .EditText
        Editor1.TextRTF = m.Function.GetRTF(wrkText)
        Editor1.ExprIsFormatted = True
    
        'Save verify settings
        m.Function.FunctionIDs = .GetFIDs
        m.Function.Formatted = .EditText
        m.Function.FormattedWithFillWords = .Preview
        m.Function.CodedText = .CodedText
        m.Function.FunctionIDs = .GetFIDs
        m.Function.DataTypeID = .FunctionReturnType
        m.Function.ReturnTypeID = .FunctionReturnType
        
        'Save Late calculating flags...
        If .LateCondition Then
            m.Function.LateCalculating = True
        Else
            m.Function.LateCalculating = False
        End If
    
    End With
        
    RefreshInputs Expr.Inputs
    If Not Expr.Inputs Is Nothing Then
        ShowParmLine TradeSense
        m.Function.TradeSenseUsage = TradeSense.Tag
    End If
    
    'Restore saved values over the current values
    'Set m.Inputs = Expr.Inputs
    'LoadGrid
    'strMsg = RestoreGridValues(tmpInputs)
    'If Len(strMsg) > 0 Then
    '    'ask about new inputs
    '    strMsg = "Unrecognized as existing functions or inputs:|" _
    '            & strMsg & "|Add the above as new INPUTS to this function?|"
    '    If AskBox("i=? ; b=+Add|-Cancel ; h=Add Inputs ; " & strMsg) = "C" Then
    '        'consider it "unverified"
    '        Err.Raise vbObjectError + 1000, , "Need to fix unrecognized functions."
    '    Else
    '        vsIndexTab1.CurrTab = 1
    '    End If
    'End If
    
    tbToolbar.Tools("ID_Verify").Enabled = False
    EnableToolbar True
    
    Screen.MousePointer = vbDefault
    m.ReturnValue = LockWindowUpdate(0)
    
'=================================================
ShowTheTree:
    Screen.MousePointer = vbDefault
    If Not Expr.Trees Is Nothing Then
        With frmTrees
            .CodedText = Expr.CodedText
            .EditText = Expr.EditText
            .Preview = Expr.Preview
            .Trees = Expr.Trees
            .LoadTrees
        End With
        ShowForm frmTrees, True
    End If
    Set Expr = Nothing
    Exit Sub

'=================================================
ErrSection:
    Screen.MousePointer = vbDefault
    m.ReturnValue = LockWindowUpdate(0)
    
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        svErr = Err.Number
        svSource = Err.Source
        svErrDesc = Err.Description
        
        'Highlight error in advanced editor...
        If Expr.EditText <> "" Then
            With Editor1
                .AppPath = App.Path
                .TurnOffEditing
                .ExprIsFormatted = False
                wrkText = Expr.EditText
                .TextRTF = m.Function.GetRTF(wrkText)
                .ExprIsFormatted = True
            End With
            Editor1.TurnOnEditing
        End If
        
        MsgBox svErrDesc & Chr(13) & Chr(10), vbInformation, "Message"
        Resume ShowTheTree:
    Else
        Set Expr = Nothing
        RaiseError "frmFunctionMgrCT.VerifyFunctionDebug"
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FunctionExists
'' Description: Determine whether the Function Exists
'' Inputs:      Function Name to look up
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FunctionExists(ByVal strFuncName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim QryDef As QueryDef
        
    Set QryDef = g.dbNav.QueryDefs("qryFunctionIDFromName")
    QryDef.Parameters(0).Value = strFuncName
    Set rs = QryDef.OpenRecordset
    
    FunctionExists = False
    If rs.RecordCount <> 0 Then
        If rs!FunctionID <> m.Function.FunctionID Then
            FunctionExists = True
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Set QryDef = Nothing
    Exit Function

ErrSection:
    RaiseError "frmFunctionMgrCT.FunctionExists"
    Resume ErrExit:

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the Print Preview form
'' Inputs:      Arguments form the Print Preview form
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lIndex&, lNumInputs&
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .FontUnderline = True
        .Text = vbLf & "Function:"
        .FontUnderline = False
        .Text = "    " & Trim(m.strName) & vbCrLf '& vbCrLf
        .Font.Bold = False
        .Font.Size = 12
        .Text = "Description: " & Trim(txtDesc.Text) & vbCrLf
        .Text = "Category: " & Trim(cboCategory.Text) & vbCrLf '& vbCrLf
        .Text = "Text: " & Trim(Editor1.Text) & vbCrLf & vbCrLf
        
        .Font.Size = 14
        .Font.Bold = True
        .FontUnderline = True
        .Text = "Default Values for Inputs:" & vbCrLf
        .FontUnderline = False
        .Font.Bold = False
        .Font.Size = 12
        
        lNumInputs = 0&
        For lIndex = 1 To vsInputs.Rows - 1
            If vsInputs.RowHidden(lIndex) = False Then
                lNumInputs = lNumInputs + 1
                .Text = vsInputs.Cell(flexcpText, lIndex, GDCol(eGDCol_InputName)) & " = "
                .Text = vsInputs.Cell(flexcpText, lIndex, GDCol(eGDCol_DefaultValue)) & vbLf
            End If
        Next lIndex
        
        If lNumInputs = 0& Then
            .Text = "(No Inputs for this Function)" & vbLf
        End If
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.GenerateReport"
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: If the value of the cell is negative, color it red else black
'' Inputs:      Row and Column of the Cell
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ColorCell(Row As Long, Col As Long)
On Error GoTo ErrExit:

    If ValOfText(vsInputs.TextMatrix(Row, Col)) < 0 Then
        vsInputs.Cell(flexcpForeColor, Row, Col) = vbRed
    Else
        vsInputs.Cell(flexcpForeColor, Row, Col) = vbBlack
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.ColorCell"
    Resume ErrExit

End Sub

Private Sub EnableToolbar(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    tbToolbar.Tools("ID_Save").Enabled = bEnable
    tbToolbar.Tools("ID_SaveAs").Enabled = (Trim(m.strName) <> "")
    tbToolbar.Tools("ID_Rename").Enabled = (Trim(m.strName) <> "")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.EnableToolbar"
    
End Sub

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    Dim bSkipAutoIf As Boolean
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        If WindowState = vbMinimized Then WindowState = vbNormal
    
        Set g.ActiveEditor = Nothing
        Editor1.RemoveTradeSense
        bSkipAutoIf = m.bSkipAutoIf
        m.bSkipAutoIf = True
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        m.bSkipAutoIf = bSkipAutoIf
        Select Case strResponse
            Case "C"
                AskToSave = True
            Case "Y"
                Save "ID_Save"
        End Select
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError Me.Name & ".AskToSave"

End Function

Public Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .FocusRect = flexFocusNone ' = flexFocusSolid
        '.HighLight = flexHighlightWithFocus
        .TabBehavior = flexTabCells
        .Ellipsis = flexEllipsisEnd
        .Editable = flexEDKbdMouse
        .ExtendLastCol = False
        .ExplorerBar = flexExMoveRows
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .Cols = GDCol(eGDCol_NumCols)
        .Rows = 1
        .FixedCols = 1
        .FixedRows = 1
        
        .TextMatrix(0, GDCol(eGDCol_InputName)) = "Input"
        .TextMatrix(0, GDCol(eGDCol_DefaultValue)) = "Default Value"
        .TextMatrix(0, GDCol(eGDCol_ParmDesc)) = "Description (optional)"
        .TextMatrix(0, GDCol(eGDCol_FromVal)) = "Min Value"
        .TextMatrix(0, GDCol(eGDCol_ToVal)) = "Max Value"
        .TextMatrix(0, GDCol(eGDCol_ParmTypeID)) = "Data Type"
        .TextMatrix(0, GDCol(eGDCol_Req)) = "Required"
        .TextMatrix(0, GDCol(eGDCol_Expression)) = "Expression"
        
        .ColAlignment(GDCol(eGDCol_InputName)) = flexAlignLeftCenter
        .ColAlignment(GDCol(eGDCol_DefaultValue)) = flexAlignLeftCenter
        
        .ColDataType(GDCol(eGDCol_Req)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_Expression)) = flexDTBoolean
        
        '6/2001 Out dated fields
        .ColHidden(GDCol(eGDCol_FromVal)) = True
        .ColHidden(GDCol(eGDCol_ToVal)) = True
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.InitGrid"
    
End Sub

Public Sub LoadGrid()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lRedraw As Long

    'Leave if no inputs exist in collection
    If m.Function.Inputs Is Nothing Then Exit Sub
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDBuffered
        
        .Rows = m.Function.Inputs.Count + .FixedRows

        For lIndex = 1 To m.Function.Inputs.Count
            AddRowToGrid lIndex, m.Function.Inputs.Item(lIndex)
        Next lIndex
        
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                .Row = lIndex
                .Col = GDCol(eGDCol_DefaultValue)
                
                Exit For
            End If
        Next lIndex
        
        RefreshGrid
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.LoadGrid"

End Sub

Public Sub AddRowToGrid(ByVal lRow As Long, ByVal Parm As cInput)
On Error GoTo ErrSection:
    
    Dim lRedraw As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDBuffered
        
        .TextMatrix(lRow, GDCol(eGDCol_InputName)) = Parm.ParmName
        If Parm.ParmName = "Market1" Then
            .TextMatrix(lRow, GDCol(eGDCol_Sort)) = " "
        Else
            .TextMatrix(lRow, GDCol(eGDCol_Sort)) = Parm.ParmName
        End If
        .TextMatrix(lRow, GDCol(eGDCol_ParmTypeID)) = Parm.ParmTypeID
        .TextMatrix(lRow, GDCol(eGDCol_ParmID)) = Parm.ParmID
        .TextMatrix(lRow, GDCol(eGDCol_ParmDesc)) = Parm.ParmDesc
        '.TextMatrix(lRow, GDCol(eGDCol_FromVal)) = Parm.FromValue
        If Parm.ParmTypeID = 5 Then
            .TextMatrix(lRow, GDCol(eGDCol_FromVal)) = 3
        Else
            .TextMatrix(lRow, GDCol(eGDCol_FromVal)) = 2
        End If
        .TextMatrix(lRow, GDCol(eGDCol_ToVal)) = Parm.ToValue
        .TextMatrix(lRow, GDCol(eGDCol_Req)) = Parm.Required
        .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = ""
        
        'Set the value (or default if one doesn't exist).  The bars and
        'trades type structure is always "Market1" and "Trades"
        Select Case Parm.ParmTypeID
        
            Case kSN_RetBars
                .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = Parm.ParmName
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = False
        
            Case kSN_RetTrades
                .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = Parm.ParmName
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = False
                    
            Case kSN_RetNumericConstant
                .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = FormatNum(Val(Parm.DefaultValue))
                ColorCell lRow, GDCol(eGDCol_DefaultValue)
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = False
                
            Case kSN_RetTrueFalseConstant
                If IsNumeric(Parm.DefaultValue) Then
                    .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = FormatNum(Val(Parm.DefaultValue))
                    ColorCell lRow, GDCol(eGDCol_DefaultValue)
                Else
                    .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = Parm.DefaultValue
                End If
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = False
                
            Case kSN_RetNumeric
                If IsNumeric(Parm.DefaultValue) Then
                    .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = FormatNum(Val(Parm.DefaultValue))
                    ColorCell lRow, GDCol(eGDCol_DefaultValue)
                Else
                    .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = Parm.DefaultValue
                End If
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = True
                
            Case kSN_RetTrueFalse
                If IsNumeric(Parm.DefaultValue) Then
                    .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = FormatNum(Val(Parm.DefaultValue))
                    ColorCell lRow, GDCol(eGDCol_DefaultValue)
                Else
                    .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = Parm.DefaultValue
                End If
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = True
            
            Case kSN_RetText
                .TextMatrix(lRow, GDCol(eGDCol_DefaultValue)) = Parm.DefaultValue
                ColorCell lRow, GDCol(eGDCol_DefaultValue)
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = False
            
            Case Else
        End Select
        
        If Parm.ParmTypeID = kSN_RetBars Then .RowHidden(lRow) = True
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.AddRowToGrid"

End Sub

Public Sub RefreshGrid()
On Error GoTo ErrSection:
    
    Dim lRedraw As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDBuffered
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        SetBackColors vsInputs
        
        EnableButtons
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.RefreshGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      nResizeCol (passed only by the AfterUserResize event)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Optional ByVal nResizeCol As Long = -1)
On Error GoTo ErrSection:
    
    Dim i&, nTotal&, nDiff&, nExtCol&
    
    ' column number of the custom extended column
    If chkAdvanced Then
        nExtCol = GDCol(eGDCol_ParmDesc)
    Else
        nExtCol = GDCol(eGDCol_InputName)
    End If
      
    With vsInputs
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= nExtCol Then
            .Redraw = flexRDNone
            nDiff = .ColWidth(nResizeCol) - m.nPrevColWidth
            For i = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = .ColWidth(i) - nDiff
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(nExtCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        If nTotal > 0 Then .ColWidth(nExtCol) = nTotal
        .ColHidden(nExtCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
       
ErrSection:
    RaiseError "frmFunctionMgrCT.ExtendCustomColumn"
End Sub

Public Sub AdvancedDisplay()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        If chkAdvanced.Value = vbChecked Then
            vsIndexTab1.TabVisible(2) = True
            .ColHidden(GDCol(eGDCol_ParmTypeID)) = True
            .ColHidden(GDCol(eGDCol_ParmID)) = True
            .ColHidden(GDCol(eGDCol_FromVal)) = True 'False
            .ColHidden(GDCol(eGDCol_ToVal)) = True 'False
            .ColHidden(GDCol(eGDCol_ParmDesc)) = False ' True
            .ColHidden(GDCol(eGDCol_Sort)) = True
            .ColHidden(GDCol(eGDCol_Req)) = False
            .ColHidden(GDCol(eGDCol_Expression)) = False
            .AutoSize GDCol(eGDCol_InputName)
        Else
            If vsIndexTab1.CurrTab = 2 Then vsIndexTab1.CurrTab = 0
            vsIndexTab1.TabVisible(2) = False
            .ColHidden(GDCol(eGDCol_ParmTypeID)) = True
            .ColHidden(GDCol(eGDCol_ParmID)) = True
            .ColHidden(GDCol(eGDCol_FromVal)) = True
            .ColHidden(GDCol(eGDCol_ToVal)) = True
            .ColHidden(GDCol(eGDCol_ParmDesc)) = True
            .ColHidden(GDCol(eGDCol_Sort)) = True
            .ColHidden(GDCol(eGDCol_Req)) = True
            .ColHidden(GDCol(eGDCol_Expression)) = True
        End If
        
        ExtendCustomColumn
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.AdvancedDisplay"

End Sub

Private Sub SetColumnWidths()
On Error GoTo ErrSection:

    With vsInputs
        If .ColWidth(GDCol(eGDCol_InputName)) > 2500 Then
            .ColWidth(GDCol(eGDCol_InputName)) = 2500
        End If
        
        If .ColWidth(GDCol(eGDCol_DefaultValue)) < 1200 Then
            .ColWidth(GDCol(eGDCol_DefaultValue)) = 1200
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.SetColumnWidths"

End Sub

Private Sub RefreshInputs(NewInputs As cInputs, Optional ByVal bShowMsg As Boolean = True)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index for a for loop
    Dim lInput As Long                  ' Index for a for loop
    Dim bFound As Boolean               ' Was the input found in the grid?
    Dim strNewInputs As String          ' New Inputs to display to the user
    Dim Parm As New cInput              ' Temporary input variable
    Dim strMsg As String                ' Message to display to the user
    Dim alNewInputs As cGdArray         ' Array of new input indexes
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lRow2 As Long
    
    Set alNewInputs = New cGdArray
    alNewInputs.Create eGDARRAY_Longs
    
    ' Walk through the new inputs and check them against the grid...
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lInput = 1 To NewInputs.Count
            Set Parm = NewInputs.Item(lInput)
            bFound = False
            For lRow = .FixedRows To .Rows - 1
                If .TextMatrix(lRow, GDCol(eGDCol_InputName)) = Parm.ParmName Then
                    Select Case Parm.ParmTypeID
                        Case kSN_RetNumericConstant
                            If CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) Then
                                Parm.ParmTypeID = kSN_RetNumeric
                            End If
                        
                        Case kSN_RetTrueFalseConstant
                            If CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) Then
                                Parm.ParmTypeID = kSN_RetTrueFalse
                            End If
                            
                        Case kSN_RetNumeric
                            If Not CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) Then
                                Parm.ParmTypeID = kSN_RetNumericConstant
                            End If
                        
                        Case kSN_RetTrueFalse
                            If Not CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) Then
                                Parm.ParmTypeID = kSN_RetTrueFalseConstant
                            End If
                            
                        Case Else
                            CheckedCell(vsInputs, lRow, GDCol(eGDCol_Expression)) = False
                        
                    End Select
                    .TextMatrix(lRow, GDCol(eGDCol_ParmTypeID)) = Parm.ParmTypeID
                    
                    bFound = True
                    Exit For
                End If
            Next lRow
            
            If Not bFound Then
                alNewInputs.Add lInput
                If NewInputs.Item(lInput).ParmTypeID <> 5 Then
                    strNewInputs = strNewInputs & NewInputs.Item(lInput).ParmName & vbCrLf
                End If
            End If
        Next lInput
        
        ' Make sure that inputs that no longer exist get deleted from the grid...
        For lRow = .Rows - 1 To .FixedRows Step -1
            bFound = False
            For lInput = 1 To NewInputs.Count
                If NewInputs.Item(lInput).ParmName = .TextMatrix(lRow, GDCol(eGDCol_InputName)) Then
                    bFound = True
                    Exit For
                End If
            Next lInput
            
            If Not bFound Then .RemoveItem lRow
        Next lRow
        
        ' If there are any new inputs to add, verify with the user, then add them...
        If Len(strNewInputs) > 0 And cboCategory.Text <> "Spread" Then
            strMsg = "Unrecognized as existing functions or inputs:|" _
                    & strNewInputs & "|Add the above as new INPUTS to this function?|"
            If AskBox("i=? ; b=+Add|-Cancel ; h=Add Inputs ; " & strMsg) = "C" Then
                Err.Raise vbObjectError + 1000, , "Need to fix unrecognized functions."
            End If
        End If
                
        If alNewInputs.Size > 0 Then
            For lInput = 0 To alNewInputs.Size - 1
                .Rows = .Rows + 1
                AddRowToGrid .Rows - 1, NewInputs.Item(alNewInputs(lInput))
            Next lInput
            
            If Len(strNewInputs) > 0 Then
                vsIndexTab1.CurrTab = 1
            End If
        End If
        
        ' Walk through and hide the Market type paramaters and move them to the top...
        For lRow = .FixedRows To .Rows - 1
            If CLng(ValOfText(.TextMatrix(lRow, GDCol(eGDCol_ParmTypeID)))) = kSN_RetBars Then
                .RowHidden(lRow) = True
                CheckedCell(vsInputs, lRow, GDCol(eGDCol_Req)) = True
                If UCase(.TextMatrix(lRow, GDCol(eGDCol_InputName))) = "MARKET1" Then
                    .RowPosition(lRow) = .FixedRows
                Else
                    ' TLB 5/17/05: add to end of markets so end up at top but in original order
                    For lRow2 = .FixedRows To lRow - 1
                        If CLng(ValOfText(.TextMatrix(lRow2, GDCol(eGDCol_ParmTypeID)))) <> kSN_RetBars Then
                            .RowPosition(lRow) = lRow2
                            Exit For
                        End If
                    Next
                    'If UCase(.TextMatrix(.FixedRows, GDCol(eGDCol_InputName))) = "MARKET1" Then
                    '    .RowPosition(lRow) = .FixedRows + 1
                    'Else
                    '    .RowPosition(lRow) = .FixedRows
                    'End If
                End If
            End If
        Next lRow
        
        .Redraw = lRedraw
    End With
    
    RefreshGrid

ErrExit:
    Exit Sub
    
ErrSection:
    vsInputs.Redraw = lRedraw
    RaiseError "frmFunctionMgrCT.RefreshInputs"
    
End Sub

Private Sub EnableButtons()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lFirstRow As Long

    With vsInputs
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                lFirstRow = lIndex
                Exit For
            End If
        Next lIndex
        
        cmdMoveUp.Enabled = (.Rows > .FixedRows) And (.Row > lFirstRow)
        cmdMoveDown.Enabled = (.Rows > .FixedRows) And (.Row < .Rows - 1)
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.EnableButtons"
    
End Sub

Private Function VerifyUsage(ByVal hRefsArray As Long, Optional ByVal bShowMsg As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim alFunctions As cGdArray         ' Array of Function ID's used in rule
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim astrFunctions As cGdArray       ' Array of Functions that cannot be used
    Dim strName As String               ' Name of a Function
    Dim lPos As Long                    ' Position to insert into Function Name array
    
    Set alFunctions = New cGdArray
    alFunctions.Create eGDARRAY_Longs
    Set astrFunctions = New cGdArray
    astrFunctions.Create eGDARRAY_Strings
    
    If alFunctions.CopyFromHandle(hRefsArray) Then
        For lIndex = 0 To alFunctions.Size - 1
            VerifyFunctionUsage alFunctions(lIndex), astrFunctions
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionRefs] " & _
                    "WHERE [FunctionID]=" & Str(alFunctions(lIndex)) & ";", dbOpenDynaset)
            Do While Not rs.EOF
                If VerifyFunctionUsage(rs!FunctionIDRef) = False Then
                    strName = FunctionNameFromID(alFunctions(lIndex))
                    If astrFunctions.BinarySearch(strName, lPos) = False Then
                        astrFunctions.Add strName, lPos
                    End If
                End If
                rs.MoveNext
            Loop
        Next lIndex
    End If
    
    If (astrFunctions.Size > 0) And (chkSystemTesting = vbChecked) Then
        VerifyUsage = False
        If bShowMsg Then
            InfBox "The following function(s) cannot be used in|strategy testing:||" & _
                astrFunctions.JoinFields("|") & "||Either change the expression or " & _
                "uncheck the|Strategy Testing box", , , "Validation Error"
        End If
    Else
        VerifyUsage = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionMgrCT.VerifyUsage"
    
End Function

Private Function VerifyFunctionUsage(ByVal lID As Long, Optional astrFunctions As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim rsFunction As Recordset
    Dim lPos As Long

    VerifyFunctionUsage = True
    
    Set rsFunction = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
            "WHERE [FunctionID]=" & Str(lID) & ";", dbOpenDynaset)
    If Not (rsFunction.BOF And rsFunction.EOF) Then
        If GetBit(rsFunction!Usage, 2) = False Then
            If Not astrFunctions Is Nothing Then
                If astrFunctions.BinarySearch(rsFunction!FunctionName, lPos) = False Then
                    astrFunctions.Add rsFunction!FunctionName, lPos
                End If
                If m.Function.FunctionCategoryID = kScoringCategory Then chkSystemTesting.Value = 0
            End If
            VerifyFunctionUsage = False
        End If
        
'aardvark 3586 fix
        If rsFunction!FunctionCategoryID = kScoringCategory And Not astrFunctions Is Nothing Then
            If rsFunction!FunctionName <> "IFF" Then
                If GetBit(rsFunction!Usage, 1) = False Then chkMM.Value = 0
                If GetBit(rsFunction!Usage, 3) = False Then chkCharting.Value = 0
                If GetBit(rsFunction!Usage, 4) = False Then chkCriteria.Value = 0
            End If
        End If
    End If

ErrExit:
    Set rsFunction = Nothing
    Exit Function
    
ErrSection:
    Set rsFunction = Nothing
    RaiseError "frmFunctionMgrCT.VerifyFunctionUsage"
    
End Function

Private Sub EnablePlanetGrid(ByVal bShow As Boolean)
On Error GoTo ErrSection:

    Dim i&, nNewRow&, nValue#
    
    If bShow = False Then
        fgPlanet.Visible = False
        Exit Sub
    End If
    
    Label1(0).Visible = False       'Usage
    TradeSense.Visible = False
    chkAdvanced.Visible = False
    chkAutoMultiplier.Visible = False
    vsIndexTab1.Visible = False
    fgPlanet.Visible = True
    fraFgButtons.Visible = True

    nNewRow = -1
    With fgPlanet
        'process body1 change
        If .Col = ePLCol_Body1 Then
            If .TextMatrix(.Row, ePLCol_Body1) <> "-999" Then
                .Cell(flexcpChecked, .Row, ePLCol_Use) = flexChecked
                'set defaults if necessary
                If .TextMatrix(.Row, ePLCol_Body2) = "" Then .TextMatrix(.Row, ePLCol_Body2) = -999
                If .TextMatrix(.Row, ePLCol_PlanetSystem) = "" Then .TextMatrix(.Row, ePLCol_PlanetSystem) = 0
                If .TextMatrix(.Row, ePLCol_Value) = "" Then .TextMatrix(.Row, ePLCol_Value) = "Longitude"
                If .TextMatrix(.Row, ePLCol_Weight) = "" Then .TextMatrix(.Row, ePLCol_Weight) = 1
                If .TextMatrix(.Row, ePLCol_Offset) = "" Then .TextMatrix(.Row, ePLCol_Offset) = 0
                If .TextMatrix(.Row, ePLCol_Harmonic) = "" Then .TextMatrix(.Row, ePLCol_Harmonic) = 1
                If .TextMatrix(.Row, ePLCol_Orb) = "" Then .TextMatrix(.Row, ePLCol_Orb) = 0
            End If
        End If
        'validate values for weight, offset, harmonic and orb
        nValue = ValOfText(.TextMatrix(.Row, ePLCol_Weight))
        If nValue > 1000# Or nValue < -1000# Then
            MsgBox "Weight must be between -1000 and 1000"
            .TextMatrix(.Row, ePLCol_Weight) = 1#
            Exit Sub
        End If
        nValue = ValOfText(.TextMatrix(.Row, ePLCol_Offset))
        If nValue > 360# Or nValue < -360# Then
            MsgBox "Offset must be between -360 and 360"
            .TextMatrix(.Row, ePLCol_Offset) = 0#
            Exit Sub
        End If
        nValue = ValOfText(.TextMatrix(.Row, ePLCol_Harmonic))
        If nValue > 1000# Or nValue < -1000# Then
            MsgBox "Harmonic must be between -1000 and 1000"
            .TextMatrix(.Row, ePLCol_Harmonic) = 1#
            Exit Sub
        End If
        nValue = ValOfText(.TextMatrix(.Row, ePLCol_Orb))
        If nValue > 180# Or nValue < -180# Then
            MsgBox "Offset must be between -180 and 180"
            .TextMatrix(.Row, ePLCol_Orb) = 0#
            Exit Sub
        End If
        'look for available 'new' row
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, ePLCol_Body1) = "-999" Then
                nNewRow = i
                Exit For
            End If
        Next
        
        'set values for new row
        If nNewRow = -1 Then
            .Rows = .Rows + 1
            nNewRow = .Rows - 1
        End If
        .Cell(flexcpChecked, nNewRow, ePLCol_Use) = flexUnchecked
        .TextMatrix(nNewRow, ePLCol_Body1) = -999
        For i = ePLCol_Body2 To kPlanetCols - 1
            .TextMatrix(nNewRow, i) = ""
        Next
        
        'disallow changing to another category
        If .Rows > 2 Then
            cboCategory.Enabled = False
        End If
        
    End With
    
    EnableFgButtons -1
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.EnablePlanetGrid"

End Sub

Private Sub InitPlanetGrid()
On Error GoTo ErrSection:
        
    With fgPlanet
        .Redraw = flexRDNone
        SetupGrid Me.fgPlanet, eGridMode_Grid
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 1
        .Cols = kPlanetCols
        'show/no show check box
        .ColWidth(ePLCol_Use) = kZeroColWidth
        .ColDataType(ePLCol_Use) = flexDTBoolean
        'combolist dropdown
        .ColComboList(ePLCol_Body1) = kBodyList
        .ColComboList(ePLCol_Body2) = kBodyList
        .ColComboList(ePLCol_PlanetSystem) = kSystemList
        'format for weight, offset, harmonic & orb data
        .ColFormat(ePLCol_Weight) = ".00"
        .ColFormat(ePLCol_Offset) = ".00"
        .ColFormat(ePLCol_Harmonic) = ".00"
        .ColFormat(ePLCol_Orb) = ".00"
        'column headers
        .TextMatrix(0, ePLCol_Use) = "Use"
        .TextMatrix(0, ePLCol_Body1) = "Body 1"
        .TextMatrix(0, ePLCol_Body2) = "Body 2"
        .TextMatrix(0, ePLCol_PlanetSystem) = "System"
        .TextMatrix(0, ePLCol_Value) = "Value"
        .TextMatrix(0, ePLCol_Weight) = "Weight"
        .TextMatrix(0, ePLCol_Offset) = "Offset(deg)"
        .TextMatrix(0, ePLCol_Harmonic) = "Harmonic"
        .TextMatrix(0, ePLCol_Orb) = "Orb(deg)"
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.InitPlanetGrid"

End Sub

Private Sub BuildPlanetText()
On Error GoTo ErrSection:

    Dim i&, nValue&, strFunc$
    Dim strInclude$, strBody1$, strBody2$
    Dim strSystem$, strValue$, strOffset$
    Dim strHarmonic$, strOrb$, strWeight$

    With fgPlanet
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, ePLCol_Use) = flexChecked Then
                strInclude = "1"
            Else
                strInclude = "0"
            End If
            strBody1 = .TextMatrix(i, ePLCol_Body1)
            strBody2 = .TextMatrix(i, ePLCol_Body2)
            strSystem = .TextMatrix(i, ePLCol_PlanetSystem)
            strValue = .TextMatrix(i, ePLCol_Value)
            strOffset = .TextMatrix(i, ePLCol_Offset)
            strHarmonic = .TextMatrix(i, ePLCol_Harmonic)
            strOrb = .TextMatrix(i, ePLCol_Orb)
            strWeight = .TextMatrix(i, ePLCol_Weight)
            
            If strBody1 = "None" Then strBody1 = "-999"
            If strBody2 = "None" Then strBody2 = "-999"
            If strSystem = "Geocentric" Then strSystem = "0"
            If strValue = "Longitude" Or strValue = "Rectascension" Then strValue = "0"
                        
            'convert value string to numeric
            Select Case strValue
                Case "Longitude", "Rectascension"
                    nValue = 0
                Case "Latitude", "Declination"
                    nValue = 1
                Case "Distance"
                    nValue = 2
                Case "Speed"
                    nValue = 3
                Case "Acceleration"
                    nValue = 4
                Case "Aspect"
                    nValue = 5
                Case Else
                    nValue = 0
            End Select
            strValue = Str(nValue)
            
            If strBody1 <> "-999" Then
                If Len(strFunc) > 0 Then
                    strFunc = strFunc & " + "
                End If
                strFunc = strFunc & "IFF(" & strInclude & ", " & strWeight & " * " & "Planet Position(" _
                          & Chr(34) & strBody1 & Chr(34) & ", " & Chr(34) & strBody2 & Chr(34) & ", " _
                          & strSystem & ", " & strValue & ", " & strOffset _
                          & ", " & strHarmonic & ", " & strOrb & "), 0)"
            End If

        Next
    End With
    
    InitEditor
    Editor1.Text = strFunc
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.BuildPlanetText"

End Sub

Private Sub LoadPlanetGrid()
On Error GoTo ErrSection:

    Dim i&, j&, nSystem&, strField$, strExpr$
    Dim aData As New cGdArray
    Dim aFields As New cGdArray
    
    'Sample function string:
    ' IFF(1, 3.5 * Planet Position("4", "5", 0, 0, 0, 3, 2), 0) + IFF(1, 4 * Planet Position("0", "-999", 0, 0, 0, 3, 2), 0)
    'Splitting field at "+" places IFF(...) string arrays into aData.
    'Chop at right paren, strip quotes, replace left paren and asterisk with comma,
    'then split each string of aData into aFields at comma results in 11 strings as follows:
    ' aFields[0] = IFF
    ' aFields[1] = 1         '1=value of check box column
    ' aFields[2] = 3.5       ' weight
    ' aFields[3] = Planet Position
    ' aFields[4] = 4         ' body1
    ' aFields[5] = 5         ' body2
    ' aFields[6] = 0         '0=planet system (heliocentric etc.)
    ' aFields[7] = 0         '0=value (longitude etc.)
    ' aFields[8] = 0         '0=offset
    ' aFields[9] = 3         '3=harmonic
    ' aFields[10] = 2        '2=orb
   
    aData.Clear
    aData.SplitFields Editor1.Text, "+"
    
    If aData.Size < 1 Then Exit Sub

    With fgPlanet
        .Redraw = flexRDNone
        .Rows = 1
        For i = 0 To aData.Size - 1
            .Rows = .Rows + 1
            strExpr = Trim(aData(i))
            ' chop at right paren (in case other optional args had been added later)
            j = InStr(strExpr, ")")
            If j > 0 Then strExpr = Trim(Left(strExpr, j - 1))
            ' strip out quotes
            strExpr = Replace(strExpr, Chr(34), "")
            ' replace left paren and asterisk with a comma, then split at commas
            strExpr = Replace(strExpr, "(", ",")
            strExpr = Replace(strExpr, "*", ",")
            aFields.SplitFields strExpr, ","
            If aFields.Size >= 11 Then
                ' if used or not
                If Val(aFields(1)) = 0 Then
                    .Cell(flexcpChecked, .Rows - 1, ePLCol_Use) = flexUnchecked
                Else
                    .Cell(flexcpChecked, .Rows - 1, ePLCol_Use) = flexChecked
                End If
                
                ' weight
                .TextMatrix(.Rows - 1, ePLCol_Weight) = Val(aFields(2))
                
                ' planet bodies
                .TextMatrix(.Rows - 1, ePLCol_Body1) = Val(aFields(4))
                strField = Trim(aFields(5))
                If Len(strField) = 0 Then
                    .TextMatrix(.Rows - 1, ePLCol_Body2) = -999
                Else
                    .TextMatrix(.Rows - 1, ePLCol_Body2) = Val(strField)
                End If
                
                'parse system
                nSystem = Val(aFields(6))
                .TextMatrix(.Rows - 1, ePLCol_PlanetSystem) = nSystem
                
                'convert numeric value to string description
                strField = ""
                Select Case Val(aFields(7))
                    Case 0
                        If nSystem = 2 Then
                            strField = "Rectascension"
                        Else
                            strField = "Longitude"
                        End If
                    Case 1
                        If nSystem = 2 Then
                            strField = "Declination"
                        Else
                            strField = "Latitude"
                        End If
                    Case 2
                        strField = "Distance"
                    Case 3
                        strField = "Speed"
                    Case 4
                        strField = "Acceleration"
                    Case 5
                        strField = "Aspect"
                End Select
                .TextMatrix(.Rows - 1, ePLCol_Value) = strField
                
                'parse offset, harmonic, and orb
                .TextMatrix(.Rows - 1, ePLCol_Offset) = Val(aFields(8))
                .TextMatrix(.Rows - 1, ePLCol_Harmonic) = Val(aFields(9))
                .TextMatrix(.Rows - 1, ePLCol_Orb) = Val(aFields(10))
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    
    EnablePlanetGrid True
     
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.LoadPlanetGrid"

End Sub

Private Sub RemoveCategory(ByVal strCat$)
On Error GoTo ErrSection:

    Dim i&
        
    For i = 0 To cboCategory.ListCount - 1
        If cboCategory.List(i) = strCat Then
            cboCategory.RemoveItem (i)
            Exit For
        End If
    Next

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.RemoveCategory"
    
End Sub

Private Sub InitEditor()
On Error GoTo ErrSection:
    
    With Editor1
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = False
        .ShowNewFunction = True
        .Usage = 2             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.InitEditor"
End Sub

Private Sub EnableFgButtons(ByVal Row&)
On Error GoTo ErrSection:
    
    Dim fg As VSFlexGrid
'    Dim bSpread As Boolean
    
    If fgPlanet.Visible Then
        Set fg = fgPlanet
    ElseIf fgScoring.Visible Then
        Set fg = fgScoring
    Else
        cmdFgAdd.Enabled = False
        cmdFgDelete.Enabled = False
        cmdFgUp.Enabled = False
        cmdFgDown.Enabled = False
        Exit Sub
    End If
            
    If fg.Rows - fg.FixedRows > 1 Then
        cmdFgDelete.Enabled = True
        cmdFgAdd.Enabled = True
        With fg
            If Row > .FixedRows Then
                cmdFgUp.Enabled = True
            Else
                cmdFgUp.Enabled = False
            End If
            If Row < .Rows - .FixedRows Then
                cmdFgDown.Enabled = True
            Else
                cmdFgDown.Enabled = False
            End If
        End With
    Else
        cmdFgAdd.Enabled = False
        cmdFgDelete.Enabled = False
        cmdFgUp.Enabled = False
        cmdFgDown.Enabled = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.EnableButtons"
    
End Sub

Private Sub InitScoringGrid()
On Error GoTo ErrSection:
        
    m.strConditions = BooleanConditions()
    
    With fgScoring
        .Redraw = flexRDNone
        SetupGrid Me.fgScoring, eGridMode_Grid
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 1
        .Cols = kScoringCols
        .ColComboList(eSCRCol_Condition) = m.strConditions
        'show/no show check box
        .ColWidth(eSCRCol_Use) = kZeroColWidth
        .ColDataType(eSCRCol_Use) = flexDTBoolean
        .ColDataType(eSCRCol_Points) = flexDTDouble
        'column headers
        .TextMatrix(0, eSCRCol_Use) = "Use"
        .TextMatrix(0, eSCRCol_Points) = "Points"
        .TextMatrix(0, eSCRCol_Condition) = "Condition"
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.InitScoringGrid"

End Sub

Private Sub EnableScoringGrid(ByVal bShow As Boolean, Optional ByVal bLeaveVisible As Boolean = False)
On Error GoTo ErrSection:

    Dim i&, nNewRow&, nValue#
    
    If Not bLeaveVisible Then
        If bShow = False Then
            fgScoring.Visible = False
            Exit Sub
        End If
        
        Label1(0).Visible = False
        TradeSense.Visible = False
        chkAdvanced.Visible = False
        chkAutoMultiplier.Visible = False
        vsIndexTab1.Visible = False
        fraFgButtons.Visible = True
    End If
        
    nNewRow = -1
    With fgScoring
        .Visible = True
        
        'look for available 'new' row
        For i = .FixedRows To .Rows - 1
            If Len(.TextMatrix(i, eSCRCol_Condition)) < 1 Then
                nNewRow = i
                Exit For
            End If
        Next
        
        'set values for new row
        If nNewRow = -1 Then
            .Rows = .Rows + 1
            nNewRow = .Rows - 1
        End If
        .Cell(flexcpChecked, nNewRow, ePLCol_Use) = flexUnchecked
        .TextMatrix(nNewRow, eSCRCol_Points) = "1"
        .TextMatrix(nNewRow, eSCRCol_Condition) = ""
        
        'disallow changing to another category
        If .Rows > 2 Then
            cboCategory.Enabled = False
        End If
    End With
    
    EnableFgButtons -1
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.EnableScoringGrid"

End Sub

Private Sub EditScoringCondition(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strUserText$, strCodedText$
    Static rc&
    
    
    strUserText = m.strSaveCondtion
    
    rc = rc + 1
    If rc < 2 Then
        frmCustomFunction.Caption = "Condition"
        rc = frmCustomFunction.ShowMe(strUserText, strCodedText, , , False)
        'rc: -1=Cancelled, 0=Invalid, 1=Numeric, 2=Boolean
        If rc = 2 Then
            fgScoring.TextMatrix(Row, eSCRCol_Condition) = strUserText
            fgScoring.Cell(flexcpChecked, Row, eSCRCol_Use) = flexChecked
            EnableScoringGrid True, True
            tbToolbar.Tools("ID_Verify").Enabled = True
            tbToolbar.Tools("ID_Save").Enabled = True
        ElseIf rc = 1 Then
            MsgBox "Only conditions returning a true or false value can be used."
        End If
        
        If fgScoring.TextMatrix(Row, eSCRCol_Condition) = "<Custom Condition>" Then
            If Len(m.strSaveCondtion) > 0 Then
                fgScoring.TextMatrix(Row, eSCRCol_Condition) = m.strSaveCondtion
            Else
                If Row = fgScoring.Rows - 1 Then
                    fgScoring.TextMatrix(Row, eSCRCol_Condition) = ""
                Else
                    'user added a row with the add row button then cancelled or entered an invalid custom condition
                    fgScoring.RemoveItem Row
                End If
            End If
        End If
        
        rc = 0
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.EditScoringCondition"

End Sub

Private Sub BuildScoringText()
On Error GoTo ErrSection:

    Dim i&, strFunc$
    Dim strUse$, strCondition$, strPoints$

    With fgScoring
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, ePLCol_Use) = flexChecked Then
                strUse = "1"
            Else
                strUse = "0"
            End If
            strCondition = .TextMatrix(i, eSCRCol_Condition)
            If Len(strCondition) > 0 Then
                strPoints = ValOfText(.TextMatrix(i, eSCRCol_Points))
                If Len(strFunc) > 0 Then
                    strFunc = strFunc & " + "
                End If
                strFunc = strFunc & "IFF(" & strUse & " and (" & strCondition & "), " & strPoints & ", 0)"
            End If
        Next
    End With
    
    InitEditor
    Editor1.Text = strFunc
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.BuildScoringText"

End Sub

Private Sub LoadScoringGrid()
On Error GoTo ErrSection:

    Dim i&, j&, nLen&, nParenLevel&
    Dim strExpr$, strChar$
    Dim aFields As New cGdArray
    Dim aData As New cGdArray
    Dim bParenFound As Boolean
      
    strExpr = Editor1.Text
    nLen = Len(strExpr)
    If nLen < 1 Then Exit Sub
    
    'Sample function string:
    'selected left/right parentheses and commas are replaced with tabs and return chars as shown below
    'IFF(1 and (condittion string), 1.5, 0) + IFF(0 and (condittion string), 0.5, 0)
    '   ^      ^                 ^^    ^  ^      ^      ^                 ^        ^
    '   T      T                 TT    T  R      T      T                 TT    T  R
    For i = 1 To nLen
        strChar = Mid(strExpr, i, 1)
        Select Case strChar
            Case "("
                nParenLevel = nParenLevel + 1
                If nParenLevel = 1 Or nParenLevel = 2 Then Mid(strExpr, i, 1) = vbTab
            Case ")"
                nParenLevel = nParenLevel - 1
                If nParenLevel = 0 Then
                    Mid(strExpr, i, 1) = vbCrLf
                ElseIf nParenLevel = 1 Then
                    Mid(strExpr, i, 1) = vbTab
                End If
        End Select
    Next
   
    'Splitting field at return char places ..IFF(...) string arrays into aData.
    aData.Clear
    aData.SplitFields strExpr, vbCrLf
    
    'Splitting each string of aData into aFields at [tab] results in 6 strings as follows:
    '  aFields[0] = IFF
    '  aFields[1] = 1 and    '1=value of use check box column
    '  aFields[2] = text     'condition string
    '  aFields[3] = blank    'ignored
    '  aFields[4] = 1.5      'points
    '  aFields[5] = 0        'ignored
    With fgScoring
        .Redraw = flexRDNone
        .Rows = 1
        For i = 0 To aData.Size - 1
            strExpr = Trim(aData(i))
            If Len(strExpr) > 0 Then
                .Rows = .Rows + 1
                'reverse the string
                strExpr = StrReverse(strExpr)
                'replace first 2 commas with tab
                strExpr = Replace(strExpr, ",", vbTab, , 1)
                strExpr = Replace(strExpr, ",", vbTab, , 1)
                'reverse string to normal
                strExpr = StrReverse(strExpr)
                'split string at tabs
                aFields.SplitFields strExpr, vbTab
                If aFields.Size >= 6 Then
                    'if used or not
                    If Val(Left(aFields(1), 1)) = 0 Then
                        .Cell(flexcpChecked, .Rows - 1, eSCRCol_Use) = flexUnchecked
                    Else
                        .Cell(flexcpChecked, .Rows - 1, eSCRCol_Use) = flexChecked
                    End If
                    'condition
                    .TextMatrix(.Rows - 1, eSCRCol_Condition) = aFields(2)
                    'points
                    .TextMatrix(.Rows - 1, eSCRCol_Points) = Val(aFields(4))
                End If
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    
    EnableScoringGrid True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionMgrCT.LoadScoringGrid"

End Sub

Private Function BooleanConditions() As String
On Error GoTo ErrSection:

    Dim strCond$
    Dim Criteria As cCriteria
    
    For Each Criteria In g.SymbolPool.Criterias
        With Criteria
            If .IsBoolean Then
                strCond = strCond & "|" & .Name
            End If
        End With
    Next
    
    BooleanConditions = strCond

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmFunctionMgrCT.BooleanConditions"

End Function

Private Function BooleanConditionText(ByVal strCond$) As String
On Error GoTo ErrSection:

    Dim Criteria As cCriteria

    BooleanConditionText = ""
    For Each Criteria In g.SymbolPool.Criterias
        With Criteria
            If .IsBoolean And .Name = strCond Then
                BooleanConditionText = Criteria.EnglishText
                Exit For
            End If
        End With
    Next
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmFunctionMgrCT.BooleanConditionText"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidMarket
'' Description: Validate any "Symbol,Period" markets
'' Inputs:      Market
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidMarket(ByVal strMarket As String) As Boolean
On Error GoTo ErrSection:
    
    Dim strSymbol As String             ' Symbol of the given market
    Dim strPeriod As String             ' Period of the given market
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    ValidMarket = False
    strMarket = UCase(Trim(strMarket))
    If Left(strMarket, 1) = Chr(34) And Right(strMarket, 1) = Chr(34) Then
        strSymbol = Parse(Replace(strMarket, Chr(34), ""), ",", 1)
        strPeriod = Parse(Replace(strMarket, Chr(34), ""), ",", 2)
        ' TLB: we can just ignore Daily/Weekly since allowed everywhere
        If strPeriod = "WEEKLY" Or strPeriod = "DAILY" Then
            strPeriod = ""
        End If
        
        If Len(strSymbol) > 0 And Len(strPeriod) > 0 Then
            Bars.Prop(eBARS_PeriodicityStr) = strPeriod
            If Bars.Prop(eBARS_Periodicity) < ePRD_Days Then
                DM_GetBars Bars, strSymbol, strPeriod, LastDailyDownload - 5
            Else
                DM_GetBars Bars, strSymbol, strPeriod
            End If
            If Bars.Size > 0 Then ValidMarket = True
        ElseIf Len(strSymbol) > 0 Then
            If Right(strSymbol, 1) <> ":" And Left(strSymbol, 1) <> "-" Then
                DM_GetBars Bars, strSymbol, "Daily"
                If Bars.Size > 0 Then ValidMarket = True
            ElseIf chkCriteria.Value = vbUnchecked Or strSymbol = "-067" Then
                ValidMarket = True ' TLB: we now allow "-067" for criteria as a special case
            End If
        ElseIf Len(strPeriod) > 0 Then
            If chkCriteria.Value = vbUnchecked Then
                ValidMarket = True
            End If
        End If
    Else
        ValidMarket = True
    End If
    Set Bars = Nothing

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionMgrCT.ValidMarket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPeriodInMarkets
'' Description: Fix the Period in "Of" expressions surrounded by quotes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixPeriodInMarkets()
On Error GoTo ErrSection:

    Dim astrTokens As New cGdArray      ' Array of space delimited tokens
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol of the market variable
    Dim strPeriod As String             ' Period of the market variable
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    astrTokens.SplitFields Editor1.Text, " "
    For lIndex = 0 To astrTokens.Size - 1
        If UCase(astrTokens(lIndex)) = "OF" Then
            If lIndex + 1 < astrTokens.Size Then
                If Left(astrTokens(lIndex + 1), 1) = Chr(34) And Right(astrTokens(lIndex + 1), 1) = Chr(34) Then
                    strSymbol = Parse(Replace(astrTokens(lIndex + 1), Chr(34), ""), ",", 1)
                    strPeriod = Parse(Replace(astrTokens(lIndex + 1), Chr(34), ""), ",", 2)
                    
                    If Len(strPeriod) > 0 Then
                        Bars.Prop(eBARS_PeriodicityStr) = strPeriod
                        strPeriod = Bars.Prop(eBARS_PeriodicityStr)
                        
                        astrTokens(lIndex + 1) = Chr(34) & strSymbol & "," & strPeriod & Chr(34)
                    Else
                        astrTokens(lIndex + 1) = Chr(34) & strSymbol & Chr(34)
                    End If
                End If
            End If
        End If
    Next lIndex
    Editor1.Text = astrTokens.JoinFields(" ")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.FixPeriodInMarkets"
    
End Sub

Private Sub EnableSpreadGrid(ByVal bShow As Boolean)
On Error GoTo ErrSection:

    Dim i&, nNewRow&, nValue#
    
    If bShow = False Then
        fgSpread.Visible = False
        Exit Sub
    End If

    Label1(0).Visible = False
    TradeSense.Visible = False
    chkAdvanced.Visible = False
    chkAutoMultiplier.Visible = True
    vsIndexTab1.Visible = False
    fgSpread.Visible = True
    fraFgButtons.Visible = True

    With fgSpread
        .Rows = .FixedRows + 1
        .TextMatrix(.FixedRows, 0) = ""
        .TextMatrix(.FixedRows, 1) = ""
        .TextMatrix(.FixedRows, 2) = "1"
        .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.EnableSpreadGrid"
  
End Sub

Private Sub InitSpreadGrid()
On Error GoTo ErrSection:

    With fgSpread
        .Redraw = flexRDNone
        SetupGrid Me.fgSpread, eGridMode_Grid
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 1
        .Cols = 5
        'column headers
        .TextMatrix(0, 0) = "Plus/Minus/Divide (+ - /)"
        .TextMatrix(0, 1) = "Symbol"
        .TextMatrix(0, 2) = "Multiplier"
        .TextMatrix(0, 3) = "#Contracts"
        .TextMatrix(0, 4) = "Multiplier User"    'hidden column that saves multiplier input by user
        'button & dropdown for columns
        .ColComboList(0) = kOpAll
        .ColComboList(1) = "..."
        'alignment
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        'hidden columns
        .ColHidden(4) = True
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.InitSpreadGrid"
  
End Sub

Private Function BuildSpreadText() As Boolean
On Error GoTo ErrSection:

    Dim bSuccess As Boolean
    Dim strExpr As String
    
    strExpr = BuildSpreadExpr(fgSpread)
    If Len(strExpr) > 0 Then
        bSuccess = True
        'strip off the last +/-
        InitEditor
        Editor1.Text = strExpr
    End If
    
    BuildSpreadText = bSuccess

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionMgrCT.BuildSpreadText"

End Function

Private Function LoadSpreadGrid() As String
On Error GoTo ErrSection:

    Dim i&
    Dim strExpr$
    Dim aTemp As New cGdArray
    Dim tbData As cGdTable
    
    Dim bDivide As Boolean
      
    strExpr = Editor1.Text
    If Len(strExpr) < 1 Then Exit Function
    
    Set tbData = SpreadExprToTable(strExpr, bDivide)
    If tbData Is Nothing Then Exit Function
    If tbData.NumRecords = 0 Then Exit Function
           
    strExpr = ""
    If bDivide Then
        If tbData.NumRecords = 2 Then
            'ratio spreads can only have 2 symbols
            With fgSpread
                .Rows = .FixedRows + 2
                .TextMatrix(.FixedRows, 0) = ""
                .TextMatrix(.FixedRows, 1) = tbData(1, 0)
                .TextMatrix(.FixedRows, 2) = tbData(2, 0)
                .TextMatrix(.FixedRows, 3) = tbData(3, 0)
                .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                .TextMatrix(.FixedRows + 1, 0) = tbData(0, 1)
                .TextMatrix(.FixedRows + 1, 1) = tbData(1, 1)
                .TextMatrix(.FixedRows + 1, 2) = tbData(2, 1)
                .TextMatrix(.FixedRows + 1, 3) = tbData(3, 1)
                'save multiplier to hidden column
                .TextMatrix(.FixedRows, 4) = .TextMatrix(.FixedRows, 2)
                .TextMatrix(.FixedRows + 1, 4) = .TextMatrix(.FixedRows + 1, 2)
                'rebuild the spread text for input into function for determining auto multiplier
                strExpr = .TextMatrix(.FixedRows, 0) & "," & .TextMatrix(.FixedRows, 1) & "," & .TextMatrix(.FixedRows, 2) & ";"
                strExpr = .TextMatrix(.FixedRows + 1, 0) & "," & .TextMatrix(.FixedRows + 1, 1) & "," & .TextMatrix(.FixedRows + 1, 2) & ";"
            End With
        End If
        LoadSpreadGrid = strExpr
        cmdFgAdd.Enabled = False
        Exit Function
    End If
    
    aTemp.Clear
    Set aTemp = tbData.CreateSortedIndex(0, eGdSort_Descending Or eGdSort_Stable)
    If aTemp.Size < 1 Then
        Set tbData = Nothing
        Set aTemp = Nothing
        Exit Function
    End If
    
    fgSpread.Rows = fgSpread.FixedRows
    For i = 0 To aTemp.Size - 1
        With fgSpread
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = tbData(0, aTemp(i))
            .TextMatrix(.Rows - 1, 1) = tbData(1, aTemp(i))
            .TextMatrix(.Rows - 1, 2) = tbData(2, aTemp(i))
            .TextMatrix(.Rows - 1, 3) = tbData(3, aTemp(i))
            'save multiplier to hidden column
            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 2)
            'rebuild the spread text for input into function for determining auto multiplier
            strExpr = strExpr & .TextMatrix(.Rows - 1, 0) & "," & .TextMatrix(.Rows - 1, 1) & "," & .TextMatrix(.Rows - 1, 2) & ";"
        End With
    Next
    
    With fgSpread
        If .Rows >= .FixedRows Then
            .TextMatrix(.FixedRows, 0) = ""
            .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
        End If
    End With
        
    AddBlankRow
    cboCategory.Enabled = False
    
    Set tbData = Nothing
    Set aTemp = Nothing
    
    LoadSpreadGrid = strExpr

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionMgrCT.LoadSpreadGrid"

End Function

Private Function AddBlankRow(Optional ByVal bNewGrid As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim strOp$, strSym$, strMult$, i&
    Dim bNewRow As Boolean
    
    If bNewGrid Then
        With fgSpread
            .Rows = .FixedRows + 1
            .Cell(flexcpText, .FixedRows, 0, .FixedRows, .Cols - 1) = "Click here to get started ..."
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.FixedRows) = True
            .Cell(flexcpBackColor, .FixedRows, 0, .FixedRows, .Cols - 1) = vbWhite
        End With
        Exit Function
    End If
    
    'see if new blank row should be added
    bNewRow = True
    With fgSpread
        If .MergeRow(.Rows - 1) = True Then
            bNewRow = False
            'if user originally selected + or minus for the second row
            'then changed it to divide, then need to remove the blank row
            If .Rows > .FixedRows + 1 Then
                If .TextMatrix(.FixedRows + 1, 0) = kDivide Then
                    .RemoveItem .Rows - 1
                End If
            End If
            AddBlankRow = .Rows - 1
        Else
            For i = .FixedRows To .Rows - 1
                If i = .FixedRows Then
                    strOp = kPlus
                Else
                    strOp = .TextMatrix(i, 0)
                End If
                strSym = .TextMatrix(i, 1)
                strMult = .TextMatrix(i, 2)
                If Len(strOp) = 0 Or Len(strSym) = 0 Or Len(strMult) = 0 Or strOp = kDivide Then
                    bNewRow = False
                    AddBlankRow = i
                    Exit For
                End If
            Next
        End If
    End With
    
    If bNewRow Then
        With fgSpread
            .Rows = .Rows + 1
            If .Rows - 1 = .FixedRows Then
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to get started ..."
            Else
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to add another row ..."
            End If
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
        End With
    End If
    
    With fgSpread
        If .Rows > .FixedRows + 1 Then
            If .TextMatrix(.FixedRows + 1, 0) = kDivide Then
                cmdFgAdd.Enabled = False
            Else
                cmdFgAdd.Enabled = True
            End If
        Else
            cmdFgAdd.Enabled = True
        End If
    End With
    
    cmdFgDelete.Enabled = True
    
    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".AddBlankRow"
        
End Function

Private Sub NewGridRow()
On Error GoTo ErrSection:
    
    With fgSpread
        If .MergeRow(.Rows - 1) = True Then
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = ""
            .MergeRow(.Rows - 1) = False
            If .Rows > .FixedRows + 2 Then
                .ColComboList(0) = kOpPlusMinus
            Else
                .ColComboList(0) = kOpAll
            End If
            .Row = .Rows - 1
            .Col = 0
            If .Row = .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                .Col = 2
                fgSpread_CellButtonClick .FixedRows, 2
            Else
                .EditCell
                SendKeys "{F4}" '(to dropdown the combo list)
            End If
        End If
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".NewGridRow"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ResaveLog
'' Description: Logging for the resave functions mode
'' Inputs:      String to Log
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ResaveLog(ByVal strToLog As String)
On Error Resume Next

    Dim fh As Integer                   ' File handle to open file with

    fh = FreeFile
    Open AddSlash(App.Path) & "Resave\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strToLog
        Close #fh
    End If

End Sub

