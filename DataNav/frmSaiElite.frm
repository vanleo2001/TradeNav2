VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSaiElite 
   Caption         =   "SAI Report"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   14190
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   4560
      Top             =   1020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "frmSaiElite.frx":0000
      ToolBars        =   "frmSaiElite.frx":3431
   End
   Begin HexUniControls.ctlUniFrameWL fraDate 
      Height          =   1575
      Left            =   420
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
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
      Caption         =   "frmSaiElite.frx":35EC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSaiElite.frx":3638
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSaiElite.frx":3658
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRunForDate 
         Height          =   375
         Left            =   420
         TabIndex        =   2
         Top             =   960
         Width           =   2175
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
         Caption         =   "frmSaiElite.frx":3674
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSaiElite.frx":36C0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSaiElite.frx":36E0
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdReportDate 
         Height          =   315
         Left            =   420
         TabIndex        =   1
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         AllowWeekends   =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   4020
      Left            =   -60
      TabIndex        =   3
      Top             =   60
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   7091
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
      Caption         =   $"frmSaiElite.frx":36FC
      Align           =   0
      Appearance      =   1
      CurrTab         =   2
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
      Begin VSFlex7LCtl.VSFlexGrid fgReport 
         Height          =   3645
         Index           =   0
         Left            =   -14865
         TabIndex        =   4
         Top             =   330
         Width           =   13920
         _cx             =   24553
         _cy             =   6429
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
      Begin VSFlex7LCtl.VSFlexGrid fgReport 
         Height          =   3645
         Index           =   1
         Left            =   -14565
         TabIndex        =   5
         Top             =   330
         Width           =   13920
         _cx             =   24553
         _cy             =   6429
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
      Begin VSFlex7LCtl.VSFlexGrid fgReport 
         Height          =   3645
         Index           =   2
         Left            =   45
         TabIndex        =   6
         Top             =   330
         Width           =   13920
         _cx             =   24553
         _cy             =   6429
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
      Begin VSFlex7LCtl.VSFlexGrid fgReport 
         Height          =   3645
         Index           =   3
         Left            =   14655
         TabIndex        =   7
         Top             =   330
         Width           =   13920
         _cx             =   24553
         _cy             =   6429
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
      Begin VSFlex7LCtl.VSFlexGrid fgReport 
         Height          =   3645
         Index           =   4
         Left            =   14955
         TabIndex        =   8
         Top             =   330
         Width           =   13920
         _cx             =   24553
         _cy             =   6429
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
      Begin VSFlex7LCtl.VSFlexGrid fgReport 
         Height          =   3645
         Index           =   5
         Left            =   15255
         TabIndex        =   9
         Top             =   330
         Width           =   13920
         _cx             =   24553
         _cy             =   6429
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
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
Attribute VB_Name = "frmSaiElite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kMaxSymbols As Long = 500
Private Const kExtraSpace As Long = -60

Private Enum eVSCurrTab
    eVSTab_FA = 0 ' Futures Aggressive
    eVSTab_FC ' Futures Conservative
    eVSTab_XA ' Forex Aggressive
    eVSTab_XC ' Forex Conservative
    eVSTab_UA ' US stocks Aggressive
    eVSTab_IA ' Intl stocks Aggressive
    eVSTab_Old
End Enum

Private Type mPrivate
    bForexAllowed As Boolean
    bFuturesAllowed As Boolean
    bStocksAllowed As Boolean
    RowIDs As cGdArray
    dFontSize As Double
    LogoImage As IPictureDisp
    
    nSessionDate As Long
    SymbolIDs As cGdArray ' ID's flagged as negative are not shown in the report
    DefaultSymbolIDs As cGdArray
    
    bElite As Boolean
    'iReportType As Long ' 0 = SAI, 1 = Elite
    
    bRunning As Boolean
    bUnloading As Boolean
End Type
Private m As mPrivate

Private Sub cmdRunForDate_Click()
On Error GoTo ErrSection

    If m.nSessionDate <> gdReportDate.Value Or fgReport(vsTab.CurrTab).Rows <= 1 Then
        m.nSessionDate = gdReportDate.Value
        LoadGrid
    End If
    
    fraDate.Visible = False
    vsTab.Visible = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.cmdRunForDate_Click"
End Sub

#If 0 Then
Private Sub fg_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection

    Dim i&, j&, nMovedID&, nPriorID&, iMovedTo&

    ' symbol was moved
    nMovedID = GetSymbolID(fg.TextMatrix(0, Position))
    If nMovedID <> 0 Then
        If Position > fg.FixedCols Then
            nPriorID = GetSymbolID(fg.TextMatrix(0, Position - 1))
        End If
        ' find moved symbol
        For i = 0 To m.SymbolIDs.Size - 1
            If Abs(m.SymbolIDs(i)) = nMovedID Then
                m.SymbolIDs.Remove i
                ' find where to move to
                iMovedTo = 0 ' (default = move to the beginning)
                If nPriorID <> 0 Then
                    For j = 0 To m.SymbolIDs.Size - 1
                        If Abs(m.SymbolIDs(j)) = nPriorID Then
                            iMovedTo = j + 1
                            Exit For
                        End If
                    Next
                End If
                m.SymbolIDs.Add nMovedID, iMovedTo
                'frmTest.AddList GetSymbol(nMovedID) & vbTab & Str(iMovedTo)
                Exit For
            End If
        Next
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.fg_AfterMoveColumn"
End Sub
#End If

Private Sub fgReport_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    GridTooltip fgReport(Index)

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection

    Dim strText$, i&
    
    g.Styler.StyleForm Me

    For i = 0 To vsTab.NumTabs - 1
        fgReport(i).Rows = 1
    Next

    vsTab.CurrTab = 0
    vsTab.BoldCurrent = True
    vsTab.FrontTabColor = 14342874              'light gray
    vsTab.FrontTabForeColor = 10485760          'dark blue

    strText = GetIniFileProperty("SAIE_Report", "", "Placement", g.strIniFile)
    If Len(strText) > 0 Then
        SetFormPlacement Me, strText, "LTHW"
    Else
        CenterTheForm Me
    End If

    Me.Icon = Picture16(ToolbarIcon("kSAI"), , True)
    With tbToolbar
        .Tools("ID_Symbols").Picture = Picture16("kChangeSymbol")
        .Tools("ID_SelectDate").Picture = Picture16("kCalendar")
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        '.Tools("ID_ZoomIn").Picture = Picture16("kTextIncrease")
        '.Tools("ID_ZoomOut").Picture = Picture16("kTextDecrease")
        .Tools("ID_Close").Picture = Picture16("kCancel")
        
        .Tools("ID_Print").Visible = False
        .Tools("ID_UserGuide").Visible = False
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    Dim s$
    s = m.SymbolIDs.JoinFields(",")
    SetIniFileProperty "SymbolIDs", s, "SAIE", g.strIniFile
    SetIniFileProperty "FontSize", m.dFontSize, "SAIE", g.strIniFile
    SetIniFileProperty "CurrTab", vsTab.CurrTab, "SAIE", g.strIniFile
    SetIniFileProperty "SAIE_Report", GetFormPlacement(Me), "Placement", g.strIniFile
    
    ' if running, give chance for running routine to exit
    m.bUnloading = True
    If m.bRunning And UnloadMode = vbFormControlMenu Then
        Me.Hide
        Cancel = True
    End If

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim t&, l&
    
    With vsTab
        l = 0
        t = .Top
        .Move l, t, Me.ScaleWidth - l * 2, Me.ScaleHeight - t - l
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    Set m.RowIDs = Nothing
    Set m.SymbolIDs = Nothing
    Set m.DefaultSymbolIDs = Nothing

End Sub

Public Sub ShowMe()
On Error GoTo ErrSection

    Dim i&

    m.bRunning = False
    m.bUnloading = False

    ' see what's allowed based on enablements
    m.bForexAllowed = HasModule("SAIE_AX,SAIE_CX")
    m.bFuturesAllowed = HasModule("SAIE_AF,SAIE_CF")
    m.bStocksAllowed = HasModule("SAIE_AS,SAIE_CS")
    If Not m.bForexAllowed And Not m.bFuturesAllowed And Not m.bStocksAllowed Then Exit Sub
    
    ' and show tabs based on enablements
    vsTab.TabVisible(eVSTab_FA) = HasModule("SAIE_AF")
    vsTab.TabVisible(eVSTab_FC) = HasModule("SAIE_CF")
    vsTab.TabVisible(eVSTab_XA) = HasModule("SAIE_AX")
    vsTab.TabVisible(eVSTab_XC) = HasModule("SAIE_CX")
    vsTab.TabVisible(eVSTab_UA) = HasModule("SAIE_AS")
    vsTab.TabVisible(eVSTab_IA) = HasModule("SAIE_AS")

    m.dFontSize = GetIniFileProperty("FontSize", 0, "SAIE", g.strIniFile)
    If m.dFontSize < 3 Then m.dFontSize = 9
    
    ' get default date
    m.nSessionDate = 0
    ShowDate
    fraDate.Visible = False
    
m.bElite = True

    ' default to last shown tab
    vsTab.FirstTab = 0
    i = GetIniFileProperty("CurrTab", 0, "SAIE", g.strIniFile)
    If i < 0 Or i >= vsTab.NumTabs Then
        i = 0
    End If
    If vsTab.TabVisible(i) Then
        vsTab.CurrTab = i
    Else
        ' but if last shown tab is not visible (i.e. no longer enabled), then just find first visible tab
        For i = 0 To vsTab.NumTabs - 1
            If vsTab.TabVisible(i) Then
                vsTab.CurrTab = i
                Exit For
            End If
        Next
    End If
    
    ' show form and load grid using default date
    GetSymbols
    InitGrid
    ShowForm Me, eForm_Nonmodal, frmMain
    LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.ShowMe"
End Sub

Private Sub ShowDate()
On Error GoTo ErrSection
    
    Dim nMaxDate&, dGMT#

    ' by default, set MaxDate to next business day after last daily download
    nMaxDate = LastDailyDownload + 1
    Do While Not IsWeekday(nMaxDate)
        nMaxDate = nMaxDate + 1
    Loop
    If m.nSessionDate = 0 Then
        m.nSessionDate = nMaxDate
    End If
    
    ' but if after 10pm GMT on that date, then allow going forward one more business day
    dGMT = ConvertTimeZone(Now, "", "GMT")
    If dGMT > nMaxDate + 22# / 24# Then
        nMaxDate = nMaxDate + 1
        Do While Not IsWeekday(nMaxDate)
            nMaxDate = nMaxDate + 1
        Loop
    End If
    gdReportDate.MaxDate = nMaxDate
    If m.nSessionDate <= nMaxDate Then
        gdReportDate.Value = m.nSessionDate
    Else
        gdReportDate.Value = nMaxDate
    End If
    fraDate.Visible = True
    vsTab.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.ShowDate"
End Sub

' returns True if this symbol is allowed based on the SAI enablements
Public Function SymbolAllowed(ByVal vSymbol As Variant) As Boolean
On Error GoTo ErrSection

    Dim strSymbol$
    
    strSymbol = GetSymbol(vSymbol)
    If Len(strSymbol) > 0 Then
        Select Case SecurityType(strSymbol)
        Case "F"
            SymbolAllowed = m.bFuturesAllowed
        Case "S"
            SymbolAllowed = m.bStocksAllowed
        Case Else
            If IsForex(strSymbol) Then
                SymbolAllowed = m.bForexAllowed
            ElseIf m.bForexAllowed Or m.bFuturesAllowed Or m.bStocksAllowed Then
                SymbolAllowed = True ' other indices are allowed if any other type is allowed
            End If
        End Select
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiElite.SymbolAllowed"
End Function

Private Sub GetSymbols()
On Error GoTo ErrSection

    Dim i&, nID&, s$
    Dim aSymbols As New cGdArray

    ' setup list of default symbols (based on allowed security types)
    Set m.DefaultSymbolIDs = New cGdArray
    m.DefaultSymbolIDs.Create eGDARRAY_Longs, 0, 0
    's = "ES-067,ZB-067,GC3-067,SI3-067,PL3-067,HG3-067,CL3-067,$AUD-USD,$EUR-JPY,$USD-CAD,$CHF-JPY,$EUR-USD,$GBP-JPY,$GBP-USD,$NZD-USD,$USD-JPY,$USD-CHF,$AUD-JPY,$EUR-GBP,$GBP-CHF,$GBP-CAD,$EUR-AUD,$CAD-JPY,$GBP-AUD,$AUD-CAD,MSFT,JPM"
    
    ' Aussie stocks
    s = "AMP,ANZ,BHP,BXB,CBA,CSL,MQG,NAB,NCM,ORG,QBE,RIO,STO,SUN,TLS,WBC,WDC,WES,WOW,WPL,"
    s = Replace(s, ",", "@ASX,")
    ' US stocks
    s = s & "AAPL,AIG,ANZN,BAC,BIDU,C,CSCO,GE,GILD,GOOGL,INTC,JPM,MA,MSFT,ORCL,QCOM,NFLX,"
    ' Futures
    s = s & "GC3-,QO-,SI3-,HG3-,PA3-,PL3-,ES-,NQ-,YM-,NK3-,WP-,TF-,CL3-,HO3-,RB3-,NG3-,JO-,HE-,GF-,LE-,ZC-,ZS-,ZM-,ZL-,ZW-,ZO-,ZR-,CT-,SB-,CC-,KC-,G6A-,G6B-,G6C-,G6E-,G6J-,G6N-,G6S-,G6M-,DX-,GE-,ZB-,ZN-,ZF-,EBM-,EBI-,GX-,LL-,"
    s = Replace(s, "-", "-067")
    ' Forex
    s = s & "$EUR-USD,$GBP-USD,$USD-CHF,$GBP-CHF,$AUD-USD,$CHF-JPY,$GBP-JPY,$AUD-JPY,$CAD-JPY,$EUR-GBP,$EUR-AUD,$NZD-JPY,$AUD-CHF,$GBP-AUD,$EUR-NZD,$AUD-CAD,$USD-CAD,$EUR-JPY,$USD-JPY,$NZD-USD,$AUD-NZD,$EUR-CAD,$USD-SGD,$EUR-CHF,"
    ' put in array and sort
    aSymbols.SplitFields s, ","
    aSymbols.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    For i = 0 To aSymbols.Size - 1
        s = aSymbols(i)
        'If SymbolAllowed(s) Then
            nID = GetSymbolID(s)
            If nID > 0 Then
                m.DefaultSymbolIDs.Add nID
            End If
        'End If
    Next
    
    ' get user's list of symbols
    Set m.SymbolIDs = New cGdArray
    m.SymbolIDs.Create eGDARRAY_Longs, 0, 0
    s = GetIniFileProperty("SymbolIDs", "", "SAIE", g.strIniFile)
    aSymbols.SplitFields s, ","
    For i = 0 To aSymbols.Size - 1
        nID = Val(aSymbols(i))
        If SymbolAllowed(Abs(nID)) Then
            m.SymbolIDs.Add nID
        End If
    Next
    
    ' if empty list, set to defaults
    If m.SymbolIDs.Size = 0 Then
        Set m.SymbolIDs = m.DefaultSymbolIDs.MakeCopy
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.GetSymbols"
End Sub

' to format list of symbols for Selection dialog
Private Function FormattedSymbolList(SymbolIDs As cGdArray) As cGdArray
On Error GoTo ErrSection

    Dim i&, nID&, strSymbol$, strDesc$
    Dim bActive As Boolean
    Dim aSymbols As New cGdArray
    
    aSymbols.Create eGDARRAY_Strings
    For i = 0 To SymbolIDs.Size - 1
        nID = SymbolIDs(i)
        bActive = (nID > 0) ' ID flagged as negative means to not show in the report
        nID = Abs(nID)
        strSymbol = GetSymbol(nID)
        If SymbolAllowed(strSymbol) Then
            strDesc = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbolID(nID))
            aSymbols.Add "1" & vbTab & strSymbol & vbTab & " " & strDesc & vbTab & Str(Abs(bActive))
        End If
    Next
    
    Set FormattedSymbolList = aSymbols

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiElite.FormattedSymbolList"
End Function

' to allow user to select and arrange the symbols (columns)
Private Sub ManageSymbols()
On Error GoTo ErrSection

    Dim i&, s$, nID&, strSymbol$, nActive&
    Dim aSymbols As cGdArray, aDefaultSymbols As cGdArray
    
    ' build the string arrays to hand off to dialog
    Set aSymbols = FormattedSymbolList(m.SymbolIDs)
    Set aDefaultSymbols = FormattedSymbolList(m.DefaultSymbolIDs)
       
    ' portfolio dialog
    If frmQuoteBoardFields.ShowMe(aSymbols, eQbfMode_SaiElite, aDefaultSymbols) Then
        aSymbols.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
        nActive = 0
        m.SymbolIDs.Size = 0
        For i = 0 To aSymbols.Size - 1
            nID = GetSymbolID(Parse(aSymbols(i), vbTab, 2))
            If nID > 0 And nActive < kMaxSymbols Then
                If Val(Parse(aSymbols(i), vbTab, 4)) = 0 Or nActive >= kMaxSymbols Then
                    m.SymbolIDs.Add -nID ' "inactive" = not shown in report
                Else
                    m.SymbolIDs.Add nID ' "active" = show in report
                    nActive = nActive + 1
                End If
            End If
        Next
        
        LoadGrid
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.ManageSymbols"
End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
    
    Dim s$
    
    Select Case Tool.ID
    Case "ID_Symbols"
        ManageSymbols
        
    Case "ID_SelectDate"
        If 0 Then
            s = InfBox("Display the SAI Report for:", "?", , "SAI Report Date", , , , , , "d", DateFormat(m.nSessionDate))
            If Len(s) > 0 Then
                If DateOf(s) <> m.nSessionDate Then
                    m.nSessionDate = DateOf(s)
                    LoadGrid
                End If
            End If
        Else
            ShowDate
        End If
            
    Case "ID_Print"
        fraDate.Visible = False
        vsTab.Visible = True
        PrintMe
        
    Case "ID_ZoomIn"
        ChangeFontSize 1
    
    Case "ID_ZoomOut"
        ChangeFontSize -1
        
    Case "ID_Disclaimer"
        s = "@" & App.Path & "\Info\SAI_Disclaimer.rtf"
        frmMessage.ShowMe "Strategic Analysis Indicator", s ', eModalMessage
                    
    Case "ID_Close"
        Unload Me
        
    Case "ID_UserGuide"
        s = App.Path & "\Info\SAI-manual.pdf"
        If FileExist(s) Then
            RunProcess s
        Else
            InfBox "User Guide not found", "e", , "Strategic Analysis Indicator"
        End If
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.tbToolbar_ToolClick"
    Resume ErrExit
End Sub

Public Sub GridTextIncrease()
    ChangeFontSize 1
End Sub

Public Sub GridTextDecrease()
    ChangeFontSize -1
End Sub

Private Sub ChangeFontSize(ByVal dUpDown#)

    Dim i&, iRow&, iCol&, dSize#, dRowHeight#

    If Not IsIDE Then
        On Error Resume Next
    End If
    
    m.dFontSize = m.dFontSize + dUpDown
    If m.dFontSize < 4 Then m.dFontSize = 4
    
    For i = 0 To vsTab.NumTabs - 1
        fgReport(i).Font.Size = m.dFontSize
    Next

#If 0 Then
    With fg
        If .Visible Then
            m.dFontSize = m.dFontSize + dUpDown
            If m.dFontSize < 4 Then m.dFontSize = 4
            .Font.Size = m.dFontSize
            dRowHeight = .RowHeight(.Rows - 1) ' height of typical row
    
            ' since some rows have custom font-stuff (e.g. bolding),
            ' we have to reset font size of each cell
            For iRow = 0 To .Rows - 1
                Select Case UCase(Parse(.TextMatrix(iRow, 0), " ", 1))
                Case "DAILY", "MONTHLY", "WEEKLY"
                    dSize = m.dFontSize + 1
                    .RowHeight(iRow) = dRowHeight * 1.25
                Case Else
                    dSize = m.dFontSize
                End Select
                .Cell(flexcpFontSize, iRow, 0, iRow, 0) = dSize
                .Cell(flexcpFontSize, iRow, 1, iRow, .Cols - 1) = m.dFontSize
            Next
            .AutoSize 0, .Cols - 1, , kExtraSpace
        End If
    End With
#End If

End Sub


' returns the bar# for the last completed data bar prior to the session date
Private Function GetBarNumberAndFixClose(Bars As cGdBars) As Long
On Error GoTo ErrSection
    
    Dim nBar&, nDate&, nBarEndDate&, strSymbol$, dTime#, i&, dClose#, dStartTC#
    Dim MinuteBars As New cGdBars
    
    If Bars.Size = 0 Or m.nSessionDate <= 0 Then
        nBar = -1
    Else
        ' Get bar# of data completed prior to this session date
        nBar = Bars.FindDateTime(m.nSessionDate) - 1
        Do While nBar >= 0
            If Bars(eBARS_Close, nBar) <> kNullData Then
                Exit Do
            End If
            nBar = nBar - 1
        Loop
        
' TLB: was testing this per Gary, but probably won't need it now?
' (hopefully not, since this would make the report inconsistent with the SAI chart indicators)
#If False Then
        ' for Forex: fix the closing price (set to price at 5pm NY yearround)
        strSymbol = Bars.Prop(eBARS_Symbol)
        If IsForex(strSymbol) Then
            nBarEndDate = Bars(eBARS_DateTime, nBar)
            If nBarEndDate > 0 Then
                dStartTC = gdTickCount
                ' go back until we have a day with data (e.g. if starting from monthly bars)
                For nDate = nBarEndDate To nBarEndDate - 6 Step -1
                    If IsWeekday(nDate) Then
                        DM_GetBars MinuteBars, strSymbol, "60 minute", nDate, nDate
                        If MinuteBars.Size > 0 Then
                            ' now go back until we get the hourly bar at 5pm NY
                            For i = MinuteBars.Size - 1 To 0 Step -1
                                ' convert time to NY and round to the nearest hour
                                dTime = MinuteBars.DateTimeConvert(i, "NY")
                                If Hour(dTime + 0.5 / 24) <= 17 Then
                                    dClose = MinuteBars(eBARS_Close, i)
                                    If dClose > 0 Then
                                        ' replace the Close of the Bars with this price
                                        Bars(eBARS_Close, nBar) = dClose
                                    End If
                                    Exit For
                                End If
                            Next
                            Exit For
                        End If
                    End If
                Next
                If IsIDE Then
                    'frmTest.AddList strSymbol & vbTab & Format(gdTickCount - dStartTC, "####0")
                End If
            End If
        End If
#End If
    End If
    
    GetBarNumberAndFixClose = nBar
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiElite.GetBarNumberAndFixClose"
End Function

' Pass in 0 for BS/SS, or 1-4 for Profit Level
Private Function CalcPL(Bars As cGdBars, ByVal nBar&, ByVal iProfitLevel&, ByVal bSell As Boolean) As Double
On Error GoTo ErrSection
    
    Dim dMult#, dRange#, dClose#
    
    dClose = Bars(eBARS_Close, nBar)
    dRange = Bars(eBARS_High, nBar) - Bars(eBARS_Low, nBar)
    
    If iProfitLevel < 0 Or iProfitLevel > 4 Or dClose = kNullData Or dRange < 0 Then
        CalcPL = 0
    Else
        ' Get the multiplier for this profit level
        'BS/SS = 0.073, PL1 = 0.309, PL2 = 0.545, PL3 = 0.691, PL4 = 0.927
        Select Case iProfitLevel
        Case 0
            dMult = 0.073
        Case 1
            dMult = 0.309
        Case 2
            dMult = 0.545
        Case 3
            dMult = 0.691
        Case 4
            dMult = 0.927
        Case Else
            dMult = 0 ' undefined
        End Select
        If bSell Then
            dMult = -dMult
        End If
        
        CalcPL = dClose + dRange * dMult
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiElite.CalcPL"
End Function

Private Sub InitGrid()
On Error GoTo ErrSection

    Dim s$, strEach$, iRow&, iType&, dRowHeight#, iTab&
    Dim fg As VSFlexGrid
    
    ' init the Row ID's
    'strEach = "RISK ; BUY PL4 ; BUY PL3 ; BUY PL2 ; BUY PL1 ; BUY STOP ; SELL STOP ; SELL PL1 ; SELL PL2; SELL PL3 ; SELL PL4 ; PPC ;"
    strEach = " BUY PL4 ; BUY PL3 ; BUY PL2 ; BUY PL1 ; BUY STOP ; RISK ; SELL STOP ; SELL PL1 ; SELL PL2; SELL PL3 ; SELL PL4 ; PPC ;"
    s = " ; DAILY ;" & strEach & " MTR ; IND ; WEEKLY ;" & strEach & " OTE (weekly) ; MONTHLY ;" & strEach & " OTE (monthly) ;"
    Set m.RowIDs = New cGdArray
    m.RowIDs.SplitFields s, ";"

    If m.bElite Then
        For iTab = 0 To vsTab.NumTabs - 1
            Set fg = fgReport(iTab)
            With fg
                SetupGrid fg, eGridMode_Grid
                .Font.Size = m.dFontSize
                .SelectionMode = flexSelectionFree
                .AllowSelection = True
                .ExtendLastCol = False
                .ExplorerBar = flexExMove
                '.BackColorFixed = RGB(255, 244, 216)
                '.BackColorFixed = RGB(224, 224, 224)
                .BackColorFixed = &HDDFAF9
                .ExplorerBar = flexExSortShow
                .Cols = 9
                .Rows = 1
                '.FixedCols = 1
                .FixedRows = 1
                .TextMatrix(0, 0) = "Symbol"
                .TextMatrix(0, 1) = "Position"
                .TextMatrix(0, 2) = "Entry Date"
                .TextMatrix(0, 3) = "Entry Price"
                .TextMatrix(0, 4) = "Current Price"
                .TextMatrix(0, 5) = "Net Points"
                .TextMatrix(0, 6) = "Open Equity"
                .TextMatrix(0, 7) = "Current Orders"
                .TextMatrix(0, 8) = "If filled, next order ..."
                .ColHidden(5) = True
                
                .ColAlignment(1) = flexAlignCenterTop
                .ColAlignment(2) = flexAlignCenterTop
                .ColAlignment(3) = flexAlignCenterTop
                .ColAlignment(4) = flexAlignCenterTop
                .ColAlignment(5) = flexAlignCenterTop
                .ColAlignment(6) = flexAlignCenterTop
                .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
                
                .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
                .AutoSize 0, .Cols - 1, , 180
            End With
        Next
    Else
        Set fg = fgReport(0)
        With fg
            SetupGrid fg, eGridMode_Grid
            .Font.Size = m.dFontSize
            .SelectionMode = flexSelectionFree
            .AllowSelection = True
            .ExtendLastCol = False
            .ExplorerBar = flexExMove
            '.BackColorFixed = RGB(255, 244, 216)
            '.BackColorFixed = RGB(224, 224, 224)
            .BackColorFixed = &HDDFAF9
            .ExplorerBar = flexExMove
            .Cols = 1
            .FixedCols = 1
            .Rows = m.RowIDs.Size
            dRowHeight = .RowHeight(.Rows - 1) ' height of typical row (e.g. last row)
            For iRow = 0 To m.RowIDs.Size - 1
                s = Trim(m.RowIDs(iRow))
                m.RowIDs(iRow) = s
                
                Select Case UCase(Parse(s, " ", 1))
                Case "DAILY", "MONTHLY", "WEEKLY"
                    .Cell(flexcpForeColor, iRow, 0) = vbWhite
                    .Cell(flexcpBackColor, iRow, 0) = RGB(1, 1, 1) ' so will be black
                    .Cell(flexcpFontSize, iRow, 0) = m.dFontSize + 1
                    .RowHeight(iRow) = dRowHeight * 1.25
                Case "BUY"
                    .Cell(flexcpForeColor, iRow, 0) = vbBlue
                Case "SELL"
                    .Cell(flexcpForeColor, iRow, 0) = vbRed
                Case "OTE"
                    s = "OTE"
                    .Cell(flexcpBackColor, iRow, 0) = RGB(144, 192, 240)
                End Select
                
                Select Case UCase(Parse(s, " ", 2))
                Case "PL1"
                    s = "Profit Level 1"
                Case "PL2"
                    s = "Profit Level 2"
                Case "PL3"
                    s = "Profit Level 3"
                Case "PL4"
                    s = "Profit Level 4"
                End Select
                            
                .TextMatrix(iRow, 0) = s
            Next
            .ColAlignment(0) = flexAlignCenterCenter
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
            .AutoSize 0, 0
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.InitGrid"
End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection

    Dim i&, iSymbol&, iCol&, iRow&, iLevel&, nBar&, dValue#, dMTR#, iCount&, dDailyClose#, iIndDir&, dBS#, dSS#
    Dim s$, strSymbol$
    Dim bBold As Boolean, bSectionHdr As Boolean
    Dim Bars As cGdBars, Daily As New cGdBars, Weekly As New cGdBars, Monthly As New cGdBars
    Dim fg As VSFlexGrid
              
    'm.nSessionDate = gdForDate.Value
    If m.nSessionDate <= 0 Or m.nSessionDate > LastDailyDownload + 4 Then
        m.nSessionDate = LastDailyDownload + 1
    End If
    Do While Not IsWeekday(m.nSessionDate)
        m.nSessionDate = m.nSessionDate + 1
    Loop
    
    Me.Caption = "Strategic Analysis Indicator for " & DateFormat(m.nSessionDate)
    
    fraDate.Visible = False
    vsTab.Visible = True
    'fg.Redraw = flexRDBuffered
    
    If m.bElite Then
        RunElite
    Else
        Set fg = fgReport(0)
        iCol = fg.FixedCols
        For iSymbol = 0 To m.SymbolIDs.Size - 1
            ' ID flagged as negative means don't show symbol in report
            If m.SymbolIDs(iSymbol) > 0 Then
                strSymbol = GetSymbol(m.SymbolIDs(iSymbol))
                If Not SymbolAllowed(strSymbol) Then
                    strSymbol = ""
                End If
            Else
                strSymbol = ""
            End If
            If Len(strSymbol) > 0 Then
                ' load daily/weekly/monthly bars (but load enough to get a 55-bar moving average)
                Set Daily = New cGdBars
                Set Weekly = New cGdBars
                Set Monthly = New cGdBars
                DM_GetBars Daily, strSymbol, 0, m.nSessionDate - 88, m.nSessionDate + 6
                If g.RealTime.Active And m.nSessionDate > LastDailyDownload Then
                    g.RealTime.SpliceBars Daily
                End If
                Daily.AddForecastBars 1
                Weekly.BuildBars "Weekly", Daily.BarsHandle
                Monthly.BuildBars "Monthly", Daily.BarsHandle
                
                ' calc the MTR from a 55-bar MA
                Set Bars = Daily
                nBar = GetBarNumberAndFixClose(Bars)
                iIndDir = 0
                iCount = 0
                dMTR = 0
                For i = nBar To 0 Step -1
                    dValue = Bars(eBARS_Close, i)
                    If dValue <> kNullData Then
                        dMTR = dMTR + dValue
                        iCount = iCount + 1
                        If iCount >= 55 Then
                            Exit For
                        End If
                    End If
                Next
                dDailyClose = Bars(eBARS_Close, nBar)
                If iCount >= 55 Then
                    dMTR = dMTR / iCount
                    If dMTR > dDailyClose Then
                        iIndDir = 1 ' IND = Buy
                    ElseIf dMTR < dDailyClose Then
                        iIndDir = -1 ' IND = Sell
                    End If
                Else
                    dMTR = kNullData
                End If
                
                With fg
                    .Cols = iCol + 1
                    .TextMatrix(0, iCol) = strSymbol
                    .ColAlignment(iCol) = flexAlignCenterCenter
                    
                    For iRow = .FixedRows To m.RowIDs.Size - 1
                        bBold = False
                        bSectionHdr = False
                        dValue = kNullData
                        .Cell(flexcpForeColor, iRow, iCol) = .Cell(flexcpForeColor, iRow, 0) ' default
                        
                        s = m.RowIDs(iRow)
                        iLevel = Val(Right(s, 1))
                        Select Case UCase(Parse(s, " ", 1))
                        Case "DAILY"
                            bSectionHdr = True
                            dValue = Bars(eBARS_DateTime, nBar + 1)
                            .TextMatrix(iRow, iCol) = DateFormat(dValue)
                        Case "WEEKLY"
                            bSectionHdr = True
                            Set Bars = Weekly
                            nBar = GetBarNumberAndFixClose(Bars)
                            dValue = Bars(eBARS_DateTime, nBar + 1)
                            ' backup to the previous Monday
                            For i = Int(dValue) To 1 Step -1
                                If Not IsWeekday(i - 1) Then
                                    .TextMatrix(iRow, iCol) = DateFormat(i, M_D) & "-" & DateFormat(dValue, M_D)
                                    Exit For
                                End If
                            Next
                        Case "MONTHLY"
                            bSectionHdr = True
                            Set Bars = Monthly
                            nBar = GetBarNumberAndFixClose(Bars)
                            dValue = Bars(eBARS_DateTime, nBar + 1)
                            .TextMatrix(iRow, iCol) = DateFormat(dValue, MMM_YY)
                            
                        Case "BUY"
                            dValue = CalcPL(Bars, nBar, iLevel, False)
                            If iLevel = 0 Then bBold = True
                        Case "SELL"
                            dValue = CalcPL(Bars, nBar, iLevel, True)
                            If iLevel = 0 Then bBold = True
                        Case "RISK"
                            dValue = CalcPL(Bars, nBar, iLevel, False) - CalcPL(Bars, nBar, iLevel, True)
                            bBold = True
                        Case "PPC"
                            dValue = Bars(eBARS_Close, nBar)
                        Case "MTR"
                            dValue = dMTR
                        Case "IND"
                            If iIndDir > 0 Then
                                .TextMatrix(iRow, iCol) = "Buy"
                                .Cell(flexcpForeColor, iRow, iCol) = vbBlue
                            ElseIf iIndDir < 0 Then
                                .TextMatrix(iRow, iCol) = "Sell"
                                .Cell(flexcpForeColor, iRow, iCol) = vbRed
                            Else
                                .TextMatrix(iRow, iCol) = ""
                            End If
                        Case "OTE"
                            ' TLB: per our understanding of the formula given to us by Gary ...
                            dBS = CalcPL(Bars, nBar, 0, False) ' Monthly/Weekly Buy Stop
                            dSS = CalcPL(Bars, nBar, 0, True) ' Monthly/Weekly Sell Stop
                            If dDailyClose > dBS Then
                                dValue = dDailyClose - dBS
                                bBold = True
                            ElseIf dDailyClose < dSS Then
                                dValue = dSS - dDailyClose
                                bBold = True
                            Else
                                ' undefined when between BS and SS?
                                dValue = kNullData
                                .TextMatrix(iRow, iCol) = ""
                            End If
                            .Cell(flexcpBackColor, iRow, iCol) = .Cell(flexcpBackColor, iRow, 0)
                        End Select
                        
                        If bSectionHdr Then
                            ' new timeframe section
                            .Cell(flexcpBackColor, iRow, iCol) = .Cell(flexcpBackColor, iRow, 0)
                        ElseIf dValue <> kNullData Then
                            ' display as price
                            .TextMatrix(iRow, iCol) = Bars.PriceDisplay(dValue)
                        End If
                        .Cell(flexcpFontBold, iRow, iCol) = bBold
                    Next
                End With
                iCol = iCol + 1
            End If
        Next
        
        With fg
            .Cols = iCol
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
            .AutoSize 0, .Cols - 1, , kExtraSpace
            If .Cols > .FixedCols Then
                .ShowCell .FixedRows, .FixedCols
            End If
            .Select 0, 0
        End With
    End If
    
    Set Bars = Nothing
    Set Daily = Nothing
    Set Weekly = Nothing
    Set Monthly = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.LoadGrid"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Send the grid to the Print Preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    If m.dFontSize < 7 Then
        PrintMe = frmPrintPreview.ShowMe("CNV SaiElite", frmSaiElite, , 1.6, 0.5, 0.4, 0.3, True, , , True)
    Else
        PrintMe = frmPrintPreview.ShowMe("CNV SaiElite", frmSaiElite, , 1.6, 0.5, 0.4, 0.3, False, , , True)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSaiElite.PrintMe"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Arguments into the Print Preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    Dim lFirstCol&, lLastCol&
    Dim strText As String
    
#If 0 Then
    ' see if user has selected multiple columns to print (i.e. if not all columns)
    With fg
        If .ColSel > .Col Then
            lFirstCol = .Col
            lLastCol = .ColSel
        ElseIf .ColSel < .Col Then
            lFirstCol = .ColSel
            lLastCol = .Col
        Else ' print all columns
            lFirstCol = 0
            lLastCol = 0
        End If
        If lLastCol > lFirstCol Then
            ' hide the non-selected columns
            For lCol = .FixedCols To .Cols - 1
                If lCol >= lFirstCol And lCol <= lLastCol Then
                    .ColHidden(lCol) = False
                Else
                    .ColHidden(lCol) = True
                End If
            Next
        End If
    End With
    
    strText = App.Path & "\Info\SAI.jpg"
    If FileExist(strText) Then
        Set m.LogoImage = LoadPicture(strText)
    Else
        Set m.LogoImage = Nothing
    End If
    
    With frmPrintPreview.vp
        .Clear
        .StartDoc
        If 0 Then
            DoPrintHeader 8
        Else
            .LineSpacing = 100
            .HdrFontName = fg.Font.Name ' "Times New Roman"
            .HdrFontSize = 10
            strText = "|Trade Navigator" & vbCrLf & "Genesis Financial Technologies - "
            .Header = " "
            '.Header = strText & GetProvidedProperty("Website", , True)
            .Footer = "  Powered by Genesis Financial Technologies - TradeNavigator.com||Page: %d    "
        End If
        
    If 0 Then
        .TextAlign = taCenterMiddle
        .Font.Name = fg.Font.Name ' "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        '.FontUnderline = True
        .Text = "Strategic Analysis Indicator Report for " & DateFormat(m.nSessionDate) & vbLf
        .Font.Size = 12
        .FontUnderline = False
        .Font.Bold = False
        .TextAlign = taLeftMiddle
        .Text = vbLf
    End If
              
        
        'fg.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fg
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fg.hWnd
        End If
        'fg.ExtendLastCol = True
        
        strText = FileToString(App.Path & "\Info\SAI_Disclaimer.rtf")
        If Len(strText) > 10 Then
            .NewPage
            .Text = vbLf
            .TextRTF = strText
        End If
        
        .EndDoc
    End With

    If lLastCol > lFirstCol Then
        ' show all columns again
        For lCol = fg.FixedCols To fg.Cols - 1
            fg.ColHidden(lCol) = False
        Next
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.GenerateReport", eGDRaiseError_Raise
End Sub

Public Sub AfterHeaderEvent(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim strText As String
    
    With frmPrintPreview.vp
        If Not frmPrintPreview.GoingToFile Then
            If .CurrentPage = 1 Then
                strText = strText
            End If
            
            .CurrentY = "0.5in"
            .TextAlign = taLeftMiddle
            .Font.Name = fgReport(0).Font.Name ' "Times New Roman"
                        
            .Font.Bold = False
            .Font.Size = 8
            .Text = "Copyright by Strategic Analysis -- Available by Subscription Only -- NOT For Distribution" & vbLf '& vbLf & vbLf
            
            .Font.Size = 14
            .Font.Bold = True
            .Text = vbLf & "Strategic Analysis Indicator for " & DateFormat(m.nSessionDate) & vbLf
            .Font.Size = 12
            .Font.Bold = False
            .Text = "With 4 pre-defined profit levels!" & vbLf
            '.Font.Size = 8
            '.Text = vbLf & "Copyright by Strategic Analysis (www.strategic-analysis.biz) -- NOT for Distribution" & vbLf
            
            If Not m.LogoImage Is Nothing Then
                .DrawPicture m.LogoImage, "6in", "0.4in", "1in", "1in", vppaZoom ' vppaRightTop
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.AfterHeaderEvent", eGDRaiseError_Raise
End Sub

Private Sub RunElite()
On Error GoTo ErrSection:

    Dim i&, iTab&, iRow&, iSymbol&, nSystemID&, iToDate&, nEntryDate&
    Dim iRule&, iParm&, nNumTicks&, nCount&
    Dim dTime#, dEntryPrice#, dOrderPrice#, dLastPrice#, dNet#, dProfit#, dSum#
    Dim s$, strSymbol$, strPos$, strOrder$, strSecType$, strOrder2$
    Dim Strategy As New cSystem
    Dim Parm As cInput
    Dim Bars As New cGdBars
    Dim aFile As New cGdArray
    Dim aFields As New cGdArray
    Dim fg As VSFlexGrid
    
    
    iTab = vsTab.CurrTab
    If iTab < 0 Then Exit Sub
    Set fg = fgReport(iTab)
        
    s = ""
    Select Case iTab
    Case eVSTab_FA
        s = "SAI Aggressive"
        strSecType = "F"
    Case eVSTab_FC
        s = "SAI Conservative"
        strSecType = "F"
    Case eVSTab_XA
        s = "SAI Aggressive"
        strSecType = "X"
    Case eVSTab_XC
        s = "SAI Conservative"
        strSecType = "X"
    Case eVSTab_UA
        s = "SAI Aggressive"
        strSecType = "U" ' US stocks
    Case eVSTab_IA
        s = "SAI Aggressive"
        strSecType = "S"
    End Select
    If s = "" Then Exit Sub
    nSystemID = SystemIDForName(s)
    If nSystemID = 0 Then
        ' if the strategy was not found, then they likely do not yet have the proper library
        s = "SAI Elite library not found.|(you may need to download 'Upgrade'| and Re-import the libraries)"
        InfBox s, "e", , "ERROR"
        Exit Sub
    End If
     
    'iToDate = LastDailyDownload + 1
    iToDate = m.nSessionDate
    Do While Not IsWeekday(iToDate)
        iToDate = iToDate + 1
    Loop
    
    dTime = gdTickCount
    Strategy.Load nSystemID
    Strategy.ToEndOfData = False
    Strategy.ToDate = iToDate
    'Strategy.FromDate = iToDate - 250
    Strategy.FromDate = -350 ' to allow overriding the Load From date (in cSystem)
    
    s = DateFormat(iToDate, MM_DD_YYYY)
    Me.Caption = "SAI Elite Report for " & s
    If Mid(s, 4, 1) = "0" Then s = Left(s, 3) & Mid(s, 5)
    If Left(s, 1) = "0" Then s = Mid(s, 2)
    fg.Rows = 1
    fg.TextMatrix(0, fg.Cols - 2) = "Orders for " & s
    
    m.bRunning = True
    dSum = 0
    nCount = 0
    For iSymbol = 0 To m.SymbolIDs.Size - 1
        ' see if need to abort
        If m.bUnloading Then GoTo ErrSection
        
        strOrder = ""
        strOrder2 = ""
        nEntryDate = 0
        
        ' ID flagged as negative means don't show symbol in report
        If m.SymbolIDs(iSymbol) > 0 Then
            strSymbol = GetSymbol(m.SymbolIDs(iSymbol))
            If Not SymbolAllowed(strSymbol) Then
                strSymbol = ""
            Else
                Select Case SecurityType(strSymbol)
                Case "F"
                    If strSecType <> "F" Then strSymbol = ""
                Case "S"
                    If InStr(strSymbol, "@") = 0 Then
                        ' US stocks
                        If strSecType <> "U" Then strSymbol = ""
                    Else
                        ' Other stocks
                        If strSecType <> "S" Then strSymbol = ""
                    End If
                Case "I"
                    If IsForex(strSymbol) Then
                        If strSecType <> "X" Then strSymbol = ""
                    Else
                        If strSecType <> "S" Then strSymbol = ""
                    End If
                Case Else
                    strSymbol = ""
                End Select
            End If
        Else
            strSymbol = ""
        End If
        If Len(strSymbol) > 0 Then
        
#If 1 Then
            Set Bars = Nothing
            'Strategy.LoadBarsForMultRun 0, strSymbol, "Daily", Strategy.FromDate, Strategy.ToEndOfData, Strategy.ToDate
            Strategy.LoadMarket1Bars strSymbol, "Daily", True
            Set Bars = Strategy.Bars
#Else
            Strategy.FromDate = iToDate - 250
            DM_GetBars Bars, strSymbol, "Daily", Strategy.FromDate, Strategy.ToDate
            If Bars.Size > 0 Then
                'g.RealTime.SpliceBars Bars
                Bars.AddForecastBars 1
                Strategy.Bars = Bars
                Strategy.ChangeMarket1 Bars
            End If
#End If

            If Bars.Size > 0 Then
                ' set NumTicks for strategy rules
                nNumTicks = 5
                If 0 Then
                    For iRule = 1 To Strategy.Rules.Count
                        For iParm = 1 To Strategy.Rules.Item(iRule).Inputs.Count
                            Set Parm = Strategy.Rules.Item(iRule).Inputs.Item(iParm)
                            If Not Parm Is Nothing Then
                                If UCase(Parm.ParmName) = "NUMTICKS" Then
                                    Parm.Value = nNumTicks
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                End If
                
                Strategy.NextBarReport eGDNextBarMode_RunMult, iToDate
                s = App.Path & "\Trades\NB-" & Str(nSystemID) & "-" & Str(Bars.Prop(eBARS_SymbolID)) & ".txt"
                aFile.FromFile s
                If aFile.Size >= 6 Then
                    ' Get current position from 3rd line of NB file:
                    ' S   1   41493   15409   SAI Aggr Short Entry
                    aFields.SplitFields aFile(2), vbTab
                    strPos = aFields(0) ' most likely either "L" (long) or "S" (short)
                    nEntryDate = Val(aFields(2))
                    dEntryPrice = Val(aFields(3))
                    'If nEntryDate > 0 And dEntryPrice > 0 And Len(strPos) = 1 Then
                        
                        ' Get next bar order from 6th line of NB file:
                        '  1 EL  6234    S   15584               1
                        aFields.SplitFields aFile(5), vbTab
                        dOrderPrice = Val(aFields(4))
                        If dOrderPrice > 0 And aFields(3) = "S" Then ' should be a Stop order
                            Select Case aFields(1)
                            Case "EL", "XS"
                                strOrder = "Buy Stop at " & Bars.PriceDisplay(dOrderPrice)
                            Case "ES", "XL"
                                strOrder = "Sell Stop at " & Bars.PriceDisplay(dOrderPrice)
                            End Select
                        
                            ' Get the conditional order (if nbo gets filled) from 7th line of NB file:
                            '  2 XL  6234    S   15584               1
                            aFields.SplitFields aFile(6), vbTab
                            dOrderPrice = Val(aFields(4))
                            If dOrderPrice > 0 And aFields(3) = "S" Then ' should be a Stop order
                                Select Case aFields(1)
                                Case "EL", "XS"
                                    strOrder2 = "Buy Stop at " & Bars.PriceDisplay(dOrderPrice)
                                Case "ES", "XL"
                                    strOrder2 = "Sell Stop at " & Bars.PriceDisplay(dOrderPrice)
                                End Select
                            End If
                        End If
                    'End If
                End If
            End If
        End If
        
        ' add a new row if there is a position and/or order for this symbol
        If nEntryDate > 0 Or Len(strOrder) > 0 Then
            dNet = kNullData
            For i = Bars.Size - 1 To 0 Step -1
                dLastPrice = Bars(eBARS_Close, i)
                If dLastPrice <> kNullData Then
                    Exit For
                End If
            Next
            With fg
                iRow = .Rows
                .Rows = .Rows + 1
                If UCase(Left(strPos, 1)) = "S" Then
                    's = "Sold"
                    s = "Short"
                    dNet = dEntryPrice - dLastPrice
                    .Cell(flexcpForeColor, iRow, 1) = vbRed
                ElseIf UCase(Left(strPos, 1)) = "L" Then
                    's = "Bought"
                    s = "Long"
                    dNet = dLastPrice - dEntryPrice
                    .Cell(flexcpForeColor, iRow, 1) = vbBlue
                Else
                    s = "Flat"
                    dNet = kNullData
                    dEntryPrice = kNullData
                    .Cell(flexcpForeColor, iRow, 1) = vbBlack
                End If
                's = s & " on " & DateFormat(nEntryDate) & " at " & Bars.PriceDisplay(dEntryPrice)
                .TextMatrix(iRow, 0) = strSymbol
                .TextMatrix(iRow, 1) = s
                .TextMatrix(iRow, 2) = DateFormat(nEntryDate, MM_DD_YYYY)
                .TextMatrix(iRow, 3) = Bars.PriceDisplay(dEntryPrice)
                .TextMatrix(iRow, 4) = Bars.PriceDisplay(dLastPrice)
                .TextMatrix(iRow, 5) = Bars.PriceDisplay(dNet)
                s = ""
                If Bars.Prop(eBARS_TickMove) > 0 And dNet <> kNullData Then
                    'dProfit = g.Profit.Profit(strSymbol, dNet)
                    If Bars.SecurityType = "F" Then
                        dNet = dNet / Bars.TickMove * Bars.TickValue
                        s = Format(dNet, "$#,##0")
                    ElseIf IsForex(strSymbol) Then
                        dNet = dNet / Bars.TickMove
                        s = Str(Round(dNet, 0)) & " pips"
                    ElseIf dEntryPrice > 0 Then
                        dNet = dNet / dEntryPrice * 100
                        s = Format(dNet, "#0.000") & "%"
                    End If
                    dSum = dSum + dNet
                    nCount = nCount + 1
                End If
                .TextMatrix(iRow, 6) = s
                If dNet < 0 Then
                    .Cell(flexcpForeColor, iRow, 5) = vbRed
                    .Cell(flexcpForeColor, iRow, 6) = vbRed
                Else
                    .Cell(flexcpForeColor, iRow, 5) = RGB(0, 128, 0)
                    .Cell(flexcpForeColor, iRow, 6) = RGB(0, 128, 0)
                End If
                
                If dEntryPrice <> kNullData Then
                    ' Stop-and-reverse order to place ...
                    .TextMatrix(iRow, 7) = strOrder
                    If Left(strOrder, 1) = "S" Then
                        .Cell(flexcpForeColor, iRow, 7) = vbRed
                    Else
                        .Cell(flexcpForeColor, iRow, 7) = vbBlue
                    End If
                    
                    ' If filled, then place next order ...
                    .TextMatrix(iRow, 8) = strOrder2
                    If Left(strOrder2, 1) = "S" Then
                        .Cell(flexcpForeColor, iRow, 8) = vbRed
                    Else
                        .Cell(flexcpForeColor, iRow, 8) = vbBlue
                    End If
                    .MergeRow(iRow) = False
                Else
                    ' If position is Flat, should place both orders ...
                    s = strOrder & "  AND  " & strOrder2
                    .TextMatrix(iRow, 7) = s
                    .TextMatrix(iRow, 8) = s
                    .Cell(flexcpForeColor, iRow, 7) = vbBlack
                    .Cell(flexcpForeColor, iRow, 8) = vbBlack
                    .MergeCells = flexMergeFree
                    .MergeRow(iRow) = True
                End If
                
                If nCount Mod 10 = 1 Then
                    .AutoSize 0, .Cols - 1, , 180
                End If
            End With
        End If
    Next
    dTime = gdTickCount - dTime
    'AddList Str(dTime) & " ms to run"
    Set Strategy = Nothing
    Set Bars = Nothing
    
    s = ""
    If nCount > 0 Then
        iRow = fg.Rows
        fg.Rows = fg.Rows + 1
        If strSecType = "F" Then
            s = Format(dSum, "$#,##0")
            fg.TextMatrix(iRow, 0) = "Total OE ="
            fg.TextMatrix(iRow, 6) = s
        ElseIf strSecType = "X" Then
            s = Str(Round(dSum, 0)) & " pips"
            fg.TextMatrix(iRow, 0) = "Total OE ="
            fg.TextMatrix(iRow, 6) = s
        Else
            ' average %
            dSum = dSum / nCount
            s = Format(dSum, "#0.000") & "%"
            fg.TextMatrix(iRow, 0) = "Average OE ="
            fg.TextMatrix(iRow, 6) = s
        End If
        fg.Cell(flexcpAlignment, iRow, 0) = flexAlignRightCenter
        fg.Cell(flexcpFontBold, iRow, 0) = True
        fg.Cell(flexcpFontBold, iRow, 6) = True
    End If
    fg.AutoSize 0, fg.Cols - 1, , 180
    
    'Me.Caption = "Total OE = " & s

ErrExit:
    Set fg = Nothing
    Set Strategy = Nothing
    Set Bars = Nothing
    m.bRunning = False
    On Error Resume Next
    If m.bUnloading Then
        Unload Me
    End If
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.RunElite"
End Sub

Private Sub vsTab_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If fgReport(vsTab.CurrTab).Rows <= 1 Then
            RunElite
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSaiElite.vsTab_Click"
End Sub

Private Sub vsTab_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)

    'RunElite

End Sub

