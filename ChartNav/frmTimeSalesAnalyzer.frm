VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTimeSalesAnalyzer 
   Caption         =   "Time & Sales Analyzer"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraTradeSpeed 
      Height          =   4420
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   6740
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
      Caption         =   "frmTimeSalesAnalyzer.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSalesAnalyzer.frx":0036
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesAnalyzer.frx":0056
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtTradePeriod 
         Height          =   315
         Left            =   4260
         TabIndex        =   7
         Top             =   240
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTimeSalesAnalyzer.frx":0072
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
         Tip             =   "frmTimeSalesAnalyzer.frx":0094
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":00B4
      End
      Begin HexUniControls.ctlUniComboImageXP cboTradeInterval 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2535
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
         Tip             =   "frmTimeSalesAnalyzer.frx":00D0
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":00F0
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgTrade 
         Height          =   2500
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6135
         _cx             =   10821
         _cy             =   4410
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
         Rows            =   14
         Cols            =   9
         FixedRows       =   1
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
      Begin HexUniControls.ctlUniLabelXP lblTradePeriod 
         Height          =   255
         Left            =   2880
         Top             =   270
         Width           =   1275
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
         Caption         =   "frmTimeSalesAnalyzer.frx":010C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesAnalyzer.frx":0148
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":0168
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.Timer tmr 
      Interval        =   500
      Left            =   8160
      Top             =   5520
   End
   Begin HexUniControls.ctlUniFrameWL fraBidAskVol 
      Height          =   4415
      Left            =   120
      TabIndex        =   1
      Top             =   100
      Width           =   6740
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
      Caption         =   "frmTimeSalesAnalyzer.frx":0184
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSalesAnalyzer.frx":01C0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesAnalyzer.frx":01E0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRange 
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   210
         Width           =   950
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
         Caption         =   "frmTimeSalesAnalyzer.frx":01FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesAnalyzer.frx":0224
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":0244
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboVolInterval 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
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
         Tip             =   "frmTimeSalesAnalyzer.frx":0260
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":0280
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtVolPeriod 
         Height          =   315
         Left            =   4020
         TabIndex        =   5
         Top             =   240
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTimeSalesAnalyzer.frx":029C
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
         Tip             =   "frmTimeSalesAnalyzer.frx":02BE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":02DE
      End
      Begin VSFlex7LCtl.VSFlexGrid fgVol 
         Height          =   2500
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   6135
         _cx             =   10821
         _cy             =   4410
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
         Rows            =   14
         Cols            =   9
         FixedRows       =   1
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
      Begin HexUniControls.ctlUniLabelXP lblVolPeriod 
         Height          =   255
         Left            =   2640
         Top             =   270
         Width           =   1275
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
         Caption         =   "frmTimeSalesAnalyzer.frx":02FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesAnalyzer.frx":0336
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":0356
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblVolRange 
         Height          =   375
         Left            =   2740
         Top             =   210
         Visible         =   0   'False
         Width           =   2475
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
         Caption         =   "frmTimeSalesAnalyzer.frx":0372
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesAnalyzer.frx":03D0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesAnalyzer.frx":03F0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTimeSales 
      Height          =   4350
      Left            =   7200
      TabIndex        =   0
      Top             =   195
      Width           =   2775
      _cx             =   4895
      _cy             =   7673
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
      Left            =   9360
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   4
      DisplayContextMenu=   0   'False
      Tools           =   "frmTimeSalesAnalyzer.frx":040C
      ToolBars        =   "frmTimeSalesAnalyzer.frx":0572
   End
End
Attribute VB_Name = "frmTimeSalesAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kFormWd = 10725
Private Const kFormHt = 4845

Private Const kTimeRangeText = "Custom Time Range"
Private Const kEditPrompt = "Click here to add start/end time ..."

Private Const kCaptionBase = "AMPT Time & Sales Analyzer"

Private Type mPrivate
    Data As cTSVData
    strVolPeriod As String
    strTradePeriod As String
    
    nSymID As Long
    strSym As String
    
    dSumAvgLotBid As Double     'for re-calculating daily totals RT
    dSumAvgLotAsk As Double
    dSumVol As Double
    dSumVolBid As Double
    dSumVolAsk As Double
    
    nArrange As Long            '0=vert, 1=horz

    bInitInprog As Boolean
    bTimerInProg As Boolean
    bReloadData As Boolean
End Type

Private m As mPrivate

Public Sub ShowMe(ByVal strSym$)
On Error GoTo ErrSection:
    
    Dim strText$
    
    tbToolbar.Tools("ID_Arrange").ComboBox.ListIndex = m.nArrange
       
    InitTimeSalesGrid
    InitVolGrid
    InitTradeGrid
    
    If cboVolInterval.ListCount > 0 Then InitCbo True
    If cboTradeInterval.ListCount > 0 Then InitCbo False
    
    LoadNewSymData strSym
    
    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If
    
    ShowForm Me, eForm_Nonmodal, frmMain
    
    tmr.Enabled = g.RealTime.Active
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.ShowMe"

End Sub

Private Sub InitTimeSalesGrid()
On Error GoTo ErrSection:
        
    m.bInitInprog = True
    
    With fgTimeSales
        .FixedCols = 0
        .FixedRows = 1
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
    End With
            
    m.bInitInprog = False
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.InitTimeSalesGrid"

End Sub

Private Sub InitVolGrid()
On Error GoTo ErrSection:
        
    m.bInitInprog = True
    
    With fgVol
        .FixedCols = 0
        .FixedRows = 1
        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
        .Rows = 3
                
        .MergeRow(0) = True
        .MergeRow(2) = True
        .MergeCells = flexMergeFree
        
        'header row 1
        .TextMatrix(0, 0) = "INTERVAL"
        .TextMatrix(0, 1) = "INTERVAL"
        .TextMatrix(0, 2) = "BOUGHT"
        .TextMatrix(0, 3) = "BOUGHT"
        .TextMatrix(0, 4) = "BOUGHT"
        .TextMatrix(0, 5) = Space(5)
        .TextMatrix(0, 6) = "SOLD"
        .TextMatrix(0, 7) = "SOLD"
        .TextMatrix(0, 8) = "SOLD"
        
        'header row 2
        .TextMatrix(1, 0) = "Start"
        .TextMatrix(1, 1) = "End"
        .TextMatrix(1, 2) = "Avg" & vbCrLf & "Lots"
        .TextMatrix(1, 3) = "Percent"
        .TextMatrix(1, 4) = "Volume"
        .TextMatrix(1, 5) = "Total" & vbCrLf & "Volume"
        .TextMatrix(1, 6) = "Volume"
        .TextMatrix(1, 7) = "Percent"
        .TextMatrix(1, 8) = "Avg" & vbCrLf & "Lots"
        .RowHeight(1) = .RowHeight(0) * 2
        
        'row 3
        .TextMatrix(2, 0) = "Daily Totals"
        .TextMatrix(2, 1) = "Daily Totals"
                
        .ColWidthMin = 600
        .Cell(flexcpFontBold, 0, 2, 0, .Cols - 1) = True
        .Cell(flexcpFontBold, 2, 0, 2, .Cols - 1) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        
        .Width = fraBidAskVol.Width - 250
        .Height = fraBidAskVol.Height - .Top - 100
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    
    m.bInitInprog = False
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.InitVolGrid"

End Sub

Private Sub InitTradeGrid()
On Error GoTo ErrSection:
        
    m.bInitInprog = True
    
    With fgTrade
        .FixedCols = 0
        .FixedRows = 1
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
        .ExtendLastCol = True
        .Rows = 3
        .Cols = 10
                
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        
        'header row 1
        .TextMatrix(0, 0) = "INTERVAL"
        .TextMatrix(0, 1) = "BOUGHT"
        .TextMatrix(0, 2) = "BOUGHT"
        .TextMatrix(0, 3) = "BOUGHT"
        .TextMatrix(0, 4) = "BOUGHT"
        .TextMatrix(0, 5) = Space(5)
        .TextMatrix(0, 6) = "SOLD"
        .TextMatrix(0, 7) = "SOLD"
        .TextMatrix(0, 8) = "SOLD"
        .TextMatrix(0, 9) = "SOLD"
        
        'header row 2
        .TextMatrix(1, 0) = "User" & vbCrLf & "Defined"
        .TextMatrix(1, 1) = "Avg" & vbCrLf & "Lots"
        .TextMatrix(1, 2) = "Percent"
        .TextMatrix(1, 3) = "Speed"
        .TextMatrix(1, 4) = "Volume"
        .TextMatrix(1, 5) = "Total" & vbCrLf & "Volume"
        .TextMatrix(1, 6) = "Volume"
        .TextMatrix(1, 7) = "Speed"
        .TextMatrix(1, 8) = "Percent"
        .TextMatrix(1, 9) = "Avg" & vbCrLf & "Lots"
        .RowHeight(1) = .RowHeight(0) * 2

        .ColWidthMin = 850
        .Cell(flexcpFontBold, 0, 1, 0, .Cols - 1) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        
        .Width = fraTradeSpeed.Width - 250
        .Height = fraTradeSpeed.Height - .Top - 100
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    
    m.bInitInprog = False
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.InitTradeGrid"

End Sub

Private Sub cboVolInterval_Click()
On Error Resume Next

    FixIntervalCtrls cboVolInterval, lblVolRange, lblVolPeriod, txtVolPeriod, fgVol
    
End Sub

Private Sub cboTradeInterval_Click()
On Error Resume Next

    FixIntervalCtrls cboTradeInterval, lblTradePeriod, lblTradePeriod, txtTradePeriod, fgTrade

End Sub

Private Sub InitCbo(ByVal bVolCtrl As Boolean)
On Error GoTo ErrSection:

    Dim strPeriod$, strCaption$, strToParse$
    Dim lblCtrl As ctlUniLabelXP, txtCtrl As ctlUniTextBoxXP 'TextBox 'RH changed from Textbox
    Dim nIdx&

    If bVolCtrl Then
        strToParse = m.strVolPeriod
        Set lblCtrl = lblVolPeriod
        Set txtCtrl = txtVolPeriod
    Else
        strToParse = m.strTradePeriod
        Set lblCtrl = lblTradePeriod
        Set txtCtrl = txtTradePeriod
    End If
    
    If Len(strToParse) > 0 Then
        strPeriod = Mid(strToParse, 1, Len(strToParse) - 1)
        strCaption = Right(strToParse, 1)
        If strCaption = "m" Then
            strCaption = "Minutes per bar:"
        ElseIf strCaption = "t" Then
            strCaption = "Ticks per bar:"
            nIdx = 1
        ElseIf strCaption = "v" Then
            strCaption = "Volume per bar:"
            nIdx = 2
        ElseIf bVolCtrl Then
        Else
            strCaption = "Minutes per bar:"     'default for trade speed grid
            strPeriod = "30"
        End If
    Else
        strCaption = "Minutes per bar:"
        If bVolCtrl Then
            strPeriod = "30"
        Else
            strPeriod = "5"
        End If
    End If
    
    lblCtrl.Caption = strCaption
    txtCtrl.Text = strPeriod
    
    If bVolCtrl Then
        lblVolRange.Visible = False
        cboVolInterval.ListIndex = nIdx
    Else
        cboTradeInterval.ListIndex = nIdx
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.InitCbo"

End Sub

Private Function IsValidTime(ByVal strTime) As Boolean
On Error GoTo ErrSection:

    Dim bValid As Boolean
    Dim strHour$, strMinute$
    Dim nHour&, nMinute&
    
    bValid = True       'assume success
    If Len(strTime) <> 5 Then
        bValid = False
    Else
        strHour = Left(strTime, 2)
        strMinute = Right(strTime, 2)
        
        If IsDigit(strHour) And IsDigit(strHour, 2) And IsDigit(strMinute) And IsDigit(strMinute, 2) Then
            nHour = ValOfText(strHour)
            nMinute = ValOfText(strMinute)
        Else
            bValid = False
        End If
    End If
    
    IsValidTime = bValid
    
    Exit Function

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.IsValidTime"

End Function

Private Sub DisplayRangeTotals()
On Error GoTo ErrSection:

    Dim i&, j&
    Dim dVolBid#, dAvgLotBid#, dPercentBid#
    Dim dVolAsk#, dAvgLotAsk#, dPercentAsk#
    Dim dVol#
    
    With fgVol
        j = .Rows - 3
        If j > 0 Then
            For i = 3 To .Rows - 1
                If .MergeRow(i) Then Exit For
                dAvgLotAsk = dAvgLotAsk + ValOfText(.TextMatrix(i, 2))
                dPercentAsk = dPercentAsk + ValOfText(.TextMatrix(i, 3))
                dVolAsk = dVolAsk + ValOfText(.TextMatrix(i, 4))
                dVol = dVol + ValOfText(.TextMatrix(i, 5))
                dVolBid = dVolBid + ValOfText(.TextMatrix(i, 6))
                dPercentBid = dPercentBid + ValOfText(.TextMatrix(i, 7))
                dAvgLotBid = dAvgLotBid + ValOfText(.TextMatrix(i, 8))
            Next
            
            dAvgLotBid = RoundNum(dAvgLotBid / j)
            dPercentBid = RoundNum(dPercentBid / j)
            dVolBid = RoundNum(dVolBid / j)
            dVolAsk = RoundNum(dVolAsk / j)
            dPercentAsk = RoundNum(dPercentAsk / j)
            dAvgLotAsk = RoundNum(dAvgLotAsk / j)
            
            .TextMatrix(2, 2) = Str(Int(dAvgLotAsk))
            .TextMatrix(2, 3) = Str(Int(dPercentAsk))
            .TextMatrix(2, 4) = Str(Int(dVolAsk))
            .TextMatrix(2, 5) = Str(Int(dVol))
            .TextMatrix(2, 6) = Str(Int(dVolBid))
            .TextMatrix(2, 7) = Str(Int(dPercentBid))
            .TextMatrix(2, 8) = Str(Int(dVolBid))
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.DisplayRangeTotals"

End Sub


Private Sub cmdRange_Click()
On Error GoTo ErrSection:

    Dim aBars As cGdArray
    Dim Bars As cGdBars
    Dim i&

    Dim dVolBid#, dAvgLotBid#, dPercentBid#
    Dim dVolAsk#, dAvgLotAsk#, dPercentAsk#
    Dim dVol#, dDontCare#
    
    If cmdRange.Caption = "Edit" Then
        With fgVol
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = kEditPrompt
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
        End With
        cmdRange.Caption = "Done"
        Exit Sub
    End If
    
    With fgVol
        For i = .Rows - 1 To 3 Step -1
            If .MergeRow(i) Then
                .Rows = .Rows - 1
            Else
                Exit For
            End If
        Next
        m.Data.ResetTimeRangeBars fgVol, 3, fgVol.Rows - 1
    End With
    
    Set aBars = m.Data.TimeRangeBars
    If aBars Is Nothing Then Exit Sub
    
    For i = 0 To aBars.Size - 1
        Set Bars = aBars(i)
        
        If Bars.Size = 1 Then
            CalcVolData 0, Bars, dVolBid, dAvgLotBid, dPercentBid, dVolAsk, dAvgLotAsk, dPercentAsk, dVol, dDontCare, dDontCare
            With fgVol
                .TextMatrix(i + 3, 2) = Str(Int(dAvgLotAsk))
                .TextMatrix(i + 3, 3) = Str(Int(dPercentAsk))
                .TextMatrix(i + 3, 4) = Str(Int(dVolAsk))
                .TextMatrix(i + 3, 5) = Str(Int(dVol))
                .TextMatrix(i + 3, 6) = Str(Int(dVolBid))
                .TextMatrix(i + 3, 7) = Str(Int(dPercentBid))
                .TextMatrix(i + 3, 8) = Str(Int(dAvgLotBid))
            End With
        Else
            With fgVol
                .TextMatrix(i + 3, 2) = ""
                .TextMatrix(i + 3, 3) = ""
                .TextMatrix(i + 3, 4) = ""
                .TextMatrix(i + 3, 5) = ""
                .TextMatrix(i + 3, 6) = ""
                .TextMatrix(i + 3, 7) = ""
                .TextMatrix(i + 3, 8) = ""
            End With
        End If
    Next
    
    DisplayRangeTotals

    Set Bars = Nothing
    
    cmdRange.Caption = "Edit"
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.cmdRange_Click"

End Sub

Private Sub fgVol_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    If Col = 0 Then
        If Not IsValidTime(fgVol.TextMatrix(Row, Col)) Then
            fgVol.TextMatrix(Row, Col) = ""
            fgVol.Col = 0
            fgVol.EditCell
        End If
    ElseIf Col = 1 Then
        With fgVol
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = kEditPrompt
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
        End With
    End If

End Sub

Private Sub fgVol_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next

    If cboVolInterval.Text <> kTimeRangeText Then
        Cancel = True
    ElseIf Row <= 2 Then
        Cancel = True
    ElseIf Col <> 0 And Col <> 1 Then
        Cancel = True
    ElseIf cboVolInterval.Text = kTimeRangeText And Not fgVol.MergeRow(fgVol.Rows - 1) Then
        If cmdRange.Caption = "Edit" Then Cancel = True
    End If

End Sub

Private Sub fgVol_Click()
On Error Resume Next
    
    Dim lMouseRow As Long

    If cboVolInterval.Text = kTimeRangeText Then
        With fgVol
            lMouseRow = .MouseRow
            If lMouseRow >= .FixedRows Then
                If .MergeRow(lMouseRow) = True Then
                    .Cell(flexcpText, lMouseRow, 0, lMouseRow, .Cols - 1) = ""
                    .MergeRow(lMouseRow) = False
                    .Col = 0
                    .EditCell
                End If
            End If
        End With
    End If

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("ID_TimeSalesAnalyzer"))
    
    g.Styler.StyleForm Me
    
    With tbToolbar
        .Tools("ID_Symbol").Picture = Picture16(ToolbarIcon("ID_Symbol"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_Settings").Picture = Picture16(ToolbarIcon("ID_Settings"))   'want new toolbar to use kSettings for consistency
    End With
    
    'caption for time range label
    lblVolRange.Caption = "Specify start/end times in grid." & vbCrLf & "Click button when done."
    
    'interval controls for bid/ask vol data
    With cboVolInterval
        .Clear
        .AddItem "Minute Bars"
        .AddItem "Tick Bars"
        .AddItem "Volume Bars"
        .AddItem kTimeRangeText
        .ListIndex = 0
    End With
    InitCbo True
        
    'interval controls for trade speed data
    With cboTradeInterval
        .Clear
        .AddItem "Minute Bars"
        .AddItem "Tick Bars"
        .AddItem "Volume Bars"
        .ListIndex = 0
    End With
    InitCbo False
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.Form_Load"

End Sub

Private Function VolFrameCtlPos() As Long
On Error GoTo ErrSection:

    Dim iHeight&, iWidth&
    
    iWidth = cboVolInterval.Width + lblVolRange.Width + txtVolPeriod.Width + cmdRange.Width / 2
    
    If iWidth > fraBidAskVol.Width Then
        lblVolRange.Move cboVolInterval.Left, cboVolInterval.Top + cboVolInterval.Height + 80
        lblVolPeriod.Move cboVolInterval.Left, cboVolInterval.Top + cboVolInterval.Height + 100
        txtVolPeriod.Move lblVolRange.Left + lblVolPeriod.Width + 50, lblVolRange.Top
        iHeight = cboVolInterval.Height + lblVolRange.Height + 80
    Else
        lblVolRange.Move cboVolInterval.Left + cboVolInterval.Width + 50, cboVolInterval.Top
        lblVolPeriod.Move lblVolRange.Left, lblVolRange.Top + 50
        txtVolPeriod.Move lblVolRange.Left + lblVolPeriod.Width + 50, lblVolRange.Top
    End If
    
    cmdRange.Move fraBidAskVol.Width - cmdRange.Width - 125, lblVolRange.Top

    VolFrameCtlPos = iHeight
    
    Exit Function
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.VolFrameCtlPos"

End Function

Private Function TradeFrameCtlPos() As Long
On Error GoTo ErrSection:

    Dim iHeight&, iWidth&
    
    iWidth = cboTradeInterval.Width + lblTradePeriod.Width + txtTradePeriod.Width
    
    If iWidth > fraTradeSpeed.Width Then
        lblTradePeriod.Move cboTradeInterval.Left, cboTradeInterval.Top + cboTradeInterval.Height + 100
        txtTradePeriod.Move lblTradePeriod.Left + lblTradePeriod.Width + 50, lblTradePeriod.Top - 40
        iHeight = cboTradeInterval.Height + txtTradePeriod.Height + 80
    Else
        lblTradePeriod.Move cboTradeInterval.Left + cboTradeInterval.Width + 50, cboTradeInterval.Top + 50
        txtTradePeriod.Move lblTradePeriod.Left + lblTradePeriod.Width + 50, cboTradeInterval.Top
    End If
    
    TradeFrameCtlPos = iHeight
    
    Exit Function

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.TradeFrameCtlPos"

End Function

Private Sub Form_Resize()
On Error Resume Next

    Dim iVolCtlHt&, iTradeCtlHt&
    
    'If LimitFormSize(Me, kFormWd, kFormHt) Then Exit Sub
    
    
    If m.nArrange = 0 Then
        fgTimeSales.Left = Me.Width - fgTimeSales.Width - 250
        
        fraBidAskVol.Width = Me.Width - fgTimeSales.Width - 450
        fraBidAskVol.Height = Int(Me.ScaleHeight / 2) - 100
        
        iVolCtlHt = VolFrameCtlPos()
        
        fgVol.Top = cboVolInterval.Top + cboVolInterval.Height + 100
        fgVol.Width = fraBidAskVol.Width - 250
        fgVol.Height = fraBidAskVol.Height - cboVolInterval.Height - 500
        
        fgTimeSales.Height = fraBidAskVol.Height - 100
        
        fraTradeSpeed.Move fraBidAskVol.Left, _
                           fraBidAskVol.Top + fraBidAskVol.Height + 50, _
                           Me.ScaleWidth - 200, _
                           fraBidAskVol.Height
                           
        iTradeCtlHt = TradeFrameCtlPos()
                           
        fgTrade.Width = fraTradeSpeed.Width - 250
        fgTrade.Height = fraTradeSpeed.Height - cboVolInterval.Height - 500
    Else
        fraBidAskVol.Width = Int((Me.Width - fgTimeSales.Width) * 0.44)
        fraBidAskVol.Height = Me.ScaleHeight - 150
        
        fgTimeSales.Left = fraBidAskVol.Left + fraBidAskVol.Width + 150
        fgTimeSales.Height = fraBidAskVol.Height - 100
        
        iVolCtlHt = VolFrameCtlPos()
        
        If iVolCtlHt > 0 Then
            fgVol.Top = cboVolInterval.Top + iVolCtlHt + 50
        Else
            fgVol.Top = cboVolInterval.Top + cboVolInterval.Height + 100
        End If
        fgVol.Width = fraBidAskVol.Width - 250
        fgVol.Height = fraBidAskVol.Height - cboVolInterval.Height - 500
        
        fraTradeSpeed.Move fgTimeSales.Left + fgTimeSales.Width + 150, _
                           fraBidAskVol.Top, _
                           Me.Width - fraBidAskVol.Width - fgTimeSales.Width - 500, _
                           fraBidAskVol.Height
        
        iTradeCtlHt = TradeFrameCtlPos()
        
        If iTradeCtlHt > 0 Then
            fgTrade.Top = cboTradeInterval.Top + iTradeCtlHt + 50
        Else
            fgTrade.Top = cboTradeInterval.Top + cboTradeInterval.Height + 100
        End If
        
        fgTrade.Width = fraTradeSpeed.Width - 250
        fgTrade.Height = fgVol.Height
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Set m.Data = Nothing
    
    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile

End Sub

Private Sub tbToolbar_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim i&
    
    i = Tool.ComboBox.ListIndex
    If i <> m.nArrange Then
        m.nArrange = i
        Form_Resize
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.tbToolbar_ComboCloseUp"

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
        
    Select Case Tool.ID
        Case "ID_Symbol"
            LoadNewSymData ""
        Case "ID_Close"
            Unload Me
        Case "ID_Settings"
    End Select
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.tbToolbar_ToolClick"

End Sub

Private Sub LoadTimeSalesData(ByVal strSym$, ByVal nSymID&)
On Error GoTo ErrSection:

    Dim i&
    
    If g.bUnloading Then Exit Sub

    m.Data.ResetTickBars strSym, nSymID, Me
    With fgTimeSales
        .Redraw = flexRDNone
        
        .FlexDataSource = m.Data
        .AutoSize 0, .Cols - 1
        
        For i = .FixedRows To .Rows - 1
            If "1" = .TextMatrix(i, 3) Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = kAskColor
            ElseIf "2" = .TextMatrix(i, 3) Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = kBidColor
            End If
        Next
        
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalsesAnalyzer.LoadTimeSalesData"

End Sub

Private Sub CalcVolData(ByVal iBar&, Bars As cGdBars, _
    dBidVol#, dAvgLotBid#, dPercentBid#, _
    dAskVol, dAvgLotAsk#, dPercentAsk#, dVol#, _
    dBidTrades#, dAskTrades#)
On Error GoTo ErrSection:
        
    dVol = Bars(eBARS_Vol, iBar)            'vol of all trades
    If dVol = kNullData Then dVol = 0
    
    dBidVol = Bars(eBARS_BidVol, iBar)      'vol of all trades at bid
    If dBidVol = kNullData Then dBidVol = 0
    
    dAskVol = Bars(eBARS_AskVol, iBar)      'vol of all trades at ask
    If dAskVol = kNullData Then dAskVol = 0
    
    dBidTrades = Bars(eBARS_UpTicks, iBar)        'see note in cTSVData.cls for this
    If dBidTrades = kNullData Then dBidTrades = 0
    
    dAskTrades = Bars(eBARS_DownTicks, iBar)              'count of trades at ask
    If dAskTrades = kNullData Then dAskTrades = 0
        
    If dBidVol > 0 And dBidTrades > 0 Then dAvgLotBid = RoundNum(dBidVol / dBidTrades)
    
    If dAskVol > 0 And dAskTrades > 0 Then dAvgLotAsk = RoundNum(dAskVol / dAskTrades)
    
    If dVol > 0 Then
        If dBidVol > 0 Then dPercentBid = RoundNum(dBidVol / dVol, 2) * 100  'percentages
        If dAskVol > 0 Then dPercentAsk = RoundNum(dAskVol / dVol, 2) * 100
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.CalcVolData"

End Sub

Private Sub LoadVolData(ByVal strSym$, ByVal nSymID&)
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim i&, j&, iStart&
    Dim dBidVol#, dAvgLotBid#, dPercentBid#
    Dim dAskVol#, dAvgLotAsk#, dPercentAsk#, dVol#
    Dim dDontCare#, dDateTime#
    'for daily totals
    Dim dSumBidVol#, dSumAskVol#, dSumVol#
    Dim dSumAvgBid#, dSumAvgAsk#
        
    If g.bUnloading Then Exit Sub
    
    'reset
    m.dSumAvgLotBid = 0
    m.dSumAvgLotAsk = 0
    m.dSumVol = 0
    m.dSumVolBid = 0
    m.dSumVolAsk = 0
    'bid/ask volume grid
    If Len(m.strVolPeriod) = 0 Then m.strVolPeriod = "30m"
    m.Data.ResetVolBars m.strVolPeriod
    Set Bars = m.Data.BidAskVolBars
    
    If Bars Is Nothing Then
        ClearGrid True, False
        Exit Sub
    ElseIf Bars.Size = 0 Then
        ClearGrid True, False
        Exit Sub
    End If
    
    iStart = Bars.Size - 1
    With fgVol
        .Redraw = flexRDNone
        .Rows = 3
        For i = iStart To 0 Step -1
            .Rows = .Rows + 1
            'start/end time
            dDateTime = Bars(eBARS_DateTime, i)
            If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            .TextMatrix(.Rows - 1, 1) = DateFormat(dDateTime, NO_DATE, HH_MM)
            If i = 0 Then
                '.TextMatrix(.Rows - 1, 0) = DateFormat(Bars.Prop(eBARS_StartTime) / 1440, NO_DATE, HH_MM)  - save for reference
                dDateTime = m.Data.TickBars(eBARS_DateTime, 0)
                If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                .TextMatrix(.Rows - 1, 0) = DateFormat(dDateTime, NO_DATE, HH_MM)
            Else
                dDateTime = Bars(eBARS_DateTime, i - 1)
                If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                .TextMatrix(.Rows - 1, 0) = DateFormat(dDateTime, NO_DATE, HH_MM)
            End If
            
            CalcVolData i, Bars, dBidVol, dAvgLotBid, dPercentBid, dAskVol, dAvgLotAsk, dPercentAsk, _
                dVol, dDontCare, dDontCare
            
            .TextMatrix(.Rows - 1, 2) = Str(Int(dAvgLotAsk))
            .TextMatrix(.Rows - 1, 3) = Str(Int(dPercentAsk))
            .TextMatrix(.Rows - 1, 4) = Str(Int(dAskVol))
            .TextMatrix(.Rows - 1, 5) = Str(Int(dVol))
            .TextMatrix(.Rows - 1, 6) = Str(Int(dBidVol))
            .TextMatrix(.Rows - 1, 7) = Str(Int(dPercentBid))
            .TextMatrix(.Rows - 1, 8) = Str(Int(dAvgLotBid))
            
            .Cell(flexcpBackColor, .Rows - 1, 7, .Rows - 1, 8) = vbWhite
            .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, 3) = vbWhite
            
            If dBidVol > dAskVol Then
                .Cell(flexcpBackColor, .Rows - 1, 7, .Rows - 1, 8) = kFrameShort
            ElseIf dAskVol > dBidVol Then
                .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, 3) = kFrameLong
            End If
            
            'sums for daily totals
            dSumVol = dSumVol + dVol
            dSumBidVol = dSumBidVol + dBidVol
            dSumAskVol = dSumAskVol + dAskVol
            dSumAvgBid = dSumAvgBid + dAvgLotBid
            dSumAvgAsk = dSumAvgAsk + dAvgLotAsk
            'exclude last bar data from saved sums (i.e. row at top of grid below daily totals row)
            If i < iStart Then
                m.dSumAvgLotBid = m.dSumAvgLotBid + dAvgLotBid
                m.dSumAvgLotAsk = m.dSumAvgLotAsk + dAvgLotAsk
                m.dSumVol = m.dSumVol + dVol
                m.dSumVolBid = m.dSumVolBid + dBidVol
                m.dSumVolAsk = m.dSumVolAsk + dAskVol
            End If
        Next
        
        'daily totals
        If .Rows - 3 > 0 Then
            j = RoundNum(dSumAvgBid / (.Rows - 3))
            .TextMatrix(2, 2) = Str(j)         'avg lots
            j = RoundNum(dSumAvgAsk / (.Rows - 3))
            .TextMatrix(2, 8) = Str(j)
        End If
        
        .TextMatrix(2, 4) = Str(Int(dSumAskVol))    'volumes
        .TextMatrix(2, 6) = Str(Int(dSumBidVol))
        .TextMatrix(2, 5) = Str(Int(dSumVol))
        
        .Cell(flexcpBackColor, 2, 7, 2, 8) = vbWhite    'clear previous coloring
        .Cell(flexcpBackColor, 2, 2, 2, 3) = vbWhite
        
        If dSumBidVol > dSumAskVol Then
            .Cell(flexcpBackColor, 2, 7, 2, 8) = kFrameShort
        ElseIf dSumAskVol > dSumBidVol Then
            .Cell(flexcpBackColor, 2, 2, 2, 3) = kFrameLong
        End If
        
        If dSumVol > 0 Then
            dSumBidVol = RoundNum(dSumBidVol / dSumVol, 2) * 100
            dSumAskVol = RoundNum(dSumAskVol / dSumVol, 2) * 100
        End If
        .TextMatrix(2, 3) = Str(Int(dSumAskVol))
        .TextMatrix(2, 7) = Str(Int(dSumBidVol))

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSize 0, .Cols - 1
        
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.LoadVolData"

End Sub

Private Sub LoadTradeData(ByVal strSym$, ByVal nSymID&)
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim dBidVol#, dAvgLotBid#, dPercentBid#
    Dim dAskVol#, dAvgLotAsk#, dPercentAsk#, dVol#
    Dim dSeconds#, dSpeed#, dBidTrades#, dAskTrades#, dDateTime#
    Dim i&, j&
    
    If g.bUnloading Then Exit Sub
    
    'trade speed grid
    If Len(m.strTradePeriod) = 0 Then m.strTradePeriod = "1m"
    
    m.Data.ResetTradeBars m.strTradePeriod
    Set Bars = m.Data.TradeSpeedBars
    If Bars Is Nothing Then
        ClearGrid False, True
        Exit Sub
    ElseIf Bars.Size = 0 Then
        ClearGrid False, True
        Exit Sub
    End If
        
    txtTradePeriod.Text = Bars.Prop(eBARS_PeriodsPerBar)
    With fgTrade
        .Redraw = flexRDNone
        .Rows = 2
        For i = Bars.Size - 1 To 0 Step -1
            .Rows = .Rows + 1
            'interval
            dDateTime = Bars(eBARS_DateTime, i)
            If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            .TextMatrix(.Rows - 1, 0) = DateFormat(dDateTime, NO_DATE, HH_MM)
            
            dSeconds = 0
            CalcVolData i, Bars, dBidVol, dAvgLotBid, dPercentBid, dAskVol, dAvgLotAsk, dPercentAsk, _
                dVol, dBidTrades, dAskTrades
            
            'volume etc.
            .TextMatrix(.Rows - 1, 5) = Str(Int(dVol))
            .TextMatrix(.Rows - 1, 4) = Str(Int(dAskVol))
            .TextMatrix(.Rows - 1, 6) = Str(Int(dBidVol))
            .TextMatrix(.Rows - 1, 1) = Str(Int(dAvgLotAsk))
            .TextMatrix(.Rows - 1, 9) = Str(Int(dAvgLotBid))
            .TextMatrix(.Rows - 1, 2) = Str(Int(dPercentAsk))
            .TextMatrix(.Rows - 1, 8) = Str(Int(dPercentBid))
            
            If dBidVol > dAskVol Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 8) = kFrameShort
            ElseIf dAskVol > dBidVol Then
                .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, 4) = kFrameLong
            End If
            
            'trade speed
            dSeconds = SecsBetweenBars(i, Bars)
            If dSeconds = 0 Then
                .TextMatrix(.Rows - 1, 3) = "0.00"
                .TextMatrix(.Rows - 1, 7) = "0.00"
            ElseIf dSeconds > 0 Then
                dSpeed = RoundNum(dAskTrades / dSeconds, 2)
                .TextMatrix(.Rows - 1, 3) = Format(dSpeed, "#0.00")
                dSpeed = RoundNum(dBidTrades / dSeconds, 2)
                .TextMatrix(.Rows - 1, 7) = Format(dSpeed, "#0.00")
            End If
        Next
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.LoadTradeData"

End Sub

Private Sub ClearGrid(ByVal bClearVol As Boolean, ByVal bClearTrade As Boolean, _
    Optional ByVal bClearTimeSales As Boolean = False)
On Error Resume Next
        
    If bClearVol Then
        With fgVol
            .Redraw = flexRDNone
            .Rows = 3
            .TextMatrix(2, 2) = ""
            .TextMatrix(2, 3) = ""
            .TextMatrix(2, 4) = ""
            .TextMatrix(2, 5) = ""
            .TextMatrix(2, 6) = ""
            .TextMatrix(2, 7) = ""
            .TextMatrix(2, 8) = ""
            If cboVolInterval.Text = kTimeRangeText Then
                .Rows = 4
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = kEditPrompt
                .MergeCells = flexMergeRestrictRows
                .MergeRow(.Rows - 1) = True
            End If
            .Redraw = flexRDBuffered
        End With
    End If
    
    If bClearTrade Then
        With fgTrade
            .Redraw = flexRDNone
            .Rows = 2
            .Redraw = flexRDBuffered
        End With
    End If
    
    'this is a virtual grid, will only clear if data object has no data
    If bClearTimeSales Then
        With fgTimeSales
            .Redraw = flexRDNone
            .FlexDataSource = Nothing
            .Rows = .FixedRows
            .Redraw = flexRDBuffered
        End With
    End If

End Sub

Private Sub LoadNewSymData(ByVal strSym$)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray
    Dim nSymbolID&, strSymbol$, strMsg$
    
    If g.bUnloading Then Exit Sub
    
    'reset grid & data
    Set m.Data = New cTSVData
    ClearGrid True, True, True
    DoEvents            'let grid repaint
        
    If Left(strSym, 1) = "$" Then 'And Not IsForex(strSym) Then
        LoadGridNoVol strSym
        Exit Sub
    End If
    
    m.strSym = ""
    m.nSymID = 0
    
    If Len(strSym) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", False)
        If astrSymbols.Size > 0 Then strSymbol = astrSymbols(0)
    Else
        strSymbol = strSym
    End If
    
    nSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
    
    If nSymbolID > 0 Then
        m.strSym = strSymbol
        m.nSymID = nSymbolID
        
        m.bInitInprog = True
                
        InfBox "Loading data.  Please wait...", , , kCaptionBase, True
        
        LoadTimeSalesData strSym, nSymbolID
        
        If m.Data.TickBars.Size > 0 Then
            LoadVolData strSym, nSymbolID
            LoadTradeData strSym, nSymbolID
            Me.Caption = kCaptionBase & " for " & m.strSym & " " & DateFormat(m.Data.SessionDate, MM_DD_YYYY, NO_TIME)
        Else
            LoadGridNoVol m.strSym
        End If
                
        m.bInitInprog = False
        
        InfBox ""
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.LoadNewSymData"

End Sub

Private Sub UpadateRangeRT(ByVal nNumNewTicks&)
On Error GoTo ErrSection:

    Dim TickBars As cGdBars
    Dim Bars As cGdBars
    Dim aRangeBars As cGdArray
    Dim dGridDate1#, dGridDate2#, dBarDate#
    Dim dTickDate#, nTickIdx&, i&
    Dim bUpdate As Boolean

    'for calculating updated values
    Dim dBidVol#, dAvgLotBid#, dPercentBid#
    Dim dAskVol, dAvgLotAsk#, dPercentAsk#, dVol#, dDontCare#
    
    If cmdRange.Caption = "Done" Then Exit Sub
    
    Set TickBars = m.Data.TickBars
    If TickBars Is Nothing Then Exit Sub            'precautionary, theoretically should never happen
    
    Set aRangeBars = m.Data.TimeRangeBars
    If aRangeBars Is Nothing Then Exit Sub          'precautionary, theoretically should never happen
    
    'subtract number of new ticks to get bar index of where new tick starts
    nTickIdx = TickBars.Size - nNumNewTicks
    'get the date time of earliest new tick
    dTickDate = TickBars(eBARS_DateTime, nTickIdx)
    
    'loop through array of time-range bars to see which bar need updating
    'bars contain end-date value as double (if end date < newest tick date then no need to update)
    fgVol.Redraw = flexRDNone
    For i = 0 To aRangeBars.Size - 1
        Set Bars = aRangeBars(i)
        dBarDate = Bars(eBARS_DateTime, 0)
        
        bUpdate = False
        If dTickDate > 0 And dBarDate > 0 Then
            If Second(dTickDate) > 0 Then
                'disregard seconds value in tick time
                If Hour(dTickDate) > Hour(dBarDate) Then
                    bUpdate = True
                ElseIf Hour(dTickDate) = Hour(dBarDate) Then
                    If Minute(dTickDate) <= Minute(dBarDate) Then
                        bUpdate = True
                    End If
                End If
            ElseIf dBarDate <= dTickDate Then
                bUpdate = True
            End If
        End If
        bUpdate = True      'testcode
        
        If bUpdate Then
            dGridDate1 = DateValue(Now) + TimeValue(fgVol.TextMatrix(i + 3, 0))     'start time
            dGridDate2 = DateValue(Now) + TimeValue(fgVol.TextMatrix(i + 3, 1))     'end time
            
            m.Data.TimeRangeBarData Bars, dGridDate1, dGridDate2, nTickIdx, TickBars.Size - 1, 0
            CalcVolData 0, Bars, dBidVol, dAvgLotBid, dPercentBid, dAskVol, dAvgLotAsk, dPercentAsk, dVol, dDontCare, dDontCare
            
            With fgVol
                 .TextMatrix(i + 3, 2) = Str(Int(dAvgLotAsk))
                 .TextMatrix(i + 3, 3) = Str(Int(dPercentAsk))
                 .TextMatrix(i + 3, 4) = Str(Int(dAskVol))
                 .TextMatrix(i + 3, 5) = Str(Int(dVol))
                 .TextMatrix(i + 3, 6) = Str(Int(dBidVol))
                 .TextMatrix(i + 3, 7) = Str(Int(dPercentBid))
                 .TextMatrix(i + 3, 8) = Str(Int(dAvgLotBid))
            End With
            DisplayRangeTotals
        End If
    Next
    fgVol.Redraw = flexRDBuffered
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.UpdateRangeRT"

End Sub

Private Sub UpdateVolRT(ByVal nPrevSize&, ByVal bNewBar As Boolean)
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim strStart$, strEnd$
    Dim bReload As Boolean
    
    Dim dVol#, dDontCare#
    Dim dVolBid#, dAvgLotBid#, dPercentBid#
    Dim dVolAsk#, dAvgLotAsk#, dPercentAsk#
    'for calculating daily totals
    Dim dNewVol#, dNewVolBid#, dNewVolAsk#
    Dim dNewAvgLotBid#, dNewAvgLotAsk#
    Dim dDateTime#
           
    If bNewBar Then
        bReload = True
    Else
        Set Bars = m.Data.BidAskVolBars
        If Bars Is Nothing Then
            Exit Sub                'theoretically should never get here ...
        ElseIf Bars.Size <> nPrevSize Then
            bReload = True          '... or here ...
        Else
            dDateTime = Bars(eBARS_DateTime, Bars.Size - 1)
            If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            strEnd = DateFormat(dDateTime, NO_DATE, HH_MM)
            dDateTime = Bars(eBARS_DateTime, Bars.Size - 2)
            If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            strStart = DateFormat(dDateTime, NO_DATE, HH_MM)
            With fgVol
                If .TextMatrix(3, 0) <> strStart Or .TextMatrix(3, 1) <> strEnd Then
                    bReload = True  '... or here ...
                End If
            End With
        End If
    End If
    
    If bReload Then
        LoadVolData m.strSym, m.nSymID
        Exit Sub
    End If
        
    CalcVolData Bars.Size - 1, Bars, dVolBid, dAvgLotBid, dPercentBid, _
        dVolAsk, dAvgLotAsk, dPercentAsk, dVol, dDontCare, dDontCare
    
    With fgVol
        .TextMatrix(3, 2) = Str(Int(dAvgLotAsk))
        .TextMatrix(3, 3) = Str(Int(dPercentAsk))
        .TextMatrix(3, 4) = Str(Int(dVolAsk))
        .TextMatrix(3, 5) = Str(Int(dVol))
        .TextMatrix(3, 6) = Str(Int(dVolBid))
        .TextMatrix(3, 7) = Str(Int(dPercentBid))
        .TextMatrix(3, 8) = Str(Int(dAvgLotBid))
        
        'color cells
        .Cell(flexcpBackColor, 3, 7, 3, 8) = vbWhite
        .Cell(flexcpBackColor, 3, 2, 3, 3) = vbWhite
        If dVolBid > dVolAsk Then
            .Cell(flexcpBackColor, 3, 7, 3, 8) = kFrameShort
        ElseIf dVolAsk > dVolBid Then
            .Cell(flexcpBackColor, 3, 2, 3, 3) = kFrameLong
        End If
        'add new values to saved sums to do new daily totals
        dNewVol = m.dSumVol + dVol
        dNewVolBid = m.dSumVolBid + dVolBid
        dNewVolAsk = m.dSumVolAsk + dVolAsk
        dNewAvgLotBid = m.dSumAvgLotBid + dAvgLotBid
        dNewAvgLotAsk = m.dSumAvgLotAsk + dAvgLotAsk
        'volume
        .TextMatrix(2, 4) = Str(Int(dNewVolAsk))
        .TextMatrix(2, 5) = Str(Int(dNewVol))
        .TextMatrix(2, 6) = Str(Int(dNewVolBid))
        If dNewVolBid > dNewVolAsk Then
            .Cell(flexcpBackColor, 2, 7, 2, 8) = kFrameShort
            .Cell(flexcpBackColor, 2, 2, 2, 3) = .Cell(flexcpBackColor, 1, 0)
        ElseIf dNewVolAsk > dNewVolBid Then
            .Cell(flexcpBackColor, 2, 2, 2, 3) = kFrameLong
            .Cell(flexcpBackColor, 2, 7, 2, 8) = .Cell(flexcpBackColor, 1, 0)
        End If
        'percentages
        If dNewVol > 0 Then
            dNewVolBid = RoundNum(dNewVolBid / dNewVol, 2) * 100
            dNewVolAsk = RoundNum(dNewVolAsk / dNewVol, 2) * 100
            .TextMatrix(2, 3) = Str(Int(dNewVolAsk))
            .TextMatrix(2, 7) = Str(Int(dNewVolBid))
        Else
            .TextMatrix(2, 3) = ""
            .TextMatrix(2, 7) = ""
        End If
        'avg lot
        If .Rows - 3 > 0 Then
            .TextMatrix(2, 2) = Str(RoundNum(dNewAvgLotAsk / (.Rows - 3)))
            .TextMatrix(2, 8) = Str(RoundNum(dNewAvgLotBid / (.Rows - 3)))
        End If
    End With
    
    Exit Sub
        
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.UpdateVolRT"

End Sub

Private Function SecsBetweenBars(ByVal iBar&, Bars As cGdBars) As Long
On Error GoTo ErrSection:

    Dim nSeconds&
    Dim dTimeBar#, dTimePrevBar#
    Dim dTimeDiff#, dTimeNow#
    
    If iBar = 0 Then
        dTimeBar = Bars(eBARS_DateTime, iBar)
        dTimePrevBar = m.Data.TickBars(eBARS_DateTime, 0)
        nSeconds = (dTimeBar - dTimePrevBar) * 86400
    ElseIf iBar > 0 And iBar < Bars.Size Then
        dTimeBar = Bars(eBARS_DateTime, iBar)
        dTimePrevBar = Bars(eBARS_DateTime, iBar - 1)
        nSeconds = (dTimeBar - dTimePrevBar) * 86400
    End If
    
    If iBar = Bars.Size - 1 Then
        dTimeNow = m.Data.TickBars(eBARS_DateTime, m.Data.TickBars.Size - 1)
        dTimeDiff = Abs(dTimeNow - dTimePrevBar) * 86400
        If Int(dTimeDiff) <= 0 Then
            nSeconds = 1
        ElseIf Int(dTimeDiff) < nSeconds Then
            nSeconds = Int(dTimeDiff)
        End If
        'StatusMsg Str(nSeconds) & "," & DateFormat(dTimePrevBar, NO_DATE, HH_MM_SS) & "," & DateFormat(dTimeNow, NO_DATE, HH_MM_SS)
    End If

    SecsBetweenBars = nSeconds
    
    Exit Function

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.SecsBetweenBars"

End Function

Private Sub UpdateTradeRT(ByVal nPrevSize&, ByVal bNewBar As Boolean)
On Error GoTo ErrSection:

    Dim bReload As Boolean
    Dim Bars As cGdBars
    Dim strTime$, dVol#, dSeconds#, dSpeed#, dDateTime#
    
    Dim dVolBid#, dAvgLotBid#, dPercentBid#, dBidTrades#
    Dim dVolAsk#, dAvgLotAsk#, dPercentAsk#, dAskTrades#
    
    If bNewBar Then
        bReload = True
    Else
        Set Bars = m.Data.TradeSpeedBars
        If Bars Is Nothing Then
            Exit Sub                'theoretically should never get here ...
        ElseIf Bars.Size <> nPrevSize Then
            bReload = True          '... or here ...
        Else
            dDateTime = Bars(eBARS_DateTime, Bars.Size - 1)
            If g.bShowInLocalTimeZone Then dDateTime = ConvertTimeZone(dDateTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
            strTime = DateFormat(dDateTime, NO_DATE, HH_MM)
            If fgTrade.TextMatrix(2, 0) <> strTime Then
                bReload = True      '... or here ...
            End If
        End If
    End If
    
    If bReload Then
        LoadTradeData m.strSym, m.nSymID
        Exit Sub
    End If

    CalcVolData Bars.Size - 1, Bars, dVolBid, dAvgLotBid, dPercentBid, _
        dVolAsk, dAvgLotAsk, dPercentAsk, dVol, dBidTrades, dAskTrades
        
    dSeconds = SecsBetweenBars(Bars.Size - 1, Bars)
    
    With fgTrade
        .Redraw = flexRDNone
        
        .TextMatrix(2, 1) = Str(Int(dAvgLotAsk))
        .TextMatrix(2, 2) = Str(Int(dPercentAsk))
        .TextMatrix(2, 4) = Str(Int(dVolAsk))
        .TextMatrix(2, 5) = Str(Int(dVol))
        .TextMatrix(2, 6) = Str(Int(dVolBid))
        .TextMatrix(2, 8) = Str(Int(dPercentBid))
        .TextMatrix(2, 9) = Str(Int(dAvgLotBid))
        
        'color cells
        .Cell(flexcpBackColor, 2, 6, 2, 8) = vbWhite
        .Cell(flexcpBackColor, 2, 2, 2, 4) = vbWhite
        If dVolBid > dVolAsk Then
            .Cell(flexcpBackColor, 2, 6, 2, 8) = kFrameShort
        ElseIf dVolAsk > dVolBid Then
            .Cell(flexcpBackColor, 2, 2, 2, 4) = kFrameLong
        End If
        
        'trade speed
        If dSeconds = 0 Then
            .TextMatrix(2, 3) = "0.00"
            .TextMatrix(2, 7) = "0.00"
        ElseIf dSeconds > 0 Then
            dSpeed = RoundNum(dAskTrades / dSeconds, 2)
            .TextMatrix(2, 3) = Format(dSpeed, "#0.00")
            dSpeed = RoundNum(dBidTrades / dSeconds, 2)
            .TextMatrix(2, 7) = Format(dSpeed, "#0.00")
        Else
            .TextMatrix(2, 3) = ""
            .TextMatrix(2, 7) = ""
        End If
        
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.UpdateTradeRT"

End Sub

Private Sub tmr_Timer()
On Error GoTo ErrSection:
        
    Dim i&, nVolBarSize&, nTradeBarSize&
    Dim bNewBarVol As Boolean, bNewBarTrade As Boolean
    Dim bRebuildTable As Boolean
    Dim Bars As cGdBars
    
    If m.bInitInprog Or m.bTimerInProg Or g.bUnloading Then
        Exit Sub
    ElseIf Not g.RealTime.Active Then
        m.bTimerInProg = False
        tmr.Enabled = False
        Exit Sub
    End If
    
    m.bTimerInProg = True
    
    Set Bars = m.Data.BidAskVolBars
    If Not Bars Is Nothing Then nVolBarSize = Bars.Size
    
    Set Bars = m.Data.TradeSpeedBars
    If Not Bars Is Nothing Then nTradeBarSize = Bars.Size
    
    If g.RealTime.Active And g.RealTime.FeedTime > 0 Then
        i = m.Data.UpdateDataRT(bNewBarVol, bNewBarTrade, bRebuildTable)
        If i > 0 Then
            If Not g.bUnloading Or m.bReloadData Then
                With fgTimeSales
                    .Redraw = flexRDNone
                    .FlexDataSource = m.Data
                    .Redraw = flexRDBuffered
                    DoEvents
                    .Redraw = flexRDNone
                    For i = .FixedRows To .BottomRow
                        If "1" = .TextMatrix(i, 3) Then
                            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = kAskColor
                        ElseIf "2" = .TextMatrix(i, 3) Then
                            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = kBidColor
                        End If
                    Next
                    .Redraw = flexRDBuffered
                End With
                If cmdRange.Visible Then
                    UpadateRangeRT i
                Else
                    UpdateVolRT nVolBarSize, bNewBarVol
                End If
                UpdateTradeRT nTradeBarSize, bNewBarTrade
            End If
        End If
    End If
    
    If i = -1 Or m.bReloadData Then
        LoadTimeSalesData m.strSym, m.nSymID
        LoadVolData m.strSym, m.nSymID
        LoadTradeData m.strSym, m.nSymID
        m.bReloadData = False
    End If
    
    m.bTimerInProg = False
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.tmr_Timer"

End Sub

'RH changed from VB6 controls to Hexagora
'Private Sub FixIntervalCtrls(cbo As ComboBox, lblRange As Label, lblPeriod As Label, txtPeriod As TextBox, fg As VSFlexGrid)
Private Sub FixIntervalCtrls(cbo As ctlUniComboImageXP, lblRange As ctlUniLabelXP, lblPeriod As ctlUniLabelXP, txtPeriod As ctlUniTextBoxXP, fg As VSFlexGrid)
On Error GoTo ErrSection:

    With cbo
        If .ListIndex = 3 Then
            cmdRange.Visible = True
            cmdRange.Caption = "Done"
            lblRange.Visible = True
            lblPeriod.Visible = False
            txtPeriod.Visible = False
            ClearGrid True, False
        Else
            cmdRange.Visible = False
            lblRange.Visible = False
            lblPeriod.Visible = True
            txtPeriod.Visible = True
            Select Case cbo.ListIndex
                Case 0:
                    lblPeriod.Caption = "Minutes per bar:"
                Case 1:
                    lblPeriod.Caption = "Ticks per bar:"
                Case 2:
                    lblPeriod.Caption = "Volume per bar:"
            End Select
            MoveFocus txtPeriod
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.FixIntervalCtrls"

End Sub

Private Sub txtTradePeriod_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim strInterval$
    Dim bLoadData As Boolean
    
    If KeyCode = 13 Then
        If m.strTradePeriod <> txtTradePeriod.Text Then
            Set Bars = m.Data.TradeSpeedBars
            If Not Bars Is Nothing Then
                strInterval = cboTradeInterval.Text
                If InStr(strInterval, "Volume") <> 0 Then
                    m.strTradePeriod = txtTradePeriod.Text & "v"
                    bLoadData = True
                ElseIf InStr(strInterval, "Minute") <> 0 Then
                    m.strTradePeriod = txtTradePeriod.Text & "m"
                    bLoadData = True
                ElseIf InStr(strInterval, "Tick") <> 0 Then
                    m.strTradePeriod = txtTradePeriod.Text & "t"
                    bLoadData = True
                End If
                If bLoadData Then LoadTradeData Bars.Prop(eBARS_Symbol), Bars.Prop(eBARS_SymbolID)
            End If
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.txtTradePeriod_KeyUp"

End Sub

Private Sub txtVolPeriod_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim strInterval$
    Dim bLoadData As Boolean

    If KeyCode = 13 Then
        If m.strVolPeriod <> txtVolPeriod.Text Then
            Set Bars = m.Data.BidAskVolBars
            If Not Bars Is Nothing Then
                strInterval = cboVolInterval.Text
                If InStr(strInterval, "Volume") <> 0 Then
                    m.strVolPeriod = txtVolPeriod.Text & "v"
                    bLoadData = True
                ElseIf InStr(strInterval, "Minute") <> 0 Then
                    m.strVolPeriod = txtVolPeriod.Text & "m"
                    bLoadData = True
                ElseIf InStr(strInterval, "Tick") <> 0 Then
                    m.strVolPeriod = txtVolPeriod.Text & "t"
                    bLoadData = True
                End If
                If bLoadData Then LoadVolData Bars.Prop(eBARS_Symbol), Bars.Prop(eBARS_SymbolID)
            End If
        End If
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.txtVolPeriod_KeyUp"

End Sub

Public Sub RefreshData()
On Error GoTo ErrSection:

    If m.bInitInprog Then Exit Sub
    
    If m.bTimerInProg Then
        m.bReloadData = True
    Else
        m.bInitInprog = True
        tmr.Enabled = False
        LoadTimeSalesData m.strSym, m.nSymID
        LoadVolData m.strSym, m.nSymID
        LoadTradeData m.strSym, m.nSymID
        tmr.Enabled = True
        m.bInitInprog = False
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesAnalyzer.RefreshData"

End Sub

Private Sub LoadGridNoVol(ByVal strSym$)
On Error GoTo ErrSection:

    Dim i&
    
    Me.Caption = kCaptionBase & " for " & strSym
    
    With fgVol
        .Rows = fgVol.Rows + 1
        .MergeRow(fgVol.Rows - 1) = True
        For i = 0 To .Cols - 1
            .TextMatrix(.Rows - 1, i) = kFootPrintNoVol
        Next
    End With
            
    Exit Sub
    
ErrSection:
    RaiseError "frmTimeSalesAnalyzer.LoadGridNoVol"

End Sub

