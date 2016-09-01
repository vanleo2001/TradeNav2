VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptimizer 
   Caption         =   "Optimizer"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11910
   Icon            =   "frmOptimizer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   11910
   Begin vsOcx6LibCtl.vsElastic vsElastic2 
      Height          =   450
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5610
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   794
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
      Begin HexUniControls.ctlUniRichTextBoxXP rtfColDesc 
         Height          =   210
         Left            =   165
         TabIndex        =   6
         Top             =   105
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   370
         BackColor       =   -2147483648
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmOptimizer.frx":014A
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
         Tip             =   "frmOptimizer.frx":016A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":018A
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
      Begin HexUniControls.ctlUniLabelXP sbField 
         Height          =   255
         Index           =   5
         Left            =   8535
         Top             =   90
         Width           =   4095
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
         Caption         =   "frmOptimizer.frx":01A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptimizer.frx":01C6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":01E6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   7
         X1              =   8490
         X2              =   8490
         Y1              =   75
         Y2              =   370
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   7
         X1              =   8490
         X2              =   12690
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   7
         X1              =   8490
         X2              =   12690
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   7
         X1              =   12690
         X2              =   12690
         Y1              =   75
         Y2              =   375
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   3
         X1              =   8415
         X2              =   8415
         Y1              =   60
         Y2              =   360
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   6
         X1              =   105
         X2              =   8415
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   6
         X1              =   90
         X2              =   90
         Y1              =   60
         Y2              =   360
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   6
         X1              =   90
         X2              =   8415
         Y1              =   60
         Y2              =   60
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP Corner 
      Height          =   375
      Left            =   11475
      TabIndex        =   0
      Top             =   6225
      Visible         =   0   'False
      Width           =   1335
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
      Caption         =   "frmOptimizer.frx":0202
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmOptimizer.frx":022E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmOptimizer.frx":024E
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic1 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6060
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   873
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
      Begin VSFlex7LCtl.VSFlexGrid vsStatusBar 
         Height          =   255
         Left            =   135
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1365
         _cx             =   5080
         _cy             =   5080
         _ConvInfo       =   1
         Appearance      =   1
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
         BackColor       =   8421504
         ForeColor       =   16777215
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8421504
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   8421504
         GridColor       =   -2147483639
         GridColorFixed  =   -2147483639
         TreeColor       =   -2147483632
         FloodColor      =   16711680
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin HexUniControls.ctlUniLabelXP sbField 
         Height          =   255
         Index           =   4
         Left            =   10470
         Top             =   120
         Width           =   2160
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
         Caption         =   "frmOptimizer.frx":026A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptimizer.frx":028A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":02AA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP sbField 
         Height          =   255
         Index           =   3
         Left            =   8535
         Top             =   120
         Width           =   1740
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
         Caption         =   "frmOptimizer.frx":02C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptimizer.frx":02E6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":0306
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP sbField 
         Height          =   255
         Index           =   2
         Left            =   6660
         Top             =   120
         Width           =   1710
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
         Caption         =   "frmOptimizer.frx":0322
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptimizer.frx":0342
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":0362
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP sbField 
         Height          =   255
         Index           =   1
         Left            =   3945
         Top             =   105
         Width           =   2550
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
         Caption         =   "frmOptimizer.frx":037E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptimizer.frx":039E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":03BE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP sbField 
         Height          =   255
         Index           =   0
         Left            =   1710
         Top             =   120
         Width           =   2040
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
         Caption         =   "frmOptimizer.frx":03DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOptimizer.frx":03FA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOptimizer.frx":041A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   5
         X1              =   75
         X2              =   75
         Y1              =   90
         Y2              =   390
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   5
         X1              =   10410
         X2              =   12705
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   4
         X1              =   10410
         X2              =   10410
         Y1              =   90
         Y2              =   420
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   6
         X1              =   12690
         X2              =   12690
         Y1              =   75
         Y2              =   405
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   5
         X1              =   10410
         X2              =   12675
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   5
         X1              =   10335
         X2              =   10335
         Y1              =   90
         Y2              =   420
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   4
         X1              =   8505
         X2              =   10335
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   4
         X1              =   8490
         X2              =   10335
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   2
         X1              =   8490
         X2              =   8490
         Y1              =   90
         Y2              =   400
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   4
         X1              =   8415
         X2              =   8415
         Y1              =   90
         Y2              =   420
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   3
         X1              =   6615
         X2              =   6615
         Y1              =   90
         Y2              =   400
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   3
         X1              =   6615
         X2              =   8430
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   3
         X1              =   6615
         X2              =   8430
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   2
         X1              =   1560
         X2              =   1560
         Y1              =   90
         Y2              =   420
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   2
         X1              =   90
         X2              =   1545
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   2
         X1              =   75
         X2              =   1560
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   1
         X1              =   3885
         X2              =   3885
         Y1              =   90
         Y2              =   400
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   1
         X1              =   3885
         X2              =   6555
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   1
         X1              =   6540
         X2              =   6540
         Y1              =   90
         Y2              =   420
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   1
         X1              =   3885
         X2              =   6525
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line BottomLine 
         BorderColor     =   &H80000009&
         Index           =   0
         X1              =   1650
         X2              =   3840
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line RightLine 
         BorderColor     =   &H80000009&
         Index           =   0
         X1              =   3810
         X2              =   3810
         Y1              =   75
         Y2              =   405
      End
      Begin VB.Line TopLine 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   1620
         X2              =   3810
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Line LeftLine 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   1620
         X2              =   1620
         Y1              =   90
         Y2              =   400
      End
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic3 
      Height          =   5610
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   9895
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
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      Appearance      =   0
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   1
      GridCols        =   2
      _GridInfo       =   $"frmOptimizer.frx":0436
      Begin ActiveToolBars.SSActiveToolBars tbToolbar 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131083
         ToolBarsCount   =   1
         ToolsCount      =   13
         DisplayContextMenu=   0   'False
         Tools           =   "frmOptimizer.frx":047B
         ToolBars        =   "frmOptimizer.frx":15C4
      End
      Begin VSFlex7LCtl.VSFlexGrid vsTests 
         Height          =   5430
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   11730
         _cx             =   20690
         _cy             =   9578
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
         BackColorAlternate=   12648447
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuUseValues 
         Caption         =   "&Use Values"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "&Reports"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export to File"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOptimizer.frm
'' Description: Form for displaying an optimization run
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/16/2011   DAJ         Added code for the Highlight Bar Reporter
'' 06/22/2011   DAJ         Changed "# Days" to "# Bars", fixed tooltip on headers
'' 06/27/2011   DAJ         If HighlightBar Reporter, show chart after optimization done
'' 05/01/2013   DAJ         Shadow Trading
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Optimizer As cOptimizer
    bAbort As Boolean
    lTotalIterations As Long
    bDrillDown As Boolean
    bInProgress As Boolean
    bExporting As Boolean
    
    strSystemName As String
    strSymbol As String
    strSymbolDesc As String
    strSystemType As String
    strDataType As String
    strDeveloper As String
    bRanUnloadCode As Boolean
    
    RptBridges As cGdTree
    astrTradeFiles As cGdArray
    
    strReportName As String
    bPyramid As Boolean
    hTradeFiles As Long
    hTblRptRules As Long
    
    Mode As eGDOptMode
    oGame As cGameMode      'game object for doing reports etc. for game results
End Type
Private m As mPrivate

Property Let SystemName(ByVal pData As String)
    m.strSystemName = pData
End Property
Property Get SystemName() As String
    SystemName = m.strSystemName
End Property

Property Let Symbol(ByVal pData As String)
    m.strSymbol = pData
End Property
Property Get Symbol() As String
    Symbol = m.strSymbol
End Property

Property Let SymbolDesc(ByVal pData As String)
    m.strSymbolDesc = pData
End Property
Property Get SymbolDesc() As String
    SymbolDesc = m.strSymbolDesc
End Property

Property Let SystemType(ByVal pData As String)
    m.strSystemType = pData
End Property
Property Get SystemType() As String
    SystemType = m.strSystemType
End Property

Property Let DataType(ByVal pData As String)
    m.strDataType = pData
End Property
Property Get DataType() As String
    DataType = m.strDataType
End Property

Property Let Developer(ByVal pData As String)
    m.strDeveloper = pData
End Property
Property Get Developer() As String
    Developer = m.strDeveloper
End Property

Property Let Abort(ByVal bAbort As Boolean)
    m.bAbort = Abort
End Property
Property Get Abort() As Boolean
    Abort = m.bAbort
End Property

Public Property Get TestRuleID() As Long
    TestRuleID = m.Optimizer.TestRuleID
End Property
Public Property Let TestRuleID(ByVal lRuleID As Long)
    m.Optimizer.TestRuleID = lRuleID
End Property

Public Property Get TestRunMode() As eGDRunMode
    TestRunMode = m.Optimizer.TestRunMode
End Property
Public Property Let TestRunMode(ByVal eTestRunMode As eGDRunMode)
    m.Optimizer.TestRunMode = eTestRunMode
End Property

Public Property Get InProgress() As Boolean
    InProgress = m.bInProgress
End Property

Public Property Get OptMode() As eGDOptMode
    OptMode = m.Mode
End Property
Public Property Let OptMode(ByVal nMode As eGDOptMode)
    m.Mode = nMode
End Property

'Called by the EngineCalled Back routine (controlled by Engine).  This starts
'up the form and initializes testing...
Public Function Init(pTotalIterations As Long, pSystemInputs As cGdArray, _
    Optional Mode As eGDOptMode = eGDOptMode_Optimization, _
    Optional GameObj As cGameMode = Nothing, Optional ByVal bGuruMode As Boolean = False) As Boolean
On Error GoTo ErrSection:

    m.lTotalIterations = pTotalIterations
    m.Mode = Mode
    Set m.oGame = GameObj

    SetFormCaption
    
    Select Case m.Mode
        Case eGDOptMode_Optimization, eGDOptMode_GameMode
            Me.Icon = Picture16(ToolbarIcon("kOptimizer"), , True)
            vsElastic2.Visible = True
            tbToolbar.Tools("ID_UseValues").ChangeAll ssChangeAllName, "&Use Values"
            tbToolbar.Tools("ID_Chart").Visible = True
            tbToolbar.Tools("ID_MergedReports").Visible = False
        
        Case eGDOptMode_MultipleRun, eGDOptMode_StrategyBasket
            Me.Icon = Picture16(ToolbarIcon("ID_StrategyBaskets"), , True)
            vsElastic2.Visible = False
            tbToolbar.Tools("ID_UseValues").ChangeAll ssChangeAllName, "Edit S&trategy"
            tbToolbar.Tools("ID_UseValues").Visible = Not bGuruMode
            tbToolbar.Tools("ID_Chart").Visible = False
            tbToolbar.Tools("ID_MergedReports").Visible = True
        
        Case eGDOptMode_HighlightBarReport
            Me.Icon = Picture16(ToolbarIcon("kHBReporter"), , True)
            vsElastic2.Visible = False
            tbToolbar.Tools("ID_UseValues").Visible = False
            tbToolbar.Tools("ID_Chart").Visible = True
            tbToolbar.Tools("ID_MergedReports").Visible = False
    
    End Select
        
    ' Hide the System Manager and show this form non-modally
    ''frmSystemManager.Optimizing = True
    InfBox ""
    ShowForm Me, False, , , ALT_GRID_ROW_COLOR
    ''SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    'Set flag to incidate if at least one system was drilled down to.  When
    'leaving the optimizer, prompt to restore the optimization from/to/step
    'settings
    m.bDrillDown = False
    
    'Enable "Abort"...
    m.bInProgress = True
    EnableButtons False
    SetBtn "ID_Stop", True
    m.bAbort = False
    Init = True
    
    'Initialize optimization (including test results grid)
    Set m.Optimizer = New cOptimizer
    With m.Optimizer
        .Mode = Mode
        .vsTests = vsTests
        .TotalIterations = pTotalIterations
        .SystemInputs = pSystemInputs
        .Init
    End With
    
    Me.Refresh
    
ErrExit:
    Exit Function

ErrSection:
    Init = False
    ''EndOptimization
    RaiseError "frmOptimizer.Init", eGDRaiseError_Show
    EndOptimization
    Resume ErrExit:

End Function

'The Engine calls this after each system test is complete.  The test#,
'Trades array and inputs used in the test are passed...
Public Function Add(ByVal pIteration As Long, pTrades As cGdArray, _
    pParmValues As cGdArray) As Integer
On Error GoTo ErrSection:
    
    Dim rc    As Boolean
    Dim nGameModeRc As Long
    
    'User pressed "Stop" button.  Abort optimization:
    ' - Filter/resort current tests
    ' - Cleanup "EndOptimization"
    m.Optimizer.FormatVisibleRows
    DoEvents
    If m.bAbort Then
        Add = 1     'Engine needs to shut down gracefully here
        m.Optimizer.RefreshColumns
        EndOptimization
        Exit Function
    End If
    
    If m.Mode = eGDOptMode_GameMode Then
        'JM 11-03-2011: fix for instant replay issues 6475, 6476
        'rc is of type boolean as declared above
        'the optimizer object returns numeric value of 0, 1 or 2
        'with 2 being an error condition, the gamemode object never gets the 2
        'by using a numeric to receive the optimizer object's return code
        'the error condition gets passed to the gamemode allowing it exit the processing loop
        nGameModeRc = m.Optimizer.Add(pIteration, pTrades, pParmValues)
        Add = nGameModeRc
        
        If nGameModeRc = kSN_OPTIMIZATION_COMPLETED Or nGameModeRc = kSN_OPTIMIZATION_ERROR Then
            EndOptimization         'don't think need this, but just in case
        End If
        
    Else
        rc = m.Optimizer.Add(pIteration, pTrades, pParmValues)
        Select Case rc
            Case kSN_OPTIMIZATION_IN_PROGRESS
                Add = kSN_OPTIMIZATION_IN_PROGRESS
                
            Case kSN_OPTIMIZATION_COMPLETED
                EndOptimization
                
            Case kSN_OPTIMIZATION_ERROR
                Add = kSN_OPTIMIZATION_ERROR  'Engine needs to shut down gracefully here
                EndOptimization
                
        End Select
        m.Optimizer.FormatVisibleRows
    End If
 
ErrExit:
    Exit Function

ErrSection:
    Add = kSN_OPTIMIZATION_ERROR
    EndOptimization
    RaiseError "frmOptimizer.Add", eGDRaiseError_Show
    Resume ErrExit:

End Function

Private Sub EndOptimization()
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    EnableButtons True
    SetBtn "ID_Stop", False
    vsStatusBar.Visible = False
    Screen.MousePointer = vbDefault
    m.bInProgress = False
    
    If OptMode = eGDOptMode_HighlightBarReport Then
        ShowChart
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.EndOptimization"
    
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
    RaiseError "frmOptimizer.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText$
    Dim strFont As String
    
    Me.Icon = Picture16(ToolbarIcon("ID_News"), , True)
    
    g.Styler.StyleForm Me
    
    With tbToolbar
        '.Tools("ID_Stop") = Picture16(ToolbarIcon("kStop"))
        .Tools("ID_UseValues").Picture = Picture16(ToolbarIcon("ID_Strategies"))
        .Tools("ID_Reports").Picture = Picture16(ToolbarIcon("ID_Performance"))
        .Tools("ID_MergedReports").Picture = Picture16(ToolbarIcon("ID_Performance"))
        .Tools("ID_Settings").Picture = Picture16(ToolbarIcon("ID_Settings"))   'want new toolbar to use kSettings for consistency
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Export").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Cancel").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_Chart").Picture = Picture16(ToolbarIcon("kBarChart"))
        .Tools("ID_Colored").Picture = Picture16(ToolbarIcon("kMarketDepth"))
        .Tools("ID_Colored").Enabled = False
        .Tools("ID_Colored").Visible = False
    End With
    
    strText = GetIniFileProperty("Optimizer", "", "Placement", g.strIniFile)
    If strText = "" Then
        ReSizeMDIChildForm Me, Corner
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LHTW"
    End If
    
    Me.Refresh
    
    Screen.MousePointer = vbDefault
        
    EnableButtons True
    SetBtn "ID_Stop", False
    
    ''If FormIsLoaded("frmSystemManager") Then
    ''    frmSystemManager.WindowState = vbMinimized
    ''End If
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("Optimizer", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsTests.Font, strFont
    
    Set m.astrTradeFiles = New cGdArray
    m.astrTradeFiles.Create eGDARRAY_Strings
    
    Set m.RptBridges = New cGdTree
    
    m.bExporting = False
    
    vsTests.BackColorAlternate = ALT_GRID_ROW_COLOR
    
ErrExit:
    Exit Sub

ErrSection:
    EndOptimization
    RaiseError "frmOptimizer.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetBtn(ByVal strButton As String, ByVal bOn As Boolean)
On Error GoTo ErrSection:

    tbToolbar.Tools(strButton).Enabled = bOn

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.SetBtn", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    ' TLB 8/16/2000: make user abort first before closing form
    If UnloadMode = 0 Then
        If tbToolbar.Tools("ID_Stop").Enabled Then
            Beep
            Cancel = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    SetIniFileProperty "Optimizer", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "Optimizer", FontToString(vsTests.Font), "Fonts", g.strIniFile
    
    m.Optimizer.SaveColumns
    Set m.Optimizer = Nothing
    ''frmSystemManager.Optimizing = False

    ''If FormIsLoaded("frmSystemManager") Then
    ''    frmSystemManager.WindowState = vbNormal
    ''End If
    
    If FormIsLoaded("frmOptChart") Then Unload frmOptChart
    
    For lIndex = 0 To m.astrTradeFiles.Size - 1
        KillFile m.astrTradeFiles(lIndex)
    Next lIndex
    Set m.astrTradeFiles = Nothing
    
    For lIndex = m.RptBridges.Count To 1 Step -1
        Set m.RptBridges(lIndex) = Nothing
    Next lIndex
    Set m.RptBridges = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub EnableButtons(ByVal bOn As Boolean)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    With tbToolbar
        .Tools("ID_Stop").Enabled = bOn
        If m.Mode = eGDOptMode_GameMode Then
            .Tools("ID_UseValues").Enabled = False
            .Tools("ID_Delete").Visible = True
            .Tools("ID_Rename").Visible = True
        Else
            .Tools("ID_UseValues").Enabled = bOn
            .Tools("ID_Delete").Visible = False
            .Tools("ID_Rename").Visible = False
        End If
        .Tools("ID_Reports").Enabled = bOn
        .Tools("ID_MergedReports").Enabled = bOn
        .Tools("ID_Settings").Enabled = bOn
        .Tools("ID_Print").Enabled = bOn
        .Tools("ID_Export").Enabled = bOn
        .Tools("ID_Toolbox").Enabled = bOn
        .Tools("ID_Cancel").Enabled = bOn
        'Disable chart button for strategy runs without Rules or Inputs
        'until further discussion with Tim. Aardvark issue 866.
        If bOn = True Then
            'check for "I-" or "R-" columns
            .Tools("ID_Chart").Enabled = False
            For lIndex = 0 To vsTests.Cols - 1
                If IsOptimizedColumn(lIndex) Then
                    .Tools("ID_Chart").Enabled = True
                    Exit For
                End If
            Next lIndex
        Else
            .Tools("ID_Chart").Enabled = bOn
        End If
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.EnableButtons", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Private Sub Export()
On Error GoTo ErrSection:

    Dim X As Long
    Dim lRow As Long
    Dim lCol As Long
    Dim strBuffer As String
    Dim strFileName As String
    Dim fh As Integer
    Dim sb As New cStatusBar
    
    m.bExporting = True
        
    strFileName = CommonDialogFile(frmMain.CommonDialog1, True, "CSV Files (*.csv)|*.csv")
    fh = FreeFile
    
    If Len(strFileName) = 0 Then GoTo ErrExit           '6001
    
    Open strFileName For Output As #fh
    
    Screen.MousePointer = vbHourglass
    With vsTests
        If .Rows > 1000 Then
            sb.StatusBarControl = vsStatusBar
            vsStatusBar.Visible = True
            SetBtn "ID_Stop", True
        End If
        m.bAbort = False
        
        If m.Mode = eGDOptMode_HighlightBarReport Then
            Print #fh, "Highlight Bar Report"
        Else
            Print #fh, "Strategy Optimization Report"
            If m.Mode = eGDOptMode_Optimization Then
                Print #fh, "Strategy: " & g.CurrentSystem.SystemName
            End If
        End If
        
        For X = 0 To .Cols - 1
            If IsOptimizedColumn(X, False) Then
                Print #fh, m.Optimizer.ColumnDesc(X)
            End If
        Next X
        
        For lRow = 0 To .Rows - 1
            If m.bAbort Then Exit For
            If .Rows > 1000 Then sb.Value = ((lRow + 1) / .Rows) * 100
            If lRow Mod 100 = 0 Then DoEvents
            
            strBuffer = ""
            If .RowHidden(lRow) = False Then
                For lCol = 0 To .Cols - 1
                    If .ColHidden(lCol) = False Then
                        strBuffer = strBuffer & Chr(34) & .Cell(flexcpTextDisplay, lRow, lCol) & Chr(34) & ","
                    End If
                Next lCol
                If Len(strBuffer) > 0 Then
                    strBuffer = Left(strBuffer, Len(strBuffer) - 1)
                    Print #fh, strBuffer
                End If
            End If
        Next lRow
    End With
    Close #fh

ErrExit:
    m.bAbort = False
    SetBtn "ID_Stop", False
    m.bExporting = False
    vsStatusBar.Visible = False
    Set sb = Nothing
    Screen.MousePointer = 0
    Exit Sub

ErrSection:
    m.bAbort = False
    SetBtn "ID_Stop", False
    m.bExporting = False
    Set sb = Nothing
    Screen.MousePointer = 0
    RaiseError "frmOptimizer.Export", eGDRaiseError_Raise

End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsTests

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuExport_Click()
On Error GoTo ErrSection:

    Export

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.mnuExport.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "SNV Optimizer", frmOptimizer, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.mnuPrint.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuReports_Click()
On Error GoTo ErrSection:

    m.Optimizer.ShowReport

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.mnuReports.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuSettings_Click()
On Error GoTo ErrSection:

    If frmOptCustomize.ShowMe Then m.Optimizer.RefreshColumns

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.mnuSettings.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuUseValues_Click()
On Error GoTo ErrSection:

    If m.Mode = eGDOptMode_Optimization Then
        m.bDrillDown = True
        m.Optimizer.SaveTest
        Unload Me
    Else
        m.Optimizer.LoadSystem
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.mnuUseValues.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strFile$
    
    Select Case Tool.ID
        Case "ID_Stop"
            m.bAbort = True
            DoEvents
            
        Case "ID_UseValues"
            If m.Mode = eGDOptMode_Optimization Then
                m.bDrillDown = True
                m.Optimizer.SaveTest
                Unload Me
            Else
                m.Optimizer.LoadSystem
            End If
            
        Case "ID_Reports"
            ShowReport
            
        Case "ID_MergedReports"
            ShowMergedReports
            
        Case "ID_Settings"
            If frmOptCustomize.ShowMe Then m.Optimizer.RefreshColumns
            
        Case "ID_Print"
            frmPrintPreview.ShowMe "SNV Optimizer", frmOptimizer, 0
        
        Case "ID_Export"
            Export
            
        Case "ID_Toolbox"
            frmToolbox.ShowMe
            
        Case "ID_Chart"
            ShowChart
            
        Case "ID_Cancel"
            Me.Hide
            Unload Me
            
        Case "ID_Rename"        'game mode only
            If Not m.oGame Is Nothing Then
                strFile = vsTests.TextMatrix(vsTests.Row, 0)
                strFile = m.oGame.RenameResult(strFile)
                If Len(strFile) > 0 Then
                    vsTests.TextMatrix(vsTests.Row, 0) = strFile
                End If
            End If
            
        Case "ID_Delete"        'game mode only
            If Not m.oGame Is Nothing Then
                strFile = vsTests.TextMatrix(vsTests.Row, 0) & ".txt"
                
'JM 11-04-2011: delete & sort issue 6475
'   this is a virtual grid, just clear grid & let all the supproting table, arrays etc. also clear
'   gamemode object will reload game results when it deletes a result file
                vsTests.FlexDataSource = Nothing
                Set m.Optimizer = Nothing
                m.oGame.DeleteResult strFile

'JM 11-04-2011: original code, leave awhile then remove if all ok
'                vsTests.RowHidden(vsTests.Row) = True
            End If
            
        Case "ID_Colored"
            If Not m.Optimizer Is Nothing Then
                m.Optimizer.ColorSymbols = Tool.State
            End If
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsTests_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:
    
    vsTests.FlexDataSource = m.Optimizer
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.vsTests.AfterMoveColumn", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsTests_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)

    m.Optimizer.FormatVisibleRows
    
End Sub

'Check for column resorting
Private Sub vsTests_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsTests
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow = 0 And Button = vbLeftButton Then
            Cancel = True
            m.Optimizer.SortOnCol lMouseCol
        ElseIf Button = vbRightButton And lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
            
            If Not tbToolbar.Tools("ID_Stop").Enabled Then
                PopupMenu mnuPopUp
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.vsTests.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'Always keep column 1 where it is
Private Sub vsTests_BeforeMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo ErrSection:
    
    If Col = 0 Then Position = 0
    If Col <> 0 And Position = 0 Then Position = Col
    
    m.Optimizer.SortOnCol -1 ' set back to previous sort
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.vsTests.BeforeMoveColumn", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub vsTests_BeforeScrollTip(ByVal Row As Long)
On Error GoTo ErrSection:
    
    m.Optimizer.BeforeScrollTip Row
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.vsTests.BeforeScrollTip", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub vsTests_DblClick()
On Error GoTo ErrSection:
    
    If Not m.bInProgress Then
        ShowReport
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.vsTests_DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub vsTests_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsTests
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
    
        If lMouseRow >= 0 And lMouseCol >= 0 Then
            If lMouseRow = 0 Then
                .ToolTipText = SORT_BY_PREFIX & Trim(Parse(m.Optimizer.ColumnDesc(lMouseCol), "=", 2))
            ElseIf IsOptimizedColumn(lMouseCol, False) Then
                .ToolTipText = m.Optimizer.ColumnDesc(lMouseCol)
            Else
                .ToolTipText = ""
            End If
        Else
            .ToolTipText = ""
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.vsTests.MouseMove", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsTests_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    If vsTests.Row >= vsTests.FixedRows And vsTests.Row < vsTests.Rows - 1 Then
        vsTests.RowSel = vsTests.Row
    End If
    m.Optimizer.vsTests_AfterRowColChange OldRow, OldCol, NewRow, NewCol
    'If vsStatusBar.Visible Then vsStatusBar.SetFocus
    
ErrExit:
    Exit Sub

ErrSection:
    EndOptimization
    RaiseError "frmOptimizer.vsTests.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Variant set of arguments from the Print Preview control
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim X As Integer                    ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strText As String               ' Text to add to the printer control
    
    With frmPrintPreview.vp
        .StartDoc
        
        ' Header and Footer
        DoPrintHeader
        
        ' Report Heading and date/time...
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 14
        .TextAlign = taCenterMiddle
        
        If m.Mode = eGDOptMode_HighlightBarReport Then
            .Text = "Highlight Bar Report" & vbLf & vbLf
        Else
            .Text = "Strategy Optimization Report" & vbLf & vbLf
        End If
        
        .TextAlign = taLeftMiddle
        .Font.Size = 12
        .Font.Bold = False
        
        ' Blank line after printing heading, then draw a line under title
        .DrawLine .MarginLeft, .CurrentY, .PageWidth - .MarginRight, .CurrentY
        
        ' System name and iteration descriptions
        If m.Mode = eGDOptMode_Optimization Then
            .Text = vbLf & "Strategy: " & g.CurrentSystem.SystemName & vbLf
        ElseIf m.Mode = eGDOptMode_HighlightBarReport Then
            .Text = vbLf & "Enter " & Parse(g.CurrentSystem.HighlightBarInfo, "|", 1)
            .Text = " When: " & Parse(g.CurrentSystem.HighlightBarInfo, "|", 2) & vbLf & vbLf
        End If
        For X = 0 To vsTests.Cols - 1
            If IsOptimizedColumn(X, False) Then
                .Text = m.Optimizer.ColumnDesc(X) & vbLf
            End If
        Next X
        .Text = vbLf
        
        ' Draw line between header information and report
        .DrawLine .MarginLeft, .CurrentY, .PageWidth - .MarginRight, .CurrentY
        .Text = vbLf
        
        ' If printing or previewing show the grid, otherwise if printing to file
        ' walk through the grid outputing tab delimeted strings
        If frmPrintPreview.GoingToFile = False Then
            .RenderControl = vsTests.hWnd
        Else
            With vsTests
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If .ColHidden(lCol) = False Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.GenerateReport", eGDRaiseError_Raise
    Resume ErrExit

End Sub

Public Sub StopRun()
On Error GoTo ErrSection:
        
    m.bAbort = True
    m.Optimizer.RefreshColumns
    EndOptimization

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.StopRun", eGDRaiseError_Raise
    
End Sub

Private Sub ShowReport()
On Error GoTo ErrSection:

    Dim strFile$

    Select Case m.Mode
    Case eGDOptMode_GameMode
        strFile = m.oGame.ResultFilePath & vsTests.TextMatrix(vsTests.Row, 0) & ".txt"
        m.oGame.ShowGameReport strFile
    Case eGDOptMode_StrategyBasket, eGDOptMode_MultipleRun
        ShowMergedReports True
    Case Else
        m.Optimizer.ShowReport
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.ShowReport"
End Sub

Private Sub ShowMergedReports(Optional ByVal bEvenIfJustOne As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim RptBridge As New cRptBridge     ' Report bridge to pass information to the reports
    Dim astrFiles As New cGdArray       ' Files to process for the reports

    Screen.MousePointer = vbHourglass
    For lIndex = m.RptBridges.Count To 1 Step -1
        If m.RptBridges(lIndex).IsLoaded = False Then
            m.RptBridges.Remove lIndex
        End If
    Next lIndex
    
    Set RptBridge = New cRptBridge
    Set astrFiles = m.Optimizer.SelectedFiles
    
    If astrFiles.Size > 1 Or bEvenIfJustOne Then
        mSysNav.ShowMergedReports RptBridge, m.strReportName, m.bPyramid, astrFiles.ArrayHandle, m.hTblRptRules
    Else
        mSysNav.ShowMergedReports RptBridge, m.strReportName, m.bPyramid, m.hTradeFiles, m.hTblRptRules
    End If
    m.RptBridges.Add RptBridge
    Screen.MousePointer = vbDefault

ErrExit:
    Set astrFiles = Nothing
    Exit Sub
    
ErrSection:
    Set astrFiles = Nothing
    RaiseError "frmOptimizer.ShowMergedReports", eGDRaiseError_Raise
    
End Sub

Public Sub SetUpMergedRun(ByVal strReportName As String, ByVal bPyramid As Boolean, ByVal hTradeFiles As Long, Optional ByVal hTblRptRules As Long = 0&)
On Error GoTo ErrSection:

    m.strReportName = strReportName
    m.bPyramid = bPyramid
    
    m.astrTradeFiles.CopyFromHandle hTradeFiles
    m.hTradeFiles = m.astrTradeFiles.ArrayHandle
    
    m.hTblRptRules = hTblRptRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.SetUpMergedRun", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetFormCaption
'' Description: Set the form caption based on the mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormCaption()
On Error GoTo ErrSection:

    Select Case m.Mode
        Case eGDOptMode_Optimization
            Caption = "Optimizer"
        Case eGDOptMode_MultipleRun
            Caption = "Strategy Basket Run"
        Case eGDOptMode_StrategyBasket
            Caption = "Strategy Basket Run"
        Case eGDOptMode_GameMode
            Caption = "Game Results"
        Case eGDOptMode_HighlightBarReport
            Caption = "Highlight Bar Report"
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.SetFormCaption"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOptimizedColumn
'' Description: Determine if the given column is an optimized column
'' Inputs:      Column, From Grid?
'' Returns:     True if optimized Column, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsOptimizedColumn(ByVal lColumn As Long, Optional ByVal bFromGrid As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strColumnName As String         ' Column name
    
    bReturn = False
    If bFromGrid Then
        strColumnName = UCase(vsTests.TextMatrix(0, lColumn))
    Else
        strColumnName = UCase(Parse(m.Optimizer.ColumnDesc(lColumn), "=", 1))
    End If
    
    If (m.Mode = eGDOptMode_HighlightBarReport) And (strColumnName = "# BARS") Then
        bReturn = True
    ElseIf (Left(strColumnName, 2) = "I-") Or (Left(strColumnName, 2) = "R-") Then
        bReturn = True
    End If
    
    IsOptimizedColumn = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptimizer.IsOptimizedColumn"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowChart
'' Description: Show the chart form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowChart()
On Error GoTo ErrSection:
    
    frmOptChart.ShowMe vsTests, m.Optimizer, sbField(1), sbField(4)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptimizer.ShowChart"
    
End Sub

