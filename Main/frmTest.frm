VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test 1"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   9435
   Begin HexUniControls.ctlUniTextBoxXP txtZip 
      Height          =   285
      Left            =   4380
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTest.frx":0000
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
      Tip             =   "frmTest.frx":002A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTest.frx":004A
   End
   Begin HexUniControls.ctlUniCheckXP chkTest 
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5640
      Width           =   855
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
      Caption         =   "frmTest.frx":0066
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483634
      Pressed         =   0   'False
      Tip             =   "frmTest.frx":008E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmTest.frx":00AE
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Width           =   1275
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
      Tip             =   "frmTest.frx":00CA
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmTest.frx":00EA
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   5400
   End
   Begin HexUniControls.ctlUniListBoxXP lst 
      Height          =   1425
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   2955
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmTest.frx":0106
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmTest.frx":0126
      ManualStart     =   0   'False
      Columns         =   0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
      Caption         =   "frmTest.frx":0142
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTest.frx":016E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTest.frx":018E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP Command10 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4560
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
         Caption         =   "frmTest.frx":01AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":01E2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":0202
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command9 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4140
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
         Caption         =   "frmTest.frx":021E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":025C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":027C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command8 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3600
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
         Caption         =   "frmTest.frx":0298
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":02D2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":02F2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command7 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3180
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
         Caption         =   "frmTest.frx":030E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":0340
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":0360
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command6 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2640
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
         Caption         =   "frmTest.frx":037C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":03B8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":03D8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command5 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2220
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
         Caption         =   "frmTest.frx":03F4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":0430
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":0450
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command4 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
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
         Caption         =   "frmTest.frx":046C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":049E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":04BE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command3 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1260
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
         Caption         =   "frmTest.frx":04DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":050E
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":052E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
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
         Caption         =   "frmTest.frx":054A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":057E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":059E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
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
         Caption         =   "frmTest.frx":05BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTest.frx":05EE
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTest.frx":060E
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4335
      Left            =   3120
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   5415
      _cx             =   9551
      _cy             =   7646
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
      BackColorAlternate=   13695231
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
      Cols            =   14
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
   Begin vsOcx6LibCtl.vsElastic vseDetach 
      Height          =   240
      Left            =   1800
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Symbol Link"
      Top             =   5460
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ForeColor       =   16777215
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Picture         =   "frmTest.frx":062A
      Caption         =   ""
      Align           =   0
      Appearance      =   1
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
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
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   375
      Left            =   5640
      Top             =   6000
      Visible         =   0   'False
      Width           =   3495
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
      Caption         =   "frmTest.frx":0784
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTest.frx":07D4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTest.frx":07F4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TAS_IndicatorInit Lib "TASIndicators.dll" (ByVal strPgm$, ByVal strFunc$, ByVal strSymbol$) As Long
Private Declare Function TAS_IndicatorSetParameter Lib "TASIndicators.dll" (ByVal nFuncID&, ByVal nParmID&, ByVal dParmValue#) As Long
Private Declare Function TAS_IndicatorSetBar Lib "TASIndicators.dll" (ByVal nFuncID&, ByVal nBar&, ByVal nYYYYMMDD&, ByVal nHHMM&, ByVal dOpen#, ByVal dHigh#, ByVal dLow#, ByVal dClose#, ByVal dVol#, ByVal dOI#) As Long
Private Declare Function TAS_IndicatorValue Lib "TASIndicators.dll" (ByVal nFuncID&, ByVal nReturnID&) As Double


'int DumpBySymbol( gdArrayStr *inMessage ) -- pass symbol as 1st string
Private Declare Function DumpBySymbol Lib "SalmonClient.dll" (ByVal hStrings As Long) As Long

'Private Declare Function JMAUT Lib "JRS_UT.dll" (ByVal dSeries#, ByVal dSmooth#, ByVal dPhase#, ByVal pdOutput&, ByVal iDestroy&, piSeriesID&, ByVal iSameBar&) As Long
Private Declare Function JMAUT Lib "JRS_UT.dll" (ByVal dSeries#, ByVal dSmooth#, ByVal dPhase#, pdOutput#, ByVal iDestroy&, piSeriesID&, ByVal iSameBar&) As Long

'BOOL WINAPI SetProcessDEPPolicy(__in  DWORD dwFlags);
Private Declare Function SetProcessDEPPolicy Lib "kernel32" (ByVal dwFlags As Long) As Long
Private Declare Function GetProcessDEPPolicy Lib "kernel32" (ByVal hProcess As Long, dwFlags As Long, dwPerm As Long) As Long
Private Declare Function GetSystemDEPPolicy Lib "kernel32" () As Long

Private Sub Command1_Click()
On Error GoTo ErrSection:
    
    Dim i&, j&, n&, s$, d#, d2#, s2$, strTemp$, strText$, strSymbol$, strMsg$
    Dim b As New cGdBars, b2 As New cGdBars
    Dim nSymbolID&, nDate&
    Static bToggle As Boolean
    
    Dim dCurTime#
    Static dPrevTime#

    Dim mb As New cMemBuffer
    Dim aBytes() As Byte
    Dim aStrings() As String
    Dim aFiles As New cGdArray
    
    Dim bb As Boolean
    bb = True
    
    Dim a As New cGdArray
    Dim t As New cGdTable
    Dim Tree As cGdTree
    Dim frm As Form
    Dim frmC As frmChart
    
    Dim TimeInfo As TradeTimeInfo
    Dim strFeedSymbol$, strFeedExchange$
    
    Dim aHwnd() As Long
    
    Dim F As Single
      
    Dim dHoursSoFar#, iDone&, iLeft&, iTotal&, dHoursLeft#
    iTotal = 50
    For i = 0 To iTotal - 1
        iDone = i
        iLeft = iTotal - iDone
        dHoursSoFar = i * 0.1
        If iDone > 0 Then
            dHoursLeft = dHoursSoFar * iLeft / iDone
            AddList Format(dHoursLeft, "#0.00")
        End If
    Next
Exit Sub

    s = "28fe085400361ae2d63c41752566cdf1af2f0d3e379ae8c7984769610416703f299113648be66b799e777bc53acfe57c55dd"
    s = DecryptFromHex(s)
    AddList s
    
    For i = Date To Date + 365
        If gdIsHoliday(i, "NKPGMJLTX") <> 0 Then
            AddList "Holiday = " & DateFormat(i)
        End If
    Next
    
    s = g.SymbolPool.GetSymbolsForGroup("All Stocks.GRP")
    a.SplitFields s, vbTab
    i = a.Size
    
    'GetCont67Group
    AddList Str(PhysicalRAM(True)) & vbTab & Str(PhysicalRAM(False))
Exit Sub

    s = "@AllTradeable.GRP"
    s = g.SymbolPool.GetSymbolsForGroup(s)
    a.SplitFields s, vbTab
    For i = 0 To a.Size - 1
        AddList a(i)
    Next
    AddList "count = " & Str(a.Size)

Exit Sub

    txtZip.Visible = True
    txtZip.ZOrder
    DoEvents
    MoveFocus txtZip
        
    SendKeys "abc"
Exit Sub
    
    's = Trim(GetRegistryValue(rkLocalMachine, "Software\Trader Workstation", "jtspath", ""))
    s = Trim(GetRegistryValue(rkClassesRoot, "VBS", "", ""))
    If InStr(UCase(GetRegistryValue(rkClassesRoot, "VBS", "", "")), "VBx") > 0 Then
        i = Len(s)
    End If
    

Exit Sub
    ChkDJIA
Exit Sub
    
    s = "es-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
    s = "bo-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
    s = "cc-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
    s = "f-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
    s = "fb-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
    s = "uc-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
    s = "xyz-067"
    AddList s & vbTab & GetFuturesCategory(s)
    
Exit Sub
    
    FindEhlersData
    'ChkQM
    s = WindowsVersionStr
Exit Sub
       
    s = "www.TradeNavigator.com/Analysis/default.aspx?U=*&P=*&T=Seas"
    s = "https://www.tradenavigator.com/seasonalsweetspots/?U=*&P=*"
    s = "https://www.tradenavigator.com/stockscreener/?U=*&P=*"
    
    s = "28fe085400361ae2d63c41752566cdf1af2f0d3e379ae8c7984769610416703f299113648be66b799e777bc53acfe57c55dd"

    s = EncryptToHex(s)
    s = DecryptFromHex(s)
       
    GetCont67Group
Exit Sub
       
    s = ""
    For i = 0 To 133
        s = s & "IFF (&n = " & Str(i) & " , &p" & Str(i) & ") Or "
    Next
    FileFromString "c:\temp.txt", s
       
    'BuildDataFix
       
Exit Sub

    'SetAppColors
    'CheckForTradeNavMessages
    s = "�"
    aFiles(0) = s
    s = aFiles(0)
    AddList s
    
    DM_GetBars b, "ES-067", "1min", Date - 3, Date
    i = b.Size
    b2.BuildBars "5min", b.BarsHandle
    i = b2.Size
       
Exit Sub

    s = GetAllDrives
    AddList s
    For i = 1 To Len(s)
        s2 = Mid(s, i, 1) & ":"
        d = GetDiskFreeSpace(s2)
        d2 = GetDiskSize(s2)
        n = GetDriveType(s2)
        AddList s2 & " = type " & Str(n) & ", " & Str(d) & " bytes free out of " & Str(d2)
    Next
    AddList "done"
    
    's = mMain.GenesisCDInDrive(True)
    
    'i = MoveFiles("c:\Xttm\* /s /g=5000", "c:\test\")
    'AddList "CopyFiles = " & Str(i)
    
Exit Sub

    For i = 0 To 25
        s = Chr(Asc("A") + i)
        If DirExist(s & ":\") Then
            AddList s
        End If
    Next
    AddList "done"
Exit Sub
       
    s = "NQ-067"
    nSymbolID = GetSymbolID(s)
    SetBarProperties b, s
    'b.Prop(eBARS_Periodicity) = ePRD_Minutes + 15
' CRASHES:
bb = DM_GetTickBars(g.DMS, nSymbolID, 15, DateSerial(2000, 4, 1), DateSerial(2000, 7, 1), b.BarsHandle)
    AddList Str(b.Size) & " bars"
    'ChkSetBarProps
Exit Sub

    s = "QQQ"
    DM_GetBars b, s
    Set t = TAS_CalcMarketMap(b, 70, 40, 50)
    s = t.ToString(, , True)
    FileFromString "c:\MarketMap.txt", s
    
    Set t = TAS_CalcMarketMap(b, 40, 70, 50)
    s = t.ToString(, , True)
    FileFromString "c:\MarketMap2.txt", s
    
    'CreateTickFiles
    
Exit Sub
    
    'CheckProfileBars
    'CheckFunctions
    'TestTAS
    
    s = "EX-067"
    AddList g.RealTime.GetSymbolOffsetForGenesisRT(s, True)
    
    s = "Please wait until the Microsoft .NET dependencies have been installed,|then hit 'Continue' ..."
    If InfBox(s, "i", "+Continue|-Quit", ".NET Framework dependency check") = "Q" Then
    End If

Exit Sub
    
    GetVisionMargins
    
Exit Sub
    s = "We recommend you PAUSE the Historical Data Downloader while data streaming is turned on."
    'InfBox s, "i", , "Realtime Streaming"
    
    s = "Link in the downloaded historical data now?||(or can select 'Install Data' under 'File' menu later)"
    If InfBox(s, "?", "+OK|-Not Now", "Data Install", , 10) = "O" Then
    End If
    
    s = "www.TradeNavigator.com/DataInst/DataInst.txt"
    's = "www.TradeNavigator.com/DataInst.htm"
    's = GetWebPageData(s)
    'FileFromString "c:\d.cfg", s
    
    'AddList "start"
    's = frmDataInstall2.GetFtpDataset
    'AddList "cfg = " & s
    
Exit Sub

    If 1 Then
        s = "AAPL"
        n = 20140813
    Else
        s = "ZB-067"
        n = 0
    End If
    Set aFiles = mDataNav.GetOptionChainBidAskData(s, n)
    lst.Clear
    For i = 0 To aFiles.Size - 1
        AddList aFiles(i)
    Next
Exit Sub

    DM_GetBars b, "nq-067", , Date - 7
    d = GetPrevCloseForQB(b)
    AddList b.PriceDisplay(d)
Exit Sub
    'CheckCRC
    'CheckFractZen
    
    'GetCont67Group
    
    strSymbol = "ES-067"
    AddList g.FractZen.GetSpeedInfo(strSymbol)
    g.FractZen.SetSpeedInfo strSymbol, 14
    AddList g.FractZen.GetSpeedInfo(strSymbol)
    g.FractZen.SetSpeedInfo strSymbol
    AddList g.FractZen.GetSpeedInfo(strSymbol)
    
    d = gdTickCount
    n = LastDailyDownload
    For i = 1 To 100
        s = Str(g.FractZen.GetFractZenRange(strSymbol, n))
        's = g.FractZen.GetSpeedInfo(strSymbol)
    Next
    d = gdTickCount - d
    AddList "FZ = " & s & ", ms = " & Str(Round(d))
Exit Sub
    
    DM_GetBars b, "DX-067", "d", Date - 5, Date
    i = b.Size
    i = b.Prop(eBARS_FractZen)
    b.Prop(eBARS_FractZen) = 1
    i = b.Prop(eBARS_FractZen)
    DM_GetBars b, "DX-067", "30b", Date - 5, Date
    i = b.Prop(eBARS_FractZen)
    DM_GetBars b, "DX-067", "30m", Date - 5, Date
    i = b.Prop(eBARS_FractZen)
    
    DM_GetBars b, "DX-067", "FractZen", Date - 10, Date
    
Exit Sub
    
    'GetCont67Group
    strSymbol = "GX-067"
    SetBarProperties b, strSymbol
    d = g.Profit.Profit(strSymbol, b.Prop(eBARS_TickMove))
    AddList strSymbol & " = $" & Str(d)

    strSymbol = "$EUR-USD"
    SetBarProperties b, strSymbol
    d = g.Profit.Profit(strSymbol, b.Prop(eBARS_TickMove), 100000)
    AddList strSymbol & " = $" & Str(d)

    strSymbol = "$USD-JPY"
    SetBarProperties b, strSymbol
    d = g.Profit.Profit(strSymbol, b.Prop(eBARS_TickMove), 100000)
    AddList strSymbol & " = $" & Str(d)

    strSymbol = "$CAD-CHF"
    SetBarProperties b, strSymbol
    d = g.Profit.Profit(strSymbol, b.Prop(eBARS_TickMove), 100000)
    AddList strSymbol & " = $" & Str(d)

Exit Sub
    
    Set frm = ActiveChart
    frm.Chart.TestFZ
    Set frm = Nothing
Exit Sub
    MakeMidCmd
    i = g.RealTime.GetSecondsBehind
    AddList Str(i) & " seconds behind"
Exit Sub
    
    'DM_GetBars b, "NQ1-067", "15 min"
    AddList Str(b.Size)
    CheckContHistory
Exit Sub
    'ChkFormIteration
    s = "http://www.paycheckcity.com/coadp4/netpaycalculator.asp"
    s = GetWebPageData(s)
    FileFromString "c:\test.htm", s
Exit Sub
    
    'ChkPitRolls
    'CheckSectors
    AddList "starting"
    frmSectorTree.ShowMe "", True
    AddList "done", True
Exit Sub
    DM_GetBars b, "US-067"
    For i = DateSerial(2013, 1, 1) To DateSerial(2014, 1, 5)
        If b.IsHoliday(i) Then
            AddList DateFormat(i)
        End If
    Next
'Exit Sub
    
    SetBarProperties b, "CC9-067"
    i = b.Prop(eBARS_CsiNumber)
        
    s = "SSAIX"
    i = GetSymbolID(s)
    n = g.SymbolPool.PoolRecForSymbolID(i)
    SetBarProperties b, s
    DM_GetBars b, s
    AddList s & vbTab & Str(i) & vbTab & Str(n) & vbTab & Str(g.SymbolPool.SecType(n)) & vbTab & Str(b.Size)
    
    'Set b = LoadYahooStockHist("AA")
    
    'AddList Str(Forms.Count) & " forms"
        
    'IntradayMarketScope
    'MinutizeHistory
    'CheckContHistory
    'RunSAI
    'CheckOldStockHist
    'CheckSectors
    CheckEtfTP
    If IsIDE Then
    '    frmSaiElite.ShowMe
    End If
    'CheckForexPips
       
Exit Sub

    If 0 Then
        n = 0
    Else
        s = AddSlash(App.Path) & "Test.LOG"
        'n = FreeFile
        'Open s For Append As #n
        n = gdFileOpen(s, "a+")
    End If
    d = gdTickCount
    For i = 1 To 10 * 150&
        s = "A new test message #" & Str(i) & " for logging purposes."
        DumpDebugLog "A", s, n
        'DumpDebugLog "B", s
        'DumpDebugLog "C", s
        'DumpDebugLog "D", s
        'DumpDebugLog "E", s
    Next
    d = gdTickCount - d
    AddList Str(d) & " ms"
    If n <> 0 Then
        'Close n
        gdFileClose n
    End If
    
Exit Sub
    
    s = "Decrypted from Hex"
    s = mGdDll.EncryptToHex(s)
    FileFromString "c:\EncryptedHex.txt", s
Exit Sub
    
    AddList "=============================="
    n = GetWindow(frmMain.hWnd, GW_CHILD)
    LockWindowUpdate n
    d = gdTickCount
    For i = 1 To 10
        Set frmC = New frmChart
        frmC.Chart.TemplateLoad "Default"
        frmC.Chart.FromDate = Date + 7
        frmC.Chart.SetSymbol "$DJIA"
        n = frmC.Chart.Bars.Size
        frmC.Show
    Next
    AddList Str(gdTickCount - d) & " ms"
    LockWindowUpdate 0
    DoEvents
    AddList Str(gdTickCount - d) & " ms"
Exit Sub
    AddList "=============================="
    
    Set Tree = GetChartsInZorder
    For i = 1 To Tree.Count
        Set frm = Tree(i)
        AddList Str(i) & " = " & frm.Caption
    Next
    Set Tree = Nothing
        
Exit Sub

    GetWindowHandles aHwnd, 0
    For i = LBound(aHwnd) To UBound(aHwnd)
        n = aHwnd(i)
        s = vbGetWindowText(n)
        AddList Str(n) & "  " & s
        'If Left(Trim(s), 3) = "ES-" Then
        If InStr(s, "-067") > 0 Then
            n = n
        End If
    Next
    
Exit Sub
    
    AddList "=============================="
    n = GetWindow(frmMain.hWnd, GW_CHILD)
    s = vbGetWindowText(n)
    AddList Str(n) & "  " & s
    
    n = GetWindow(n, GW_CHILD)
    
    'n = frmMain.hWnd
    n = GetWindow(0, GW_CHILD)
    n = GetTopWindow(frmMain.hWnd)
    n = GetTopWindow(n)
    n = GetTopWindow(0)
    Do While n <> 0
        s = vbGetWindowText(n)
        AddList Str(n) & "  " & s
        If Left(Trim(s), 3) = "ES-" Then
            n = n
        End If
        'n = GetWindow(n, GW_HWNDNEXT)
        n = GetNextWindow(n, GW_HWNDNEXT)
    Loop
    
Exit Sub
    
       
    s = "43218606050069ab630cf48154d184da6203c2df97478d22e19b6fe09c14799c61ece26934be4d0e9096304ec9f12e36b1425bae557c4a41b1c9"
    s = FixURL(s)
    
    'Set t = GetPivotFarmData("ES-067", "debug")
    't.Serialize "c:\temp\test.gdt", True
    
    If 0 Then
        Set t = GetRollsTable("ES-067")
        For i = 0 To t.NumRecords - 1
            d = t(2, i)
        Next
    End If
    
Exit Sub
    
    SetBarProperties b, "$GBP-JPY"
    d = b.Prop(eBARS_TickValue)
    d = b.Prop(eBARS_TickMove)
    
    SetBarProperties b, "ZC1-067"
    i = b.Prop(eBARS_StartTime)
    
    s = "HE-201302-S1"
    bb = IsSpreadSymbol(s)
Exit Sub
    'g.RealTime.Reconnect 0, True
    
    i = 41258
    SetBarProperties b, "ZC1-067"
    'i = DM_GetTickBars(g.DMS, GetSymbolID("ZC1-067"), 30, i - 6, i, b.BarsHandle)
    DM_GetBars b, "ZC1-067", "30 min", -420, LastDailyDownload
    i = b.Size
    
Exit Sub
    If 0 Then
        
        s = "http://www.TradeNavigator.com/SharedChartPages/index.php?U=*&P=*" 'U=4fc83ae5&P=a791367fcfa151"
        s = s & "&E=,SCP_WAYNE,SCP_BRANDT,"
        s = FixURL(s)
        RunWebReport "Shared Chart Pages", s, "", 0
    Else
        PublishSharedChartPage
    End If

Exit Sub
    
    s = App.Path & "\charts\pages\Shared Test2.gzp"
    s = FileToString(s, , , True)
    'mb.FromFile s
    's = mb.GetStr
    If Len(s) > 0 Then
        s = EncryptToHex(s)
        FileFromString "c:\Shared Test.SCP", s
    End If
    
Exit Sub
    
    i = 0
    SetupGrid fg, eGridMode_Grid
    With fg
        .Cols = 15
        For i = 1 To .Cols - 1
            .TextMatrix(0, i) = "Column " & Str(i)
        Next
        For j = 1 To .Rows - 1
            .TextMatrix(j, 0) = "Row " & Str(j)
            For i = 1 To .Cols - 1
                .TextMatrix(j, i) = Str(j) & " , " & Str(i)
            Next
        Next
        '.AllowSelection = True
        '.DragMode = 0
        '.ScrollBars = flexScrollBarVertical
        '.ScrollTrack = True
        .Visible = True
        .ZOrder
    End With
    
    'SetAppBackColor 8421504, True
    'SetAppBackColor RGB(64, 72, 80), True
    
    'SetAppBackColor &H806040, True
    
    'SetAppBackColor RGB(50, 60, 70), True

Exit Sub

If 1 Then
    j = 20120817
    
    s = "ES-067"
    s = "BAC"
    n = GetSymbolID(s)
    j = Date
    bb = BuildProfileBars(b, n, j, j)
    i = b.Size
    
    j = Date - 1
    'bb = BuildProfileBars(b, n, 0, 0)
    i = b.Size
Else
    s = "ES-067"
    d = gdTickCount
    j = 20120816
    bb = BuildProfileBars(b, GetSymbolID(s), 20120813, 20120822)
    d = gdTickCount - d
    i = b.Size
End If

    a.Size = 0
    For i = 0 To b.Size - 1
        s = BarDisplay(b, i)
        a.Add s
        'AddList s
    Next
    a.ToFile "c:\export\Profile.txt"
    i = b.Size
    
Exit Sub
    
    For j = DateSerial(2012, 8, 16) To DateSerial(2012, 8, 21)
        If IsWeekday(j) Then
            DM_GetBars b, "ES-067", "Each tick", j, j
            s = "c:\Export\ES-" & Format(j, "YYYYMMDD") & ".tck"
            b.SerializeDataArrays s, True
        End If
    Next
    i = 0
Exit Sub
    
    s = "BAC"
    DM_GetBars b, s, "30 min", Date - 90
    
    If GetVolumeIterators(b) Then
        For i = 0 To b.Size - 1
            If b(eBARS_Flags, i) > 0 Then
                AddList Str(i) & vbTab & Str(b(eBARS_Vol, i))
            End If
        Next
    End If
Exit Sub
    If b2.BuildSMPBars(b, 12, 9) Then
        For i = 0 To b2.Size - 1
            AddList BarDisplay(b2, i)
        Next
    End If
    
    
Exit Sub
    
    DM_GetBars b, "FB", "Daily"
    AddList "Daily = " & Str(b.Size) & " bars"
    For i = 0 To b.Size - 1
        AddList BarDisplay(b, i)
    Next
    
    DM_GetBars b, "FB", "5 min"
    AddList "5 min = " & Str(b.Size) & " bars"
    For i = 0 To b.Size - 1
        AddList BarDisplay(b, i)
    Next
    
    DM_GetBars b, "FB", "Each Tick"
    AddList "Each Tick = " & Str(b.Size) & " bars"
    For i = 0 To b.Size - 1
        If i < 10 Or i >= b.Size - 10 Then
            AddList BarDisplay(b, i)
        End If
    Next
    
    'SetAppBackColor RGB(90, 105, 120), True
    
    'frmSaiReport.ShowMe
Exit Sub
    
    s = Format(RI_GetDataServiceID, "#000000:000")
    AddList s
    s = Format(RI_GetDataServiceID / 1000, "#000000") & ":" & Format(RI_GetDataServiceID Mod 1000, "000")
    AddList s
    i = 0
Exit Sub
    'SetAppBackColor RGB(224, 232, 240), True
    'SetAppBackColor RGB(96, 128, 160), True
    'SetAppBackColor RGB(96, 176, 216), True
    'SetAppBackColor RGB(64, 176, 192), True
    SetAppBackColor RGB(90, 105, 120), True
    
    'SetAppBackColor RGB(112, 128, 144), True
    'SetAppBackColor RGB(120, 124, 136), True
    'SetAppBackColor &HE0D0C0    ', True
    'SetAppBackColor RGB(212, 224, 236) ', True
    
Exit Sub
    
Exit Sub
        
    Do While True
        s = InfBox("symbol", "?", , "Enter symbol", , , , , , "s")
        If Len(s) = 0 Then Exit Do
        AddList s & vbTab & BaseForAutoExitFavorites(s)
    Loop
Exit Sub

    frmMain.PlaySound ".\Provided\BullConfirm.wav"
    frmMain.PlaySound ".\Provided\BullPTP.wav"
    frmMain.PlaySound ".\Provided\BearConfirm.wav"
    frmMain.PlaySound ".\Provided\BearPTP.wav"
    AddList "done"
Exit Sub

    fg.Visible = True
    fg.ZOrder
    fg.Cell(flexcpBackColor, 0, 1, 0, 1) = vbGreen
    fg.Cell(flexcpBackColor, 1, 1, 1, 1) = vbGreen
Exit Sub

    DM_GetBars b, "es-067"
    If DM_CalendarSpread(a, b, 1) Then
        AddList Str(a.Size)
    End If
        
Exit Sub
        
    
    For j = 0 To g.SymbolPool.NumRecords - 1
        i = g.SymbolPool.SymbolID(j)
        s = g.SymbolPool.SymbolForID(i)
        DM_GetBars b, s, 0, 20030101, 20030228
        If b.Size > 0 Then
            d = gdMinValue(b.ArrayHandle(eBARS_Low), 0, b.Size - 1)
            If d < 75 And d > 73 Then
                d2 = gdMaxValue(b.ArrayHandle(eBARS_High), 0, b.Size - 1)
                If d2 > 86 And d2 < 95 Then
                    AddList s
                End If
            End If
        End If
        If j Mod 1000 = 0 Then
            AddList Str(j)
            DoEvents
        End If
    Next
    AddList "done"

Exit Sub
    i = 0
    'SetAppBackColor RGB(216, 224, 238) ', True
    SetAppBackColor RGB(224, 232, 240) ', True
    
    'SetAppBackColor RGB(255, 240, 200) ', True
    
    'SetAppBackColor RGB(216, 228, 242) ', True
    'SetAppBackColor RGB(180, 214, 224) ', True
        
Exit Sub
    If TranslateSymbol("ABC@ASX", "S", "S", strFeedSymbol, strFeedExchange, d, i, j, False) Then
        AddList strFeedSymbol
    End If
    
Exit Sub
        
    AddList "1 min bars from DM_GetBars"
    DM_GetBars b, "zc-067", "1 min", 20101202, 20101202
    For i = 0 To 10
        AddList DateFormat(b(eBARS_DateTime, i), MM_DD_YYYY, HH_MM_SS) & vbTab & b.PriceDisplay(b(eBARS_Open, i)) & vbTab & b.PriceDisplay(b(eBARS_High, i)) & vbTab & b.PriceDisplay(b(eBARS_Low, i)) & vbTab & b.PriceDisplay(b(eBARS_Close, i))
    Next

    AddList "Each tick"
    DM_GetBars b2, "zc-067", "Each tick", 20101202, 20101202
    For i = 0 To 10
        AddList DateFormat(b2(eBARS_DateTime, i), MM_DD_YYYY, HH_MM_SS) & vbTab & b2.PriceDisplay(b2(eBARS_Close, i))
    Next
    AddList "1 min bars from ticks"
    b.Size = 0
    b.BuildBars "1 min", b2.BarsHandle
    For i = 0 To 10
        AddList DateFormat(b(eBARS_DateTime, i), MM_DD_YYYY, HH_MM_SS) & vbTab & b.PriceDisplay(b(eBARS_Open, i)) & vbTab & b.PriceDisplay(b(eBARS_High, i)) & vbTab & b.PriceDisplay(b(eBARS_Low, i)) & vbTab & b.PriceDisplay(b(eBARS_Close, i))
    Next

Exit Sub
    lst.Clear
    TestPriceCluster

Exit Sub
        
    s = "80110a832c1a086a2def98d664de553883698541d2e61f6aae927ef026aebdaa6f9b43d2e5cb1cf1"
    s = DecryptFromHex(s)
    AddList s
    
    s = "8eaffef9b44ec41e5e99a602c0c881b8a1729f0a7a886983301081c353f9804deb6a35eecda8eec32793b8b85b8ecc628c56"
    s = DecryptFromHex(s)
    AddList s
Exit Sub
    
    'SetAppBackColor RGB(214, 218, 222)
    'SetAppBackColor RGB(224, 224, 216)
    
    'SetAppBackColor RGB(120, 155, 180)
    'SetAppBackColor RGB(96, 128, 160), True
    'SetAppBackColor RGB(96, 176, 216), True
    
    'SetAppBackColor RGB(220, 228, 240) ', True
    'SetAppBackColor RGB(64, 176, 192), True
    
    'SetAppBackColor RGB(90, 135, 190), True
SetAppBackColor RGB(200, 216, 236) ', True
    'SetAppBackColor RGB(125, 158, 192) ', True

   
    'InfBox "test this", , "test"
    
Exit Sub
    a(3) = 77
    aFiles(2) = 88
    t.AttachField a
    t.AttachField aFiles
    'i = PFP_IndicatorMatches(t.TableHandle, 5, 15)
Exit Sub

    s = "A:ES-201006\tES-201009\tES-201006 C1500\tES-201006 P1500\crB:ES-\crR:CL3-201006\tCL3-201009\tCL3-201006 C1500\cr"
    s = Replace(s, "\t", vbTab)
    s = Replace(s, "\cr", vbCrLf)
    'SyncOptNavSymbolsWithSalmon s, True
Exit Sub
    s = "MSFT 20100506 P40.0"
    's = "ES-201004 C850"
    's = "MSFT a P40.0"
    Set aFiles = frmSymbolSelector.ShowMe(s, , , , , , True)
    s = aFiles(0)
    'i = ConvertStockOptionDatFile(App.Path & "\ftp\data.dat")
Exit Sub
    
    d = WindowsVersion
    AddList "WinVer = " & Str(d) & ", " & WindowsVersionStr
    AddList "98 = " & Str(Is9598orMe)
    AddList "XP = " & Str(IsAtLeastXP)
    AddList "Vista = " & Str(IsAtLeastVista)
    d = WindowsVersion
Exit Sub
    
    'Dim SymInfo As cSymbolInfo
    s = "http://football.fantasysports.yahoo.com/f1/524941/matchup?week=11&mid1=4&mid2=9"
    s = GetWebPageData(s)
    s = Replace(s, vbCrLf, Chr(10))
    s = Replace(s, Chr(10), vbCrLf)
    FileFromString "c:\temp\test.htm", s
    
    'aFiles.SplitFields s, Chr(10)
Exit Sub
    s = Replace(s, Chr(10), vbCrLf)
    FileFromString "c:\temp\test.htm", s
    
    s = FileToString("c:\dvlp\ffl\in4.htm")
    s = UCase(Replace(s, vbCrLf, ""))
    FileFromString "c:\temp\test2.htm", s
    
Exit Sub
    aFiles.SplitFields s, Chr(10) ' vbCrLf
    For i = aFiles.Size - 1 To 0 Step -1
        aFiles(i) = Trim(aFiles(i))
        If Len(aFiles(i)) = 0 Then
            aFiles.Remove i
        End If
    Next
    aFiles.ToFile "c:\temp\test.htm"
    
Exit Sub
    s = "smtp.email.msn.com"
    RunProcess App.Path & "\smtp.exe", Chr(34) & App.Path & "\Email.txt" & Chr(34) & " " & Chr(34) & s & Chr(34), True, vbMinimizedNoFocus, i
    AddList "ExitCode = " & Str(i)
Exit Sub
    
    s = Chr(34) & "Abd Xyz.gzp  " & Chr(34)
    AddList s & "|"
    s = Trim(StripStr(s, Chr(34)))
    AddList s & "|"

Exit Sub
    
    j = 1000
    d = gdTickCount
    For i = 1 To j
        s = "ES-200912 C" & Str(i)
        nDate = g.RealTime.SymbolInfo(s).LastTradedSession
    Next
    d = gdTickCount - d
    AddList "Created = " & Str(d)
    
    d = gdTickCount
    For i = 1 To j
        s = "ES-200912 C" & Str(i)
        g.RealTime.SymbolInfo s, True
    Next
    d = gdTickCount - d
    AddList "Removed = " & Str(d)
    
Exit Sub
    
    DM_GetBars b, "ES1-067", "each tick", 20090928, 20090929
    i = b.Size
    b.ArrayMask = eBARS_TickByTick
    b.DumpToFile "t:\temp\", , True
Exit Sub
    
    bToggle = Not bToggle
    d2 = gdTickCount
    
    s = "GRP:ALL STOCKS.GRP"
    If bToggle Then
        s2 = g.SymbolPool.GetSymbolsForGroup(s)
        d2 = gdTickCount - d2
        AddList "Num = " & Str(j) & ", Len = " & Str(Len(s2)) & ", ms = " & Str(Int(d2))
    Else
        j = g.SymbolPool.FieldNumForID(s)
        Set a = g.SymbolPool.ArrayTable.FieldArray(j)
        If Not a Is Nothing Then
            s2 = ""
            j = 0
            For i = 0 To a.Size - 1
                If a(i) > 0 Then
                    j = j + 1
                    s2 = s2 & g.SymbolPool.Symbol(i) & vbTab
                End If
            Next
            d2 = gdTickCount - d2
            AddList "Num = " & Str(j) & ", Len = " & Str(Len(s2)) & ", ms = " & Str(Int(d2))
        End If
    End If
   
Exit Sub
    
    s = "If you have another data disk to install from, please put it in now.  Then select 'Next Disk' ..."
    If InfBox(s, "?", "+Next Disk|-Finished", "Install Next Disk?") = "N" Then
        'GoTo DoNextDisk
    End If
Exit Sub

    s = "ES-067"
    d = g.RealTime.LastKnownPrice(s)
    AddList "LastKnowPrice for: " & s & vbTab & Str(d)
Exit Sub
    SetBarProperties b, "ibm"
    b.DumpToFile "c:\temp\"
Exit Sub
    
    s = "ES-067;D;10"
    GetPriceDataForOptNav s
Exit Sub
    
    s = "YI2-067"
    If SU_GetTimeInfo(g.SU, GetSymbolID(s), Date - 0, TimeInfo) <> 0 Then
        If TimeInfo.cFeedTime <> "N" Then
            i = TimeInfo.iLocalToGmtOffset
        End If
    End If
    
Exit Sub
   
    AddList RollSymbolForDate("ES-083", Date)
    AddList RollSymbolForDate("ES-067")
    
#If 0 Then
    For i = DateSerial(1950, 1, 1) To Date + 60
        s = RollSymbolForDate("ES-067", i)
        If s <> s2 Then
            AddList DateFormat(i) & vbTab & s
        End If
        s2 = RollSymbolForDate2("ES-067", i)
        If s <> s2 Then
            AddList "ERROR"
        End If
    Next
#End If
    
Exit Sub

    a.FromFile App.Path & "\Provided\Cont067.grp"
    For j = 0 To a.Size - 1
        s2 = Parse(a(j), vbTab, 2)
        If Len(s2) > 0 Then
            s2 = Parse(s2, "-", 1)
            s = s2 & ":"
            For i = 0 To 4
                s = s & vbTab & ConvertFutureSymbol(s2, i) & vbTab
            Next
            AddList s
        End If
    Next
Exit Sub
    
    strTemp = "It is recommended that you periodically Archive your work (settings, charts, library items, etc.).||Start the Archive program after shutting down?"
    'strTemp = "Would you like to Backup/Restore all of your custom settings and items you have created or modified?"
    'strTemp = "Start the Archive program after shutting down?"
    If InfBox(strTemp, "?", "+Exit Now|Archive|-Cancel", "Exit " & g.strTitle) = "A" Then
        RunProcess App.Path & "\TNArchive.exe", Chr(34) & g.strTitle & Chr(34)
    End If
    'd = ZipExecute("u", "t:\tradenav\DotNetSetup.gzp", "c:\temp\ii\", , True, True)
Exit Sub
    
    b.Size = 0
    d = gdMinValue(b.ArrayHandle(eBARS_Close), 0, b.Size - 1)
    
    DM_GetBars b, "ES-067", "each tick", LastDailyDownload, LastDailyDownload
    d = gdMinValue(b.ArrayHandle(eBARS_Close), 0, b.Size - 1)
    If b.Size > 0 Then
        If b.SerializeDataArrays("c:\export\test.bin", True) Then
            If b2.SerializeDataArrays("c:\export\test.bin", False) Then
                For i = 0 To b.Size - 1
                    If b(eBARS_DateTime, i) <> b2(eBARS_DateTime, i) Then
                        d = d
                    End If
                    If b(eBARS_Close, i) <> b2(eBARS_Close, i) Then
                        d = d
                    End If
                Next
            End If
        End If
    End If
    
Exit Sub
    
    SetDepOff
Exit Sub
    AddList Str(frmMain.ScaleHeight) & " twips," & Str(frmMain.ScaleHeight / Screen.TwipsPerPixelY) & " pixels"
Exit Sub
    
    SetBarProperties b2, "ES-067"
    b2.Prop(eBARS_StartTime) = 360
    b2.Prop(eBARS_EndTime) = 600
    DM_GetBars b2, "ES-067", "5 min", LastDailyDownload, LastDailyDownload
    
    DM_GetBars b, "ES-067", "Eachtick", LastDailyDownload + 1
    i = b.Size
    d = b(eBARS_DateTime, 0)
    d2 = b(eBARS_DateTime, b.Size - 1)
    
    b.Prop(eBARS_StartTime) = b2.Prop(eBARS_StartTime)
    b.Prop(eBARS_EndTime) = b2.Prop(eBARS_EndTime)
    i = b2.Size
    b2.BuildBars "5 min", b.BarsHandle, True
    d = b2(eBARS_DateTime, 0)
    d = b2(eBARS_DateTime, i - 1)
    d = b2(eBARS_DateTime, i)
    d2 = b2(eBARS_DateTime, b2.Size - 1)
    d = b2(eBARS_DownTicks, b2.Size - 2) + b2(eBARS_UpTicks, b2.Size - 2)
    d2 = b2(eBARS_DownTicks, b2.Size - 1) + b2(eBARS_UpTicks, b2.Size - 1)
    
    For i = 0 To b2.Size - 1
        AddList Format(b2(eBARS_DateTime, i), "MM/DD HH:NN") & vbTab & Str(b2(eBARS_Close, i)) & vbTab & Str(b2(eBARS_UpTicks, i) + b2(eBARS_DownTicks, i))
    Next
    
    'g.RealTime.CreateTransFile ' NO LONGER works for base symbols with no 199912 contract
Exit Sub
    
    d2 = DateSerial(2008, 11, 20) + TimeSerial(15, 59, 0)
    DM_GetBars b2, "SPY", "Eachtick", Int(d2), Int(d2)
    d = 0
    i = b2.FindDateTime(d2)
    For i = i To b2.Size
        d = d + 1
        If RoundToSecond(b2(eBARS_DateTime, i)) <> RoundToSecond(b2(eBARS_DateTime, i + 1)) Then
            AddList Format(RoundToSecond(b2(eBARS_DateTime, i)), "hh:mm:ss") & vbTab & Str(d)
            If RoundToSecond(b2(eBARS_DateTime, i + 1)) > d2 + 1 / 1440# Then
                Exit For
            End If
            d = 0
        End If
    Next
Exit Sub
    
    
    i = MonthNumber("may")
                InfBox "Your computer's Date, Time or Time Zone may need to be adjusted.  It is currently set to:||" _
                    & DateFormat(Now, MM_DD_YYYY, H_MM, AMPM_LOWER), "!", , "PLEASE CHECK"
    'ReadNYTime
Exit Sub
    'AskForActivate True
            
'Me.Caption = "TradeNavigator.com � 1999-" & Str(Year(Date)) & " All rights reserved"
    
    AddList "Sleep 10"
    Sleep 10
    Exit Sub
    
    s = "ES-200803"
    's = "GX-200803"
    's = "YI2-200803"
    gdResetProfiles 250, 299
    gdStartProfile 250
    'd = LastKnownPrice(s, True, True, dCurTime)
    gdStopProfile 250
    gdStartProfile 251
    d2 = g.RealTime.LastKnownPrice(s, , True, dPrevTime)
    If d2 <> d Then
        d = d
    ElseIf Abs(dPrevTime - dCurTime) > 0.0001 Then
        d = d
    End If
    gdStopProfile 251
    
    AddList s & vbTab & Str(d2) & vbTab & Format(dPrevTime, "hh:mm:ss")
    
    s = gdGetProfiles(250, 299)
    AddList s
    
    d = g.RealTime.LastKnownPrice("ES-200803")
    
Exit Sub
    
    i = b.Prop(eBARS_StartTime)
    SetBarProperties b, "IBM"
    i = b.Prop(eBARS_StartTime)
    b.Prop(eBARS_StartTime) = 600
    SetBarProperties b, "MSFT"
    i = b.Prop(eBARS_StartTime)
    SetBarProperties b, "US-067"
    i = b.Prop(eBARS_StartTime)
    
    DM_GetBars b, "IBM", "daily", Date - 7, Date
    i = b.Prop(eBARS_StartTime)
    i = b.Prop(eBARS_DefaultStartTime)
    b.Prop(eBARS_StartTime) = 600
    DM_GetBars b, "MSFT", "daily", Date - 7, Date
    DM_GetBars b, "US-067", "daily", Date - 7, Date
    
    DM_GetBars b, "IBM", "30m", Date - 7, Date
    i = b.Prop(eBARS_StartTime)
    i = b.Prop(eBARS_DefaultStartTime)
    b.Prop(eBARS_StartTime) = 600
    DM_GetBars b, "MSFT", "30m", Date - 7, Date
    DM_GetBars b, "US-067", "30m", Date - 7, Date
    

Exit Sub
    i = frmMain.apmRTClient.CreateMessage("GenesisRT", 22, "-1")
    If i <> 0 Then AddList "Storage OFF"
Exit Sub

    bToggle = Not bToggle
    'g.RealTime.SetRTPriority bToggle

    'i = UpdateDBConfig
    'i = IsFullTickDB
Exit Sub

    i = DM_GetBars(b, "es1-067", "Each Tick", Date, Date)
    b2.BuildBars "1000v", b.BarsHandle, True
    For i = 0 To b2.Size - 1
        AddList Str(i) & vbTab & Format(b2(eBARS_DateTime, i), "hh:mm") & vbTab & Str(b2(eBARS_Vol, i))
    Next

Exit Sub

        If frmMain.Enabled Then
            s = ""
            Set a = frmSymbolSelector.ShowMe("")
            If Not a Is Nothing Then
                For i = 0 To a.Size - 1
                    If Len(s) = 0 Then
                        s = a(i)
                    Else
                        s = s & vbTab & a(i)
                    End If
                Next
            End If
        Else
            s = "."
        End If
Exit Sub
    
    DM_GetBars b, "SP-067"
    d = gdTickCount
    Set b2 = b.CreateHeikinAshi
    d = gdTickCount - d
    For i = b.Size - 10 To b.Size - 1
        If b(eBARS_DateTime, i) = b2(eBARS_DateTime, i) Then
            AddList Str(i) & vbTab & Str(b(eBARS_Open, i)) & vbTab & Str(b(eBARS_High, i)) & vbTab & Str(b(eBARS_Low, i)) & vbTab & Str(b(eBARS_Close, i))
            AddList Str(i) & vbTab & Str(b2(eBARS_Open, i)) & vbTab & Str(b2(eBARS_High, i)) & vbTab & Str(b2(eBARS_Low, i)) & vbTab & Str(b2(eBARS_Close, i))
        End If
    Next
    AddList "ms = " & Str(d)
Exit Sub
    
    
    Me.Hide
    Sleep 2
    SetupBrokerLayout True
Exit Sub
    
    
Timer1.Enabled = True
Exit Sub
   
    d = gdTickCount
    Set aFiles = GetAllowedList("T")
    d = gdTickCount - d
    For i = 0 To aFiles.Size - 1
        AddList Str(i) & vbTab & aFiles(i)
    Next
    AddList Str(d) & " ms"
Exit Sub
    
    i = aFiles.GetMatchingFiles("c:\dvlp\gd\*.cpp /s", True, True, True, d)
    AddList Str(i) & " files = " & Str(d) & " bytes"
    For i = 0 To aFiles.Size - 1
        AddList Str(i) & vbTab & aFiles(i)
    Next
    
Exit Sub
    AddList Str(PhysicalRAM)
    AddList Str(PhysicalRAM(True))
Exit Sub
    
    d = gdTickCount
    If GetPredictionLabsData Then
        d = gdTickCount - d
        AddList "PredLabs data = " & Format(d / 1000, "#0.00") & " seconds"
    Else
        AddList "ERROR getting PredLabs data"
    End If
    
Exit Sub
    
    d = RoundToMinMove(809.4, 0.1)
    AddList Str(d - 809.4)
    
    frmNewAccount.ShowMe
Exit Sub
    
    s = "A1234"
    If IsAlpha(Trim(s)) Then
        InfBox "This is not a valid Genesis Customer ID.|Please contact Genesis Sales at 800-808-3282.", "!", , "Invalid Customer ID"
    End If

Exit Sub
  
    
    AddList GetSourceCode
    
    g.RealTime.EminiOffset = g.RealTime.EminiOffset + 0.25
    AddList "off = " & Str(g.RealTime.EminiOffset)
    Exit Sub
    
    Dim aShow As New cGdArray
      
    d = gdTickCount(False)
    DM_GetBars b, "ES-067", "60m", 20060426 ' 20060614
    d = gdTickCount(False) - d
    AddList Str(b.Size) & " bars, ms = " & Str(d)
    'i = b.Size

    d = gdTickCount(False)
    DM_GetBars b, "G6E-200609", "Each", 20060623, 20060623
    d = gdTickCount(False) - d
    AddList Str(b.Size) & " bars, ms = " & Str(d)
    For i = 0 To b.Size - 1
        If b(eBARS_DateTime, i) >= 38890.99 And b(eBARS_DateTime, i) <= 38891.02 Then
            AddList DateFormat(b(eBARS_DateTime, i), MM_DD_YYYY, HH_MM) & vbTab & Str(b(eBARS_Close, i)) _
                & vbTab & Str(b(eBARS_Vol, i)) & vbTab & Str(b(eBARS_UpTicks, i)) & vbTab & Str(b(eBARS_DownTicks, i))
        End If
    Next

Exit Sub

    d = gdTickCount(False)
nDate = FreeFile
Open "c:\test.txt" For Output As #nDate
    For i = 0 To b.Size - 1
        Print #nDate, Str(b(eBARS_Close, i))
    Next
Close #nDate
    d = gdTickCount(False) - d
    AddList Str(b.Size) & " bars XX, ms = " & Str(d)
      
    d = gdTickCount(False)
    i = 0
nDate = FreeFile
Open "c:\test.txt" For Input As #nDate
    Do While Not EOF(nDate)
        Line Input #nDate, s
        i = i + 1
    Loop
Close #nDate
    d = gdTickCount(False) - d
    AddList Str(i) & " bars ZZ, ms = " & Str(d)
      
Exit Sub
      
    For i = 1 To 100
        AddList Str(gdTickCount) & vbTab & Str(gdTickCount(False))
    Next
Exit Sub
      
    d = gdTickCount
    i = HeapCompact(GetProcessHeap, 0)
    AddList "HeapCompact = " & Format(i, "#,##0") & " -- time = " & Format(gdTickCount - d, "0.##")
    
      
    'strTemp = InfBox("You need to upgrade the program in order to connect to the real-time data stream.||Would you like to Upgrade now?", "?", "+Upgrade|-Not now", "Upgrade Required")
Exit Sub
    frmTTSummary.Caption = "TRADE CONSOLE  (info for live accounts is unofficial -- verify with your broker)"
Exit Sub
      
    If g.RealTime.UseBrokerFeed = 0 Then
        g.RealTime.UseBrokerFeed = eTT_AccountType_TransAct
        AddList "Transact feed"
    Else
        g.RealTime.UseBrokerFeed = 0
        AddList "GenesisRT feed"
    End If
    g.Transact.ResyncSubscriptionList
Exit Sub
      
      
    Timer1.Interval = 5000
    Timer1.Enabled = False
Exit Sub
    InfBox "test", , , "test", True
Exit Sub
    Set aShow = g.RealTime.BrokerSubscriptionList(eTT_AccountType_TransAct)
    For i = 0 To aShow.Size - 1
        AddList aShow(i)
    Next
    
    If g.RealTime.UseBrokerFeed = 0 Then
        g.RealTime.UseBrokerFeed = eTT_AccountType_TransAct
    Else
        g.RealTime.UseBrokerFeed = 0
    End If
    
Exit Sub
      
      
    strText = "The online agreement will be displayed in your browser.  After completing the agreement, a connection with Genesis is required in order to validate your new account enablements."
    strText = InfBox(strText, "i", "+Validate|-Abort", "Live Trading Setup")
    
    InfBox strText, "i", , "no wait", True
Exit Sub
      
    frmMain.tmrQuickStart.Enabled = True
    
    Sleep 12
    InfBox "Test", "i"
    
Exit Sub
      
    aShow.FromFile App.Path & "\Toolbar.sho"
    For i = aShow.Size - 1 To 0 Step -1
        Select Case UCase(aShow(i))
        Case "ID_SNAPSHOT", "ID_CHAIN"
            aShow.Remove i
        End Select
    Next
    aShow.ToFile App.Path & "\Toolbar.sho"
    ToolbarReset True
  
    'frmNewAccount.ShowMe
Exit Sub

    strTemp = "Please write down your new account| information for future reference:||Customer ID = " & "32100" & _
        "|Data Service ID = " & "001" & "|Password = " & "test123"
    InfBox strTemp, "!", , "New Account Information"
Exit Sub
    
    s = "ETA,GOLD"
    AddList s & vbTab & Str(HasModule(s))
    s = "-ETA,GOLD"
    AddList s & vbTab & Str(HasModule(s))
    s = "-ETAX,GOLD"
    AddList s & vbTab & Str(HasModule(s))
    s = "GOLD,-ETA"
    AddList s & vbTab & Str(HasModule(s))
    s = "GXOLD,-ETA"
    AddList s & vbTab & Str(HasModule(s))
Exit Sub
      
      
    If DM_GetBars(b, 41126, "daily", 23697, 38645) Then
        i = b.Size
        d = b(eBARS_DateTime, 0)
    End If
Exit Sub
      
    s = FileToString("c:\temp\longmsg.txt")
    InfBox s, "i", , "Long Msg"
Exit Sub
      
    s = "CL-067"
    i = g.SymbolPool.TickFirstDate(s)
    d = GetFirstTickDate(s)
    AddList Str(d) & vbTab & Str(i)
    Exit Sub
    
    SetBarProperties b, "SP-067"
    AddList b.MinMove
    AddList b.MinMove(DateSerial(1990, 1, 1))
    AddList b.MinMove(DateSerial(2000, 1, 1))
    Exit Sub
    
    s = "LP-199801"
    Do
        s = GetNextContract(s)
        AddList s
    Loop While Len(s) > 0
Exit Sub
    
    AddList Hex(GetSysColor(COLOR_3DHILIGHT))
    AddList Hex(GetSysColor(COLOR_3DLIGHT))
    AddList Hex(GetSysColor(COLOR_3DSHADOW))
    AddList Hex(GetSysColor(COLOR_3DDKSHADOW))
    AddList Hex(GetSysColor(COLOR_3DFACE))
  
    
Exit Sub
    
    
     
    AddList Str(g.RealTime.CanHaveMarketDepth("IBM"))
    AddList Str(g.RealTime.CanHaveMarketDepth("SP-067"))
    AddList Str(g.RealTime.CanHaveMarketDepth("MSFT"))
    AddList Str(g.RealTime.CanHaveMarketDepth("IBM"))
    Exit Sub
      
    AddList lst.Font.Name & "  " & Str(lst.Font.Size)
    AddList Str(MaxSymbolsAllowed)
    AddList Str(MaxSymbolsAllowed(True))
    Exit Sub
    
    AddList Command1.Font.Name
    AddList Me.lst.Font.Name
    AddList Me.Frame1.Font.Name
    
    Dim Font As StdFont
    Set Font = Command1.Font
    s = Font.Name
    Font.Name = "abcdx"
    s = Font.Name
    
    s = CheckSSFont
    AddList s
    
    Exit Sub
        
    's = "!$*,!*-*,!*@*"
    s = "!$#@-"
    's = "*@ASX,*@SGX"
    's = "[!$]*-*"
    
    s = "*@ASX,*@SGX,!$#@-"
    
   
Exit Sub
        
    i = GetSymbolID("IBM")
    d = g.SymbolPool.TickFirstDate("ibm")
    AddList d
    d = g.SymbolPool.TickLastDate(i)
    AddList d
    d = g.SymbolPool.EodFirstDate(i)
    AddList d
    d = g.SymbolPool.EodLastDate(i)
    AddList d
        
Exit Sub
    strTemp = "By default, Trade Navigator now displays| all trade times in your local time zone |(e.g. quote board, minute bars, trade reports).||However, this is a setting which can be changed on the 'Misc' tab of the Program Settings."
    InfBox strTemp, "i", , "Local Time Zone Display"
        
Exit Sub

DM_GetBars b, "IBM"
s = b.Prop(eBARS_ExchangeTimeZoneInf)
'B.Prop(eBARS_ExchangeTimeZoneInf) = "NY"
d = b(eBARS_DateTime, b.Size - 1)
d = b.DateTimeConvert(b.Size - 1)
AddList Format(CDate(d))
        
DM_GetBars b, "IBM", "30", Date - 7
s = b.Prop(eBARS_ExchangeTimeZoneInf)
d = b(eBARS_DateTime, b.Size - 1)
d = b.DateTimeConvert(b.Size - 1)
AddList Format(CDate(d))

DM_GetBars b, "$EUR-USD", "30", Date - 7
s = b.Prop(eBARS_ExchangeTimeZoneInf)
'B.Prop(eBARS_ExchangeTimeZoneInf) = "GMT"
d = b(eBARS_DateTime, b.Size - 1)
d = b.DateTimeConvert(b.Size - 1)
AddList Format(CDate(d))

DM_GetBars b, "EBI-067", "30", Date - 7
s = b.Prop(eBARS_ExchangeTimeZoneInf)
'B.Prop(eBARS_ExchangeTimeZoneInf) = "GMT"
d = b(eBARS_DateTime, b.Size - 1)
d = b.DateTimeConvert(b.Size - 1)
AddList Format(CDate(d))

Exit Sub
        
        
    g.RealTime.CreateTransFile
    Exit Sub
        
    s = "AUK BF" & vbTab & "38371.5" & vbTab & "2.75" & vbTab & "9"
    g.RealTime.TestMsg s
    Exit Sub
        
        
    If IsAlpha("1ab") Then
        i = i
    End If
    If IsAlpha("1ab", 1) Then
        i = i
    End If
    If IsAlpha("1ab", 2) Then
        i = i
    End If
        
        
    SetBarProperties b, "C-067"
    s = b.PriceDisplay(30.5)
    d = b.PriceFromString(s)
        
Exit Sub

mb.Buffer = "Test this"
If mb.FromFile("c:\ns.zip") Then
    If mb.ToFile("c:\ns2.zip") Then
        MsgBox "yes"
    End If
End If
Exit Sub
        
    AddList FormatNum(34)
    AddList FormatNum(34.2)
    AddList FormatNum(1234)
    AddList FormatNum(-1234.345)
    AddList FormatNum(-1234.3, 6, , True)
    AddList FormatNum(1.234567, 3)
    AddList FormatNum(1.234567, -3)
    
Exit Sub
    
    
    s = ChrW(66) & ChrW(193) & ChrW(174) & ChrW(202) & ChrW(226) & ChrW(41) & ChrW(161) & ChrW(143) & ChrW(118) & ChrW(61) & ChrW(156)
    s2 = Chr(66) & Chr(193) & Chr(174) & Chr(202) & Chr(226) & Chr(41) & Chr(161) & Chr(143) & Chr(118) & Chr(61) & Chr(156)
    
    s = ChrW(61) & ChrW(156)
    s2 = Chr(61) & Chr(156)
    
    If s <> s2 Then
        i = Len(s)
    End If
    
    s = ChrB(70) & ChrB(61) & ChrB(156)
    i = Len(s)
    i = LenB(s)

    s = "Test this"
    mb.Buffer = s
    aBytes = mb.Bytes
    i = UBound(aBytes)
    s = mb.Buffer
    mb.Bytes = aBytes
    s = mb.Buffer

Exit Sub
    
    mb.Bytes = aBytes
    
    aStrings = Split("test,this", ",")
    i = LBound(aStrings)
    i = UBound(aStrings)
    
    aStrings = Split("", ",")
    If IsNull(aStrings) Then
        i = i
    End If
    If IsEmpty(aStrings) Then
        i = i
    End If
    i = LBound(aStrings)
    i = UBound(aStrings)
    
    
    If IsNull(aBytes) Then
        i = i
    End If
    If IsEmpty(aBytes) Then
        i = i
    End If
    If IsNull(aBytes) Then
        i = i
    End If
    
    
    s = "Test this"
    aBytes = s
    i = aBytes(0)
    s = ""
    s = aBytes
    

    s = "Test this"
    mb.Buffer = s
    i = mb.Length
    s = mb.Buffer
    
    Set mb = New cMemBuffer
    aBytes = mb.Bytes
    
    If IsNull(aBytes) Then
        i = i
    End If
    If IsEmpty(aBytes) Then
        i = i
    End If
    mb.Bytes = aBytes
   
Exit Sub
    s = "test"
    dCurTime = gdTickCount * 1000
'dPrevTime = dCurTime
    If Len(s) > 0 Then
        AddList CStr(Int(dCurTime - dPrevTime)) & "  " & s
    End If
    dPrevTime = dCurTime
    
    'SaveChartPage "Save Test"
Exit Sub

    AddList CStr(bToggle)
    AddList Str(Now)
    AddList Str(Date)
    AddList Str(True)
    AddList Str(3000.8)
    
    DM_GetBars b, "sp-067", "100/3p"
    i = b.Size
    If b.Prop(eBARS_PeriodType) = ePRD_EodPF Then
        d = b.Prop(eBARS_PeriodsPerBar)
        i = b.Prop(eBARS_PeriodsPerBar) Mod 1000
        d = Int(b.Prop(eBARS_PeriodsPerBar) / 1000) * b.Prop(eBARS_TickMove) '* b.Prop(eBARS_MinMoveInTicks)
        
    End If
    
Exit Sub

    s = ChrW(130)
    If Asc(Chr(130)) <> 130 Then
        s = ""
    End If
    
    SaveChartPage "Save Test"
    Exit Sub
    
    's = "'Coast Trading Package' -- www.fibtrader.com -- �2003 CIS, Inc. All rights reserved worldwide."
    'FileFromString "c:\temp\c.txt", s
    
    Timer1.Enabled = Not Timer1.Enabled
    
    Exit Sub
    
    AddList "Start"
    i = g.RealTime.SymbolDelay("IBM")
    AddList CStr(i), True
    Exit Sub
    
    s = "eSignal"
    s = "c:\dvlp\genesis\RealTime\" & s _
        & "\" & s & "RT.exe"
    RunProcess s, , , vbMinimizedNoFocus
    
    'ChangePath FilePath(s)
    'Shell s, vbMinimizedNoFocus
    'ChangePath App.Path
    Exit Sub
    
    
    bToggle = Not bToggle
    BenchMark
    For i = 0 To 100000
        d = i
        If bToggle Then
            gdStartProfile 1
            gdStopProfile 1
        End If
    Next
    BenchMark "Done"
    Exit Sub
       
       
    SetBarProperties b, "SP-199003"
    
    s = "SP-200212"
    nSymbolID = GetSymbolID(s)
    d = DateSerial(2002, 6, 16)
    'd = Date - 1
    
    Set b = New cGdBars
    AddList "DM_GetTickBars ..."
    'i = DM_GetBars(b, nSymbolID, "90min", d, Date, , , , , False)
    AddList b.Size & " bars", True
    For i = 0 To 15
        AddList BarDisplay(b, i)
    Next
    
#If 0 Then
    Set b = New cGdBars
    AddList "BuildBars ..."
    'i = DM_GetBars(b, nSymbolID, "90min", d, Date, , , , , 1)
    AddList b.Size & " bars", True
    For i = 0 To 15
        AddList BarDisplay(b, i)
    Next
#End If
    
    Set b = New cGdBars
    AddList "BuildBars ..."
    'i = DM_GetBars(b, nSymbolID, "90min", d, Date, , , , , 2)
    AddList b.Size & " bars", True
    For i = 0 To 15
        AddList BarDisplay(b, i)
    Next
    
    Set b = New cGdBars
    b.ArrayMask = eBARS_TickByTick
    b.Size = 200000
    b.Size = 0
    AddList "DM_GetTickData - ALL days at once ..."
    If DM_GetTickData2(g.DMS, nSymbolID, d, Date, b.BarsHandle, 0) <> 0 Then
        AddList b.Size & " ticks", True
    End If
    
    Set b = New cGdBars
    AddList "DM_GetTickData -- ONE day at a time ..."
    i = 0
    For nDate = d To Date
        If nDate Mod 7 >= 2 Then
            b.Size = 0
            If DM_GetTickData2(g.DMS, nSymbolID, nDate, nDate, b.BarsHandle, 0) <> 0 Then
                i = i + b.Size
            End If
        End If
    Next
    AddList i & " ticks", True
    
    Exit Sub
       
    frmQuotes.TotalRefresh True
    Exit Sub

       
    d = Date + 60000# / 86400#
    d = Int(d * 86400# + 0.5) / 86400#
    AddList Format(d, "HH:MM:SS")
    b.Prop(eBARS_LastTickTime) = MinutesFromMidnight(d)
    AddList b.Prop(eBARS_LastTickTime)
    
    
    d = Date + 60060# / 86400#
    d = Int(d * 86400# + 0.5) / 86400#
    AddList Format(d, "HH:MM:SS")
    b.Prop(eBARS_LastTickTime) = (d - Int(d)) * 1440
    AddList b.Prop(eBARS_LastTickTime)
    
    Exit Sub
       
       
    BenchMark
    If RunProcess("d:\dvlp\temp\ys_rsrc.exe", , True, , i) Then
        AddList "Exit = " & CStr(i) & ",  time = " & CStr(BenchMark)
    Else
        AddList BenchMark
    End If
    Exit Sub
       
       
    s = InputBox("number")
    d = ValOfText(s)
    MsgBox CStr(d)
    Exit Sub
    
    'g.RealTime.Active = Not g.RealTime.Active
    Exit Sub
    
    AddList "Requested" & Str(GetTickCount)
    'frmMain.ChartNav.CreateMessage "Test", 5, CStr(GetTickCount)
    Exit Sub
    
    'b.FromFile "CSI", "d:\gd\back67", "tq-067"
    DM_GetBars b, "tq-067"
    For i = b.Size - 10 To b.Size - 1
        s = b.PriceDisplay(b(eBARS_Close, i), False) _
            & vbTab & b.PriceDisplay(b(eBARS_Close, i), True) _
            & vbTab & CStr(b(eBARS_Close, i))
        AddList s
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest.Command1_Click"
End Sub

Private Sub Command2_Click()
    
    Dim i&, j&, h&, s$, d#, strType$, bNewBar As Boolean
    Dim a1 As New cGdArray, a2 As cGdArray
    Dim b1 As New cGdBars, b2 As cGdBars
    
    Static bAlreadyDone As Boolean
         
    'frmTest3.ShowMe
    'CheckForStaleFundamentals
    CheckForexSpikes
Exit Sub
         
If IsIDE Then
    MaxTradesPerSecond
    Exit Sub
End If
    
    AddList "start"
    For i = 1 To 1000
        SetBarProperties b1, "ES-067", False
    Next
    AddList "finish1", True
    For i = 1 To 1000
        SetBarProperties b1, "ES-067", True
    Next
    AddList "finish2", True
Exit Sub
        
    'g.RealTime.SalmonStart
Exit Sub
    
    Set a1 = frmSymbolSelector.ShowMe()
    If a1.Size > 0 Then
        strType = InfBox("Which type of data?", "?", "+Daily|Minute|Ticks", "Data Type")
        For i = 0 To a1.Size - 1
            s = UCase(Trim(a1(i)))
            Select Case strType
            Case "D"
                If g.RealTime.SymbolInfo(s).GetDataRequestStatus(0) = eSalmonAvailable Then
                    DM_GetBars b1, s, , Date - 7, , , , , False
                    g.RealTime.SymbolInfo(s).UpdateDailyBars b1, bNewBar
                    AddList s & vbTab & "OHLC = " & b1.PriceDisplay(b1(eBARS_Open, b1.Size - 1)) _
                        & " " & b1.PriceDisplay(b1(eBARS_High, b1.Size - 1)) _
                        & " " & b1.PriceDisplay(b1(eBARS_Low, b1.Size - 1)) _
                        & " " & b1.PriceDisplay(b1(eBARS_Close, b1.Size - 1))
                Else
                    AddList s & vbTab & "(pending)"
                End If
            Case "T"
                h = ePRD_EachTick + 1
                AddList "TICKS STATUS = " & Str(g.RealTime.SymbolInfo(s).GetDataRequestStatus(h))
                Set b2 = g.RealTime.SymbolInfo(s).GetIntradayData(h)
            Case "M"
                h = ePRD_Minutes + 1
                AddList "MINUTEBAR STATUS = " & Str(g.RealTime.SymbolInfo(s).GetDataRequestStatus(h))
                Set b2 = g.RealTime.SymbolInfo(s).GetIntradayData(h)
                For j = 0 To b2.Size - 1
                    If j < 3 Or j >= b2.Size - 3 Then
                        AddList "  bar " & Str(j) & ": " & DateFormat(b2(eBARS_DateTime, j), MM_DD_YYYY, HH_MM) & vbTab _
                            & ", OHLC=" & b2.PriceDisplay(b2(eBARS_Open, j)) _
                            & " " & b2.PriceDisplay(b2(eBARS_High, j)) _
                            & " " & b2.PriceDisplay(b2(eBARS_Low, j)) _
                            & " " & b2.PriceDisplay(b2(eBARS_Close, j))
                    End If
                Next
            End Select
        Next
        AddList Format(Now, "hh:mm:ss")
    End If
    
Exit Sub
    
   
    
#If 0 Then
    d = gdTickCount
    i = DateSerial(2008, 10, 10)
    DM_GetBars b1, "ES-067", "EachTick", i, i
    i = b1.Size
    AddList "Loaded in " & Str(Int(gdTickCount - d)) & " ms"
    d = gdTickCount
    g.RealTime.FixTicks b1
    AddList "Fixed in " & Str(Int(gdTickCount - d)) & " ms"
    
Exit Sub
    s = "YM1-200812"
    SetBarProperties b1, s
    'g.RealTime.SerializeBars b1, "c:\temp\" & s & ".tck", False
    b1.SerializeDataArrays "c:\temp\" & s & ".tck", False
    i = b1.Size
Exit Sub
#End If
    
    i = DateSerial(2008, 10, 10)
    d = gdTickCount(0)
    'DM_GetBars b1, "$EUR-USD", "EachTick", i, i
    DM_GetBars b1, "ES-067", "EachTick", i, i
    d = gdTickCount(0) - d
    AddList "GetBars, #ticks = " & Str(b1.Size) & ", ms = " & Str(Int(d))
    i = b1.Size
    
    d = gdTickCount(0)
    s = App.Path & "\chk\test.bin"
    'If g.RealTime.SerializeBars(b1, s, True) Then
    If b1.SerializeDataArrays(s, True) Then
        d = gdTickCount(0) - d
        AddList "Write serialized file: ms = " & Str(Int(d))
    End If
    
    d = gdTickCount(0)
    Set b2 = b1.MakeCopy(True)
    'If g.RealTime.SerializeBars(b2, s, False) Then
    If b2.SerializeDataArrays(s, False) Then
        d = gdTickCount(0) - d
        AddList "Read serialized file: ms = " & Str(Int(d))
    End If
    
Exit Sub

    i = frmMain.apmRTClient.CreateMessage("GenesisRT", 22, "0")
    If i <> 0 Then AddList "Storage ON"
Exit Sub
    
    d = gdTickCount
    DM_GetBars b1, "es1-067", "5m", -500
    AddList Str(Int(gdTickCount - d)) & " new"
    For i = 0 To b1.Size - 1
        If i < 10 Or i > b1.Size - 10 Then
            If b1(eBARS_DateTime, i) > 0 Then
                AddList Str(i) & vbTab & Format(b1(eBARS_DateTime, i), "mm/dd/yyyy hh:mm") & vbTab & Str(b1(eBARS_Close, i))
            Else
                AddList Str(i) & vbTab & "(null)"
            End If
        End If
    Next
    
    d = gdTickCount
    DM_GetBars b1, "es1-067", "5m", Int(b1(eBARS_DateTime, 0))
    AddList Str(Int(gdTickCount - d)) & " old"
    For i = 0 To b1.Size - 1
        If i < 10 Or i > b1.Size - 10 Then
            If b1(eBARS_DateTime, i) > 0 Then
                AddList Str(i) & vbTab & Format(b1(eBARS_DateTime, i), "mm/dd/yyyy hh:mm") & vbTab & Str(b1(eBARS_Close, i))
            Else
                AddList Str(i) & vbTab & "(null)"
            End If
        End If
    Next
    
    Exit Sub
    
    'PlaySoundFile "c:\temp\t083.wav"
    PlaySoundFile "c:\temp\startup.wav", False, True
    Exit Sub
    
    DM_GetBars b1, "sp-067"
    s = SecurityType(b1) '.Prop(eBARS_SymbolID))
    i = Date - 70
    i = b1.FindDateTime(i)
    d = b1(eBARS_DateTime, i)
    
    Exit Sub
    
    
    Dim Criterias As Object
    Dim Criteria As cCriteria
    Dim lIndex As Long
    
        
        Set Criterias = g.SymbolPool.Criterias
        For lIndex = 1 To Criterias.Count
            Set Criteria = Criterias(lIndex)
            If Criteria.Custom Then
                s = Criteria.Name
                If InStr(UCase(s), "THRUST") > 0 Then
                    i = i
                End If
                If Len(Criteria.Required) > 0 Then
                    i = i
                End If
                If HasModule(Criteria.Required) Then
                    s = Criteria.Name
                End If
            End If
        Next lIndex
    
    Exit Sub
    
    a1(0) = 2.2
    a1(1) = 3.3
    a1(2) = 4.4
    
    Set a2 = a1.MakeCopy
    
    For i = 0 To a2.Size - 1
        AddList i & ":  " & a1(i) & "  " & a2(i)
    Next
    a2(1) = 3.1
    For i = 0 To a2.Size - 1
        AddList i & ":  " & a1(i) & "  " & a2(i)
    Next
    
    
    b1.FromFile "CSI", "c:\gd\back67", "sp-067", "m"
    Set b2 = b1.MakeCopy
    
    For i = 0 To 3
        AddList i & ":  " & b1(eBARS_Close, i) & "  " & b2(eBARS_Close, i)
    Next
    b2(eBARS_Close, 1) = 3.1
    For i = 0 To 3
        AddList i & ":  " & b1(eBARS_Close, i) & "  " & b2(eBARS_Close, i)
    Next
    
End Sub

Private Sub Command3_Click()

    Dim aCodedNames As New cGdArray

    Dim X As Long
    Dim lNumDays As Long
    Dim strNotKnown As String
    Dim bExtraInputs As Boolean
    
    Dim strUserText$
    Dim strCodedText$
    Dim bIsBoolean As Boolean
    Dim s$, i&, j&, d#
    Dim frm As Form, tmr As Timer
    Dim a1 As New cGdArray
    Dim b As New cGdBars
       
    Dim SymInf As cSymbolInfo
    
    Dim rs As Recordset

    
    IntradayMarketScope
Exit Sub

    'MakeMidCmd
    
    s = GetMidCmd
    s = GetMidCmd
    'frmQuotes.WebPageCheck
Exit Sub

    i = 41265
'i = Date + 7
If 0 Then
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!SessionDate <= i Then
            'rs.Edit
            rs.Delete
            'rs.Update
        End If
        rs.MoveNext
    Loop
End If
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!SessionDate <= i Then
            'rs.Edit
            rs.Delete
        End If
        rs.MoveNext
    Loop
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!SessionDate <= i Then
            'rs.Edit
            rs.Delete
        End If
        rs.MoveNext
    Loop

Exit Sub
    ' tblFills, tblOrders, tblOrderLegs - delete OrderID <= 2421
    i = 2421
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblFills];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!OrderID <= i Then
            rs.Edit
            rs.Delete
        End If
        rs.MoveNext
    Loop
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!OrderID <= i Then
            rs.Edit
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
Exit Sub

    AddList "start"
    DoPFCheck
    AddList "done", True
Exit Sub

    If FtpUploadCheck Then
        AddList "FtpUpload = TRUE"
    Else
        AddList "FtpUpload = FALSE"
    End If
    
    'ChkBigJumpsInBonds
Exit Sub

    'SectorWeb=www.TradeNavigator.com/industries/indyrank.aspx?U=*&P=*
    'NewsWeb=www.TradeNavigator.com/Analysis/default.aspx?U=*&P=*&T=News
    s = "www.TradeNavigator.com/Analysis/default.aspx?U=*&P=*&T=News"
    s = "www.TradeNavigator.com/SeasFilter/default.aspx?U=*&P=*&T=Seas"
    s = "www.TradeNavigator.com/Analysis/default.aspx?U=*&P=*&T=Seas"
    s = EncryptToHex(s)
    s = DecryptFromHex(s)
    
    
    DM_GetBars b, "TF-201109", "Each Tick", 20110816, 20110816
    i = b.Size
    For i = 0 To b.Size - 1
        d = b(eBARS_Close, i)
        If d < 100 Then
            d = d
        End If
    Next
    d = d
    
Exit Sub
    s = "SOFTWARE\Trader Workstation"
    s = GetRegistryValue(rkLocalMachine, s, "jtspath", "")
    AddList "jtspath=" & s
    If Len(Trim(s)) > 3 Then
    End If
Exit Sub
    s = "http://www.sfe.com.au/Content/reports/EODWebMarketSummary100824SFT.htm"
    s = GetWebPageData(s)
    FileFromString "c:\temp\test.txt", s
Exit Sub
    
    s = ".SCN"
    AddList GetUniqueCustomFilename(s)
Exit Sub
    
    'http://www.TradeNavigator.com/alerts/index.aspx?U=4fc83ae5&P=a791367fcfa151
    '&A=7193596323@txt.att.net&S=Trade Navigator Alert for ES1-067&C=price of ES1-067 is up to 1107.50 (23 Apr 2010 17:10:15)
    
    s = "https://www.TradeNavigator.com/orderwizard/LogIn.aspx?s=" & EncryptToHex(RI_GetDataServiceID) _
        & "&p=" & EncryptToHex(RI_GetUserPassword) & "&agree=BRKRLIVE"
    
    s = "https://www.TradeNavigator.com/orderwizard/LogIn.aspx?s=*&p=*&agree=BRKRLIVE"
    s = FixURL(GetProvidedProperty("AgreeBrokerLive", s))

    frmMain.tmrQuickStart.Enabled = True

Exit Sub
    s = FixURL("http://www.TradeNavigator.com/alerts/index.aspx?U=*&P=*")
    's = s & "&A=" & astrAction(2) & "&S=Trade Navigator Alert for " & strCaption & "&C=" & strMsg
    s = s & "&A=7193596323@txt.att.net&S=Trade Navigator Alert for ES1-067&C=price of ES1-067 is up to 1107.50 (23 Apr 2010 ??)"
    If Len(s) >= 1024 Then s = Left(s, 1022)  'cannot send more than 1024 chars in URL
    frmTest.AddList s
    s = GetWebPageData(s)
    frmTest.AddList s

Exit Sub
    
    If Not g.RealTime.SalmonIsRunning Then Exit Sub
    
    Set a1 = frmSymbolSelector.ShowMe(, , , "Remove symbol from stream")
    If a1.Size > 0 Then
        For i = 0 To a1.Size - 1
            s = UCase(Trim(a1(i)))
            Set SymInf = g.RealTime.SymbolInfo(s)
            SymInf.RemoveFromStream
            Set SymInf = Nothing
        Next
    End If

Exit Sub
       
    i = FindWindow("TradeNavStartup", "TradeNavStartup")
    If i <> 0 Then
        PostMessage i, WM_USER + 1, 2, 3
        i = i
    End If
Exit Sub

    s = "testing this just to see how long it takes"
    d = gdTickCount
    For i = 1 To 1000
        For j = i To 1000
            X = X + 1
            lNumDays = InStr(s, "long")
        Next
    Next
    d = gdTickCount - d
    AddList Str(Int(d)) & " ms, count = " & Str(X)

Exit Sub
    
    AddList "================"
'frmMain.tmrMain.Interval = 1000
'frmTTSummary.tmrBrokers.Interval = 5000
    For i = 0 To Forms.Count - 1
        Set frm = Forms(i)
        For j = 0 To frm.Controls.Count - 1
            If TypeOf frm.Controls(j) Is Timer Then
                Set tmr = frm.Controls(j)
                If tmr.Enabled Then
                    X = tmr.Interval
                Else
                    X = 0
                End If
                If X > 0 Then
                    AddList frm.Name & " " & tmr.Name & " " & Str(X)
                End If
                Set tmr = Nothing
            End If
        Next
    Next
    Set frm = Nothing
Exit Sub
    
    
    s = ConvertSynthetic("ES1", False)
    s = ConvertSynthetic("ES", True)
    s = ConvertSynthetic("ES-200603", True)
    s = ConvertSynthetic("ES1", True)
    
    s = ConvertSynthetic("ES-200603", False)
    AddList s
    s = ConvertSynthetic("ES1-200603", False)
    AddList s
    'PlaySoundFile
    Exit Sub
    
    
    MsgBox "MsgBox"
    
    Exit Sub
    
    
    BenchMark
    
    strUserText = "MovingAvg(Adx(7) - Close.5, 14) - MovingAvg(Adx(7) - Close.5, 17).3"
    'strCodedText = frmCustomFunction.ShowMe(strUserText, True)
    'strUserText = "WillVal(GC, 2, 22, 156)"
    strCodedText = frmCustomFunction.ShowMe(strUserText, True)
    
    AddList BenchMark
    AddList strUserText
    AddList strCodedText
    
    BenchMark
    X = GetFunctionIDFromCodedName("WillVal")
    AddList BenchMark & " old found " & CStr(X)
    
    BenchMark
    For X = 1 To g.Functions.Count
        aCodedNames.Add g.Functions.Item(X).CodedName & vbTab & CStr(g.Functions.Item(X).FunctionID)
        'If UCase(pCodedName) = UCase(g.Functions.Item(X).CodedName) Then
        '    GetFunctionIDFromCodedName = g.Functions.Item(X).FunctionID
        '    Exit For
        'End If
    Next X
    aCodedNames.Sort eGdSort_IgnoreCase
    AddList BenchMark & " to sort"

    BenchMark
    aCodedNames.BinarySearch "WillVal", X, eGdSort_IgnoreCase
    If X >= 0 Then
        If Parse(aCodedNames(X), vbTab, 1) = "WillVal" Then
            AddList BenchMark & " new found "
        End If
    End If
    
End Sub

Private Sub Command4_Click()

    Dim i&, s$, n&, d#, lDate&
    Dim aStrings As New cGdArray
    
    ShowDivTable
Exit Sub
    
    For n = 0 To g.SymbolPool.NumRecords - 1
        ' get subsector for stock symbol
        If g.SymbolPool.SecType(n) = eSYMType_Stock Then
            i = g.SymbolPool.SymbolID(n)
            If DM_GetSnap1(g.DMS, i, 163, d, lDate) = 0 Then
                d = 0
            End If
            If d > 0 Then
                s = g.SymbolPool.SymbolForID(i) & vbTab & Str(i) & vbTab & Str(d) & vbTab & DateFormat(lDate)
                aStrings.Add s
            End If
        End If
    Next
    aStrings.Sort
    aStrings.ToFile "c:\SubSectors.txt"
    AddList Str(aStrings.Size) & " stocks in subsectors"
    
Exit Sub
    
    With frmTTSummary.fgPositions
        AddList "Positions rows = " & Str(.Rows)
        .Rows = .FixedRows
        AddList "Positions rows = " & Str(.Rows)
    End With
Exit Sub

    g.SymbolPool.DirtyCriteria = True
    g.SymbolPool.RecalcDirtyCriteria 20111202
    'g.SymbolPool.CreateCriteriaRecalcFiles
    AddList "done"
Exit Sub
        
    EnableAero True
Exit Sub
    
    i = IsInWesternHemisphere
Exit Sub
    
    aStrings.Add "3Testing"
    aStrings.Add "6Testing"
    aStrings.Add "1Testing"
    aStrings.Add "7Testing"
    aStrings.Add "9Testing"
    aStrings.Add "2Testing"
    aStrings.Add "6Testing"
    
    aStrings.Sort eGdSort_IgnoreCase
    
    s = "6Test"
    If aStrings.BinarySearch(s, i, eGdSort_IgnoreCase Or eGdSort_MatchUsingSearchStringLength) Then
        s = aStrings(i)
    End If

End Sub

Private Sub Command5_Click()

    Dim i&, strSymbol$, strText$, dDate#, dDelta#, d#, dMult#, dMinMove#, n&, strFile$, hFile&, strPath$
    Dim Table As cGdTable
    
    Dim Bars As New cGdBars, Results As New cGdArray
    Dim iSeriesID&, dSmooth#, dPhase#, dInput#, dResult#, bGood As Boolean
    Dim hArrayI&, hArrayR&
    
    strText = UCase(Trim(InfBox("Symbol", "?", , "Get Salmon History", , , , , , "s", "$DJIA")))
    strSymbol = Parse(strText, ",", 1)
    If Len(strSymbol) > 0 Then
        n = 0
        dDate = 0
        strText = Parse(strText, ",", 2)
        i = InStr(strText, "/")
        If i > 0 Then
            ' get date for full ticks
            If InStr(i + 1, strText, "/") > 0 Then
                dDate = DateOf(strText)
            Else
                dDate = DateOf(strText & "/" & Str(Year(Date)))
                If dDate > Date Then
                    dDate = DateOf(strText & "/" & Str(Year(Date) - 1))
                End If
            End If
        Else
            n = Val(strText)
        End If
        
        If dDate > 0 Then
            strText = "Each tick"
            DM_GetBars Bars, strSymbol, strText, Date + 2, Date + 2
            Results(0) = "2"
            Results(1) = Str(JulToLong(dDate, True))
            Results(2) = Str(JulToLong(dDate, True))
            strText = DateFormat(dDate, MM_DD_YYYY)
        ElseIf n > 0 Then
            strText = Str(n) & " minute"
            DM_GetBars Bars, strSymbol, strText, Date + 2, Date + 2
            Results(0) = "1"
            Results(1) = Str(JulToLong(Date - 14, True))
            Results(2) = Str(JulToLong(LastDailyDownload - 1, True))
        Else
            strText = "Daily"
            DM_GetBars Bars, strSymbol, strText, Date + 2, Date + 2
            Results(0) = "0"
            Results(1) = "19000101"
            Results(2) = Str(JulToLong(LastDailyDownload - 1, True))
        End If
        Results(3) = ""
        Bars.Size = 0
        n = 0
        
        AddList "Calling 'GetHistory' for: " & strSymbol & ", " & strText
        n = GetSalmonHistory(Results.ArrayHandle, Bars.BarsHandle)
        AddList "Return code = " & Str(n) & ", Bars size = " & Str(Bars.Size), True
        If Bars.Size > 0 Then
            AddList BarDisplay(Bars, 0)
            AddList BarDisplay(Bars, 1)
            AddList BarDisplay(Bars, Bars.Size - 2)
            AddList BarDisplay(Bars, Bars.Size - 1)
        End If
    End If
    
Exit Sub
    Set Table = GetRollsTable(strText)
    For i = 0 To Table.NumRecords - 1
        n = Table(0, i)
        strText = SU_GetSymbol(n)
        dDate = Table(1, i)
        dDelta = Table(2, i)
        AddList CStr(i) & vbTab & strText & vbTab & _
            DateFormat(dDate) & vbTab & CStr(dDelta)
    Next
Exit Sub
    DM_GetBars Bars, "yi2-067", "each tick", LastDailyDownload, LastDailyDownload
    i = Bars.Size
    lst.Clear
    AddList Str(Bars.Size) & " ticks"
    For i = 0 To Bars.Size - 1
        If i > Bars.Size - 100000 Then
            AddList Str(i) & vbTab & DateFormat(Bars(eBARS_DateTime, i), M_D_YY, HH_MM_SS) & vbTab & Bars.PriceDisplay(Bars(eBARS_Close, i)) _
                & vbTab & Str(Bars(eBARS_Vol, i)) & vbTab & Str(Bars(eBARS_DateTime, i))
        End If
    Next
    
Exit Sub

    dMinMove = 0.01
    dMult = Int(1 / dMinMove + 0.5)
    For i = 1 To 1000
        d = i * dMinMove
        If d <> i / dMult Then
            d = d
            If i * dMinMove <> 0.35 Then
                d = d
            End If
            If i / dMult <> 0.35 Then
                d = d
            End If
        End If
    Next
    
Exit Sub
    
    
    i = DateSerial(2010, 10, 18)
    DM_GetBars Bars, "G6E-067", "EachTick", i, i
    i = Bars.Size
    Bars.FixPrices
    dDelta = Bars.MinMove
    dDelta = Int(1 / dDelta + 0.5)
    For i = 0 To Bars.Size - 1
        d = Bars(eBARS_Close, i)
        'd = Int(d / dDelta + 0.5) * dDelta
        d = Int(d * dDelta + 0.5) / dDelta
        If d <> Bars(eBARS_Close, i) Then
            d = d
        End If
    Next
    
Exit Sub
    
    AddList "Is 95, 98 or ME = " & Str(Is9598orMe)
    AddList "Is at least Vista = " & Str(IsAtLeastVista)
    InfBox "Test message", "i", , "Testing"
Exit Sub
    
    DM_GetBars Bars, "ES-200609", "Each tick", Date, Date
    i = Bars.Size
    For i = 0 To Bars.Size
        If Bars(eBARS_Flags, i) > 0 Then
            i = i
        End If
    Next
    
    DM_GetBars Bars, "ES-200609", "5 min", Date, Date
    i = Bars.Size
    For i = 0 To Bars.Size
        If Bars(eBARS_Flags, i) > 0 Then
            i = i
        End If
        If Bars(eBARS_BidVol, i) > 0 Then
            i = i
        End If
    Next
Exit Sub
    
    Bars.Prop(eBARS_PeriodicityStr) = "daily"
    i = Bars.Prop(eBARS_ArrayMask)
    Bars.Size = 10
    i = Bars.ArrayHandle(eBARS_Flags)
    i = gdGetSize(i)
    i = Bars(eBARS_Flags, 0)
    
    Set Bars = New cGdBars
    Bars.ArrayMask = eBARS_TickByTick

    i = Bars.Prop(eBARS_ArrayMask)
    Bars.Size = 10
    i = Bars(eBARS_Flags, 0)
    Bars(eBARS_Flags, 0) = 1
    i = Bars(eBARS_Flags, 0)
    i = Bars.ArrayHandle(eBARS_Flags)
    i = gdGetSize(i)
Exit Sub
    
    
    dSmooth = 10
    dPhase = -20
    iSeriesID = 0
    
    DM_GetBars Bars, "IBM"
    Results.Create eGDARRAY_Doubles, Bars.Size
    'pArrayData = gdGetDataPtr(Results.ArrayHandle)
    
    AddList "Start JMAUT", True
    hArrayI = Bars.ArrayHandle(eBARS_Close)
    hArrayR = Results.ArrayHandle
    bGood = False
    For i = 0 To Bars.Size - 1
        dInput = gdGetNum(hArrayI, i) 'Bars(eBARS_Close, i)
        dResult = -999999
        'n = JMAUT(Bars(eBARS_Close, i), dSmooth, dPhase, pArrayData + 8 * i, iDestroy, iSeriesID, iSameBar)
        n = JMAUT(dInput, dSmooth, dPhase, dResult, 0, iSeriesID, 0)
        If n <> 0 Then
            AddList "JMAUT Error = " & Str(n) & " at bar " & Str(i)
        Else
            If dResult <> dInput Then bGood = True
            If bGood Then
                gdSetNum hArrayR, i, dResult
                'Results(i) = dResult
            End If
        End If
    Next
    n = JMAUT(0, 0, 0, 0, 1, iSeriesID, 0)
    If n <> 0 Then
        AddList "JMAUT Destroy Error = " & Str(n)
    End If
    AddList "Finish JMAUT", True
    AddList Results(Results.Size - 1)
    
    For i = 0 To Bars.Size - 1
        If i < 40 Or i > Bars.Size - 15 Then
            AddList "Bar " & Str(i) & vbTab & Str(Bars(eBARS_Close, i)) & vbTab & Str(Results(i))
        End If
    Next
    
    
Exit Sub
       
           
    strFile = App.Path & "\SimTrade\Info.GZP"
    hFile = FileOpen(strFile, "rb")
    If hFile Then
        strText = Space(FileLength(strFile))
        FileBinaryIO hFile, ByVal strText, Len(strText), False
        FileClose hFile
    End If
           
    'strText = msg.Message
    strFile = App.Path & "\SimTrade\Trades.GZP"
    KillFile strFile
    hFile = FileOpen(strFile, "wb")
    If hFile Then
        FileBinaryIO hFile, ByVal strText, Len(strText), True
        FileClose hFile
        If FileLength(strFile) > 10 Then
            ZipExecute "U", strFile, AddSlash(App.Path) & "SimTrade\In", "*.CSV"
        End If
    End If
       
End Sub

Private Sub Command6_Click()

    Dim strSymbol$, strFeedSymbol$, strExchange$, strSecType$, strFeed$
    Dim dMult#, nCrossover&, nGmtOffset&, i&, h&, s$, d#
    Dim aCompSymbols As New cGdArray
    Dim aCompExchanges As New cGdArray
    
    Dim wp As WINDOWPLACEMENT
    Dim pTable&, pSymbol&, pDate&, pValues&, iSum&, iCount&, dAvg#
    
    Dim frm As Form, ctl As Control
    
    i = 0
    If frmMain.tmrMain.Enabled Then
        For i = Forms.Count - 1 To 0 Step -1
            Set frm = Forms(i)
            For h = 0 To frm.Controls.Count - 1
                Set ctl = frm.Controls(h)
                If TypeOf ctl Is Timer Then
                    If ctl.Enabled Then
                        s = frm.Name & "." & ctl.Name
                        AddList s
                        ctl.Enabled = False
                        Sleep 5
                    End If
                End If
            Next
            Set ctl = Nothing
        Next
        AddList "All timers now disabled"
    Else
        frmMain.tmrMain.Enabled = True
        AddList "frmMain.tmrMain renabled"
    End If
    Set frm = Nothing
    
Exit Sub
    frmTTSummary.tmrBrokers.Enabled = Not frmTTSummary.tmrBrokers.Enabled
    AddList "tmrBrokers = " & Str(frmTTSummary.tmrBrokers.Enabled)
Exit Sub
    DumpDailyData 'DateSerial(2010, 10, 14)

    Exit Sub
    
#If 0 Then
    dMult = 0.001
    
    d = gdTickCount
    dAvg = 0
    For i = 1 To 1000000
        dAvg = gdRoundPriceToMinMove(dAvg + dMult, dMult)
    Next
    AddList Str(Int(gdTickCount - d)) & " ms for NEW RoundMinMove"
    
    d = gdTickCount
    dAvg = 0
    For i = 1 To 1000000
        dAvg = RoundToMinMove_OLD(dAvg + dMult, dMult)
    Next
    AddList Str(Int(gdTickCount - d)) & " ms for OLD RoundMinMove"
    
    d = gdTickCount
    dAvg = 0
    For i = 1 To 1000000
        dAvg = SigDigits(i / 100#, 5)
    Next
    AddList Str(Int(gdTickCount - d)) & " ms for NEW SigDigits"
    
    d = gdTickCount
    dAvg = 0
    For i = 1 To 1000000
        dAvg = RoundToSigDigits(i / 100#, 5)
    Next
    AddList Str(Int(gdTickCount - d)) & " ms for OLD SigDigits"
    
    d = gdTickCount
    dAvg = 0
    For i = 1 To 1000000
        dAvg = gdRoundNum(i / 100#, 5)
    Next
    AddList Str(Int(gdTickCount - d)) & " ms for NEW RoundNum"
    
    d = gdTickCount
    dAvg = 0
    For i = 1 To 1000000
        dAvg = RoundNum(i / 100#, 5)
    Next
    AddList Str(Int(gdTickCount - d)) & " ms for OLD RoundNum"
#End If
    
    
    pTable = TblOpen(App.Path & "\Data\idx_ful7.dbf", False, True)
    pSymbol = fldPtr(pTable, "SymbolID")
    pDate = fldPtr(pTable, "Date")
    pValues = fldPtr(pTable, "Values")
    iSum = 0
    iCount = 0
    If TagSeek(pTable, "SymbDate", "     48552", False) Then
        Do
            If f4long(pSymbol) <> 48552 Then Exit Do
            i = f4long(pDate)
            h = f4memoLen(pValues)
            If h > 150000 Then
                h = h
            ElseIf h > 0 Then
                iSum = iSum + h
                iCount = iCount + 1
            End If
            d4skip pTable, 1
        Loop
    End If
    TblClose pTable
    If iCount > 0 Then
        dAvg = iSum / CDbl(iCount)
    End If
    
Exit Sub

    fg.BorderStyle = flexBorderNone
    fg.Cell(flexcpFloodPercent, 2, 2) = 50
    
Exit Sub
    
Sleep 10
    
    wp.Length = Len(wp)
    
    For i = 0 To Forms.Count - 1
        If TypeOf Forms(i) Is frmChart Then
            s = Forms(i).Chart.ChartName
            If InStr(s, "XOM") > 0 And InStr(s, "Daily") > 0 Then
                h = Forms(i).hWnd
                GetWindowPlacement h, wp
                dMult = wp.rcNormalPosition.Left * Screen.TwipsPerPixelX
                AddList Str(dMult)
                If dMult <> Forms(i).Left Then InfBox "error"
            End If
        End If
    Next
    
    
Exit Sub
    
    If 1 Then
        strSymbol = "ES1-200403"
        strFeed = "S"
        'strSecType = "F"
        'strExchange = "CME"
        If TranslateSymbol(strSymbol, strFeed, strSecType, strFeedSymbol, strExchange, dMult, nCrossover, nGmtOffset) Then
            dMult = dMult
        End If
    ElseIf 0 Then
        strFeedSymbol = "ES Z1"
        strFeed = "S"
        strSecType = "F"
        strExchange = "CME"
        If TranslateSymbol(strSymbol, strFeed, strSecType, strFeedSymbol, strExchange, dMult, nCrossover, nGmtOffset, True) Then
            dMult = dMult
        End If
    Else
        strFeed = "B"
        strSymbol = "ES-200112"
        strSymbol = "ABC"
        strSymbol = "$DJIA"
        'strSymbol = "TRAD"
        strSymbol = "IBM AG"
        
        strSymbol = "$EUR-USD"
        strFeed = "S"
        If TranslateSymbol(strSymbol, strFeed, strSecType, strFeedSymbol, strExchange, dMult, nCrossover, nGmtOffset) Then
            dMult = dMult
            i = Asc(strExchange)
        End If
    End If

End Sub

Function RandomNumTest&(ByVal lowerBound&, ByVal upperbound&)

    Static randomized As Boolean
    
    ' initialize random number generator once
    If Not randomized Then
        Randomize
        randomized = True
    End If
    
    ' get a random number
    If upperbound <= lowerBound Then
        RandomNumTest = lowerBound
    Else
        RandomNumTest = Int((upperbound - lowerBound + 1) * Rnd) + lowerBound
    End If
    
End Function

Private Function VbChk(ByVal n&) As Long
    VbChk = n
End Function

Private Sub Command7_Click()

    Dim Bars As New cGdBars
    Dim aList As New cGdArray
    Dim n&, i&, s$, t$, d As Date, p#, h&, ub&, lb&
    Dim strFormat$, strPath$
    Static bSocketLog As Boolean
    
    CalcSeasonals
    
    'frmTTSummary.tmrRealtime.Enabled = Not frmTTSummary.tmrRealtime.Enabled
    'AddList "tmrRealTime = " & Str(frmTTSummary.tmrRealtime.Enabled)
Exit Sub

    Dim tResults As New cGdTable
    s = CalcSeasonalChart(tResults, Bars, "ES-067", "Weekly")
    i = tResults.NumRecords
    AddList Str(i) & " records, Err = " & s
    s = tResults.ToString
    FileFromString "c:\SeasChart.dat", s, True

Exit Sub
    AddList "ABORT"
    GenZipAbort True
Exit Sub
    
    bSocketLog = Not bSocketLog
    g.RealTime.SocketLogging = bSocketLog
    AddList "SocketLog = " & Str(bSocketLog)
    Exit Sub
    
    'IdleSleep 10000, False
    'AddList "IdleSleep (full idle) done"
    
    AskForActivate
    'GetRegisterProgram
    
    Exit Sub
    
    'bars.Size = 10000
    BenchMark
    n = Bars.FromFile("AT7", "d:\temp\", "S-055.csv")
    'n = Bars.FromFile("csi", "c:\gd\back67", "tq-9967")
    'n = bars.FromFile("gt", "d:\temp", "abt.gt", "30", s)
    AddList CStr(BenchMark) & " to load: " & Str(n) & ", " & s
    n = Bars.Size

    AddList "Symbol: " & Bars.Prop(eBARS_Symbol) & ", " & Bars.Prop(eBARS_Desc)
    AddList "NumBars = " & Str(n) & ",  cf = " & Str(Bars.Prop(eBARS_ConvFactor))
    For i = 0 To n - 1
        If i < 50 Or i >= n - 50 Then
            d = Bars(eBARS_DateTime, i)
            t = DateFormat("Format", MM_DD_YY) & " "
            If d <> CLng(d) Then
                t = t & "HH:MM "
            End If
            If 0 Then
                s = Format(d, t) & Chr(9) & Str(Bars(eBARS_Close, i)) & Chr(9) & Str(Bars(eBARS_Vol, i))
            Else
                s = "0.####"
                s = Format(d, t) & Chr(9) & Format(Bars(eBARS_Open, i), s) _
                        & Chr(9) & Format(Bars(eBARS_High, i), s) _
                        & Chr(9) & Format(Bars(eBARS_Low, i), s) _
                        & Chr(9) & Format(Bars(eBARS_Close, i), s) _
                        & Chr(9) & Str(Bars(eBARS_Vol, i)) _
                        & Chr(9) & Str(Bars(eBARS_UpTicks, i)) _
                        & Chr(9) & Str(Bars(eBARS_DownTicks, i))
            End If
            AddList s
        End If
    Next
    n = Bars.Size
  
    n = Bars.ToFile("CSI", "d:\temp\", "S-055")
    AddList Str(n) & " ToFile"
  
  Exit Sub
    
    s = Bars.Prop(eBARS_Symbol)
    strFormat = "ms7"
    strPath = "d:\temp"
    n = Bars.ToFile(strFormat, strPath)
    AddList "ToFile = " & Str(n)
      
    Set Bars = New cGdBars
    n = Bars.FromFile(strFormat, strPath, s)
    'BenchMark "Loaded - " & Str(n)
    n = Bars.Size
    AddList "Symbol: " & Bars.Prop(eBARS_Symbol) & ", " & Bars.Prop(eBARS_Desc)
    AddList "NumBars = " & Str(n) & ",  cf = " & Str(Bars.Prop(eBARS_ConvFactor))
    For i = 0 To n - 1
        If i < 10 Or i >= n - 10 Then
            d = Bars(eBARS_DateTime, i)
            s = Format(d) & Chr(9) & Str(Bars(eBARS_Close, i)) & Chr(9) & Str(Bars(eBARS_Vol, i))
            AddList s
        End If
    Next
      
      
Exit Sub
    
    Dim randomized As Boolean

    If SetRegistryValue(rkLocalMachine, "software\atest\test1", "val1", "Testing this") Then
        s = GetRegistryValue(rkLocalMachine, "software\atest\test1", "val1", "default")
    
    End If
Exit Sub
    
    BenchMark
    
    n = 10000000
    
    For i = 1 To n
        h = gdCheckMemoryLeaks(0)
    Next
    
    BenchMark "DLL func"
    
    For i = 1 To n
        h = VbChk(0)
    Next
    
    BenchMark "Vb func"
    
    For i = 1 To n
        h = 0
    Next
    
    BenchMark "Vb direct"
'Exit Sub
    
    n = 1000000
    
    ub = 1000
    lb = 0
    
    For i = 1 To n
       ' initialize random number generator once
        If Not randomized Then
'            Randomize
            randomized = True
        End If
        
        If ub <= lb Then
            h = lb
        Else
            h = Int((ub - lb + 1) * Rnd) + lb
        End If
    Next
    
    BenchMark "Vb no func"
    
    For i = 1 To n
        h = RandomNum(lb, ub)
    Next
    
    BenchMark "Vb func"
    
    For i = 1 To n
        h = gdRandomNumber(lb, ub)
    Next
    
    BenchMark "C"
    
Exit Sub

    s = "c:\dvlp\test.dir\check.txt"
    lst.Clear
    AddList s
    AddList FilePath(s)
    AddList FileBase(s)
    AddList FileExt(s)
    AddList ReplaceFileExt(s, "dat")

    s = "c:\dvlp\test.dir\check"
    AddList " "
    AddList s
    AddList FilePath(s)
    AddList FileBase(s)
    AddList FileExt(s)
    AddList ReplaceFileExt(s, "dat")

    s = "c:\dvlp\test.dir\check.rrr.txt"
    AddList " "
    AddList s
    AddList FilePath(s)
    AddList FileBase(s)
    AddList FileExt(s)
    AddList ReplaceFileExt(s, "dat")

    s = "c:check.rrr.txt"
    AddList " "
    AddList s
    AddList FilePath(s)
    AddList FileBase(s)
    AddList FileExt(s)
    AddList ReplaceFileExt(s, "dat")

Exit Sub

    lst.Visible = True
    lst.ZOrder
    lst.Clear

    n = Bars.Size
    'bars.Size = 10000
    n = Bars.Size
    BenchMark
    n = Bars.FromFile("csi", "c:\dvlp\gd\", "f001.dta")
    BenchMark "Loaded - " & Str(n)
    n = Bars.Size

    AddList "NumBars = " & Str(n)
    For i = 0 To n - 1
        If i < 10 Or i >= n - 10 Then
            d = Bars(eBARS_DateTime, i)
            s = Format(d) & Chr(9) & Str(Bars(eBARS_Close, i))
            AddList s
        End If
    Next
    n = Bars.Size

    p = Bars(eBARS_Close, n - 3)
    Bars(eBARS_Close, n - 3) = 100.5
    p = Bars(eBARS_Close, n - 3)


    BenchMark
    'n = 10 * 100000
    'n = 10 * 10000&
    n = n * 10

    For i = 1 To n
        p = Bars(eBARS_Close, 1)
    Next
    BenchMark "Using bars - " & Str(n)

    h = Bars.ArrayHandle(eBARS_Close)
    If h Then
        For i = 1 To n
            p = gdGetNum(h, 1)
        Next
        BenchMark "Using array - " & Str(n)
    End If

End Sub

Private Sub Command8_Click()

    Dim i&, nDate&, bNewDay As Boolean
    Static Bars As cGdBars
      
    g.RealTime.bDisableForTesting = Not g.RealTime.bDisableForTesting
    
    'CalcSeasonals
    'DoGradientColors
Exit Sub
      
    If 0 Then
        KillFile "c:\di\test9.zip"
        i = ZipExecute("C", "c:\di\test9.zip", "c:\di\test\", , , , , , , txtZip.hWnd)
    Else
        KillFile "c:\di\test\*.*"
        i = ZipExecute("U", "c:\di\eod_i.gzp", "c:\di\test\", , , , , , , txtZip.hWnd)
    End If
    AddList "ZIP DONE"
    
Exit Sub
    
    If Not g.RealTime.Active Then Exit Sub
    
    If Bars Is Nothing Then
        Set Bars = New cGdBars
        nDate = Date + 1
        SetBarProperties Bars, "ES-067"
        Bars.Prop(eBARS_Periodicity) = ePRD_EachTick + 1
        g.RealTime.AddTickBuffer Bars
        g.RealTime.SpliceBars Bars, nDate
        'DM_GetBars Bars, "ES-067", "Each", nDate, nDate
        i = Bars.Size
        If Bars.Size = 0 Then
            nDate = nDate - 1
            g.RealTime.SpliceBars Bars, nDate
        End If
        i = Bars.Size
    End If
    
    If g.RealTime.UpdateBars(Bars, bNewDay) Then
        i = Bars.Size
    End If
    
End Sub

Private Sub Command10_Click()
On Error GoTo ErrSection:

    Dim i&, n&, c&, nc&, nStart&, d#, h&, s$
    Dim gdArray As New cGdArray
    Dim vbArray() As Double
    Dim t As cGdTable
    
    Set gdArray = frmSymbolSelector.ShowMe(, , , "DumpBySymbol", , , True)
    s = gdArray(0)
    If Len(s) > 0 Then
        i = DumpBySymbol(gdArray.ArrayHandle)
        AddList "DumpBySymbol(" & s & ") = " & Str(i)
    End If
Exit Sub
    AddList g.RealTime.DumpTickBufferInfo
Exit Sub
    g.RealTime.CreateTransFile
    'Test
    Exit Sub
    
    'd = SU_GetLocalTime(Now, SU_NewYorkTimeZone)
    'AddList "NY " & Format(CDate(d))
    d = ConvertTimeZone(Now)
    AddList "NY " & Format(CDate(d))
    d = ConvertTimeZone(d, "NY", "")
    AddList "Local " & Format(CDate(d))
    d = ConvertTimeZone(Now, "", "GMT")
    AddList "GMT " & Format(CDate(d))
    d = ConvertTimeZone(Now, "", "600|1987,10/LS,3/3S|1989,10/LS,3/1S|1995,10/LS,3/LS|2000,8/LS,3/LS|2001,10/LS,3/LS")
    AddList "GMT " & Format(CDate(d))
    
    d = Now
    If Not gdIsDaylightSavingTime(d, "") Then
        AddList "Not DS " & Format(CDate(d))
    End If
    d = Now - 6 * 30
    If gdIsDaylightSavingTime(d, "") Then
        AddList "DS " & Format(CDate(d))
    End If

Exit Sub
    
    Set t = GetMasterFileMatches("c:\export", "csi")
    For i = 0 To t.NumRecords - 1
        AddList Str(i) & vbTab & t(0, i) & vbTab & t(1, i) & vbTab & t(2, i) & vbTab & t(3, i) & vbTab & t(4, i) & vbTab & t(5, i) & vbTab & t(6, i) & vbTab & t(7, i)
    Next
    
    'InfBox "test", , , , True
Exit Sub
    UpdateGenTick
Exit Sub
    
    nc = 1 '00
    n = 5000
    ReDim vbArray(n) As Double
    gdArray.Size = n
    
    For i = 0 To n - 1
        vbArray(i) = i + 0.5
        gdArray(i) = i + 0.5
    Next
    
    'nStart = GetTickCount
    DoEvents
    AddList "Start"
    d = 0
    For c = 1 To nc
        For i = 0 To n - 1
            d = d + vbArray(i)
        Next
    Next
    AddList "VB done", True
    
    d = 0
    h = gdArray.ArrayHandle
    For c = 1 To nc
        For i = 0 To n - 1
            d = d + gdGetNum(h, i)
        Next
    Next
    AddList "gdGetNum done", True
    
    d = 0
    h = gdArray.ArrayHandle
    For c = 1 To nc
        For i = 0 To n - 1
            d = d + gdGetNum(h, gdArray.ArrayHandle)
        Next
    Next
    AddList "gdGetNum2 done", True
    
    d = 0
    For c = 1 To nc
        For i = 0 To n - 1
            d = d + gdArray.Num(i)
        Next
    Next
    AddList "GD.Num done", True
    
    d = 0
    For c = 1 To nc
        For i = 0 To n - 1
            d = d + gdArray(i)
        Next
    Next
    AddList "GD done", True


ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest.Command10_Click"
    d = 0
    Resume ErrExit
End Sub

Private Sub Command9_Click()

    Dim i&, d#, dt#, s$
    Dim a As New cGdArray, r As New cGdArray
    Dim Stats As gdArrayStatistics
    
    's = "Exported Symbol Group" & vbTab & "IBM,GOOG,MSFT,AAPL"
    
    Dim nFrom&, nTo&
    
    nFrom = DateOf(InfBox("Recalc criteria FROM:", "?", , "Recalc Criteria", , , , , , "d", DateFormat(Date)))
    If nFrom > LastDailyDownload Then Exit Sub
    nTo = DateOf(InfBox("Recalc criteria TO:", "?", , "Recalc Criteria", , , , , , "d", DateFormat(Date)))
    If nTo > LastDailyDownload Then nTo = LastDailyDownload
    For i = nFrom To nTo
        If IsWeekday(i) Then
            AddList "Calculating criteria for " & DateFormat(i)
            g.SymbolPool.DirtyCriteria = True
            g.SymbolPool.RecalcDirtyCriteria i
        End If
    Next
    AddList "DONE"
    
Exit Sub
    
    a(0) = "x"
    i = DM_TestFunction(g.DMS, a.ArrayHandle)
Exit Sub

    If Len(InternetBrowser) > 0 Then
        RunProcess InternetBrowser, Chr(34) & "c:\dvlp\genesis\navigator suite\info\news.htm" & Chr(34)
    End If
Exit Sub

    g.RealTime.Reconnect 10
Exit Sub
    
    For i = 1 To 86400000
        d = Date + (i + 0.5) / 86400000#
        dt = gdFixDateTime(d)
        If d <> dt Then
            d = d - dt
            Exit For
        End If
    Next
    AddList "Done"
    Exit Sub
    
    
    a.Add 4
    a.Add 5
    a.Add 8
    a.Add 2
    a.Add 5
    a.Add 9
    a.Add 3
    a.Add 4.5
    
    Set r = a.CalcMovingStatistic(eGdStat_StdDev, 0)
    For i = 0 To r.Size - 1
        AddList Str(i) & vbTab & r(i)
    Next
    AddList a.CalcStatistic(eGdStat_StdDev)
    If gdIsConstantValue(r.ArrayHandle) Then
        i = i
    End If
    Exit Sub
    
    r.Create eGDARRAY_Doubles
    If gdCalcMovingStatistic(r.ArrayHandle, a.ArrayHandle, eGdStat_StdDev, 4) Then
        For i = 0 To r.Size - 1
            AddList r(i)
        Next
    End If
    
    
    Exit Sub
    
    If gdCalcStatistics(a.ArrayHandle, Stats, True, 0, -1) Then
    'If a.CalcStatistics(stats) Then
        With Stats
            AddList .Average
            AddList .StdDev
            AddList .Skewness
            AddList .Kurtosis
            AddList .AvgDev
            AddList .SumOfSquares
        End With
    End If

End Sub

Private Sub fg_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    GridScrollCheck fg, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel

End Sub

Private Sub Form_Load()

    Dim i&, d#
    i = Me.hWnd
    
    g.Styler.StyleForm Me

'Me.Command1.Style = 0


    i = Command1.Style
    i = &H8000000F

    Me.BackColor = Me.Frame1.BackColor
    FixFormControls Me
    Me.Width = 7200


    For i = 1 To 20
        Combo1.AddItem "Item " & Str(i)
    Next
    Combo1.ListIndex = 0
    'Combo1.Left = -600
    i = SendMessage(Combo1.hWnd, 352, Combo1.Width / Screen.TwipsPerPixelY + 40, 0&)
    
    AddList "Ram = " & Str(Int(PhysicalRAM(True))) & " of " & Str(Int(PhysicalRAM(False))) & " mb"
    

#If 0 Then
    fg.Width = 10000
    fg.ColFormat(2) = "$#,##0.00"
    fg.ColFormat(3) = "Currency"
    fg.ColDataType(4) = flexDTCurrency
    fg.ColDataType(7) = flexDTCurrency
    For i = 1 To 9
        fg.TextMatrix(0, i) = Str(i)
    Next
    
    d = 1234.5
    For i = 1 To 3
        d = d + 10
        fg.TextMatrix(i, 2) = d
        fg.TextMatrix(i, 3) = d
        fg.TextMatrix(i, 4) = d
        fg.TextMatrix(i, 5) = Format(d, "$#,##0.00")
        fg.TextMatrix(i, 6) = Format(d, "Currency")
        fg.TextMatrix(i, 7) = Format(d, "$#,##0.00")
    Next
    
#End If

'SavePicture Image2.Picture, "c:\test.ico"

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    lst.Height = Me.ScaleHeight - lst.Top * 2
    lst.Width = Me.ScaleWidth - lst.Left - lst.Top

End Sub

Public Sub AddList(ByVal strMsg$, Optional ByVal bBenchMark As Boolean = False)
    
    Dim dTicks#, i&, iPos&, strLine$
    Static dPrevTicks#
    
    If g.bUnloading Then Exit Sub '(so next line won't reload this form when app is unloading)
    If Not IsIDE And Not Me.Visible Then Exit Sub
    
'If IsNumeric(Left(strMsg, 1)) Then Exit Sub
'strMsg = Format(Now, "hh:mm:ss ") & strMsg
    
    If bBenchMark Then
        dTicks = gdTickCount
        If Len(strMsg) > 0 Then
            strMsg = strMsg & ":  " & Format((dTicks - dPrevTicks) / 1000, "0.000") & " seconds"
        End If
    End If
    
    With lst
        Do While Len(strMsg) > 0
            If .ListCount > 25000 Then
                .AddItem "*** REMOVING LINES FROM LISTBOX ***"
                .ListIndex = .ListCount - 1
                'RH commented out .Refresh
                For i = 1000 To 0 Step -1
                    .RemoveItem i
                Next
            End If
            iPos = InStr(strMsg, vbCrLf)
            If iPos > 0 Then
                strLine = Left(strMsg, iPos - 1)
                strMsg = Mid(strMsg, iPos + 2)
            Else
                strLine = strMsg
                strMsg = ""
            End If
            .AddItem strLine
            If .ListIndex >= .ListCount - 3 Then
                .ListIndex = .ListCount - 1
                'RH commented out .Refresh
            End If
        Loop
    End With
    dPrevTicks = gdTickCount
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim i&
    i = Me.hWnd

    Timer1.Enabled = False
    
End Sub

Private Sub lst_DblClick()
    lst.Clear
End Sub

Public Function BarDisplay(Bars As cGdBars, ByVal nBar&) As String

    Dim s$, dDate#, dVol#
    
    dDate = Bars(eBARS_DateTime, nBar)
    If dDate <= 0 Then
        s = vbTab
    ElseIf IsIntraday(Bars.Prop(eBARS_Periodicity)) Then
        s = DateFormat(dDate, MM_DD_YY) & Format(Bars(eBARS_DateTime, nBar), " HH:NN")
    Else
        s = DateFormat(dDate, MM_DD_YYYY)
    End If
    dVol = Bars(eBARS_Vol, nBar)
    If dVol < 0 Then dVol = -1
    If Bars.ArrayMask = eBARS_Profiled Then
        s = s & vbTab & Bars.PriceDisplay(Bars(eBARS_Close, nBar), True) _
            & vbTab & " " & Format(dVol, "#0") & _
            vbTab & Format(Bars(eBARS_BidVol, nBar), "#0") & _
            vbTab & Format(Bars(eBARS_AskVol, nBar), "#0") & _
            vbTab & Format(Bars(eBARS_Flags, nBar), "#0")
    Else
        s = s & vbTab & Bars.PriceDisplay(Bars(eBARS_Open, nBar), True) _
            & vbTab & Bars.PriceDisplay(Bars(eBARS_High, nBar), True) _
            & vbTab & Bars.PriceDisplay(Bars(eBARS_Low, nBar), True) _
            & vbTab & Bars.PriceDisplay(Bars(eBARS_Close, nBar), True) _
            & vbTab & " " & Format(dVol, "#0")
        If IsIntraday(Bars.Prop(eBARS_Periodicity)) Then
            dVol = Bars(eBARS_UpTicks, nBar) + Bars(eBARS_DownTicks, nBar)
            If dVol < 0 Then dVol = -1
            s = s & vbTab & " " & Format(dVol, "#0")
        End If
    End If

    BarDisplay = s
End Function

Private Sub Timer1_Timer()

    Dim h&
    h = GetForegroundWindow
    StatusMsg "fg = " & Str(h)

End Sub

Private Sub txtZip_Change()

    Static strText$

    If txtZip.Text <> strText Then
        strText = txtZip.Text
        AddList strText
        DoEvents
    End If

End Sub

Private Sub DoGradientColors()

#If 0 Then
    Dim i&, iColor&
    
    lst.Visible = False
    fg.Visible = True
    fg.FixedRows = 0
    fg.Rows = 17
    For i = 0 To 16
        iColor = GradientColor(100 * i / 16, gdColorFrom.Color, gdColorTo.Color)
        fg.Cell(flexcpBackColor, i, 1) = iColor
        fg.TextMatrix(i, 0) = Str(iColor)
        fg.Cell(flexcpBackColor, i, 2) = GradientColor(100 * i / 16, gdColorFrom.Color)
    Next
#End If

End Sub


'Automated Data Checking (just after a new minute starts):
'- get list of symbols (all popular 57's and indices)
'- get eSignal refresh
'- do QB refresh
'then for each symbol:
'- build 1-minute bars from Genesis and eSignal
'- for each eSignal 1-minute bar, see if higher/lower than Genesis bars which are within 1 minute
'- for each Genesis 1-minute bar, see if higher/lower than eSignal bars which are within 1 minute
Public Function AutoDataCheck(aSymbolsToCheck As cGdArray) As cGdArray
On Error GoTo ErrSection

    Dim i&, iSymbol&, iSig&, iGen&, strSymbol$
    Dim dtBar#, dHigh#, dLow#
    Dim SigBars As New cGdBars, GenBars As New cGdBars
    Dim aMismatch As cGdArray

    ' array of mismatches to return
    Set aMismatch = New cGdArray
    aMismatch.Create eGDARRAY_Strings, 0
    
    ' for each symbol in the list
    For iSymbol = 0 To aSymbolsToCheck.Size - 1
        strSymbol = Trim(UCase(aSymbolsToCheck(iSymbol)))
        If Len(strSymbol) > 0 Then
            ' get 1-minute bars for both (eSignal and Genesis)
            SigBars.BuildBars "1 min"
            DM_GetBars GenBars, strSymbol, "1 min", LastDailyDownload + 1
            
            ' first check each 1-minute bar of the eSignal data (to check for missing Genesis data)
            For iSig = 0 To SigBars.Size - 1
                ' find the closest bar timewise
                dtBar = SigBars(eBARS_DateTime, iSig)
                iGen = GenBars.FindDateTime(dtBar)
                ' get our highest and lowest of the bars which are within 1 minute
                dHigh = -999999999
                dLow = 999999999
                For i = iGen - 2 To iGen + 2
                    If Abs(GenBars(eBARS_DateTime, i) - dtBar) < 1.1 / 1440 Then
                        If dHigh < GenBars(eBARS_High, i) Then
                            dHigh = GenBars(eBARS_High, i)
                        End If
                        If dLow < GenBars(eBARS_Low, i) Then
                            dLow = GenBars(eBARS_Low, i)
                        End If
                    End If
                Next
                ' see if eSignal has a higher high or lower low than the Genesis data
                If SigBars(eBARS_High, iSig) > dHigh + 0.00000001 Then
                    aMismatch.Add strSymbol & " at " & Format(dtBar, "HH:MM") & ": eSignal has a higher high = " & Str(SigBars(eBARS_High, iSig))
                End If
                If SigBars(eBARS_Low, iSig) < dLow - 0.00000001 Then
                    aMismatch.Add strSymbol & " at " & Format(dtBar, "HH:MM") & ": eSignal has a lower low = " & Str(SigBars(eBARS_Low, iSig))
                End If
            Next iSig
                
            ' then check each 1-minute bar of the Genesis data (to check for Genesis spikes)
            For iGen = 0 To GenBars.Size - 1
                ' find the closest bar timewise
                dtBar = GenBars(eBARS_DateTime, iGen)
                iSig = SigBars.FindDateTime(dtBar)
                ' get their highest and lowest of the bars which are within 1 minute
                dHigh = -999999999
                dLow = 999999999
                For i = iSig - 2 To iSig + 2
                    If Abs(SigBars(eBARS_DateTime, i) - dtBar) < 1.1 / 1440 Then
                        If dHigh < SigBars(eBARS_High, i) Then
                            dHigh = SigBars(eBARS_High, i)
                        End If
                        If dLow < SigBars(eBARS_Low, i) Then
                            dLow = SigBars(eBARS_Low, i)
                        End If
                    End If
                Next
                ' just ignore this check if eSignal is missing data for this timeframe
                If dHigh <> -999999999 Then
                    ' see if Genesis is more than a couple ticks higher or lower than the eSignal data
                    ' (allow for just a little higher/lower in case eSignal happens to be spotty)
                    If GenBars(eBARS_High, iGen) > dHigh + 0.00000001 + GenBars.MinMove * 2 Then
                        aMismatch.Add strSymbol & " at " & Format(dtBar, "HH:MM") & ": Genesis may have a high spike = " & Str(GenBars(eBARS_High, iGen))
                    End If
                    If GenBars(eBARS_Low, iGen) < dLow - 0.00000001 - GenBars.MinMove * 2 Then
                        aMismatch.Add strSymbol & " at " & Format(dtBar, "HH:MM") & ": Genesis may have a low spike = " & Str(GenBars(eBARS_Low, iGen))
                    End If
                End If
            Next iGen
        End If
    Next iSymbol
    
ErrExit:
    Set AutoDataCheck = aMismatch
    Exit Function
    
ErrSection:
    RaiseError "AutoDataCheck"
    Resume ErrExit
End Function

Public Function ConvertStockOptionDatFile(ByVal strFile$) As Long

    Dim i&, strText$, strCode$, nChangeCount&, nDate&, strStrike$, strStock$
    Dim aStrings As New cGdArray
    
    aStrings.FromFile strFile
    For i = 0 To aStrings.Size - 1
        strText = UCase(Trim(aStrings(i)))
        'from: @MSFT-201201/P27.5 (WMFMY)
        '  to: @WMF/20120121/P27.5 (MSFT)
        If Left(strText, 1) = "@" And InStr(strText, "-") > 0 Then
            strCode = Parse(strText, " ", 2)
            If Left(strCode, 1) = "(" And Right(strCode, 1) = ")" Then
                ' strip off parens
                strCode = Trim(Mid(strCode, 2, Len(strCode) - 2))
                If Len(strCode) <= 5 Then
                    ' strip off last 2 characters
                    strCode = Trim(Left(strCode, Len(strCode) - 2))
                    ' parse out stock, date and strike
                    strText = Parse(strText, " ", 1)
                    strStock = Mid(Parse(strText, "-", 1), 2)
                    strText = Parse(strText, "-", 2) ' 201201/P27.5
                    nDate = Val(Parse(strText, "/", 1))
                    strStrike = Parse(strText, "/", 2)
                    If Len(strStrike) > 0 And nDate > 200000 And Len(strStock) > 0 Then
                        If nDate <= 999999 Then
                            ' add day to the date
                            nDate = GetDateFromRule(Int(nDate / 100), nDate Mod 100, "3F") + 1
                            nDate = JulToLong(nDate, True)
                        End If
                        strText = "@" & strCode & "/" & Str(nDate) & "/" & strStrike _
                            & " (" & strStock & ")"
                        nChangeCount = nChangeCount + 1
                        aStrings(i) = strText
                    End If
                End If
            End If
        End If
    Next
    
    If nChangeCount > 0 Then
        aStrings.ToFile strFile
    End If
    
    ConvertStockOptionDatFile = nChangeCount
    
End Function

Public Function SetDepOff() As Boolean
    
    Dim i&, hProcess&, dwFlags&, dwPerm&, s$, bReturn As Boolean
    On Error Resume Next
    bReturn = True
    'If IsAtLeastVista Then
    If 1 Then ' IsAtLeastXP Then
        i = -1
        ' 0 = Always Off, 1 = Always On, 2 = Only selected apps, 3 = Except selected apps
        i = GetSystemDEPPolicy
        AddList "GetSystemDEPPolicy = " & Str(i)
        
        dwFlags = -1
        dwPerm = -1
        i = -1
        ' Flags: 0 = DEP disabled, >0 = DEP enabled
        i = GetProcessDEPPolicy(GetCurrentProcess, dwFlags, dwPerm)
        If i <> 0 Then AddList "GetProcessDEPPolicy = " & Str(dwFlags)
If 0 Then
        If i <> 0 And dwFlags <> 0 Then
            ' try to set DEP off for this process
            i = -1
            i = SetProcessDEPPolicy(0) ' 0 = DEP off, 1 = DEP on
            AddList "SetProcessDEPPolicy = " & Str(i)
            dwFlags = -1
            dwPerm = -1
            i = -1
            i = GetProcessDEPPolicy(hProcess, dwFlags, dwPerm)
            If i <> 0 Then AddList "GetProcessDEPPolicy = " & Str(dwFlags)
            If i <> 0 And dwFlags <> 0 Then
                ' setting it Off didn't work
#If 0 Then
                s = WinSysPath & "BCDEdit.exe"
                If FileExist(s) Then
                    ''bcdedit.exe /set nx AlwaysOff
                    ''bcdedit.exe /set nx OptOut
                    'RunProcess s, "/set nx OptOut", True, vbHide, i
                    'AddList "BCDEdit = " & Str(i)
                End If
                bReturn = False
#End If
            End If
        End If
End If
        
If 0 Then
        i = -1
        i = GetSystemDEPPolicy
        AddList "GetSystemDEPPolicy = " & Str(i)
        
        If i = 1 Then
        
        ElseIf i >= 2 Then
            hProcess = GetCurrentProcess
            dwFlags = -1
            dwPerm = -1
            i = GetProcessDEPPolicy(hProcess, dwFlags, dwPerm)
            
            i = SetProcessDEPPolicy(1)
                        
            i = GetProcessDEPPolicy(hProcess, dwFlags, dwPerm)
            
            i = SetProcessDEPPolicy(0)
            
            i = GetProcessDEPPolicy(hProcess, dwFlags, dwPerm)
            AddList "SetProcessDEPPolicy = " & Str(i)
            
        End If
End If
    End If
End Function

Private Sub Test()
On Error GoTo ErrSection:

    Dim i&

    i = 3 / i

ErrExit:
    i = 9
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & "Test"
    i = 8
    Resume ErrExit
End Sub

' Max trades/second so far: 4582 on 8/9/2011 at 16:00:00 ET  (3028 on 6/29/2010 at 10am ET)
' Max price changes per second so far: 528 on 8/9/2011 at 16:00:00 ET  (335 on 8/24/2010 at 10am ET)
' Max trades/day so far: 1,595,461 on 8/9/2011
Private Sub MaxTradesPerSecond()

    Dim i&, nMax&, nDate&, dtTick#, dtPrevTick#, nTicks&, nSeconds&
    Dim nPriceChanges&, nMaxPriceChanges&, dPrice#, dPrevPrice#, nMaxTrades&
    Dim strSymbol$
    Dim b As New cGdBars
    
    strSymbol = "ES-067"
'strSymbol = "G6E-067"
    nDate = DateSerial(2010, 2, 9)
    nDate = DateSerial(2010, 1, 13)
    nDate = DateSerial(2010, 2, 23) ' 2103 at 10am ET
    nDate = DateSerial(2010, 5, 7) ' 1800 at 4pm ET
nDate = DateSerial(2011, 8, 1)
    
nDate = DateSerial(2013, 1, 1)
    
'nDate = Date - 7
    
    For nDate = nDate To Date ' LastDailyDownload
        If IsWeekday(nDate) Then
            If nDate > LastDailyDownload Then
                GetAvailTickData b, i, strSymbol, GetSymbolID(strSymbol), nDate, 0
            Else
                DM_GetBars b, strSymbol, "Each tick", nDate, nDate
            End If
            AddList Format(nDate, "mm/dd/yyyy") & " = " & Str(b.Size) & " trades"
            If b.Size > nMaxTrades Then
                nMaxTrades = b.Size
                If nMaxTrades > 1000000 Then
                    AddList Str(nMaxTrades) & " trades on " & Format(nDate, "mm/dd/yyyy")
                End If
            End If
            
            nTicks = 0
            dPrevPrice = 0
            nPriceChanges = 0
            For i = 0 To b.Size ' make sure to get the final second
                dtTick = b(eBARS_DateTime, i)
                If dtTick <> dtPrevTick Then
                    'If nTicks > 0 Then
                    '    nSeconds = Int((dtTick - dtPrevTick) * 86400 + 0.5)
                    '    If nSeconds > 1 And nSeconds < 60 Then
                    '        nTicks = Int(nTicks / nSeconds + 0.5)
                    '    End If
                    'End If
                    If nTicks > nMax Then
                        nMax = nTicks
                        AddList Str(nMax) & " trades/sec at " & Format(dtPrevTick, "mm/dd/yyyy hh:mm:ss")
                    End If
                    If nPriceChanges > nMaxPriceChanges Then
                        nMaxPriceChanges = nPriceChanges
                        AddList Str(nMaxPriceChanges) & " price changes at " & Format(dtPrevTick, "mm/dd/yyyy hh:mm:ss")
                    End If
                    dtPrevTick = dtTick
                    dPrevPrice = 0
                    nTicks = 0
                    nPriceChanges = 0
                End If
                If i < b.Size Then
                    nTicks = nTicks + 1
                    dPrice = b(eBARS_Close, i)
                    If dPrice <> dPrevPrice Then
                        dPrevPrice = dPrice
                        nPriceChanges = nPriceChanges + 1
                    End If
                End If
            Next
            DoEvents
        End If
    Next

End Sub

Private Sub DumpDailyData(Optional ByVal nSessionDate& = 0, Optional ByVal nNumDays& = 1)

    Dim i&, iRec&, iSymID&, strSymbol$, s$, nEndDate&, iPercent&
    Dim aLines As New cGdArray
    Dim b As New cGdBars
    Static bInProgress As Boolean
    
    ' see if aborting
    If bInProgress Then
        bInProgress = False
        Exit Sub
    End If
    If nSessionDate = 0 Then
        nSessionDate = DateOf(InfBox("Dump daily data for:", "?", , "Daily Data", , , , , , "date", DateFormat(LastDailyDownload)))
        If nSessionDate <= 0 Then Exit Sub
        nNumDays = Val(InfBox("Number of days:", "?", , "Daily Data", , , , , , "number", "1"))
    End If
    If nNumDays < 1 Then Exit Sub
    bInProgress = True
    
    nEndDate = nSessionDate
    For i = 2 To nNumDays
        Do
            nEndDate = nEndDate + 1
        Loop While Not IsWeekday(nEndDate)
    Next
    
    lst.Clear
    AddList "Dump daily data for " & DateFormat(nSessionDate)
    i = g.SymbolPool.NumRecords
    iPercent = 0
    For iRec = 0 To g.SymbolPool.NumRecords - 1
        iSymID = g.SymbolPool.SymbolID(iRec)
        If iSymID > 0 Then
            DM_GetBars b, iSymID, "Daily", nSessionDate, nEndDate
            If b.Size > 0 Then
                strSymbol = b.Prop(eBARS_Symbol)
                For i = 0 To b.Size - 1
                    s = Str(JulToLong(b(eBARS_DateTime, i), True)) & vbTab & strSymbol _
                        & vbTab & b.PriceDisplay(b(eBARS_Open, i), False, nSessionDate) _
                        & vbTab & b.PriceDisplay(b(eBARS_High, i), False, nSessionDate) _
                        & vbTab & b.PriceDisplay(b(eBARS_Low, i), False, nSessionDate) _
                        & vbTab & b.PriceDisplay(b(eBARS_Close, i), False, nSessionDate) _
                        & vbTab & Str(b(eBARS_Vol, i))
                    If b(eBARS_OI, 0) > 0 Then
                        s = s & vbTab & Str(b(eBARS_OI, i)) & vbTab & Str(b(eBARS_ContVol, i)) & vbTab & Str(b(eBARS_ContOI, i))
                    End If
                    aLines.Add s
                Next
            End If
        End If
        ' display % done and check for abort
        i = Int(100# * (iRec + 1) / g.SymbolPool.NumRecords)
        If iPercent <> i Then
            iPercent = i
            AddList Str(iPercent) & "% done, # of symbols = " & Str(iRec + 1)
            DoEvents
            If Not bInProgress Then Exit For ' aborted
        End If
    Next
    
    s = App.Path & "\chk\DailyData.txt"
    If bInProgress Then
        aLines.Sort
        aLines.ToFile s
        AddList "File = " & s
    Else
        AddList "Aborted"
        KillFile s
    End If
    bInProgress = False
    
End Sub

Private Function CleanSymbolDesc(ByVal strDesc$) As String
    
    ' strip off the "Cont Liq" part of the description (Liq CAdj Cont Elec Exp)
    strDesc = strDesc & " "
    strDesc = Replace(strDesc, " Liq ", " ", , , vbTextCompare)
    strDesc = Replace(strDesc, " CAdj ", " ", , , vbTextCompare)
    strDesc = Replace(strDesc, " Cont ", " ", , , vbTextCompare)
    strDesc = Replace(strDesc, " Elec ", " ", , , vbTextCompare)
    strDesc = Replace(strDesc, " Exp ", " ", , , vbTextCompare)
    strDesc = Replace(strDesc, " Comb ", " ", , , vbTextCompare)
    strDesc = Replace(strDesc, " (", vbTab, , , vbTextCompare) & " "
    strDesc = Trim(Parse(strDesc, vbTab, 1))
    CleanSymbolDesc = strDesc

End Function

' OVERALL LOGIC for calculating Seasonal Trade data:
' first get list of all symbols (including spreads) to do
' then for each Symbol (or spread)
'   for both Long and Short (except only Long for spreads)
'     ' store trade info using 3D tables (TDOY, #DaysHeld, StopLossLevel)
'     for each Bar in history
'       for each StopLoss Level
'         for each #DaysHeld
'           ' accumulate trade info in the 3D tables
'           ' (but once stop-loss is hit, it's a loss for rest of DaysHeld)
'         next
'       next
'     next
'     ' write the trade output for each TDOY/#DaysHeld/StopLoss (from the 3D tables)
'     for each TDOY
'       for each #DaysHeld
'         for each StopLoss Level
'           ' calc the trade stats from the 3D tables
'           ' write output for this seasonal trade if it's a "good enough" winner
'           ' (but ignore it if same exact results as at the previous lower stop-loss for same TDOY and #DaysHeld)
'         next
'       next
'     next
'   next
' next
'
' STATS for each Seasonal Trade (each row in the table of results used for seasonal queries):
'   Symbol, SymbolID, LongShort, EntryTDOY, EntryDate, #DaysHeld, ExitDate, StopLoss, #Trades (i.e. #Years),
'   Win% (#Wins/#Trades), Profit/Loss Ratio (TotalWon/TotalLost), AvgTrade (% or $), AvgProfitPerDay (% or $),
'   AvgDrawdown (% or $), WorstDrawdown (% or $), Avg Annual Profit/Drawdown, Kelly% (AvgTrade/AvgWin),
'   and delimited string of the TradeHistory (for each trade: EntryDate, EntryPrice, ExitDate, ExitPrice, Profit, Drawdown)
'
' Store trade stats separately for Stocks (using %'s), Futures (using $'s), and Future Spreads (intermarket and calendar, using $'s)
'
' NUMBER of different sets of trades to calculate ...
' for each symbol: for long/short, for each TDOY, for each #DaysHeld, for each StopLoss = 2 * 250 * 125 * 10 = 625,000
' Futures = 52'ish markets (all the tradeable markets) = 33 million  (takes a couple of hours to run)
' Stocks = 5500'ish markets (with at least 5 years of data) = 3.5 billion  (requires a weekend to run, 50-60 hours)
' Spreads = 1500'ish markets (but no shorts) = 625,000 / 2 * 1500 = 500 million  (can run overnight)
' (but only need to store the "good enough" winners -- e.g. ProfitFactor > x and Win% > x -- so just a fraction of the above #'s)
'
' For STOCKS: use split-adjusted data, and calculate everything as percentages (rather than dollars)
' For FUTURES: use -067 data, calculate everything as dollars (rather than %'s), but display all entry prices using the -057 data
'       and display the exit prices using a "forward roll-adjust amount" (i.e. the -057 minus the -067 at the time of the entry)
' For all futures SPREADS:
' - cannot use intraday highs/lows for any kind of spread, so stop-loss amounts are only checked at end of day
' - and need to use the "settles" instead of the "last price", esp. for calendar spreads since a further out leg could be so
'       illiquid that the last trade could be early morning causing an unfair comparison with a closing price of the other legs,
'       so need to use Combined symbols for spreads instead of the Electronic symbols
' For Intermarket SPREADS:
' - just spread each major symbol with each of the other major symbols from within it's own category (e.g. energies, grains, etc)
' - must calculate dollars for each leg separately (i.e. would be INCORRECT to just use the up/down movement of the spread price)
' - assume that intermarket spreads can be held across rolls, so use -067 data for each leg (along with the "forward roll-adjust")
' For Calendar SPREADS (a little trickier):
' - for each major symbol, just spread each monthly contract with each of the other monthly contracts up to 1 year out
' - can use the Gann symbols (-081 thru -092) for the data
' - BUT we cannot allow a trade to span across the rolls for either leg (meaning a trade MUST exit prior to the next roll for either leg)
' - when doing the history, only allow an entry when a price exists for both legs (but can exit even if no valid prices),
'       and also make sure to always exit the trade prior to the next roll of either leg
' - determine the min/max TDOY for the roll of each leg in order to predict the next roll for each (i.e. within the forecast year)
' - then when doing the output, create "dead zones" from 5 TDOY's before the min roll TDOY until after the month of the contract expiration,
'       meaning there should be no predicted entries during the upcoming year's "dead zones", and force the exits prior to the next dead zone.
Private Sub CalcSeasonals() ' data for Seasonal Sweet Spots

    ' for the 3D tables (3d arrays)
    Const kMaxTDOY As Long = 262 ' 262 = max # of weekdays in a year
    Const kMaxDaysHeld As Long = 125 ' about 6 months (=125 trading days) for longest held seasonal trade
    Const kMaxStopLosses As Long = 10 ' max of 10 different stop-loss levels

    ' for the 2D BarsData (arrays)
    Const kTDOY As Long = 0
    Const kDate As Long = 1
    Const kClose57 As Long = 2
    Const kClose As Long = 3
    Const kHigh As Long = 4
    Const kLow As Long = 5
    Const kOpen As Long = 6
    ' when doing spreads, we can re-purpose the High/Low/Open arrays (since unused for spreads)
    Const kSpread57 As Long = 4
    Const kSpread As Long = 5
    Const kRolled As Long = 6 ' i.e. rolled on this date: 0 = none, 1 = leg#1, 2 = leg#2, 3 = both legs rolled

    ' these are the 3D tables used to accumulate all the trade info for each TDOY/#DaysHeld/StopLossLevel:
    Dim tblNumUp() As Long
    Dim tblNumDown() As Long
    Dim tblAmtUp() As Double
    Dim tblAmtDown() As Double
    Dim tblDD() As Double
    Dim tblWorstDD() As Double
    
    ' to store the Trade History (each trade of each year for each TDOY for each #DaysHeld for each StopLoss)
    ' (BUT need to be careful of memory: 262 tdoys x 125 daysheld x 10 stoploss x 50 years x 20 bytes for the 5 arrays = 330 megs)
    Dim tblTradeEntryBar() As Integer ' entry bar# (in the BarsData array, which provides both the entry date and entry price)
    Dim tblTradeExitBar() As Integer ' exit bar#
    Dim tblTradeExit() As Long ' exit price (since not always the same as the BarsData due to forward-adjusting, stopped-out, etc)
    Dim tblTradeExit2() As Long ' exit price of spread leg
    Dim tblTradeProfit() As Single ' profit of trade (as either % or $)
    Dim tblTradeDD() As Single ' drawdown of trade (as either % or $)
    
    Dim BarsData() As Double ' the 2D table used to line up all the data into a consistent TDOY grid
    Dim aTradeInfo() As String ' to store prebuilt strings with formatted date and price for trades (just for some efficiency)

    ' other variables
    Dim i&, j&, iRow&, iSymbol&, iBar&, iDaysHeld&, iStopLoss&, nSymbolID&, iTDOY&, iYear&, iNumBars&, iRoll&, iEndDateBar&
    Dim d#, dEntry#, dExit#, iNumTrades&, iLineCount&, iLineCountStart&
    Dim nEndDate&, nStartDate&, iEntryBar&, iExitBar&, strDirection$, strName$, strDesc$
    Dim bShort As Boolean, bIncludeTradeHistory As Boolean
    Dim fh&, dNet#, dAvgNet#, dWinPct#, dDollarsPerPoint#, dPF#, dAvgWin#, dAvgLoss#, dStopLoss#, dStopPrice#, dTimeStarted#
    Dim dLowest#, dHighest#, dDD#, dAvgDD#, dWorstDD#, dAAP2DD#, dChk#, dRollAdjust#, dLow#, dHigh#, dOpen#
    Dim dFileSize#, dMinWinPerc#, dMinPF#, dMinMove#
    Dim s$, strSymbol$, strText$, strTrades$, strSecType$, strSymbolInfoFile$, strDataFile$, strPath$, strExch$
    Dim bUsePercentage As Boolean, bSkip As Boolean, bSymbolInfoDone As Boolean, bStoppedOut As Boolean
    Dim bDoStocks As Boolean, bDoSpreads As Boolean, bCalendarSpread As Boolean
    Dim Bars As New cGdBars, Bars57 As New cGdBars
    Dim aSymbolInfo As New cGdArray
    Dim aSymbols As New cGdArray

    Dim nSymbolID2&
    Dim dEntry2#, dExit2#, dDollarsPerPoint2#, dRollAdjust2#, dMinMove2#
    Dim strSymbol2$, strSpread$
    Dim Bars2 As New cGdBars, Bars57s As New cGdBars
    Dim Rolls As cGdTable

    ' this allows clicking on the button again in order to STOP this process if it's currently in progress
    Static bInProgress As Boolean
    If bInProgress Then
        If InfBox("Abort Seasonals?", "?", "Abort|+-No", "Abort") = "A" Then
            bInProgress = False
        End If
        Exit Sub
    End If
    
    ' don't let normal clients run this accidentally
    If Not FileExist("c:\common\files.exe") Then Exit Sub
    
    ' Initialize things based on which type being done (get settings from the Seasonals.INI file)
    bDoStocks = False
    bDoSpreads = False
    bUsePercentage = False
    AddList "Ram = " & Str(Int(PhysicalRAM(True))) & " of " & Str(Int(PhysicalRAM(False))) & " mb"
    s = InfBox("This runs the really long process to calculate all the Seasonals data.", "?", "Stocks|Futures|S&preads", "Seasonal Calcs")
    If s = "F" Then
        strSecType = "Futures"
        strDataFile = "FutSeasonals.dat"
        strSymbolInfoFile = "FutSymbols.dat"
    ElseIf s = "S" Then
        bDoStocks = True
        bUsePercentage = True
        strSecType = "Stocks"
        strDataFile = "StkSeasonals.dat"
        strSymbolInfoFile = "StkSymbols.dat"
    Else
        bDoSpreads = True
        strSecType = "Spreads"
        strDataFile = "SpreadSeasonals.dat"
        strSymbolInfoFile = "SpreadSymbols.dat"
    End If
    strSymbol = GetIniFileProperty("Symbols", "", strSecType, App.Path & "\Seasonals.INI")
    dMinPF = GetIniFileProperty("MinPF", 0, strSecType, App.Path & "\Seasonals.INI")
    dMinWinPerc = GetIniFileProperty("MinWin%", 0, strSecType, App.Path & "\Seasonals.INI")
    iYear = GetIniFileProperty("StartYear", 1970, strSecType, App.Path & "\Seasonals.INI")
    nStartDate = DateSerial(iYear, 1, 1) ' convert StartYear to the starting date
    strPath = AddSlash(GetIniFileProperty("OutPath", "C:\", strSecType, App.Path & "\Seasonals.INI"))
    d = GetIniFileProperty("TradeHistory", 1, strSecType, App.Path & "\Seasonals.INI")
    If d = 0 Then
        bIncludeTradeHistory = False
    Else
        bIncludeTradeHistory = True
    End If
    If dMinPF < 1 Or dMinWinPerc < 50 Or Len(strSymbol) = 0 Or Len(strPath) = 0 Then
        AddList "Invalid settings from Seasonals.INI file"
        Exit Sub
    End If
    strDataFile = strPath & strDataFile
    strSymbolInfoFile = strPath & strSymbolInfoFile
        
    aSymbols.Size = 0
    strExch = ""
    If bDoSpreads Then
        ' load all the SPREAD symbols
        If InStr(strSymbol, "-") > 0 Then
            ' must just be testing with some specific symbols
            aSymbols.SplitFields strSymbol, ","
        Else
            ' otherwise this is the list of Groups (e.g. "Energies,Metals,Grains,Meats,Treasuries,Currencies,Equities,*Softs")
            strExch = strSymbol
            For d = 1 To 99
                ' get category from list
                strText = Parse(strExch, ",", d)
                If Len(strText) = 0 Then Exit For
                If Left(strText, 1) = "*" Then
                    bSkip = True ' skip Intermarket spreads for this category
                    strText = Mid(strText, 2)
                Else
                    bSkip = False
                End If
                ' get symbols for this category
                s = GetIniFileProperty(strText, "", strSecType, App.Path & "\Seasonals.INI")
                aSymbolInfo.SplitFields s, ","
                For iSymbol = 0 To aSymbolInfo.Size - 1
                    strSymbol = aSymbolInfo(iSymbol)
                    If Not bSkip Then
                        ' add an Intermarket spread with each of the other symbols in the same category
                        For i = 0 To aSymbolInfo.Size - 1
                            If i <> iSymbol Then
                                strSymbol2 = aSymbolInfo(i)
                                ' e.g. "G6E - G6J"
                                aSymbols.Add strSymbol & " - " & strSymbol2
                            End If
                        Next
                    End If
                    ' and add a Calendar spread for each combination of monthly contracts
                    For i = 1 To 12
                        ' for each month, check if a Gann contract exists
                        s = strSymbol & "-" & Format(80 + i, "000")
                        nSymbolID = GetSymbolID(s)
                        If nSymbolID > 0 Then
                            For j = 1 To 12
                                If j <> i Then
                                    s = strSymbol & "-" & Format(80 + j, "000")
                                    nSymbolID2 = GetSymbolID(s)
                                    If nSymbolID2 > 0 Then
                                        ' e.g. "G6E: Mar - Dec"
                                        s = MonthName(i, True, True) & " - " & MonthName(j, True, True)
                                        aSymbols.Add strSymbol & ": " & s
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
            Next
        End If
    Else
        ' for NON-SPREADS: load the Symbols array (e.g. IBM, DOW30.GRP, ES-067, ES-)
        aSymbols.SplitFields UCase(strSymbol), ","
        ' first replace any *.GRP entry with the list of symbols in that group
        For iSymbol = aSymbols.Size - 1 To 0 Step -1
            strSymbol = aSymbols(iSymbol)
            If UCase(Right(strSymbol, 4)) = ".GRP" Then
                aSymbols.Remove iSymbol
                s = g.SymbolPool.GetSymbolsForGroup(strSymbol)
                aSymbolInfo.SplitFields UCase(s), vbTab
                aSymbols.AppendFromArray aSymbolInfo
                aSymbolInfo.Size = 0
            End If
        Next
        ' then cleanup each symbol
        For iSymbol = aSymbols.Size - 1 To 0 Step -1
            strSymbol = aSymbols(iSymbol)
            If bDoStocks Then
                ' ignore foreign stocks and forex@broker symbols
                If InStr(strSymbol, "@") > 0 Then
                    strSymbol = ""
                End If
            Else
                ' only use 67's for futures (can handle "ES", "ES-", "ES-067", or "ES-blah")
                i = InStr(strSymbol, "-")
                If i = 0 Then
                    strSymbol = strSymbol & "-067"
                Else
                    strSymbol = Left(strSymbol, i) & "067"
                End If
                ' and only do the electronic 67's (now that our electronic daily history goes all the way back)
                s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                If Len(s) > 0 Then
                    strSymbol = s
                End If
            End If
            aSymbols(iSymbol) = UCase(strSymbol)
        Next
    End If
    ' then sort list, and delete any duplicates and blanks
    aSymbols.Sort (eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues)
    If aSymbols.Size = 0 Then
        AddList "No symbols (.GRP file is missing?)"
        Exit Sub
    End If
    If aSymbols.Size < 150 Then
        For i = 0 To aSymbols.Size - 1
            AddList aSymbols(i)
        Next
    End If
    AddList "# Symbols = " & Str(aSymbols.Size)
'If IsIDE Then Exit Sub
       
       
    ' set ending date -- usually about 3-4 months ago
    i = Month(Date) - 4
    If i < 1 Then i = i + 12
    i = Val(InfBox(Str(aSymbols.Size) & " symbols to process.||Use data through end of which month?|(usually about 3-4 months ago)", "?", , "Calc Seasonals", , , , , , "n", Str(i)))
    If i <= 0 Or i > 12 Then
        Exit Sub
    End If
    i = i + 1 ' so we can more easily get to end of that month, add a month now
    If i > 12 Then i = i - 12
    iYear = Year(Date)
    If i > Month(Date) Then
        iYear = iYear - 1 ' last year
    End If
    nEndDate = DateSerial(iYear, i, 1) - 1 ' end of prior month
    
    
    'To better align data bars with TDOY's:
    '- use a 2D array of doubles (always init to -999999) -- and string array (same size)
    '- fields: date, tdoy, o, h, l, c, c57 -- and string of formatted date,price57
    '- fields: date, tdoy, c, c57, s, s57 -- and string of formatted date,price57-spread57
    '- size: # of all tdoy's from start date to end date (excluding normal holidays)
    '- put all prices from gdBars into this table (for primary and spread data), merging high/low data for holidays
    '- then can walk through the table, for each row as entry, and each # of rows beyond for #DaysHeld, and store stats in TDOY table
    iBar = (nEndDate - nStartDate) * 5 / 7# + 600
    ReDim BarsData(kOpen, iBar) As Double
    ' For this kind of analysis, probably best to remove even the past "normal holidays"
    ' to provide better consistency (so all the TDOY's "line up" better throughout the year)
    strExch = "NKPGMJLTX" ' the normal 9 major holidays (at least back to 1998)
    iBar = 0
    iEndDateBar = 0
    iTDOY = 0
    For d = nStartDate To nEndDate + 600 ' need to go 6 months after 1 year window
        bSkip = True
        If IsWeekday(d) Then
            If Len(strExch) = 0 Then
                bSkip = False
            ElseIf gdIsHoliday(d, strExch) <> 0 Then
                s = DateFormat(d)
                'AddList "Holiday: " & s
            Else
                bSkip = False
            End If
        End If
        If Not bSkip Then
            ' if starting a new year, then reset the TDOY
            If Year(d) <> Year(BarsData(kDate, iBar)) Then
                iTDOY = 0
            End If
            iTDOY = iTDOY + 1
            iBar = iBar + 1
            BarsData(kTDOY, iBar) = iTDOY
            BarsData(kDate, iBar) = d
            If d <= nEndDate Then
                iEndDateBar = iBar ' last bar of "historical window" (just prior to the 1-year "forecast window")
            End If
        End If
    Next
    iNumBars = iBar
    ReDim Preserve BarsData(kOpen, iNumBars) As Double
    
#If 0 Then
    'iBar = 0
    For iBar = iBar To iNumBars
        iTDOY = BarsData(kTDOY, iBar)
        d = BarsData(kDate, iBar)
        If iTDOY <= 20 Or iTDOY >= 240 Or iBar < 10 Or iBar > iNumBars - 10 Then
            s = Str(iBar) & vbTab & Str(iTDOY) & vbTab & DateFormat(d) & vbTab & WeekdayName(d)
            AddList s
            If iTDOY = 20 Then
                AddList "==================="
            End If
        End If
    Next
    Exit Sub
#End If

    ' write header line of the data file
    fh = FreeFile
    Open strDataFile For Output As #fh
    strText = "'Symbol" & vbTab & "SymbolID" & vbTab & "Long/Short" & vbTab & "TDOY" & vbTab & "EntryDate" & vbTab _
            & "DaysInTrade" & vbTab & "ExitDate" & vbTab & "StopLoss" & vbTab & "#Trades" & vbTab _
            & "Win%" & vbTab & "ProfitFactor" & vbTab & "AvgTrade" & vbTab & "AvgPerDay" & vbTab _
            & "AvgDD" & vbTab & "WorstDD" & vbTab & "AAP2DD" & vbTab & "Kelly%" & vbTab & "TradeHistory"
    Print #fh, strText
    iLineCount = 0 ' iLineCount + 1
    dFileSize = 0
   
    ' init things for running the big loop
    bInProgress = True
    dTimeStarted = gdTickCount
    aSymbolInfo.Size = 0
    For iSymbol = 0 To aSymbols.Size - 1
        If Not bInProgress Then Exit For ' in case want to abort
        bSkip = False
        
        ' get Symbol -- check to see if it's a spread (e.g. "G6E - G6J", "G6E: Sep - Dec")
        strSymbol = aSymbols(iSymbol)
        strSpread = ""
        strSymbol2 = ""
        nSymbolID2 = 0
        bCalendarSpread = False
        If Not bDoSpreads Then
            nSymbolID = GetSymbolID(strSymbol)
        Else
            strSpread = strSymbol
            ' check for a calendar spread (e.g. "G6E: Dec - Jun", with or without spaces)
            If InStr(strSpread, ":") > 0 Then
                bCalendarSpread = True
                s = Parse(strSpread, ":", 2)
                i = MonthNumber(Parse(s, "-", 1))
                j = MonthNumber(Parse(s, "-", 2))
                If i > 0 And j > 0 Then
                    strSymbol = Parse(strSpread, ":", 1) & "-" & Format(80 + i, "000")
                    strSymbol2 = Parse(strSpread, ":", 1) & "-" & Format(80 + j, "000")
                End If
            ElseIf InStr(strSpread, "-") > 0 Then
                ' or an intermarket spread (e.g. "G6E - G6J")
                strSymbol = Trim(UCase(Parse(strSpread, "-", 1))) & "-067"
                strSymbol2 = Trim(UCase(Parse(strSpread, "-", 2))) & "-067"
            End If
            
            ' BUT for spreads we MUST use the settles from the Combined or Pit symbols (at a consistent time of day),
            ' instead of using the last price from the Electronic symbols (which could be early morning for far-out contracts)
            s = ConvertFutureSymbol(strSymbol, eCombinedSymbol)
            If Len(s) > 0 Then strSymbol = s
            nSymbolID = GetSymbolID(strSymbol)
            If nSymbolID <= 0 Then
                s = ConvertFutureSymbol(strSymbol, ePitSymbol)
                If Len(s) > 0 Then strSymbol = s
                nSymbolID = GetSymbolID(strSymbol)
                If nSymbolID <= 0 Then
                    s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                    If Len(s) > 0 Then strSymbol = s
                    nSymbolID = GetSymbolID(strSymbol)
                End If
            End If
            s = ConvertFutureSymbol(strSymbol2, eCombinedSymbol)
            If Len(s) > 0 Then strSymbol2 = s
            nSymbolID2 = GetSymbolID(strSymbol2)
            If nSymbolID2 <= 0 Then
                s = ConvertFutureSymbol(strSymbol2, ePitSymbol)
                If Len(s) > 0 Then strSymbol2 = s
                nSymbolID2 = GetSymbolID(strSymbol2)
                If nSymbolID2 <= 0 Then
                    s = ConvertFutureSymbol(strSymbol2, eElectronicSymbol)
                    If Len(s) > 0 Then strSymbol2 = s
                    nSymbolID2 = GetSymbolID(strSymbol2)
                End If
            End If
        End If
        
        strSecType = SecurityType(strSymbol, True)
        If Len(strSpread) > 0 Then
            strName = strSpread
        ElseIf bDoStocks Then
            strName = strSymbol
        Else
            strName = Parse(strSymbol, "-", 1) '& "-"
        End If
        
        ' Estimated Total (time or size) = SoFar (time or size) * TotalCount / DoneCount
        If iSymbol > 0 Then
            dChk = dFileSize * aSymbols.Size / iSymbol ' total size of data file (estimated)
            d = (gdTickCount - dTimeStarted) / 3600000# ' # hours running so far
            d = d * aSymbols.Size / iSymbol - d ' estimated # hours remaining (i.e. total - done)
        Else
            dChk = 0
            d = 0
        End If
        If nSymbolID <= 0 Then
            s = "*** " & strSymbol & " does not exist ***"
            bSkip = True
        ElseIf bDoSpreads And nSymbolID2 <= 0 Then
            s = "*** " & strSymbol2 & " does not exist (" & strName & ") ***"
            bSkip = True
        Else
            s = Str(iSymbol + 1) & " of " & Str(aSymbols.Size) & " ( " & Format(100# * iSymbol / aSymbols.Size, "#0.0") & "% ), " & strName _
                & ", FileSize = " & Str(Int(dFileSize / 1000000#)) & " of " & Str(Int(dChk / 1000000#)) & " mb, " & Format(d, "#0.00") & " hrs left"
        End If
        AddList s
        s = s & ", Ram = " & Str(Int(PhysicalRAM(True))) & " mb"
        DebugLog s
        DoEvents
        
If IsIDE Then
    'bSkip = True
    If Left(strSymbol, 3) <> "ZS-" Or Not bCalendarSpread Then
        'bSkip = True
    End If
End If

        ' load Primary data
        dDollarsPerPoint = 0
        dDollarsPerPoint2 = 0
        dMinMove = 0
        dMinMove2 = 0
        bSymbolInfoDone = False
        If Not bSkip Then
            DM_GetBars Bars, nSymbolID, "Daily", nStartDate, nEndDate
            dMinMove = Bars.MinMove
            If dMinMove <= 0 Then
                bSkip = True
            ElseIf strSecType = "F" Then
                ' Futures must have a valid tick value and move
                If Bars.Prop(eBARS_TickValue) <= 0 Or Bars.Prop(eBARS_TickMove) <= 0 Then
                    bSkip = True
                Else
                    dDollarsPerPoint = Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
                    ' also load the 057 data (in order to determine the "forward roll adjust" from the closing prices)
                    If bCalendarSpread Then
                        ' except for calendar spreads (which use the Gann contracts) -- just use a copy
                        Set Bars57 = Bars.MakeCopy
                    Else
                        DM_GetBars Bars57, Parse(strSymbol, "-", 1) & "-057", "Daily", Bars(eBARS_DateTime, 0), Bars(eBARS_DateTime, Bars.Size - 1)
                    End If
                    If Bars57.Size <> Bars.Size Then
                        bSkip = True ' shouldn't happen!
                        If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Bars57 Error"
                    End If
                End If
            Else
                ' but for Stocks/Indexes the prices cannot go negative (since doing percentage changes)
                If gdMinValue(Bars.ArrayHandle(eBARS_Close), 0, Bars.Size - 1) <= 0 Then
                    bSkip = True
                End If
                ' and check for any gaps in data > 2 weeks -- if so, need to ignore
                ' what's before the gap (since could be from an entirely different company)
                ' - or if price jumps up or down by factor of 10
                For iBar = Bars.Size - 1 To 1 Step -1
                    i = Bars(eBARS_DateTime, iBar) - Bars(eBARS_DateTime, iBar - 1)
                    ' check for a large percentage gap
                    d = Bars(eBARS_Close, iBar - 1)
                    If d > 0 Then
                        d = Bars(eBARS_Close, iBar) / d
                    End If
                    If i >= 14 Or d < 0.1 Or d > 10 Then
                        Bars.DeleteFirstBars iBar ' ignore all data prior to this bar
                        Exit For
                    End If
                Next
                Bars57.Size = 0
                SetBarProperties Bars57, nSymbolID
            End If
            
            ' must be at least 5 years of data, and have fairly recent data
            If Bars(eBARS_DateTime, 0) > nEndDate - 365 * 5 Then
                bSkip = True
            ElseIf Bars(eBARS_DateTime, Bars.Size - 1) < nEndDate - 30 Then
                bSkip = True
            End If
        End If
        
        ' load Spread data
        If nSymbolID2 > 0 And Not bSkip Then
            DM_GetBars Bars2, nSymbolID2, "Daily", nStartDate, nEndDate
            dMinMove2 = Bars2.MinMove
            ' Futures must have a valid tick value and move
            If Bars2.Prop(eBARS_TickValue) <= 0 Or Bars2.Prop(eBARS_TickMove) <= 0 Or dMinMove2 <= 0 Then
                bSkip = True
            Else
                dDollarsPerPoint2 = Bars2.Prop(eBARS_TickValue) / Bars2.Prop(eBARS_TickMove)
                ' also load the 057 data (in order to determine the "forward roll adjust" from the closing prices)
                If bCalendarSpread Then
                    ' except for calendar spreads (which use the Gann contracts) -- just use a copy
                    Set Bars57s = Bars2.MakeCopy
                Else
                    DM_GetBars Bars57s, Parse(strSymbol2, "-", 1) & "-057", "Daily", Bars2(eBARS_DateTime, 0), Bars2(eBARS_DateTime, Bars2.Size - 1)
                End If
                If Bars57s.Size <> Bars2.Size Then
                    bSkip = True ' shouldn't happen!
                    If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Bars57s Error"
                End If
            End If
            
            ' must be at least 5 years of data, and have fairly recent data
            If Bars2(eBARS_DateTime, 0) > nEndDate - 365 * 5 Then
                bSkip = True
            ElseIf Bars2(eBARS_DateTime, Bars2.Size - 1) < nEndDate - 30 Then
                bSkip = True
            End If
        End If
        
        ' Place the gdBars prices into the BarsData table (so all bars will be better aligned with "real" TDOY's)
        If Not bSkip Then
            ' clear out all the BarsData prices (but not the dates)
            For iBar = 0 To iNumBars
                For i = kDate + 1 To kOpen
                    BarsData(i, iBar) = kNullData
                Next
            Next
            
            iRow = iEndDateBar ' iNumBars
            For iBar = Bars.Size - 1 To 0 Step -1
                d = Bars(eBARS_DateTime, iBar)
                ' find row in table where bar date > date of prior table row
                Do While iRow >= 1
                    If d > BarsData(kDate, iRow - 1) Then
                        Exit Do
                    End If
                    iRow = iRow - 1
                Loop
                If d <> BarsData(kDate, iRow) Then
                    s = DateFormat(d)
                    If IsIDE Then
                        'AddList "Merge holiday: " & s
                    End If
                End If
                ' if no data yet for that row, then set all the prices
                If BarsData(kClose, iRow) = kNullData Then
                    BarsData(kClose, iRow) = Bars(eBARS_Close, iBar)
                    If bDoStocks Then
                        BarsData(kClose57, iRow) = Bars(eBARS_Close, iBar)
                    Else
                        BarsData(kClose57, iRow) = Bars57(eBARS_Close, iBar)
                    End If
                    If Not bDoSpreads Then
                        BarsData(kOpen, iRow) = Bars(eBARS_Open, iBar)
                        BarsData(kHigh, iRow) = Bars(eBARS_High, iBar)
                        BarsData(kLow, iRow) = Bars(eBARS_Low, iBar)
                    End If
                ElseIf Not bDoSpreads Then
                    ' else just merge the High/Low into this row (e.g. a "half-holiday bar")
                    If BarsData(kHigh, iRow) < Bars(eBARS_High, iBar) Then
                        BarsData(kHigh, iRow) = Bars(eBARS_High, iBar)
                        ' if a higher high, then also check for a higher open (in case of an overnight gap)
                        If BarsData(kOpen, iRow) < Bars(eBARS_Open, iBar) Then
                            BarsData(kOpen, iRow) = Bars(eBARS_Open, iBar)
                        End If
                    End If
                    If BarsData(kLow, iRow) > Bars(eBARS_Low, iBar) Then
                        BarsData(kLow, iRow) = Bars(eBARS_Low, iBar)
                        ' if a lower low, then also check for a lower open (in case of an overnight gap)
                        If BarsData(kOpen, iRow) > Bars(eBARS_Open, iBar) Then
                            BarsData(kOpen, iRow) = Bars(eBARS_Open, iBar)
                        End If
                    End If
                End If
            Next
            
            If bDoSpreads Then
                iRow = iEndDateBar ' iNumBars
                For iBar = Bars2.Size - 1 To 0 Step -1
                    d = Bars2(eBARS_DateTime, iBar)
                    ' find row in table where bar date > date of prior table row
                    Do While iRow >= 1
                        If d > BarsData(kDate, iRow - 1) Then
                            Exit Do
                        End If
                        iRow = iRow - 1
                    Loop
                    ' if no data yet for that row, then set all the prices
                    ' (no open/high/low or holiday merging needed for spreads)
                    If BarsData(kSpread, iRow) = kNullData Then
                        BarsData(kSpread, iRow) = Bars2(eBARS_Close, iBar)
                        BarsData(kSpread57, iRow) = Bars57s(eBARS_Close, iBar)
                    End If
                Next
                
                ' and for calendar spreads, flag the roll dates for both legs
                ' (this is actually the "expiration dates" -- i.e. one data bar prior to the roll)
                If bCalendarSpread Then
                    ' first leg
                    Set Rolls = GetRollsTable(nSymbolID)
                    If Rolls Is Nothing Then
                        bSkip = True
                    Else
                        dLowest = 0
                        iBar = iEndDateBar ' iNumBars - 1
                        For iRoll = Rolls.NumRecords - 1 To 0 Step -1
                            d = Rolls.Num(1, iRoll) ' date of roll
                            Do While iBar > 0 And d <= nEndDate
                                ' looking backwards, find the first bar where date < roll date and has good data
                                ' (i.e. this will be the "expiration date" for one of the Gann yearly contracts)
                                If BarsData(kDate, iBar) < d And BarsData(kClose, iBar) <> kNullData Then
                                    ' for first leg, just set to 1
                                    BarsData(kRolled, iBar) = 1
                                    iTDOY = BarsData(kTDOY, iBar)
                                    ' track the "lowest TDOY" for a roll
                                    If dLowest = 0 Then
                                        ' if it hasn't been set yet
                                        dLowest = iTDOY
                                    ElseIf iTDOY > 175 And dLowest < 75 Then
                                        ' if this TDOY is at end of the year and lowest so far is at beginning of the year
                                        dLowest = iTDOY
                                    ElseIf iTDOY < 75 And dLowest > 175 Then
                                        ' but ignore this TDOY if at beginning of the year when lowest is at end of the year
                                        dLowest = dLowest
                                    ElseIf iTDOY < dLowest Then
                                        ' else use this TDOY if it's the lowest so far
                                        dLowest = iTDOY
                                    End If
                                    Exit Do
                                End If
                                iBar = iBar - 1
                            Loop
                        Next
                        AddList "For " & strSymbol & ": Lowest roll = " & Str(dLowest)
                        
                        ' now mark the "dead zones" in the forecast window
                        bSkip = False
                        j = Val(Right(strSymbol, 3)) - 80 ' contract month
                        For iBar = iNumBars To iEndDateBar + 1 Step -1 ' loop backwards from end
                            ' mark this bar as in dead zone when "bSkip" flag has been set
                            If bSkip Then
                                BarsData(kRolled, iBar) = 9
                            End If
                            ' when hit the end of the contract month, start marking a dead zone
                            If Month(BarsData(kDate, iBar)) <> j And Month(BarsData(kDate, iBar - 1)) = j Then
                                bSkip = True
                            ElseIf BarsData(kTDOY, iBar) = dLowest Then
                                ' when get to lowest TDOY roll, we can stop the dead zone after this bar
                                bSkip = False
                            End If
                        Next
                    End If
                    
                    ' second leg
                    Set Rolls = GetRollsTable(nSymbolID2)
                    If Rolls Is Nothing Then
                        bSkip = True
                    Else
                        dLowest = 0
                        iBar = iEndDateBar ' iNumBars - 1
                        For iRoll = Rolls.NumRecords - 1 To 0 Step -1
                            d = Rolls.Num(1, iRoll) ' date of roll
                            Do While iBar > 1 And d <= nEndDate
                                ' looking backwards, find the first bar where date < roll date and has good data
                                ' (i.e. this will be the "expiration date" for one of the Gann yearly contracts)
                                If BarsData(kDate, iBar) < d And BarsData(kSpread, iBar) <> kNullData Then
                                    ' for second leg, set to 2 -- except if rolled for both then set to 3
                                    ' (so 1 = 1st leg, 2 = 2nd leg, 3 = both legs rolled on this date)
                                    If BarsData(kRolled, iBar) > 0 Then
                                        BarsData(kRolled, iBar) = 3
                                    Else
                                        BarsData(kRolled, iBar) = 2
                                    End If
                                    iTDOY = BarsData(kTDOY, iBar)
                                    ' track the "lowest TDOY" for a roll
                                    If dLowest = 0 Then
                                        ' if it hasn't been set yet
                                        dLowest = iTDOY
                                    ElseIf iTDOY > 175 And dLowest < 75 Then
                                        ' if this TDOY is at end of the year and lowest so far is at beginning of the year
                                        dLowest = iTDOY
                                    ElseIf iTDOY < 75 And dLowest > 175 Then
                                        ' but ignore this TDOY if at beginning of the year when lowest is at end of the year
                                        dLowest = dLowest
                                    ElseIf iTDOY < dLowest Then
                                        ' else use this TDOY if it's the lowest so far
                                        dLowest = iTDOY
                                    End If
                                    Exit Do
                                End If
                                iBar = iBar - 1
                            Loop
                        Next
                        AddList "For " & strSymbol2 & ": Lowest roll = " & Str(dLowest)
                        
                        ' now mark the "dead zones" in the forecast window
                        bSkip = False
                        j = Val(Right(strSymbol2, 3)) - 80 ' contract month
                        For iBar = iNumBars To iEndDateBar + 1 Step -1 ' loop backwards from end
                            ' mark this bar as in dead zone when "bSkip" flag has been set
                            If bSkip Then
                                BarsData(kRolled, iBar) = 9
                            End If
                            ' when hit the end of the contract month, start marking a dead zone
                            If Month(BarsData(kDate, iBar)) <> j And Month(BarsData(kDate, iBar - 1)) = j Then
                                bSkip = True
                            ElseIf BarsData(kTDOY, iBar) = dLowest Then
                                ' when get to lowest TDOY roll, we can stop the dead zone after this bar
                                bSkip = False
                            End If
                        Next
                    End If
                    Set Rolls = Nothing
                    
                    If IsIDE And aSymbols.Size <= 3 Then
                        For iBar = 0 To iNumBars
                            Select Case BarsData(kRolled, iBar)
                            Case 1:
                                s = strSymbol
                            Case 2:
                                s = strSymbol2
                            Case 3:
                                s = strSymbol & " & " & strSymbol2
                            Case 9:
                                s = "Dead zone"
                            Case Else:
                                s = ""
                            End Select
                            If Len(s) > 0 Then
                                d = BarsData(kDate, iBar)
                                i = BarsData(kTDOY, iBar)
                                AddList "Roll TDOY = " & Format(i, "000") & ", " & DateFormat(d) & ": " & s
                            End If
                        Next
                    End If
                End If
            End If
            
            ' and prebuild the basic trade info (date,price) for each bar
            ' - it's much faster to format these strings just once for each bar (instead of 1000's of times in the nested For loops below)
            ReDim aTradeInfo(iNumBars) As String
            For iBar = 0 To iNumBars
                strText = Bars.PriceDisplay(BarsData(kClose57, iBar))
                If bDoSpreads Then
                    strText = strText & " - " & Bars2.PriceDisplay(BarsData(kSpread57, iBar))
                End If
                aTradeInfo(iBar) = Format(BarsData(kDate, iBar), "mm/dd/yyyy") & "," & strText & ","
            Next
        End If

If IsIDE Then
    'If bDoSpreads Then bSkip = True
End If
            
        ' Do seasonal calcs for both LONGS and SHORTS (except do only Longs for all Spreads)
        bShort = False
        Do While bInProgress And Not bSkip
            If bShort Then
                If bDoSpreads Then Exit Do
                strDirection = "Short"
            Else
                strDirection = "Long"
            End If
            
            ' clear all the 3D tables
            ReDim tblNumUp(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Long
            ReDim tblNumDown(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Long
            ReDim tblAmtUp(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            ReDim tblAmtDown(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            ReDim tblDD(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            ReDim tblWorstDD(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            If bIncludeTradeHistory Then
                i = Year(nStartDate)
                j = Year(nEndDate)
                ReDim tblTradeEntryBar(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses, i To j) As Integer
                ReDim tblTradeExitBar(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses, i To j) As Integer
                ReDim tblTradeExit(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses, i To j) As Long
                ReDim tblTradeExit2(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses, i To j) As Long
                ReDim tblTradeProfit(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses, i To j) As Single
                ReDim tblTradeDD(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses, i To j) As Single
            End If

            ' for each Entry bar
            For iBar = 1 To iEndDateBar 'iNumBars
                If Not bInProgress Then Exit For
                
                If (iBar Mod 2500) = 0 Or iBar = iEndDateBar Then
                    AddList "   " & strDirection & ": " & Str(iBar) & " of " & Str(iEndDateBar) & " bars, Ram = " & Str(Int(PhysicalRAM(True))) & " mb"
                    DoEvents
                End If
                
                ' make sure this bar is a valid entry (date, price, TDOY)
                iEntryBar = iBar
                iExitBar = 0
                iTDOY = BarsData(kTDOY, iBar)
                d = BarsData(kDate, iBar)
                dEntry = BarsData(kClose, iBar)
                If iTDOY <= 0 Or iTDOY > kMaxTDOY Or d = kNullData Or d > nEndDate Or dEntry = kNullData Then
                    iTDOY = 0
                ElseIf bDoSpreads Then
                    ' can only enter a spread on a date which has a valid price for BOTH legs
                    dEntry2 = BarsData(kSpread, iBar)
                    If dEntry2 = kNullData Then
                        iTDOY = 0
                    End If
                End If
                    
                If iTDOY > 0 Then
                    ' calc the forward roll-adjust for Futures (in case rolls between entry and exit)
                    dRollAdjust = BarsData(kClose57, iBar) - BarsData(kClose, iBar)
                    dEntry = dEntry + dRollAdjust
                    If bDoSpreads Then
                        dRollAdjust2 = BarsData(kSpread57, iBar) - BarsData(kSpread, iBar)
                        dEntry2 = dEntry2 + dRollAdjust2
                    End If
                    
                    ' calc stats separately for each StopLoss value
                    For iStopLoss = 1 To kMaxStopLosses
                        ' determine price of the StopLoss to use
                        If bUsePercentage Then
                            dStopLoss = iStopLoss * 2.5 / 100# ' at 2.5% increments
                            dStopPrice = dEntry * dStopLoss
                        Else
                            dStopLoss = iStopLoss * 500# ' at $500 increments
                            dStopPrice = dStopLoss / dDollarsPerPoint
                        End If
                        If Not bDoSpreads Then
                            If bShort Then
                                ' move to the next highest tick
                                dStopPrice = dEntry + dStopPrice
                                dStopPrice = RoundToSigDigits(dStopPrice / dMinMove)
                                If dStopPrice <> Int(dStopPrice) Then
                                    dStopPrice = Int(dStopPrice) + 1
                                End If
                            Else
                                ' move to the next lowest tick
                                dStopPrice = dEntry - dStopPrice
                                dStopPrice = Int(RoundToSigDigits(dStopPrice / dMinMove))
                            End If
                            dStopPrice = RoundToSigDigits(dStopPrice * dMinMove)
                        End If
                        
                        ' for each #days held
                        bStoppedOut = False
                        dLowest = dEntry
                        dHighest = dEntry
                        dDD = 0
                        iExitBar = 0
                        dExit = kNullData
                        dExit2 = kNullData
                        For iDaysHeld = 1 To kMaxDaysHeld
                            ' check if beyond end of bars (shouldn't ever happen)
                            If iBar + iDaysHeld > iNumBars Then
                                If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Invalid Exit Bar"
                                Exit For
                            End If
                            
                            ' see if this is a valid Exit bar (valid date and prices)
                            If BarsData(kDate, iBar + iDaysHeld) > nEndDate Then
                                Exit For ' can exit loop if end date is beyond end of good data
                            ElseIf bCalendarSpread Then
                                ' for calendar spreads, trades cannot be held across the "roll" of either leg
                                If BarsData(kRolled, iBar + iDaysHeld) <> kNullData Then
                                    Exit For ' can exit loop since cannot exit after this date
                                End If
                            End If
                            
                            d = BarsData(kClose, iBar + iDaysHeld)
                            If bDoSpreads Then
                                If BarsData(kSpread, iBar + iDaysHeld) = kNullData Then
                                    d = kNullData
                                End If
                            End If
                            
                            ' if already stopped out (or if not a valid Exit date), then just
                            ' use the same info for the stopped-out trade
                            If d <> kNullData And Not bStoppedOut Then
                                ' get exit bar prices -- but for Futures, add the "roll adjust"
                                ' (i.e. this is a "forward-price-adjust" if rolls between the entry and exit)
                                dExit = BarsData(kClose, iBar + iDaysHeld) + dRollAdjust
                                dExit2 = BarsData(kSpread, iBar + iDaysHeld) + dRollAdjust2
                                iExitBar = iBar + iDaysHeld
                                
                                If bDoSpreads Then
'If iTDOY = 71 And iDaysHeld = 15 And dStopLoss = 3000 Then
'    i = i
'End If
                                    ' for SPREADS, we can only check the Close (since can't really calculate High/Low of spread bars)
                                    ' calculate net profit of the spread (Buy#1 - Sell#2)
                                    dNet = (dExit - dEntry) * dDollarsPerPoint
                                    d = (dExit2 - dEntry2) * dDollarsPerPoint2
                                    dNet = RoundToSigDigits(dNet - d)
                                    ' check for newer drawdown
                                    If dNet < dDD Then
                                        dDD = dNet
                                    End If
                                    ' check if hit stop-loss
                                    If dNet <= -dStopLoss Then
                                        bStoppedOut = True
                                    End If
                                Else
                                    ' for NON-SPREADS, let's check intrabar drawdown and stop (using High/Low of bars)
                                    dOpen = BarsData(kOpen, iBar + iDaysHeld) + dRollAdjust
                                    dHigh = BarsData(kHigh, iBar + iDaysHeld) + dRollAdjust
                                    dLow = BarsData(kLow, iBar + iDaysHeld) + dRollAdjust
                                
                                    ' to check drawdown
                                    If dLow < dLowest Then
                                        ' if a lower low for a Long trade, then check if hit stop-loss
                                        If Not bShort Then
                                            If dLow <= dStopPrice Then
                                                ' see if stop-loss was hit at open, or after open
                                                If dOpen <= dStopPrice Then
                                                    dLow = dOpen
                                                Else
                                                    dLow = dStopPrice
                                                End If
                                                dExit = dLow
                                                bStoppedOut = True
                                            End If
                                        End If
                                        dLowest = dLow
                                    End If
                                    If dHigh > dHighest Then
                                        ' if a higher high for a Short trade, then check if hit stop-loss
                                        If bShort Then
                                            If dHigh >= dStopPrice Then
                                                ' see if stop-loss was hit at open, or after open
                                                If dOpen >= dStopPrice Then
                                                    dHigh = dOpen
                                                Else
                                                    dHigh = dStopPrice
                                                End If
                                                dExit = dHigh
                                                bStoppedOut = True
                                            End If
                                        End If
                                        dHighest = dHigh
                                    End If
                                    
                                    If bUsePercentage Then
                                        dNet = (dExit - dEntry) / dEntry * 100#
                                        If bShort Then
                                            dNet = -dNet
                                            dDD = (dEntry - dHighest) / dEntry * 100#
                                        Else
                                            dDD = (dLowest - dEntry) / dEntry * 100#
                                        End If
                                    Else
                                        dNet = (dExit - dEntry) * dDollarsPerPoint
                                        If bShort Then
                                            dNet = -dNet
                                            dDD = (dEntry - dHighest) * dDollarsPerPoint
                                        Else
                                            dDD = (dLowest - dEntry) * dDollarsPerPoint
                                        End If
                                    End If
                                End If
                            End If
                            
                            ' whether stopped-out or not, accumulate trade info for this 3D spot (TDOY, DaysHeld, StopLossLevel)
                            If iExitBar > 0 Then
                                If dNet > 0 Then
                                    tblNumUp(iTDOY, iDaysHeld, iStopLoss) = tblNumUp(iTDOY, iDaysHeld, iStopLoss) + 1
                                    tblAmtUp(iTDOY, iDaysHeld, iStopLoss) = tblAmtUp(iTDOY, iDaysHeld, iStopLoss) + dNet
                                Else
                                    tblNumDown(iTDOY, iDaysHeld, iStopLoss) = tblNumDown(iTDOY, iDaysHeld, iStopLoss) + 1
                                    tblAmtDown(iTDOY, iDaysHeld, iStopLoss) = tblAmtDown(iTDOY, iDaysHeld, iStopLoss) + dNet
                                End If
                                
                                tblDD(iTDOY, iDaysHeld, iStopLoss) = tblDD(iTDOY, iDaysHeld, iStopLoss) + dDD
                                If dDD < tblWorstDD(iTDOY, iDaysHeld, iStopLoss) Then
                                    tblWorstDD(iTDOY, iDaysHeld, iStopLoss) = dDD
                                End If
    
                                ' and append the results for this trade onto the TradeHistory gdString in the table
                                If bIncludeTradeHistory Then
                                    iYear = Year(BarsData(kDate, iEntryBar))
                                    tblTradeEntryBar(iTDOY, iDaysHeld, iStopLoss, iYear) = iEntryBar
                                    tblTradeExitBar(iTDOY, iDaysHeld, iStopLoss, iYear) = iExitBar
                                    tblTradeExit(iTDOY, iDaysHeld, iStopLoss, iYear) = Round(dExit / dMinMove)
                                    If bDoSpreads Then
                                        tblTradeExit2(iTDOY, iDaysHeld, iStopLoss, iYear) = Round(dExit2 / dMinMove2)
                                    End If
                                    tblTradeProfit(iTDOY, iDaysHeld, iStopLoss, iYear) = dNet
                                    tblTradeDD(iTDOY, iDaysHeld, iStopLoss, iYear) = dDD
                                End If
                            End If
                        Next
                    Next
                End If
            Next

            ' Now write all the output for this symbol
            iLineCountStart = iLineCount
            For iTDOY = 1 To kMaxTDOY
                ' find "Entry Bar" (find the next TDOY which occurs after the End Date)
                iEntryBar = 0
                For iBar = iNumBars To 0 Step -1
                    If BarsData(kDate, iBar) <= nEndDate Then
                        Exit For ' done looking
                    End If
                    If BarsData(kTDOY, iBar) = iTDOY Then
                        iEntryBar = iBar
                    End If
                Next
                
                ' for each "Exit Bar"
                For iDaysHeld = 1 To kMaxDaysHeld
                    If iEntryBar <= 0 Then
                        Exit For
                    End If
                    iExitBar = iEntryBar + iDaysHeld
                    If iExitBar > iNumBars Then
                        If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Invalid Exit Bar"
                        Exit For
                    End If
                    
                    If bCalendarSpread Then
                        ' for calendar spreads, check if either the entry or exit is in a "dead zone"
                        ' (since no predicted entries allowed in a dead zone, and all trades must be exited before the next dead zone)
                        If BarsData(kRolled, iEntryBar) <> kNullData Or BarsData(kRolled, iExitBar) <> kNullData Then
                            Exit For
                        End If
                    End If
                    
                    For iStopLoss = 1 To kMaxStopLosses
                        If bUsePercentage Then
                            dStopLoss = iStopLoss * 2.5
                        Else
                            dStopLoss = iStopLoss * 500#
                        End If
                        
                        ' to be valid, needs to be a minimum # of occurances
                        iNumTrades = tblNumUp(iTDOY, iDaysHeld, iStopLoss) + tblNumDown(iTDOY, iDaysHeld, iStopLoss)
                        If iStopLoss > 1 And iNumTrades > 1 Then
                            ' but can ignore this "duplicate" if it's the same trades as the lower stop-loss amount
                            If tblAmtUp(iTDOY, iDaysHeld, iStopLoss) = tblAmtUp(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                If tblAmtDown(iTDOY, iDaysHeld, iStopLoss) = tblAmtDown(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                    If tblNumDown(iTDOY, iDaysHeld, iStopLoss) = tblNumDown(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                        If tblNumUp(iTDOY, iDaysHeld, iStopLoss) = tblNumUp(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                            If tblDD(iTDOY, iDaysHeld, iStopLoss) = tblDD(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                                iNumTrades = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If iNumTrades >= 5 Then
                            If iTDOY = kMaxTDOY Then
                                If IsIDE Then InfBox "This shouldn't really get here!", "!", , "iTDOY = kMaxTDOY"
                            End If
                        
                            ' calc avg profit and Win%
                            dAvgNet = (tblAmtUp(iTDOY, iDaysHeld, iStopLoss) + tblAmtDown(iTDOY, iDaysHeld, iStopLoss)) / iNumTrades
                            dWinPct = 100# * tblNumUp(iTDOY, iDaysHeld, iStopLoss) / iNumTrades
                            
                            ' calc Profit Factor ($won/$lost), AvgWin, AvgLoss
                            If tblAmtDown(iTDOY, iDaysHeld, iStopLoss) < 0 Then
                                dPF = Abs(tblAmtUp(iTDOY, iDaysHeld, iStopLoss) / tblAmtDown(iTDOY, iDaysHeld, iStopLoss))
                            Else
                                dPF = 100 ' just use a really high number if no $lost
                            End If
                            If tblNumUp(iTDOY, iDaysHeld, iStopLoss) > 0 Then
                                dAvgWin = tblAmtUp(iTDOY, iDaysHeld, iStopLoss) / tblNumUp(iTDOY, iDaysHeld, iStopLoss)
                            Else
                                dAvgWin = 0
                            End If
                            If tblNumDown(iTDOY, iDaysHeld, iStopLoss) > 0 Then
                                dAvgLoss = tblAmtDown(iTDOY, iDaysHeld, iStopLoss) / tblNumDown(iTDOY, iDaysHeld, iStopLoss)
                            Else
                                dAvgLoss = 0
                            End If
                            
                            ' drawdown and ROI
                            dAvgDD = tblDD(iTDOY, iDaysHeld, iStopLoss) / iNumTrades
                            dWorstDD = tblWorstDD(iTDOY, iDaysHeld, iStopLoss)
                            If dAvgDD = 0 Then
                                dAAP2DD = 99999
                            Else
                                dAAP2DD = -dAvgNet / dAvgDD ' dWorstDD
                            End If
                            
                            If dPF >= dMinPF And dWinPct >= dMinWinPerc And dAvgWin > 0 Then
                                ' output for table
                                If bUsePercentage Then
                                    strText = Format(dAvgNet, "#0.0000") & vbTab & Format(dAvgNet / iDaysHeld, "#0.0000") _
                                         & vbTab & Format(dAvgDD, "#0.0000") & vbTab & Format(dWorstDD, "#0.0000")
                                Else
                                    'strText = "$" & Format(dAvgNet * dDollarsPerPoint, "#0.00") & vbTab & "$" & Format(dAvgNet / iDaysHeld * dDollarsPerPoint, "#0.00") _
                                         & vbTab & "$" & Format(dAvgDD * dDollarsPerPoint, "#0.00") & vbTab & "$" & Format(dWorstDD * dDollarsPerPoint, "#0.00")
                                    strText = Format(dAvgNet, "#0.00") & vbTab & Format(dAvgNet / iDaysHeld, "#0.00") _
                                         & vbTab & Format(dAvgDD, "#0.00") & vbTab & Format(dWorstDD, "#0.00")
                                End If
                                
                                ' put all the trades for this item into a semi-colon delimited string
                                strTrades = ""
                                If bIncludeTradeHistory Then
                                    For iYear = Year(nStartDate) To Year(nEndDate)
                                        i = tblTradeEntryBar(iTDOY, iDaysHeld, iStopLoss, iYear)
                                        j = tblTradeExitBar(iTDOY, iDaysHeld, iStopLoss, iYear)
                                        If i > 0 And j > 0 Then
                                            'EntryDate,EntryPrices,ExitDate,ExitPrices,NetProfit,Drawdown
                                            dEntry = BarsData(kClose57, i)
                                            dEntry2 = BarsData(kSpread57, i)
                                            dExit = tblTradeExit(iTDOY, iDaysHeld, iStopLoss, iYear) * dMinMove
                                            dExit2 = tblTradeExit2(iTDOY, iDaysHeld, iStopLoss, iYear) * dMinMove2
                                            dNet = tblTradeProfit(iTDOY, iDaysHeld, iStopLoss, iYear)
                                            dDD = tblTradeDD(iTDOY, iDaysHeld, iStopLoss, iYear)
                                            ' build trade string: EntryDate, EntryPrice, ExitDate, ExitPrice, NetProfit, Drawdown
                                            If bDoSpreads Then
                                                ' EntryDate, EntryPrice1 - EntryPrice2, ExitDate, ExitPrice1 - ExitPrice2, NetProfit, Drawdown
                                                s = aTradeInfo(i) & Left(aTradeInfo(j), 11) & Bars.PriceDisplay(dExit) & " - " & Bars2.PriceDisplay(dExit2) & ","
                                            ElseIf Abs(dExit - BarsData(kClose57, j)) > 0.0000001 Then
                                                s = aTradeInfo(i) & Left(aTradeInfo(j), 11) & Bars.PriceDisplay(dExit) & ","
                                            Else
                                                s = aTradeInfo(i) & aTradeInfo(j)
                                            End If
                                            strTrades = strTrades & ";" & s & Format(dNet, "#0.00") & "," & Format(dDD, "#0.00")
                                        End If
                                    Next
                                End If
                                                               
                                If bDoSpreads Then
                                    i = 0
                                Else
                                    i = nSymbolID
                                End If
                                strText = strName & vbTab & Str(i) & vbTab & strDirection & vbTab _
                                    & Str(iTDOY) & vbTab & Str(BarsData(kDate, iEntryBar)) & vbTab _
                                    & Str(iDaysHeld) & vbTab & Str(BarsData(kDate, iExitBar)) & vbTab _
                                    & Str(dStopLoss) & vbTab & Str(iNumTrades) & vbTab & Format(dWinPct, "#0.0") & vbTab & Format(dPF, "#0.000") & vbTab _
                                    & strText & vbTab & Format(dAAP2DD, "#0.0000") & vbTab & Format(dAvgNet / dAvgWin * 100, "#0.0") & vbTab & strTrades
                                Print #fh, strText
                                iLineCount = iLineCount + 1
                                dFileSize = dFileSize + Len(strText) + 2
                                
                                ' and also get the symbol info (only one time for each symbol)
                                If Not bSymbolInfoDone Then
                                    bSymbolInfoDone = True
                                    i = g.SymbolPool.PoolRecForSymbolID(nSymbolID)
                                    s = g.SymbolPool.Desc(i)
                                    strExch = Parse(s, "@", 2) ' exchange
                                    strDesc = Parse(s, "@", 1) ' description
                                    If strSecType = "I" Then
                                        strExch = ""
                                    ElseIf strSecType = "F" Then
                                        ' strip off the "Cont Liq" part of the description (Liq CAdj Cont Elec Exp)
                                        strDesc = CleanSymbolDesc(strDesc)
                                    End If
                                    If Not bDoSpreads Then
                                        strText = strSecType & vbTab & strName & vbTab & Str(nSymbolID) & vbTab & strDesc & vbTab & strExch
                                    Else
                                        strText = strSecType & vbTab & strName & vbTab & strSymbol & vbTab & Str(nSymbolID) & vbTab & strDesc & vbTab & strExch
                                        i = g.SymbolPool.PoolRecForSymbolID(nSymbolID2)
                                        s = g.SymbolPool.Desc(i)
                                        strExch = Parse(s, "@", 2) ' exchange
                                        strDesc = Parse(s, "@", 1) ' description
                                        If strSecType = "I" Then
                                            strExch = ""
                                        ElseIf strSecType = "F" Then
                                            ' strip off the "Cont Liq" part of the description (Liq CAdj Cont Elec Exp)
                                            strDesc = CleanSymbolDesc(strDesc)
                                        End If
                                        strText = strText & vbTab & strSymbol2 & vbTab & Str(nSymbolID2) & vbTab & strDesc & vbTab & strExch
                                    End If
                                    aSymbolInfo.Add strText
                                End If
                            End If
                        End If
                    Next
                Next
            Next
            If iLineCount > iLineCountStart Then
                AddList "#Lines for " & strDirection & " " & strName & " = " & Str(iLineCount - iLineCountStart)
            End If
        
            ' setup for Shorts (except for Spreads)
            If bShort Or bDoSpreads Then
                Exit Do
            Else
                bShort = True
            End If
        Loop ' do for both Longs and Shorts (except for Spreads)
    Next
    
    aSymbolInfo.Sort
    aSymbolInfo.ToFile strSymbolInfoFile

    Close #fh
    
    ReDim tblTradeEntryBar(0) As Integer
    ReDim tblTradeExitBar(0) As Integer
    ReDim tblTradeExit(0) As Long
    ReDim tblTradeExit2(0) As Long
    ReDim tblTradeProfit(0) As Single
    ReDim tblTradeDD(0) As Single
    ReDim tblNumUp(0) As Long
    ReDim tblNumDown(0) As Long
    ReDim tblAmtUp(0) As Double
    ReDim tblAmtDown(0) As Double
    ReDim tblDD(0) As Double
    ReDim tblWorstDD(0) As Double
    ReDim tblTradeStrings(0) As Long
    ReDim BarsData(0) As Double
    ReDim aTradeInfo(0) As String
    
    AddList "Ram = " & Str(Int(PhysicalRAM(True))) & " of " & Str(Int(PhysicalRAM(False))) & " mb"
    AddList "Finished, #Lines = " & Format(iLineCount, "#,##0") & ", FileSize = " & Str(Int(dFileSize / 1000000#)) & " mb"
    bInProgress = False

End Sub

#If 0 Then
Private Sub CalcSeasonals1() ' data for Seasonal Sweet Spots

    ' for the 3D tables (3d arrays)
    Const kMaxTDOY As Long = 262 ' 262 = max # of weekdays in a year
    Const kMaxDaysHeld As Long = 125 ' about 6 months (=125 trading days) for longest held seasonal trade
    Const kMaxStopLosses As Long = 10 ' max of 10 different stop-loss levels

    ' for the 2D BarsData (arrays)
    Const kTDOY As Long = 0
    Const kDate As Long = 1
    Const kClose57 As Long = 2
    Const kClose As Long = 3
    Const kHigh As Long = 4
    Const kLow As Long = 5
    Const kOpen As Long = 6
    ' when doing spreads, we can re-purpose the High/Low/Open arrays (since unused for spreads)
    Const kSpread57 As Long = 4
    Const kSpread As Long = 5
    Const kRolled As Long = 6 ' i.e. rolled on this date: 0 = none, 1 = leg#1, 2 = leg#2, 3 = both legs rolled

    ' these are the 3D tables used to accumulate all the trade info for each TDOY/#DaysHeld/StopLossLevel:
    Dim tblNumUp() As Long
    Dim tblNumDown() As Long
    Dim tblAmtUp() As Double
    Dim tblAmtDown() As Double
    Dim tblDD() As Double
    Dim tblWorstDD() As Double
    Dim tblTradeStrings() As Long

    Dim BarsData() As Double ' the 2D table used to line up all the data into a consistent TDOY grid

    ' other variables
    Dim i&, j&, iRow&, iSymbol&, iBar&, iDaysHeld&, iStopLoss&, nSymbolID&, iTDOY&, iYear&, iNumBars&, iRoll&, iEndDateBar&
    Dim d#, dEntry#, dExit#, iNumTrades&, iLineCount&, hTrades&
    Dim nEndDate&, nStartDate&, iEntryBar&, iExitBar&, strDirection$, strName$, strDesc$
    Dim bShort As Boolean, bIncludeTradeHistory As Boolean
    Dim fh&, dNet#, dAvgNet#, dWinPct#, dDollarsPerPoint#, dPF#, dAvgWin#, dAvgLoss#, dStopLoss#, dStopPrice#, dTimeStarted#
    Dim dLowest#, dHighest#, dDD#, dAvgDD#, dWorstDD#, dAAP2DD#, dChk#, dRollAdjust#, dLow#, dHigh#, dOpen#
    Dim dTradeHistMemSize#, dFileSize#, dMinWinPerc#, dMinPF#, dMinMove#
    Dim s$, strSymbol$, strText$, strTrade$, strSecType$, strSymbolInfoFile$, strDataFile$, strPath$, strExch$
    Dim bUsePercentage As Boolean, bSkip As Boolean, bSymbolInfoDone As Boolean, bStoppedOut As Boolean
    Dim bDoStocks As Boolean, bDoSpreads As Boolean, bCalendarSpread As Boolean
    Dim Bars As New cGdBars, Bars57 As New cGdBars
    Dim aTradeInfo() As String, aTradeInfoBar() As String
    Dim aSymbolInfo As New cGdArray
    Dim aSymbols As New cGdArray

    Dim nSymbolID2&
    Dim dEntry2#, dExit2#, dDollarsPerPoint2#, dRollAdjust2#, dPrebuilt#, dRebuilt#
    Dim strSymbol2$, strSpread$
    Dim Bars2 As New cGdBars, Bars57s As New cGdBars
    Dim Rolls As cGdTable

    ' this allows clicking on the button again in order to STOP this process if it's currently in progress
    Static bInProgress As Boolean
    If bInProgress Then
        If InfBox("Abort Seasonals?", "?", "Abort|+-No", "Abort") = "A" Then
            bInProgress = False
        End If
        Exit Sub
    End If
    
    ' don't let normal clients run this accidentally
    If Not FileExist("c:\common\files.exe") Then Exit Sub
    
    ' Initialize things based on which type being done (get settings from the Seasonals.INI file)
    bDoStocks = False
    bDoSpreads = False
    bUsePercentage = False
    AddList "Ram = " & Str(Int(PhysicalRAM(True))) & " of " & Str(Int(PhysicalRAM(False))) & " mb"
    s = InfBox("This runs the really long process to calculate all the Seasonals data.", "?", "Stocks|Futures|S&preads", "Seasonal Calcs")
    If s = "F" Then
        strSecType = "Futures"
        strDataFile = "FutSeasonals.dat"
        strSymbolInfoFile = "FutSymbols.dat"
    ElseIf s = "S" Then
        bDoStocks = True
        bUsePercentage = True
        strSecType = "Stocks"
        strDataFile = "StkSeasonals.dat"
        strSymbolInfoFile = "StkSymbols.dat"
    Else
        bDoSpreads = True
        strSecType = "Spreads"
        strDataFile = "SpreadSeasonals.dat"
        strSymbolInfoFile = "SpreadSymbols.dat"
    End If
    strSymbol = GetIniFileProperty("Symbols", "", strSecType, App.Path & "\Seasonals.INI")
    dMinPF = GetIniFileProperty("MinPF", 0, strSecType, App.Path & "\Seasonals.INI")
    dMinWinPerc = GetIniFileProperty("MinWin%", 0, strSecType, App.Path & "\Seasonals.INI")
    iYear = GetIniFileProperty("StartYear", 1970, strSecType, App.Path & "\Seasonals.INI")
    nStartDate = DateSerial(iYear, 1, 1) ' convert StartYear to the starting date
    strPath = AddSlash(GetIniFileProperty("OutPath", "C:\", strSecType, App.Path & "\Seasonals.INI"))
    d = GetIniFileProperty("TradeHistory", 1, strSecType, App.Path & "\Seasonals.INI")
    If d = 0 Then
        bIncludeTradeHistory = False
    Else
        bIncludeTradeHistory = True
    End If
    If dMinPF < 1 Or dMinWinPerc < 50 Or Len(strSymbol) = 0 Or Len(strPath) = 0 Then
        AddList "Invalid settings from Seasonals.INI file"
        Exit Sub
    End If
    strDataFile = strPath & strDataFile
    strSymbolInfoFile = strPath & strSymbolInfoFile
        
    aSymbols.Size = 0
    strExch = ""
    If bDoSpreads Then
        ' load all the SPREAD symbols
        If InStr(strSymbol, "-") > 0 Then
            ' must just be testing with some specific symbols
            aSymbols.SplitFields strSymbol, ","
        Else
            ' otherwise this is the list of Groups (e.g. "Energies,Metals,Grains,Meats,Treasuries,Currencies,Equities,*Softs")
            strExch = strSymbol
            For d = 1 To 99
                ' get category from list
                strText = Parse(strExch, ",", d)
                If Len(strText) = 0 Then Exit For
                If Left(strText, 1) = "*" Then
                    bSkip = True ' skip Intermarket spreads for this category
                    strText = Mid(strText, 2)
                Else
                    bSkip = False
                End If
                ' get symbols for this category
                s = GetIniFileProperty(strText, "", strSecType, App.Path & "\Seasonals.INI")
                aSymbolInfo.SplitFields s, ","
                For iSymbol = 0 To aSymbolInfo.Size - 1
                    strSymbol = aSymbolInfo(iSymbol)
                    If Not bSkip Then
                        ' add an Intermarket spread with each of the other symbols in the same category
                        For i = 0 To aSymbolInfo.Size - 1
                            If i <> iSymbol Then
                                strSymbol2 = aSymbolInfo(i)
                                ' e.g. "G6E - G6J"
                                aSymbols.Add strSymbol & " - " & strSymbol2
                            End If
                        Next
                    End If
                    ' and add a Calendar spread for each combination of monthly contracts
                    For i = 1 To 12
                        ' for each month, check if a Gann contract exists
                        s = strSymbol & "-" & Format(80 + i, "000")
                        nSymbolID = GetSymbolID(s)
                        If nSymbolID > 0 Then
                            For j = 1 To 12
                                If j <> i Then
                                    s = strSymbol & "-" & Format(80 + j, "000")
                                    nSymbolID2 = GetSymbolID(s)
                                    If nSymbolID2 > 0 Then
                                        ' e.g. "G6E: Mar - Dec"
                                        s = MonthName(i, True, True) & " - " & MonthName(j, True, True)
                                        aSymbols.Add strSymbol & ": " & s
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
            Next
        End If
    Else
        ' for NON-SPREADS: load the Symbols array (e.g. IBM, DOW30.GRP, ES-067, ES-)
        aSymbols.SplitFields UCase(strSymbol), ","
        ' first replace any *.GRP entry with the list of symbols in that group
        For iSymbol = aSymbols.Size - 1 To 0 Step -1
            strSymbol = aSymbols(iSymbol)
            If UCase(Right(strSymbol, 4)) = ".GRP" Then
                aSymbols.Remove iSymbol
                s = g.SymbolPool.GetSymbolsForGroup(strSymbol)
                aSymbolInfo.SplitFields UCase(s), vbTab
                aSymbols.AppendFromArray aSymbolInfo
                aSymbolInfo.Size = 0
            End If
        Next
        ' then cleanup each symbol
        For iSymbol = aSymbols.Size - 1 To 0 Step -1
            strSymbol = aSymbols(iSymbol)
            If bDoStocks Then
                ' ignore foreign stocks and forex@broker symbols
                If InStr(strSymbol, "@") > 0 Then
                    strSymbol = ""
                End If
            Else
                ' only use 67's for futures (can handle "ES", "ES-", "ES-067", or "ES-blah")
                i = InStr(strSymbol, "-")
                If i = 0 Then
                    strSymbol = strSymbol & "-067"
                Else
                    strSymbol = Left(strSymbol, i) & "067"
                End If
                ' and only do the electronic 67's (now that our electronic daily history goes all the way back)
                s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                If Len(s) > 0 Then
                    strSymbol = s
                End If
            End If
            aSymbols(iSymbol) = UCase(strSymbol)
        Next
    End If
    ' then sort list, and delete any duplicates and blanks
    aSymbols.Sort (eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues)
    If aSymbols.Size = 0 Then
        AddList "No symbols (.GRP file is missing?)"
        Exit Sub
    End If
    If aSymbols.Size < 150 Then
        For i = 0 To aSymbols.Size - 1
            AddList aSymbols(i)
        Next
    End If
    AddList "# Symbols = " & Str(aSymbols.Size)
'If IsIDE Then Exit Sub
       
       
    ' set ending date -- usually about 3-4 months ago
    i = Month(Date) - 4
    If i < 1 Then i = i + 12
    i = Val(InfBox(Str(aSymbols.Size) & " symbols to process.||Use data through end of which month?|(usually about 3-4 months ago)", "?", , "Calc Seasonals", , , , , , "n", Str(i)))
    If i <= 0 Or i > 12 Then
        Exit Sub
    End If
    i = i + 1 ' so we can more easily get to end of that month, add a month now
    If i > 12 Then i = i - 12
    iYear = Year(Date)
    If i > Month(Date) Then
        iYear = iYear - 1 ' last year
    End If
    nEndDate = DateSerial(iYear, i, 1) - 1 ' end of prior month
    
    
    'To better align data bars with TDOY's:
    '- use a 2D array of doubles (always init to -999999) -- and string array (same size)
    '- fields: date, tdoy, o, h, l, c, c57 -- and string of formatted date,price57
    '- fields: date, tdoy, c, c57, s, s57 -- and string of formatted date,price57-spread57
    '- size: # of all tdoy's from start date to end date (excluding normal holidays)
    '- put all prices from gdBars into this table (for primary and spread data), merging high/low data for holidays
    '- then can walk through the table, for each row as entry, and each # of rows beyond for #DaysHeld, and store stats in TDOY table
    iBar = (nEndDate - nStartDate) * 5 / 7# + 600
    ReDim BarsData(kOpen, iBar) As Double
    ' For this kind of analysis, probably best to remove even the past "normal holidays"
    ' to provide better consistency (so all the TDOY's "line up" better throughout the year)
    strExch = "NKPGMJLTX" ' the normal 9 major holidays (at least back to 1998)
    iBar = 0
    iEndDateBar = 0
    iTDOY = 0
    For d = nStartDate To nEndDate + 600 ' need to go 6 months after 1 year window
        bSkip = True
        If IsWeekday(d) Then
            If Len(strExch) = 0 Then
                bSkip = False
            ElseIf gdIsHoliday(d, strExch) <> 0 Then
                s = DateFormat(d)
                'AddList "Holiday: " & s
            Else
                bSkip = False
            End If
        End If
        If Not bSkip Then
            ' if starting a new year, then reset the TDOY
            If Year(d) <> Year(BarsData(kDate, iBar)) Then
                iTDOY = 0
            End If
            iTDOY = iTDOY + 1
            iBar = iBar + 1
            BarsData(kTDOY, iBar) = iTDOY
            BarsData(kDate, iBar) = d
            If d <= nEndDate Then
                iEndDateBar = iBar ' last bar of "historical window" (just prior to the 1-year "forecast window")
            End If
        End If
    Next
    iNumBars = iBar
    ReDim Preserve BarsData(kOpen, iNumBars) As Double
    
#If 0 Then
    'iBar = 0
    For iBar = iBar To iNumBars
        iTDOY = BarsData(kTDOY, iBar)
        d = BarsData(kDate, iBar)
        If iTDOY <= 20 Or iTDOY >= 240 Or iBar < 10 Or iBar > iNumBars - 10 Then
            s = Str(iBar) & vbTab & Str(iTDOY) & vbTab & DateFormat(d) & vbTab & WeekdayName(d)
            AddList s
            If iTDOY = 20 Then
                AddList "==================="
            End If
        End If
    Next
    Exit Sub
#End If

    ' write header line of the data file
    fh = FreeFile
    Open strDataFile For Output As #fh
    strText = "'Symbol" & vbTab & "SymbolID" & vbTab & "Long/Short" & vbTab & "TDOY" & vbTab & "EntryDate" & vbTab _
            & "DaysInTrade" & vbTab & "ExitDate" & vbTab & "StopLoss" & vbTab & "#Trades" & vbTab _
            & "Win%" & vbTab & "ProfitFactor" & vbTab & "AvgTrade" & vbTab & "AvgPerDay" & vbTab _
            & "AvgDD" & vbTab & "WorstDD" & vbTab & "AAP2DD" & vbTab & "Kelly%" & vbTab & "TradeHistory"
    Print #fh, strText
    iLineCount = 0 ' iLineCount + 1
    dFileSize = 0
   
    ' init things for running the big loop
    bInProgress = True
    dTimeStarted = gdTickCount
    aSymbolInfo.Size = 0
    ReDim tblTradeStrings(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Long
    For iSymbol = 0 To aSymbols.Size - 1
        If Not bInProgress Then Exit For ' in case want to abort
        bSkip = False
        
        ' get Symbol -- check to see if it's a spread (e.g. "G6E - G6J", "G6E: Sep - Dec")
        strSymbol = aSymbols(iSymbol)
        strSpread = ""
        strSymbol2 = ""
        nSymbolID2 = 0
        bCalendarSpread = False
        If Not bDoSpreads Then
            nSymbolID = GetSymbolID(strSymbol)
        Else
            strSpread = strSymbol
            ' check for a calendar spread (e.g. "G6E: Dec - Jun", with or without spaces)
            If InStr(strSpread, ":") > 0 Then
                bCalendarSpread = True
                s = Parse(strSpread, ":", 2)
                i = MonthNumber(Parse(s, "-", 1))
                j = MonthNumber(Parse(s, "-", 2))
                If i > 0 And j > 0 Then
                    strSymbol = Parse(strSpread, ":", 1) & "-" & Format(80 + i, "000")
                    strSymbol2 = Parse(strSpread, ":", 1) & "-" & Format(80 + j, "000")
                End If
            ElseIf InStr(strSpread, "-") > 0 Then
                ' or an intermarket spread (e.g. "G6E - G6J")
                strSymbol = Trim(UCase(Parse(strSpread, "-", 1))) & "-067"
                strSymbol2 = Trim(UCase(Parse(strSpread, "-", 2))) & "-067"
            End If
            
            ' BUT for spreads we MUST use the settles from the Combined or Pit symbols (at a consistent time of day),
            ' instead of using the last price from the Electronic symbols (which could be early morning for far-out contracts)
            s = ConvertFutureSymbol(strSymbol, eCombinedSymbol)
            If Len(s) > 0 Then strSymbol = s
            nSymbolID = GetSymbolID(strSymbol)
            If nSymbolID <= 0 Then
                s = ConvertFutureSymbol(strSymbol, ePitSymbol)
                If Len(s) > 0 Then strSymbol = s
                nSymbolID = GetSymbolID(strSymbol)
                If nSymbolID <= 0 Then
                    s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                    If Len(s) > 0 Then strSymbol = s
                    nSymbolID = GetSymbolID(strSymbol)
                End If
            End If
            s = ConvertFutureSymbol(strSymbol2, eCombinedSymbol)
            If Len(s) > 0 Then strSymbol2 = s
            nSymbolID2 = GetSymbolID(strSymbol2)
            If nSymbolID2 <= 0 Then
                s = ConvertFutureSymbol(strSymbol2, ePitSymbol)
                If Len(s) > 0 Then strSymbol2 = s
                nSymbolID2 = GetSymbolID(strSymbol2)
                If nSymbolID2 <= 0 Then
                    s = ConvertFutureSymbol(strSymbol2, eElectronicSymbol)
                    If Len(s) > 0 Then strSymbol2 = s
                    nSymbolID2 = GetSymbolID(strSymbol2)
                End If
            End If
        End If
        
        strSecType = SecurityType(strSymbol, True)
        If Len(strSpread) > 0 Then
            strName = strSpread
        ElseIf bDoStocks Then
            strName = strSymbol
        Else
            strName = Parse(strSymbol, "-", 1) '& "-"
        End If
        
        ' Estimated Total (time or size) = SoFar (time or size) * TotalCount / DoneCount
        If iSymbol > 0 Then
            dChk = dFileSize * aSymbols.Size / iSymbol ' total size of data file (estimated)
            d = (gdTickCount - dTimeStarted) / 3600000# ' # hours running so far
            d = d * aSymbols.Size / iSymbol - d ' estimated # hours remaining (i.e. total - done)
        Else
            dChk = 0
            d = 0
        End If
        If nSymbolID <= 0 Then
            s = "*** " & strSymbol & " does not exist ***"
            bSkip = True
        ElseIf bDoSpreads And nSymbolID2 <= 0 Then
            s = "*** " & strSymbol2 & " does not exist (" & strName & ") ***"
            bSkip = True
        Else
            s = Str(iSymbol + 1) & " of " & Str(aSymbols.Size) & " ( " & Format(100# * iSymbol / aSymbols.Size, "#0.0") & "% ), " & strName _
                & ", FileSize = " & Str(Int(dFileSize / 1000000#)) & " of " & Str(Int(dChk / 1000000#)) & " mb, " & Format(d, "#0.00") & " hrs left"
        End If
        AddList s
        s = s & ", Ram = " & Str(Int(PhysicalRAM(True))) & " mb"
        DebugLog s
        DoEvents
        
If IsIDE Then
    'bSkip = True
    If Left(strSymbol, 3) <> "ZS-" Or Not bCalendarSpread Then
        'bSkip = True
    End If
End If

        ' load Primary data
        dDollarsPerPoint = 0
        dDollarsPerPoint2 = 0
        bSymbolInfoDone = False
        If Not bSkip Then
            DM_GetBars Bars, nSymbolID, "Daily", nStartDate, nEndDate
            dMinMove = Bars.MinMove
            If strSecType = "F" Then
                ' Futures must have a valid tick value and move
                If Bars.Prop(eBARS_TickValue) <= 0 Or Bars.Prop(eBARS_TickMove) <= 0 Then
                    bSkip = True
                Else
                    dDollarsPerPoint = Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
                    ' also load the 057 data (in order to determine the "forward roll adjust" from the closing prices)
                    If bCalendarSpread Then
                        ' except for calendar spreads (which use the Gann contracts) -- just use a copy
                        Set Bars57 = Bars.MakeCopy
                    Else
                        DM_GetBars Bars57, Parse(strSymbol, "-", 1) & "-057", "Daily", Bars(eBARS_DateTime, 0), Bars(eBARS_DateTime, Bars.Size - 1)
                    End If
                    If Bars57.Size <> Bars.Size Then
                        bSkip = True ' shouldn't happen!
                        If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Bars57 Error"
                    End If
                End If
            Else
                ' but for Stocks/Indexes the prices cannot go negative (since doing percentage changes)
                If gdMinValue(Bars.ArrayHandle(eBARS_Close), 0, Bars.Size - 1) <= 0 Then
                    bSkip = True
                End If
                ' and check for any gaps in data > 2 weeks -- if so, need to ignore
                ' what's before the gap (since could be from an entirely different company)
                ' - or if price jumps up or down by factor of 10
                For iBar = Bars.Size - 1 To 1 Step -1
                    i = Bars(eBARS_DateTime, iBar) - Bars(eBARS_DateTime, iBar - 1)
                    ' check for a large percentage gap
                    d = Bars(eBARS_Close, iBar - 1)
                    If d > 0 Then
                        d = Bars(eBARS_Close, iBar) / d
                    End If
                    If i >= 14 Or d < 0.1 Or d > 10 Then
                        Bars.DeleteFirstBars iBar ' ignore all data prior to this bar
                        Exit For
                    End If
                Next
                Bars57.Size = 0
                SetBarProperties Bars57, nSymbolID
            End If
            
            ' must be at least 5 years of data, and have fairly recent data
            If Bars(eBARS_DateTime, 0) > nEndDate - 365 * 5 Then
                bSkip = True
            ElseIf Bars(eBARS_DateTime, Bars.Size - 1) < nEndDate - 30 Then
                bSkip = True
            End If
        End If
        
        ' load Spread data
        If nSymbolID2 > 0 And Not bSkip Then
            DM_GetBars Bars2, nSymbolID2, "Daily", nStartDate, nEndDate
            ' Futures must have a valid tick value and move
            If Bars2.Prop(eBARS_TickValue) <= 0 Or Bars2.Prop(eBARS_TickMove) <= 0 Then
                bSkip = True
            Else
                dDollarsPerPoint2 = Bars2.Prop(eBARS_TickValue) / Bars2.Prop(eBARS_TickMove)
                ' also load the 057 data (in order to determine the "forward roll adjust" from the closing prices)
                If bCalendarSpread Then
                    ' except for calendar spreads (which use the Gann contracts) -- just use a copy
                    Set Bars57s = Bars2.MakeCopy
                Else
                    DM_GetBars Bars57s, Parse(strSymbol2, "-", 1) & "-057", "Daily", Bars2(eBARS_DateTime, 0), Bars2(eBARS_DateTime, Bars2.Size - 1)
                End If
                If Bars57s.Size <> Bars2.Size Then
                    bSkip = True ' shouldn't happen!
                    If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Bars57s Error"
                End If
            End If
            
            ' must be at least 5 years of data, and have fairly recent data
            If Bars2(eBARS_DateTime, 0) > nEndDate - 365 * 5 Then
                bSkip = True
            ElseIf Bars2(eBARS_DateTime, Bars2.Size - 1) < nEndDate - 30 Then
                bSkip = True
            End If
        End If
        
        ' Place the gdBars prices into the BarsData table (so all bars will be better aligned with "real" TDOY's)
        If Not bSkip Then
            ' clear out all the BarsData prices (but not the dates)
            For iBar = 0 To iNumBars
                For i = kDate + 1 To kOpen
                    BarsData(i, iBar) = kNullData
                Next
            Next
            
            iRow = iEndDateBar ' iNumBars
            For iBar = Bars.Size - 1 To 0 Step -1
                d = Bars(eBARS_DateTime, iBar)
                ' find row in table where bar date > date of prior table row
                Do While iRow >= 1
                    If d > BarsData(kDate, iRow - 1) Then
                        Exit Do
                    End If
                    iRow = iRow - 1
                Loop
                If d <> BarsData(kDate, iRow) Then
                    s = DateFormat(d)
                    If IsIDE Then
                        'AddList "Merge holiday: " & s
                    End If
                End If
                ' if no data yet for that row, then set all the prices
                If BarsData(kClose, iRow) = kNullData Then
                    BarsData(kClose, iRow) = Bars(eBARS_Close, iBar)
                    If bDoStocks Then
                        BarsData(kClose57, iRow) = Bars(eBARS_Close, iBar)
                    Else
                        BarsData(kClose57, iRow) = Bars57(eBARS_Close, iBar)
                    End If
                    If Not bDoSpreads Then
                        BarsData(kOpen, iRow) = Bars(eBARS_Open, iBar)
                        BarsData(kHigh, iRow) = Bars(eBARS_High, iBar)
                        BarsData(kLow, iRow) = Bars(eBARS_Low, iBar)
                    End If
                ElseIf Not bDoSpreads Then
                    ' else just merge the High/Low into this row (e.g. a "half-holiday bar")
                    If BarsData(kHigh, iRow) < Bars(eBARS_High, iBar) Then
                        BarsData(kHigh, iRow) = Bars(eBARS_High, iBar)
                        ' if a higher high, then also check for a higher open (in case of an overnight gap)
                        If BarsData(kOpen, iRow) < Bars(eBARS_Open, iBar) Then
                            BarsData(kOpen, iRow) = Bars(eBARS_Open, iBar)
                        End If
                    End If
                    If BarsData(kLow, iRow) > Bars(eBARS_Low, iBar) Then
                        BarsData(kLow, iRow) = Bars(eBARS_Low, iBar)
                        ' if a lower low, then also check for a lower open (in case of an overnight gap)
                        If BarsData(kOpen, iRow) > Bars(eBARS_Open, iBar) Then
                            BarsData(kOpen, iRow) = Bars(eBARS_Open, iBar)
                        End If
                    End If
                End If
            Next
            
            If bDoSpreads Then
                iRow = iEndDateBar ' iNumBars
                For iBar = Bars2.Size - 1 To 0 Step -1
                    d = Bars2(eBARS_DateTime, iBar)
                    ' find row in table where bar date > date of prior table row
                    Do While iRow >= 1
                        If d > BarsData(kDate, iRow - 1) Then
                            Exit Do
                        End If
                        iRow = iRow - 1
                    Loop
                    ' if no data yet for that row, then set all the prices
                    ' (no open/high/low or holiday merging needed for spreads)
                    If BarsData(kSpread, iRow) = kNullData Then
                        BarsData(kSpread, iRow) = Bars2(eBARS_Close, iBar)
                        BarsData(kSpread57, iRow) = Bars57s(eBARS_Close, iBar)
                    End If
                Next
                
                ' and for calendar spreads, flag the roll dates for both legs
                ' (this is actually the "expiration dates" -- i.e. one data bar prior to the roll)
                If bCalendarSpread Then
                    ' first leg
                    Set Rolls = GetRollsTable(nSymbolID)
                    If Rolls Is Nothing Then
                        bSkip = True
                    Else
                        dLowest = 0
                        iBar = iEndDateBar ' iNumBars - 1
                        For iRoll = Rolls.NumRecords - 1 To 0 Step -1
                            d = Rolls.Num(1, iRoll) ' date of roll
                            Do While iBar > 0 And d <= nEndDate
                                ' looking backwards, find the first bar where date < roll date and has good data
                                ' (i.e. this will be the "expiration date" for one of the Gann yearly contracts)
                                If BarsData(kDate, iBar) < d And BarsData(kClose, iBar) <> kNullData Then
                                    ' for first leg, just set to 1
                                    BarsData(kRolled, iBar) = 1
                                    iTDOY = BarsData(kTDOY, iBar)
                                    ' track the "lowest TDOY" for a roll
                                    If dLowest = 0 Then
                                        ' if it hasn't been set yet
                                        dLowest = iTDOY
                                    ElseIf iTDOY > 175 And dLowest < 75 Then
                                        ' if this TDOY is at end of the year and lowest so far is at beginning of the year
                                        dLowest = iTDOY
                                    ElseIf iTDOY < 75 And dLowest > 175 Then
                                        ' but ignore this TDOY if at beginning of the year when lowest is at end of the year
                                        dLowest = dLowest
                                    ElseIf iTDOY < dLowest Then
                                        ' else use this TDOY if it's the lowest so far
                                        dLowest = iTDOY
                                    End If
                                    Exit Do
                                End If
                                iBar = iBar - 1
                            Loop
                        Next
                        AddList "For " & strSymbol & ": Lowest roll = " & Str(dLowest)
                        
                        ' now mark the "dead zones" in the forecast window
                        bSkip = False
                        j = Val(Right(strSymbol, 3)) - 80 ' contract month
                        For iBar = iNumBars To iEndDateBar + 1 Step -1 ' loop backwards from end
                            ' mark this bar as in dead zone when "bSkip" flag has been set
                            If bSkip Then
                                BarsData(kRolled, iBar) = 9
                            End If
                            ' when hit the end of the contract month, start marking a dead zone
                            If Month(BarsData(kDate, iBar)) <> j And Month(BarsData(kDate, iBar - 1)) = j Then
                                bSkip = True
                            ElseIf BarsData(kTDOY, iBar) = dLowest Then
                                ' when get to lowest TDOY roll, we can stop the dead zone after this bar
                                bSkip = False
                            End If
                        Next
                    End If
                    
                    ' second leg
                    Set Rolls = GetRollsTable(nSymbolID2)
                    If Rolls Is Nothing Then
                        bSkip = True
                    Else
                        dLowest = 0
                        iBar = iEndDateBar ' iNumBars - 1
                        For iRoll = Rolls.NumRecords - 1 To 0 Step -1
                            d = Rolls.Num(1, iRoll) ' date of roll
                            Do While iBar > 1 And d <= nEndDate
                                ' looking backwards, find the first bar where date < roll date and has good data
                                ' (i.e. this will be the "expiration date" for one of the Gann yearly contracts)
                                If BarsData(kDate, iBar) < d And BarsData(kSpread, iBar) <> kNullData Then
                                    ' for second leg, set to 2 -- except if rolled for both then set to 3
                                    ' (so 1 = 1st leg, 2 = 2nd leg, 3 = both legs rolled on this date)
                                    If BarsData(kRolled, iBar) > 0 Then
                                        BarsData(kRolled, iBar) = 3
                                    Else
                                        BarsData(kRolled, iBar) = 2
                                    End If
                                    iTDOY = BarsData(kTDOY, iBar)
                                    ' track the "lowest TDOY" for a roll
                                    If dLowest = 0 Then
                                        ' if it hasn't been set yet
                                        dLowest = iTDOY
                                    ElseIf iTDOY > 175 And dLowest < 75 Then
                                        ' if this TDOY is at end of the year and lowest so far is at beginning of the year
                                        dLowest = iTDOY
                                    ElseIf iTDOY < 75 And dLowest > 175 Then
                                        ' but ignore this TDOY if at beginning of the year when lowest is at end of the year
                                        dLowest = dLowest
                                    ElseIf iTDOY < dLowest Then
                                        ' else use this TDOY if it's the lowest so far
                                        dLowest = iTDOY
                                    End If
                                    Exit Do
                                End If
                                iBar = iBar - 1
                            Loop
                        Next
                        AddList "For " & strSymbol2 & ": Lowest roll = " & Str(dLowest)
                        
                        ' now mark the "dead zones" in the forecast window
                        bSkip = False
                        j = Val(Right(strSymbol2, 3)) - 80 ' contract month
                        For iBar = iNumBars To iEndDateBar + 1 Step -1 ' loop backwards from end
                            ' mark this bar as in dead zone when "bSkip" flag has been set
                            If bSkip Then
                                BarsData(kRolled, iBar) = 9
                            End If
                            ' when hit the end of the contract month, start marking a dead zone
                            If Month(BarsData(kDate, iBar)) <> j And Month(BarsData(kDate, iBar - 1)) = j Then
                                bSkip = True
                            ElseIf BarsData(kTDOY, iBar) = dLowest Then
                                ' when get to lowest TDOY roll, we can stop the dead zone after this bar
                                bSkip = False
                            End If
                        Next
                    End If
                    Set Rolls = Nothing
                    
                    If IsIDE And aSymbols.Size <= 3 Then
                        For iBar = 0 To iNumBars
                            Select Case BarsData(kRolled, iBar)
                            Case 1:
                                s = strSymbol
                            Case 2:
                                s = strSymbol2
                            Case 3:
                                s = strSymbol & " & " & strSymbol2
                            Case 9:
                                s = "Dead zone"
                            Case Else:
                                s = ""
                            End Select
                            If Len(s) > 0 Then
                                d = BarsData(kDate, iBar)
                                i = BarsData(kTDOY, iBar)
                                AddList "Roll TDOY = " & Format(i, "000") & ", " & DateFormat(d) & ": " & s
                            End If
                        Next
                    End If
                End If
            End If
            
            ' and prebuild the basic trade info (date,price) for each bar
            ' - it's MUCH faster to format these strings just once for each bar (instead of 1000's of times in the nested For loops below)
            ' - and takes significantly less memory when storing all the trade history (esp. for spreads)
            ReDim aTradeInfo(iNumBars) As String
            ReDim aTradeInfoBar(iNumBars) As String
            For iBar = 0 To iNumBars
                strText = Bars.PriceDisplay(BarsData(kClose57, iBar))
                If bDoSpreads Then
                    strText = strText & " - " & Bars2.PriceDisplay(BarsData(kSpread57, iBar))
                End If
                aTradeInfo(iBar) = Format(BarsData(kDate, iBar), "mm/dd/yyyy") & "," & strText & ","
                aTradeInfoBar(iBar) = "[" & Str(iBar) & "]" ' believe it or not, this actually saves time later!
            Next
        End If

If IsIDE Then
    'If bDoSpreads Then bSkip = True
End If
            
        ' Do seasonal calcs for both LONGS and SHORTS (except do only Longs for all Spreads)
        bShort = False
        Do While bInProgress And Not bSkip
            If bShort Then
                If bDoSpreads Then Exit Do
                strDirection = "Short"
            Else
                strDirection = "Long"
            End If
            
            ' clear all the 3D tables (except tblTradeStrings)
            ReDim tblNumUp(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Long
            ReDim tblNumDown(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Long
            ReDim tblAmtUp(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            ReDim tblAmtDown(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            ReDim tblDD(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            ReDim tblWorstDD(1 To kMaxTDOY, 1 To kMaxDaysHeld, 1 To kMaxStopLosses) As Double
            If bIncludeTradeHistory Then
                ' must clear each of the gdString of trades for each item
                For iTDOY = 1 To kMaxTDOY
                    For iDaysHeld = 1 To kMaxDaysHeld
                        For iStopLoss = 1 To kMaxStopLosses
                            hTrades = tblTradeStrings(iTDOY, iDaysHeld, iStopLoss)
                            ' if string array does not yet exist, create it
                            If hTrades = 0 Then
                                'hTrades = gdCreateArray(eGDARRAY_Strings, 0)
                                hTrades = gdCreateString(0)
                                tblTradeStrings(iTDOY, iDaysHeld, iStopLoss) = hTrades
                            Else ' otherwise just clear it
                                gdSetSize hTrades, 0, False
                                gdFreeExtra hTrades
                            End If
                        Next
                    Next
                Next
            End If
            dTradeHistMemSize = 0
            'AddList "Ram = " & Str(Int(PhysicalRAM(True))) & " of " & Str(Int(PhysicalRAM(False)))

            ' for each Entry bar
            For iBar = 1 To iEndDateBar 'iNumBars
                If Not bInProgress Then Exit For
                
                If iBar = 1 Or (iBar Mod 1000) = 0 Then
                    AddList "   " & strDirection & ": " & Str(iBar) & " of " & Str(iEndDateBar) & " bars, Ram = " & Str(Int(PhysicalRAM(True))) & " mb"
                    DoEvents
                End If
                
                ' make sure this bar is a valid entry (date, price, TDOY)
                iTDOY = BarsData(kTDOY, iBar)
                d = BarsData(kDate, iBar)
                dEntry = BarsData(kClose, iBar)
                If iTDOY <= 0 Or iTDOY > kMaxTDOY Or d = kNullData Or d > nEndDate Or dEntry = kNullData Then
                    iTDOY = 0
                ElseIf bDoSpreads Then
                    ' can only enter a spread on a date which has a valid price for BOTH legs
                    dEntry2 = BarsData(kSpread, iBar)
                    If dEntry2 = kNullData Then
                        iTDOY = 0
                    End If
                End If
                    
                If iTDOY > 0 Then
                    ' calc the forward roll-adjust for Futures (in case rolls between entry and exit)
                    dRollAdjust = BarsData(kClose57, iBar) - BarsData(kClose, iBar)
                    dEntry = dEntry + dRollAdjust
                    If bDoSpreads Then
                        dRollAdjust2 = BarsData(kSpread57, iBar) - BarsData(kSpread, iBar)
                        dEntry2 = dEntry2 + dRollAdjust2
                    End If
                    
                    ' calc stats separately for each StopLoss value
                    For iStopLoss = 1 To kMaxStopLosses
                        ' determine price of the StopLoss to use
                        If bUsePercentage Then
                            dStopLoss = iStopLoss * 2.5 / 100# ' at 2.5% increments
                            dStopPrice = dEntry * dStopLoss
                        Else
                            dStopLoss = iStopLoss * 500# ' at $500 increments
                            dStopPrice = dStopLoss / dDollarsPerPoint
                        End If
                        If Not bDoSpreads Then
                            If bShort Then
                                ' move to the next highest tick
                                dStopPrice = dEntry + dStopPrice
                                dStopPrice = RoundToSigDigits(dStopPrice / dMinMove)
                                If dStopPrice <> Int(dStopPrice) Then
                                    dStopPrice = Int(dStopPrice) + 1
                                End If
                            Else
                                ' move to the next lowest tick
                                dStopPrice = dEntry - dStopPrice
                                dStopPrice = Int(RoundToSigDigits(dStopPrice / dMinMove))
                            End If
                            dStopPrice = RoundToSigDigits(dStopPrice * dMinMove)
                        End If
                        
                        ' for each #days held
                        bStoppedOut = False
                        dLowest = dEntry
                        dHighest = dEntry
                        dDD = 0
                        strTrade = ""
                        For iDaysHeld = 1 To kMaxDaysHeld
                            ' check if beyond end of bars (shouldn't ever happen)
                            If iBar + iDaysHeld > iNumBars Then
                                If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Invalid Exit Bar"
                                Exit For
                            End If
                            
                            ' see if this is a valid Exit bar (valid date and prices)
                            If BarsData(kDate, iBar + iDaysHeld) > nEndDate Then
                                Exit For ' can exit loop if end date is beyond end of good data
                            ElseIf bCalendarSpread Then
                                ' for calendar spreads, trades cannot be held across the "roll" of either leg
                                If BarsData(kRolled, iBar + iDaysHeld) <> kNullData Then
                                    Exit For ' can exit loop since cannot exit after this date
                                End If
                            End If
                            dExit = BarsData(kClose, iBar + iDaysHeld)
                            If dExit <> kNullData Then
                                If bDoSpreads Then
                                    dExit2 = BarsData(kSpread, iBar + iDaysHeld)
                                    If dExit2 = kNullData Then
                                        dExit = kNullData
                                    End If
                                End If
                            End If
                            
                            ' if already stopped out (or if not a valid Exit date), then just
                            ' use the same info for the stopped-out trade (i.e. same dNet, dDD, strTrade)
                            If dExit <> kNullData And Not bStoppedOut Then
                                ' get exit bar prices -- but for Futures, add the "roll adjust"
                                ' (i.e. this is a "forward-price-adjust" if rolls between the entry and exit)
                                dExit = dExit + dRollAdjust
                                dExit2 = dExit2 + dRollAdjust2
                                
                                If bDoSpreads Then
'If iTDOY = 71 And iDaysHeld = 15 And dStopLoss = 3000 Then
'    i = i
'End If

                                    ' for SPREADS, we can only check the Close (since can't really calculate High/Low of spread bars)
                                    ' calculate net profit of the spread (Buy#1 - Sell#2)
                                    dNet = (dExit - dEntry) * dDollarsPerPoint
                                    d = (dExit2 - dEntry2) * dDollarsPerPoint2
                                    dNet = RoundToSigDigits(dNet - d)
                                    ' check for newer drawdown
                                    If dNet < dDD Then
                                        dDD = dNet
                                    End If
                                    ' check if hit stop-loss
                                    If dNet <= -dStopLoss Then
                                        bStoppedOut = True
                                    End If
                                    ' build trade string (EntryDate, EntryPrice1 - EntryPrice2, ExitDate, ExitPrice1 - ExitPrice2, NetProfit, Drawdown)
                                    strTrade = Format(dNet, "#0.00") & "," & Format(dDD, "#0.00")
                                    If Abs(dExit - BarsData(kClose57, iBar + iDaysHeld)) < 0.000000001 _
                                            And Abs(dExit2 - BarsData(kSpread57, iBar + iDaysHeld)) < 0.000000001 Then
                                        ' can just use the prebuilt "ExitDate,Close1 - Close2" string
                                        ' (but store bar-number-placeholders in order to use a LOT less memory!)
                                        'strTrade = aTradeInfo(iBar) & aTradeInfo(iBar + iDaysHeld) & strTrade
                                        strTrade = aTradeInfoBar(iBar) & aTradeInfoBar(iBar + iDaysHeld) & strTrade
                                        dPrebuilt = dPrebuilt + 1
                                    Else
                                        ' else can use the prebuilt ExitDate, but must replace the ExitPrices
                                        s = Left(aTradeInfo(iBar + iDaysHeld), 11) ' e.g. "01/23/1992,"
                                        If Right(s, 1) <> "," Then
                                            If IsIDE Then InfBox "This shouldn't happen!", "i", , "ERROR"
                                        End If
                                        'strTrade = aTradeInfo(iBar) & s & Bars.PriceDisplay(dExit) & " - " & Bars2.PriceDisplay(dExit2) & "," & strTrade
                                        strTrade = aTradeInfoBar(iBar) & s & Bars.PriceDisplay(dExit) & " - " & Bars2.PriceDisplay(dExit2) & "," & strTrade
                                        dRebuilt = dRebuilt + 1
                                    End If
                                Else
                                    ' for NON-SPREADS, let's check intrabar drawdown and stop (using High/Low of bars)
                                    dOpen = BarsData(kOpen, iBar + iDaysHeld) + dRollAdjust
                                    dHigh = BarsData(kHigh, iBar + iDaysHeld) + dRollAdjust
                                    dLow = BarsData(kLow, iBar + iDaysHeld) + dRollAdjust
                                
                                    ' to check drawdown
                                    If dLow < dLowest Then
                                        ' if a lower low for a Long trade, then check if hit stop-loss
                                        If Not bShort Then
                                            If dLow <= dStopPrice Then
                                                ' see if stop-loss was hit at open, or after open
                                                If dOpen <= dStopPrice Then
                                                    dLow = dOpen
                                                Else
                                                    dLow = dStopPrice
                                                End If
                                                dExit = dLow
                                                bStoppedOut = True
                                            End If
                                        End If
                                        dLowest = dLow
                                    End If
                                    If dHigh > dHighest Then
                                        ' if a higher high for a Short trade, then check if hit stop-loss
                                        If bShort Then
                                            If dHigh >= dStopPrice Then
                                                ' see if stop-loss was hit at open, or after open
                                                If dOpen >= dStopPrice Then
                                                    dHigh = dOpen
                                                Else
                                                    dHigh = dStopPrice
                                                End If
                                                dExit = dHigh
                                                bStoppedOut = True
                                            End If
                                        End If
                                        dHighest = dHigh
                                    End If
                                    
                                    If bUsePercentage Then
                                        dNet = (dExit - dEntry) / dEntry * 100#
                                        If bShort Then
                                            dNet = -dNet
                                            dDD = (dEntry - dHighest) / dEntry * 100#
                                        Else
                                            dDD = (dLowest - dEntry) / dEntry * 100#
                                        End If
                                        'strTrade = Format(dNet, "#0.00") & "%" & "," & Format(dDD, "#0.00") & "%"
                                        strTrade = Format(dNet, "#0.00") & "," & Format(dDD, "#0.00")
                                    Else
                                        dNet = (dExit - dEntry) * dDollarsPerPoint
                                        If bShort Then
                                            dNet = -dNet
                                            dDD = (dEntry - dHighest) * dDollarsPerPoint
                                        Else
                                            dDD = (dLowest - dEntry) * dDollarsPerPoint
                                        End If
                                        'strTrade = "$" & Format(dNet * dDollarsPerPoint, "#0.00") & "," & "$" & Format(dDD * dDollarsPerPoint, "#0.00")
                                        strTrade = Format(dNet, "#0.00") & "," & Format(dDD, "#0.00")
                                    End If
                                    
                                    If Abs(dExit - BarsData(kClose57, iBar + iDaysHeld)) < 0.000000001 Then
                                        ' can just use the prebuilt "ExitDate,ClosingPrice" string
                                        ' (but store bar-number-placeholders in order to use a LOT less memory!)
                                        'strTrade = aTradeInfo(iBar) & aTradeInfo(iBar + iDaysHeld) & strTrade
                                        'strTrade = "[" & Str(iBar) & "][" & Str(iBar + iDaysHeld) & "]" & strTrade
                                        strTrade = aTradeInfoBar(iBar) & aTradeInfoBar(iBar + iDaysHeld) & strTrade
                                        dPrebuilt = dPrebuilt + 1
                                    Else
                                        ' else can use the prebuilt ExitDate, but must replace the ExitPrice
                                        s = Left(aTradeInfo(iBar + iDaysHeld), 11) ' e.g. "01/23/1992,"
                                        If Right(s, 1) <> "," Then
                                            If IsIDE Then InfBox "This shouldn't happen!", "i", , "ERROR"
                                        End If
                                        'strTrade = aTradeInfo(iBar) & s & Bars.PriceDisplay(dExit) & "," & strTrade
                                        strTrade = aTradeInfoBar(iBar) & s & Bars.PriceDisplay(dExit) & "," & strTrade
                                        dRebuilt = dRebuilt + 1
                                    End If
                                End If
                            End If
                            
                            ' whether stopped-out or not, accumulate trade info for this 3D spot (TDOY, DaysHeld, StopLossLevel)
                            If Len(strTrade) > 0 Then
                                If dNet > 0 Then
                                    tblNumUp(iTDOY, iDaysHeld, iStopLoss) = tblNumUp(iTDOY, iDaysHeld, iStopLoss) + 1
                                    tblAmtUp(iTDOY, iDaysHeld, iStopLoss) = tblAmtUp(iTDOY, iDaysHeld, iStopLoss) + dNet
                                Else
                                    tblNumDown(iTDOY, iDaysHeld, iStopLoss) = tblNumDown(iTDOY, iDaysHeld, iStopLoss) + 1
                                    tblAmtDown(iTDOY, iDaysHeld, iStopLoss) = tblAmtDown(iTDOY, iDaysHeld, iStopLoss) + dNet
                                End If
                                
                                tblDD(iTDOY, iDaysHeld, iStopLoss) = tblDD(iTDOY, iDaysHeld, iStopLoss) + dDD
                                If dDD < tblWorstDD(iTDOY, iDaysHeld, iStopLoss) Then
                                    tblWorstDD(iTDOY, iDaysHeld, iStopLoss) = dDD
                                End If
    
                                ' and append the results for this trade onto the TradeHistory gdString in the table
                                hTrades = tblTradeStrings(iTDOY, iDaysHeld, iStopLoss)
                                If hTrades <> 0 Then
                                    gdSetStr hTrades, -1, ";" & strTrade ' use -1 to append
                                    dTradeHistMemSize = dTradeHistMemSize + 1 + Len(strTrade)
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            AddList "      TradeHistory = " & Str(Int(dTradeHistMemSize / 1000000)) & " mb, Ram = " & Str(Int(PhysicalRAM(True))) & " mb"

            ' Now write all the output for this symbol
            For iTDOY = 1 To kMaxTDOY
                ' find "Entry Bar" (find the next TDOY which occurs after the End Date)
                iEntryBar = 0
                For iBar = iNumBars To 0 Step -1
                    If BarsData(kDate, iBar) <= nEndDate Then
                        Exit For ' done looking
                    End If
                    If BarsData(kTDOY, iBar) = iTDOY Then
                        iEntryBar = iBar
                    End If
                Next
                
                ' for each "Exit Bar"
                For iDaysHeld = 1 To kMaxDaysHeld
                    If iEntryBar <= 0 Then
                        Exit For
                    End If
                    iExitBar = iEntryBar + iDaysHeld
                    If iExitBar > iNumBars Then
                        If IsIDE Then InfBox "This shouldn't have happened!", "!", , "Invalid Exit Bar"
                        Exit For
                    End If
                    
                    If bCalendarSpread Then
                        ' for calendar spreads, check if either the entry or exit is in a "dead zone"
                        ' (since no predicted entries allowed in a dead zone, and all trades must be exited before the next dead zone)
                        If BarsData(kRolled, iEntryBar) <> kNullData Or BarsData(kRolled, iExitBar) <> kNullData Then
                            Exit For
                        End If
                    End If
                    
                    For iStopLoss = 1 To kMaxStopLosses
                        If bUsePercentage Then
                            dStopLoss = iStopLoss * 2.5
                        Else
                            dStopLoss = iStopLoss * 500#
                        End If
                        
                        ' to be valid, needs to be a minimum # of occurances
                        iNumTrades = tblNumUp(iTDOY, iDaysHeld, iStopLoss) + tblNumDown(iTDOY, iDaysHeld, iStopLoss)
                        If iStopLoss > 1 And iNumTrades > 1 Then
                            ' but can ignore this "duplicate" if it's the same trades as the lower stop-loss amount
                            If tblAmtUp(iTDOY, iDaysHeld, iStopLoss) = tblAmtUp(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                If tblAmtDown(iTDOY, iDaysHeld, iStopLoss) = tblAmtDown(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                    If tblNumDown(iTDOY, iDaysHeld, iStopLoss) = tblNumDown(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                        If tblNumUp(iTDOY, iDaysHeld, iStopLoss) = tblNumUp(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                            If tblDD(iTDOY, iDaysHeld, iStopLoss) = tblDD(iTDOY, iDaysHeld, iStopLoss - 1) Then
                                                iNumTrades = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If iNumTrades >= 5 Then
                            If iTDOY = kMaxTDOY Then
                                i = i ' shouldn't really get here?
                            End If
                        
                            ' calc avg profit and Win%
                            dAvgNet = (tblAmtUp(iTDOY, iDaysHeld, iStopLoss) + tblAmtDown(iTDOY, iDaysHeld, iStopLoss)) / iNumTrades
                            dWinPct = 100# * tblNumUp(iTDOY, iDaysHeld, iStopLoss) / iNumTrades
                            
                            ' calc Profit Factor ($won/$lost), AvgWin, AvgLoss
                            If tblAmtDown(iTDOY, iDaysHeld, iStopLoss) < 0 Then
                                dPF = Abs(tblAmtUp(iTDOY, iDaysHeld, iStopLoss) / tblAmtDown(iTDOY, iDaysHeld, iStopLoss))
                            Else
                                dPF = 100 ' just use a really high number if no $lost
                            End If
                            If tblNumUp(iTDOY, iDaysHeld, iStopLoss) > 0 Then
                                dAvgWin = tblAmtUp(iTDOY, iDaysHeld, iStopLoss) / tblNumUp(iTDOY, iDaysHeld, iStopLoss)
                            Else
                                dAvgWin = 0
                            End If
                            If tblNumDown(iTDOY, iDaysHeld, iStopLoss) > 0 Then
                                dAvgLoss = tblAmtDown(iTDOY, iDaysHeld, iStopLoss) / tblNumDown(iTDOY, iDaysHeld, iStopLoss)
                            Else
                                dAvgLoss = 0
                            End If
                            
                            ' drawdown and ROI
                            dAvgDD = tblDD(iTDOY, iDaysHeld, iStopLoss) / iNumTrades
                            dWorstDD = tblWorstDD(iTDOY, iDaysHeld, iStopLoss)
                            If dAvgDD = 0 Then
                                dAAP2DD = 99999
                            Else
                                dAAP2DD = -dAvgNet / dAvgDD ' dWorstDD
                            End If

'HG3 - SI3   0   Long    22  42768   20  42797   2500    26  61.5    2.020   928.85  46.44   -1477.40    -6275.00    0.6287  31.1
If iTDOY = 22 And iDaysHeld = 20 And dStopLoss = 2500 Then
    i = i
End If
                            If dPF >= dMinPF And dWinPct >= dMinWinPerc And dAvgWin > 0 Then
                                ' output for table
                                If bUsePercentage Then
                                    strText = Format(dAvgNet, "#0.0000") & vbTab & Format(dAvgNet / iDaysHeld, "#0.0000") _
                                         & vbTab & Format(dAvgDD, "#0.0000") & vbTab & Format(dWorstDD, "#0.0000")
                                Else
                                    'strText = "$" & Format(dAvgNet * dDollarsPerPoint, "#0.00") & vbTab & "$" & Format(dAvgNet / iDaysHeld * dDollarsPerPoint, "#0.00") _
                                         & vbTab & "$" & Format(dAvgDD * dDollarsPerPoint, "#0.00") & vbTab & "$" & Format(dWorstDD * dDollarsPerPoint, "#0.00")
                                    strText = Format(dAvgNet, "#0.00") & vbTab & Format(dAvgNet / iDaysHeld, "#0.00") _
                                         & vbTab & Format(dAvgDD, "#0.00") & vbTab & Format(dWorstDD, "#0.00")
                                End If
                                
                                ' put all the trades from this string array into a semi-colon delimited string
                                hTrades = tblTradeStrings(iTDOY, iDaysHeld, iStopLoss)
                                strTrade = gdGetStr(hTrades)
                                ' and expand the bar-number-placeholders with the real strings
                                i = InStr(strTrade, "[")
                                Do While i > 0
                                    j = InStr(i, strTrade, "]")
                                    If j > 0 Then
                                        iBar = Val(Mid(strTrade, i + 1, j - i - 1))
                                        strTrade = Left(strTrade, i - 1) & aTradeInfo(iBar) & Mid(strTrade, j + 1)
                                    End If
                                    i = InStr(strTrade, "[")
                                Loop
                                                               
                                If bDoSpreads Then
                                    i = 0
                                Else
                                    i = nSymbolID
                                End If
                                strText = strName & vbTab & Str(i) & vbTab & strDirection & vbTab _
                                    & Str(iTDOY) & vbTab & Str(BarsData(kDate, iEntryBar)) & vbTab _
                                    & Str(iDaysHeld) & vbTab & Str(BarsData(kDate, iExitBar)) & vbTab _
                                    & Str(dStopLoss) & vbTab & Str(iNumTrades) & vbTab & Format(dWinPct, "#0.0") & vbTab & Format(dPF, "#0.000") & vbTab _
                                    & strText & vbTab & Format(dAAP2DD, "#0.0000") & vbTab & Format(dAvgNet / dAvgWin * 100, "#0.0") & vbTab & strTrade
                                Print #fh, strText
                                iLineCount = iLineCount + 1
                                dFileSize = dFileSize + Len(strText) + 2
                                
                                ' and also get the symbol info (only one time for each symbol)
                                If Not bSymbolInfoDone Then
                                    bSymbolInfoDone = True
                                    i = g.SymbolPool.PoolRecForSymbolID(nSymbolID)
                                    s = g.SymbolPool.Desc(i)
                                    strExch = Parse(s, "@", 2) ' exchange
                                    strDesc = Parse(s, "@", 1) ' description
                                    If strSecType = "I" Then
                                        strExch = ""
                                    ElseIf strSecType = "F" Then
                                        ' strip off the "Cont Liq" part of the description (Liq CAdj Cont Elec Exp)
                                        strDesc = CleanSymbolDesc(strDesc)
                                    End If
                                    If Not bDoSpreads Then
                                        strText = strSecType & vbTab & strName & vbTab & Str(nSymbolID) & vbTab & strDesc & vbTab & strExch
                                    Else
                                        strText = strSecType & vbTab & strName & vbTab & strSymbol & vbTab & Str(nSymbolID) & vbTab & strDesc & vbTab & strExch
                                        i = g.SymbolPool.PoolRecForSymbolID(nSymbolID2)
                                        s = g.SymbolPool.Desc(i)
                                        strExch = Parse(s, "@", 2) ' exchange
                                        strDesc = Parse(s, "@", 1) ' description
                                        If strSecType = "I" Then
                                            strExch = ""
                                        ElseIf strSecType = "F" Then
                                            ' strip off the "Cont Liq" part of the description (Liq CAdj Cont Elec Exp)
                                            strDesc = CleanSymbolDesc(strDesc)
                                        End If
                                        strText = strText & vbTab & strSymbol2 & vbTab & Str(nSymbolID2) & vbTab & strDesc & vbTab & strExch
                                    End If
                                    aSymbolInfo.Add strText
                                End If
                            End If
                        End If
                    Next
                Next
            Next
        
            ' setup for Shorts (except for Spreads)
            If bShort Or bDoSpreads Then
                Exit Do
            Else
                bShort = True
            End If
        Loop ' do for both Longs and Shorts (except for Spreads)
    Next
    
    aSymbolInfo.Sort
    aSymbolInfo.ToFile strSymbolInfoFile

    Close #fh
    
    ' destroy all the string arrays
    For iTDOY = 1 To kMaxTDOY
        For iDaysHeld = 1 To kMaxDaysHeld
            For iStopLoss = 1 To kMaxStopLosses
                hTrades = tblTradeStrings(iTDOY, iDaysHeld, iStopLoss)
                If hTrades <> 0 Then
                    gdDestroyString hTrades
                End If
                tblTradeStrings(iTDOY, iDaysHeld, iStopLoss) = 0
            Next
        Next
    Next
    ReDim tblNumUp(0) As Long
    ReDim tblNumDown(0) As Long
    ReDim tblAmtUp(0) As Double
    ReDim tblAmtDown(0) As Double
    ReDim tblDD(0) As Double
    ReDim tblWorstDD(0) As Double
    ReDim tblTradeStrings(0) As Long
    
    AddList "Ram = " & Str(Int(PhysicalRAM(True))) & " of " & Str(Int(PhysicalRAM(False))) & " mb"
    AddList "Prebuilt = " & Format(dPrebuilt, "#,##0") & ", Rebuilt = " & Format(dRebuilt, "#,##0")
    AddList "Finished, #Lines = " & Format(iLineCount, "#,##0") & ", FileSize = " & Str(Int(dFileSize / 1000000#)) & " mb"
    bInProgress = False

End Sub
#End If

Private Sub ChkBigJumpsInBonds()

    Dim i&, yr&, d#, s$
    Dim Bars As New cGdBars
    Dim aText As New cGdArray
    
    i = 0
    For yr = 2001 To 2011
        DM_GetBars Bars, "ZB-067", "1 min", DateSerial(yr, 1, 1), DateSerial(yr, 12, 31)
        DoEvents
        i = Bars.Size
        For i = 1 To Bars.Size - 1
            If Bars.SessionDate(i) <= Bars.SessionDate(i - 1) + 1 Then
                d = Bars(eBARS_High, i) - Bars(eBARS_Close, i - 1)
                d = Round(Abs(d) * 32, 1)
                If Abs(d) >= 8 Then
                    s = DateFormat(Bars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM) & vbTab & Str(d) & vbTab _
                        & Bars.PriceDisplay(Bars(eBARS_Close, i - 1)) & vbTab & Bars.PriceDisplay(Bars(eBARS_High, i)) _
                        & vbTab & String(Abs(d), "#")
                    AddList s
                    aText.Add s
                End If
                
                d = Bars(eBARS_Low, i) - Bars(eBARS_Close, i - 1)
                d = Round(Abs(d) * 32, 1)
                If Abs(d) >= 8 Then
                    s = DateFormat(Bars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM) & vbTab & Str(d) & vbTab _
                        & Bars.PriceDisplay(Bars(eBARS_Close, i - 1)) & vbTab & Bars.PriceDisplay(Bars(eBARS_Low, i)) _
                        & vbTab & String(Abs(d), "#")
                    AddList s
                    aText.Add s
                End If
            End If
        Next
    Next
    aText.ToFile "c:\zb.txt"
    AddList "done"

End Sub


'Intraday Market Scope:
'- average each minute for each day of week
'- for last 10 yrs and last 5 yrs
'- calc average fixed pts
Private Sub IntradayMarketScope()

    Dim i&, iBar&, nStartDate&, nEndDate&, nDate&, nMinuteOfWeek&, nPrevMinuteOfWeek&
    Dim iDay&, iHour&, iMinute&, iYear&, iNumDaysInYear&, iList&
    Dim d#, dTime#, dPrice#, dPrevPrice#, dNull#, dRunTime#
    Dim s$, strSymbol$, strPit$, strComb$, strSecType$
    Dim Bars As New cGdBars, Bars2 As New cGdBars
    Dim aPriceDiffs As New cGdArray
    Dim aCounts As New cGdArray
    Dim aAvgPriceDiffs As New cGdArray
    Dim aFile As New cGdArray
    Dim aList As cGdArray
    Static bRunning As Boolean
    
    If Not FileExist("c:\common\files.exe") Then Exit Sub
    
    If bRunning Then
        bRunning = False
        Exit Sub
    End If
    bRunning = True
    dRunTime = gdTickCount
    
    nStartDate = DateSerial(2003, 1, 1)
    'nStartDate = DateSerial(2010, 1, 1)
    nEndDate = DateSerial(Year(LastDailyDownload), 1, 1) - 1
    
    Set aList = frmSymbolSelector.ShowMe(, , , "Create Weekly Intraday Averages")
    For iList = 0 To aList.Size - 1
        strSymbol = UCase(Trim(aList(iList)))
        Select Case SecurityType(strSymbol, True)
        Case "F"
            strSymbol = Parse(strSymbol, "-", 1) & "-067"
            ' convert to electronic (if exists)
            s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
            If Len(s) > 0 Then
                strSymbol = s
            End If
            ' make sure data is new/old enough and prices don't go negative
            DM_GetBars Bars, strSymbol, "Daily", nStartDate, nEndDate + 14
            If Bars(eBARS_DateTime, Bars.Size - 1) < nEndDate Then
                strSymbol = ""
            ElseIf Bars.Size < 260 * 5 Then ' at least 5 years worth of actual data
                strSymbol = ""
            End If
        Case "S", "I"
            If InStr(strSymbol, "@") > 0 Then
                strSymbol = "" ' ignore foreign
            Else
                ' for stocks/indices, make sure data is new/old enough and prices don't go negative
                DM_GetBars Bars, strSymbol, "Daily", nStartDate, nEndDate + 14
                If Bars(eBARS_DateTime, Bars.Size - 1) < nEndDate Then
                    strSymbol = ""
                ElseIf Bars.Size < 260 * 5 Then ' at least 5 years worth of actual data
                    strSymbol = ""
                ElseIf gdMinValue(Bars.ArrayHandle(eBARS_Low), 0, Bars.Size) <= 0 Then
                    strSymbol = ""
                End If
            End If
        Case Else
            strSymbol = ""
        End Select
        aList(iList) = strSymbol
    Next
    aList.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    
    For iList = 0 To aList.Size - 1
        If Not bRunning Then Exit For
        
        strSymbol = UCase(Trim(aList(iList)))
        strSecType = SecurityType(strSymbol)
        AddList strSymbol & "  " & Str(Year(nEndDate))
        
        aPriceDiffs.Create eGDARRAY_Doubles, 1440 * 7, 0
        aCounts.Create eGDARRAY_Longs, 1440 * 7, 0
        aAvgPriceDiffs.Create eGDARRAY_Doubles, 1440 * 7, 0
        
        strPit = ""
        strComb = ""
        If strSecType = "F" Then
            strSymbol = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
            strPit = ConvertFutureSymbol(strSymbol, ePitSymbol)
            strComb = ConvertFutureSymbol(strSymbol, eCombinedSymbol)
        End If
        If Len(strComb) = 0 Then
            strComb = strSymbol
        End If
        
        dNull = kNullData
        iNumDaysInYear = 0
        iYear = Year(nEndDate)
        For nDate = nEndDate To nStartDate - 1 Step -1
            If Not bRunning Then Exit For
            
            If Year(nDate) <> iYear Or nDate = nStartDate - 1 Then
                ' make sure some data existed for the year
                If iNumDaysInYear < 50 Then
                    Exit For ' out of data
                End If
                iNumDaysInYear = 0
                
                ' write the results so far to a file
                aFile.Size = 0
                aAvgPriceDiffs.Size = 0 ' to clear this array
                aAvgPriceDiffs.Size = aPriceDiffs.Size
                For nMinuteOfWeek = 0 To aCounts.Size - 1
                    If aCounts(nMinuteOfWeek) > 0 Then
                        dPrice = aPriceDiffs.Num(nMinuteOfWeek) / aCounts.Num(nMinuteOfWeek)
                        aAvgPriceDiffs(nMinuteOfWeek) = dPrice
                        iDay = Int(nMinuteOfWeek / 1440)
                        iMinute = nMinuteOfWeek Mod 1440
                        iHour = Int(iMinute / 60)
                        iMinute = iMinute Mod 60
                        s = Str(nMinuteOfWeek) & vbTab & Str(iDay) & vbTab & Format(iHour, "0#") & ":" & Format(iMinute, "0#") _
                            & vbTab & Format(dPrice / Bars.MinMove, "#0.00") & vbTab & Format(dPrice, "#0.00000000") _
                            & vbTab & Str(aPriceDiffs.Num(nMinuteOfWeek)) & vbTab & Str(aCounts.Num(nMinuteOfWeek))
                        'AddList s
                        aFile.Add s
                    End If
                Next
                If strSecType = "F" Then
                    s = Parse(strSymbol, "-", 1) & "-"
                Else
                    s = strSymbol & "!"
                End If
                s = Left(App.Path, 2) & "\WIA\" & s & Str(Year(Date) - iYear) & "yr.txt"
                aFile.ToFile s
                aAvgPriceDiffs.Serialize ReplaceFileExt(s, ".gda"), True
                
                iYear = Year(nDate)
                If nDate < nStartDate Then
                    Exit For
                End If
                AddList strSymbol & "  " & Str(Year(nDate))
            End If
        
            Bars.Size = 0
            If IsWeekday(nDate) Then
                ' try loading combined symbol
                DM_GetBars Bars, strComb, "1 min", nDate, nDate
                If Bars.Size = 0 And strSecType = "F" Then
                    Set Bars2 = New cGdBars
                    ' due to data issues back in early 2000's, try the pit and elect if no combined
                    If Len(strPit) > 0 Then
                        DM_GetBars Bars, strPit, "1 min", nDate, nDate
                    End If
                    If Len(strSymbol) > 0 Then
                        DM_GetBars Bars2, strSymbol, "1 min", nDate, nDate
                    End If
                    If Bars2.Size > Bars.Size Then
                        Set Bars = Bars2.MakeCopy
                    End If
                    If Bars.Size > 0 Then
                        i = JulToLong(nDate, True)
                    End If
                End If
                If Bars.Size > 0 Then
                    iNumDaysInYear = iNumDaysInYear + 1
                End If
            
                dPrevPrice = Bars(eBARS_Open, 0)
                nPrevMinuteOfWeek = 0
                For iBar = 0 To Bars.Size - 1
                    dTime = Bars(eBARS_DateTime, iBar)
                    dPrice = Bars(eBARS_Close, iBar)
                    If dPrevPrice = dNull Or dPrice = dNull Or dTime = dNull Then
                        Exit For
                    End If
                    ' calc # of minutes into the week (since midnight Sat)
                    nMinuteOfWeek = Round((dTime - Int(dTime)) * 1440) + (Weekday(Int(dTime)) - 1) * 1440
                    If nMinuteOfWeek <= nPrevMinuteOfWeek Or nMinuteOfWeek >= aCounts.Size Then
                        Exit For ' something's just WRONG!!!
                    End If
                    ' in case of a 15-minute or more "break" (i.e. lapse of data),
                    ' use this bar's open instead of previous bar's close as the starting price
                    If nMinuteOfWeek >= nPrevMinuteOfWeek + 15 Then
                        dPrevPrice = Bars(eBARS_Open, iBar)
                        If iBar > 0 Then
                            AddList DateFormat(dTime, M_D_YY, H_MM) & vbTab & Str(dPrice - dPrevPrice)
                        End If
                    End If
                    If strSecType = "F" Then
                        d = dPrice - dPrevPrice
                    ElseIf dPrevPrice > 0 And dPrice > 0 Then
                        ' calc ratio difference for stocks/indices
                        d = dPrice / dPrevPrice - 1
                    Else
                        d = 0
                    End If
                    aPriceDiffs.Num(nMinuteOfWeek) = aPriceDiffs.Num(nMinuteOfWeek) + d
                    aCounts.Num(nMinuteOfWeek) = aCounts.Num(nMinuteOfWeek) + 1
                    dPrevPrice = dPrice
                    nPrevMinuteOfWeek = nMinuteOfWeek
                Next
                'AddList Format(nDate, "mm/dd/yyyy")
            End If
            DoEvents
        Next
    Next
    
    dRunTime = Round(gdTickCount - dRunTime)
    AddList "Done -- " & Str(dRunTime / 1000) & " seconds"
    bRunning = False

End Sub

Private Sub MakeMidCmd()

    Dim i&, strFile$
    Dim aList As New cGdArray
    
    strFile = App.Path & "\Info\MidCmd.txt"
    strFile = "C:\LTS\MidCmd.txt"
    aList.FromFile strFile
    For i = 0 To aList.Size - 1
        aList(i) = EncryptToHex(aList(i))
    Next
    aList.ToFile ReplaceFileExt(strFile, ".cfg")

End Sub

Private Sub CheckForStaleFundamentals()
    
    Dim bStale As Boolean
    Dim iStock&, nSymbolID&, iDataKind&, nDataKindID&, dValue#, lActiveDate&, lBarDate&, nMaxFillDays&
    Dim strSymbol$, strDataKind$, s$
    Dim aSymbols As cGdArray
    Dim aOutput As New cGdArray
            
    Dim alDataIDs As New cGdArray
    Dim adValues As New cGdArray
    Dim alDates As New cGdArray
    
    Dim iLifetime As Integer, iInactive As Integer
    Dim alLifetime As New cGdArray
    Dim aDataKindDesc As New cGdArray
    
    ' store the "Lifetime" and Desc for each data kind
    alLifetime.Create eGDARRAY_Longs, 1000, 0
    For nDataKindID = 0 To alLifetime.Size - 1
        strDataKind = DM_GetDataKindNameForID(nDataKindID)
        If DM_GetDataKindInactive(g.DMS, nDataKindID, iInactive) <> 0 Then
            Select Case strDataKind
            ' Calculated
            Case "PERatio"
                iInactive = 1
            ' Obsolete (pre-2008 for all symbols)
            Case "PRevRatio", "PBookRatio", "QuickRatio", "TotDebtEquity", "PriceCashFlow", "CashPerShare", _
                    "PEGRatio", "PERatioEst", "YearEnd", "EPSChgQ1", "EPSChg12"
                iInactive = 1
            End Select
            If iInactive = 0 Then
                If DM_GetDataKindLifetime(g.DMS, nDataKindID, iLifetime) <> 0 Then
                    alLifetime(nDataKindID) = iLifetime
                End If
                aDataKindDesc(nDataKindID) = strDataKind
            End If
        End If
    Next
            
    ' get list of stock symbols
    Set aSymbols = frmSymbolSelector.ShowMe(, True, , "Check Fundamentals for which STOCKS?")
    
    ' for each symbol
    For iStock = 0 To aSymbols.Size - 1
        strSymbol = aSymbols(iStock)
        If SecurityType(strSymbol) = "S" Then
            nSymbolID = GetSymbolID(strSymbol)
            ' get all the Snapshot data for this symbol
            If DM_GetAllSnapData(nSymbolID, alDataIDs, adValues, alDates) Then
                ' for each DataKind available
                For iDataKind = 0 To alDataIDs.Size - 1
                    ' get the value and active date
                    nDataKindID = alDataIDs(iDataKind)
                    dValue = adValues(iDataKind)
                    lActiveDate = alDates(iDataKind)
                    nMaxFillDays = alLifetime(nDataKindID)
                    strDataKind = aDataKindDesc(nDataKindID)
                    If lActiveDate > 0 And nMaxFillDays <> 0 And Len(strDataKind) > 0 Then
                        'see if date is "in range"
                        bStale = True
                        lBarDate = LastDailyDownload
                        If lBarDate = lActiveDate Then
                            'exact date match
                            bStale = False
                        ElseIf nMaxFillDays > 0 And lBarDate > lActiveDate Then
                            'post-fill data (up to "n" days after value)
                            If lBarDate <= lActiveDate + nMaxFillDays Then
                                bStale = False
                            End If
                        ElseIf nMaxFillDays < 0 And lBarDate < lActiveDate Then
                            'pre-fill data (up to "n" days prior to value)
                            If lBarDate >= lActiveDate + nMaxFillDays Then
                                bStale = False
                            End If
                        End If
                        
                        If bStale Then
                            s = Format(nDataKindID, "000") & vbTab & strDataKind & vbTab & Str(nMaxFillDays) & vbTab _
                                & strSymbol & vbTab & DateFormat(lActiveDate, MM_DD_YYYY) & vbTab & Str(dValue)
                            frmTest.AddList s
                            aOutput.Add s
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    aOutput.Sort
    aOutput.ToFile "C:\Stale.txt"
    
End Sub

Private Sub DumpDebugLog(ByVal strType$, ByVal strMessage$, Optional ByVal iFH& = 0)
On Error Resume Next

    Dim fh As Integer                   ' File handle to open file with

    If iFH = 0 Then
        fh = FreeFile
        Open AddSlash(App.Path) & Format(Now, "YYYYMMDD") & "-" & strType & ".LOG" For Append As #fh
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    Else
        strMessage = Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        'gdFileStringIO iFH, strMessage, Len(strMessage), True
        fh = gdFileWriteLine(iFH, strMessage)
        gdFileFlush iFH
    End If

End Sub

Private Sub MinutizeHistory()

    Dim i&, Y&, s$, d#
    Dim a As New cGdArray
    Dim b As New cGdBars
    
    Set a = frmSymbolSelector.ShowMe()
    If a.Size > 0 Then
        For i = 0 To a.Size - 1
            s = UCase(Trim(a(i)))
            For Y = 2000 To 2013
                DM_GetBars b, s, "1440m", Y * 10000 + 101, Y * 10000 + 1231, False
                AddList s & ": " & Str(Y) & " = " & Str(b.Size)
                DoEvents
            Next
        Next
    End If
    AddList "done"

End Sub

'//      0:  dividend kind, one of the following data kinds:
'//          Cash:   DIV_DIST    Total dividend distribution (typically a sum of the other types for a particular date)
'//                  DIV_CASH    Regular cash distribution (for ex: typical quarterly dividend)
'//                  DIV_CASHEQ  Cash equivalent distribution
'//                  DIV_SPEC    Special cash distribution
'//          Stock:  DIV_STOCK   Stock Dividend (for ex: 0.20 = 20%, 0.05 = 5%, 0.0395837 = weird dividend)
'//                                  (these are like splits, divide price by 1.2, 1.05, or 1.0395837)
'//                  DIV_CONSOL  Stock Consolidation (these actually reduce the amount of stock and increase
'//                                  the price.  For instance, 0.96 would divide price by 0.96,
'//                                  effectively increasing the price by 4%.  Kind of a weird reverse split)
'//      1:  value -- depending on dividend kind, a value for the dividend
'//              cash dividend: amount of cash dividend
'//              stock dividend: amount of stock dividend (0.20 = 20%, 0.05 = 5%, etc)
'//              stock consolidation: ratio of consolidation (between 0.0 and 1.0)
'//      2:  execution date -- date the given dividend is attributed to a stock (and the stock's owner)
'//              this is generally when it affects the stock price
'//      3:  pay dates -- date the dividend is paid (to the owner of record as of the execution date)
Private Sub ShowDivTable()
    
    Dim i&, s$, dSum#
    Dim a As New cGdArray
    Dim t As cGdTable
    
    Set a = frmSymbolSelector.ShowMe("IBM", False)
    s = a(0)
    If Len(s) > 0 Then
        Set t = GetDividendsTable(s, True)
        AddList "Dividend table for " & s & " = " & t.NumRecords & " records"
        For i = 0 To t.NumRecords - 1
            s = Str(i) & ": " & vbTab & Str(t(0, i)) & vbTab & Format(t(1, i), "#0.00000") & vbTab _
                & DateFormat(t(2, i)) & vbTab & DateFormat(t(3, i))
            AddList s
            If t(0, i) = 524 Then
                dSum = dSum + t(1, i)
            End If
        Next
    End If
    AddList "Total cash dividends = " & Format(dSum, "$#0.00")

End Sub

Private Function CheckContSymbol(ByVal strType$, ByVal strSymbol$, aOut As cGdArray) As Long

    Dim i&, nSymbolID&, nDate&, strNext$, strPrev$, strOut$
    Dim Bars As New cGdBars
    Dim Rolls As cGdTable
    
    If Len(strSymbol) > 0 And Left(strSymbol, 1) <> "-" Then
        DM_GetBars Bars, strSymbol
        CheckContSymbol = Bars(eBARS_DateTime, 0)
        strOut = strType & vbTab & Bars.Prop(eBARS_Symbol) & vbTab & Bars.Prop(eBARS_Exchange) & vbTab & DateFormat(Bars(eBARS_DateTime, 0)) _
                & vbTab & DateFormat(Bars(eBARS_DateTime, Bars.Size - 1))
    
        strPrev = Parse(strSymbol, "-", 1)
        Set Rolls = GetRollsTable(strSymbol)
        For i = 0 To Rolls.NumRecords - 1
            nSymbolID = Rolls(0, i)
            strNext = Parse(GetSymbol(nSymbolID), "-", 1)
            If strNext <> strPrev Then
                nDate = Rolls(1, i)
                strOut = strOut & vbTab & GetSymbol(nSymbolID) & " " & DateFormat(nDate)
                strPrev = strNext
            End If
        Next
        
        aOut.Add strOut
    End If

End Function

' get average of highest N daily tick volumes
Private Function CalcAvgTickVol(ByVal strSymbol$, ByVal nNumAvg&) As Double
    
    Dim i&, dAvg#, nStart&, nDate&
    Dim aVol As New cGdArray
    Dim Bars As New cGdBars
    
    nStart = Val(Parse(strSymbol, "-", 2))
    If nStart > 190101 Then
        nDate = JulFromLong(nStart * 100 + 28)
        DM_GetBars Bars, strSymbol, "1440m", nDate - 250, nDate
    End If
    'If Bars.Size > 0 Then
    If Bars.Size >= nNumAvg Then
        If nNumAvg > Bars.Size Then
            nNumAvg = Bars.Size
        End If
        aVol.Create eGDARRAY_Doubles, Bars.Size, 0
        For i = 0 To Bars.Size - 1
            aVol.Num(i) = Bars(eBARS_UpTicks, i) + Bars(eBARS_DownTicks, i)
        Next
        aVol.Sort eGdSort_Descending
        For i = 0 To nNumAvg - 1
            dAvg = dAvg + aVol.Num(i)
        Next
        dAvg = dAvg / nNumAvg
        Set aVol = Nothing
    End If
    Set Bars = Nothing
    
    CalcAvgTickVol = dAvg

End Function

Private Sub CheckContHistory()

    Dim i&, iLine&, s$, strPit$, strElec$, strComb$, strSynth$, strOut$, nDate&, nEarliest&, nType&
    Dim iYear&, iMonth&, strSymbol$, dAvg#, dAvgPit#
    Dim aFile As New cGdArray
    Dim aOut As New cGdArray
    Dim aSymbols As New cGdArray
    Dim aVol As New cGdArray
    Dim Bars As New cGdBars
    
    aFile.FromFile App.Path & "\Info\SymbolMap.csv"
    For iLine = 0 To aFile.Size - 1
        s = StripStr(aFile(iLine), " ")
        If Len(s) > 0 Then
            aSymbols.SplitFields s, ","
            strPit = aSymbols(0)
            strElec = aSymbols(1)
            'strSynth = aSymbols(2)
            strComb = aSymbols(3)
            
            If Len(strElec) > 0 Then
                SetBarProperties Bars, strElec & "-067"
                strElec = Bars.Prop(eBARS_Exchange) & "," & strElec
            Else
                strElec = "ZZZZ"
            End If
            
            aFile(iLine) = strElec & vbTab & s
        Else
            aFile(iLine) = ""
        End If
    Next
    aFile.Sort eGdSort_DeleteNullValues Or eGdSort_IgnoreCase
    
    For iLine = 0 To aFile.Size - 1
        s = Parse(aFile(iLine), vbTab, 2)
        If Len(s) > 0 Then
            'AddList s
            aSymbols.SplitFields s, ","
            strPit = aSymbols(0)
            strElec = aSymbols(1)
            strSynth = aSymbols(2)
            strComb = aSymbols(3)
            
            s = Left(strPit, 1) + Left(strElec, 1) '+ Left(strSynth, 1) + Left(strComb, 1)
            If Len(s) > 1 Then
                AddList strPit & vbTab & strComb & vbTab & strElec & vbTab & strSynth
                's = ""
            End If
            
            If Len(s) > 1 Then
                DoEvents
                If 1 Then
                    strOut = ""
                    
                    ' compare intraday tick volume for Electronic vs. Pit
                    For iYear = 2000 To 2012
                        For iMonth = 1 To 12
                            dAvg = 0
                            strSymbol = strElec & "-" & Str(iYear * 100 + iMonth)
                            If GetSymbolID(strSymbol) > 0 Then
                                dAvg = CalcAvgTickVol(strSymbol, 20)
                            End If
                            If dAvg > 0 Then
                                AddList strSymbol
                                DoEvents
                                ' compare with Pit
                                strSymbol = strPit & "-" & Str(iYear * 100 + iMonth)
                                dAvgPit = CalcAvgTickVol(strSymbol, 20)
                                If dAvg > dAvgPit Then
                                    strOut = strPit & vbTab & strComb & vbTab & strElec & vbTab & strSynth _
                                        & vbTab & Str(iYear) & "-" & Format(iMonth, "00")
                                    'aOut.Add strOut
                                    'AddList strOut
                                    iYear = 9999
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                    
                    If Len(strOut) > 0 Then
                        ' see where combined intraday data starts
                        For iYear = 2000 To 2012
                            For iMonth = 1 To 12
                                dAvg = 0
                                strSymbol = strComb & "-" & Str(iYear * 100 + iMonth)
                                If GetSymbolID(strSymbol) > 0 Then
                                    dAvg = CalcAvgTickVol(strSymbol, 20)
                                End If
                                If dAvg > 0 Or Len(strComb) = 0 Then
                                    strOut = strOut & vbTab & Str(iYear) & "-" & Format(iMonth, "00")
                                    iYear = 9999
                                    DoEvents
                                    Exit For
                                End If
                            Next
                        Next
                        
                        aOut.Add strOut
                        AddList strOut
                    End If
                Else
                
                    For nType = 55 To 57
                        
                        nEarliest = CheckContSymbol("P", strPit & "-0" & Str(nType), aOut)
                        
                        nDate = CheckContSymbol("E", strElec & "-0" & Str(nType), aOut)
                        If nEarliest <= 0 Then
                            nEarliest = nDate
                        ElseIf nDate > 0 And nDate <> nEarliest Then
                            nEarliest = 999999
                        End If
                        
                        nDate = CheckContSymbol("C", strComb & "-0" & Str(nType), aOut)
                        If nEarliest <= 0 Then
                            nEarliest = nDate
                        ElseIf nDate > 0 And nDate <> nEarliest Then
                            nEarliest = 999999
                        End If
                        
                        nDate = CheckContSymbol("S", strSynth & "-0" & Str(nType), aOut)
                        If nEarliest <= 0 Then
                            nEarliest = nDate
                        ElseIf nDate > 0 And nDate <> nEarliest Then
                            nEarliest = 999999
                        End If
                        
                        If nEarliest = 999999 Then
                            strOut = String(20, "#")
                        Else
                            strOut = String(20, "-")
                        End If
                        aOut.Add strOut
                    Next
                End If
            End If
        End If
    Next
    
    aOut.ToFile "C:\Dump.txt"
    AddList "done"

End Sub

Private Sub CheckForexPips()

    Dim iRec&, nSymbolID&, dTickMove#, d#
    Dim strSymbol$
    Dim Bars As New cGdBars

    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            If IsForex(strSymbol) Then
                If Right(strSymbol, 3) = "@IB" Or Right(strSymbol, 4) = "@CNX" Then
                    SetBarProperties Bars, Parse(strSymbol, "@", 1), True
                    dTickMove = Bars.TickMove
                    d = Bars.Prop(eBARS_MinMoveInTicks)
                    SetBarProperties Bars, strSymbol, True
                    If Bars.TickMove <> dTickMove Then
                        AddList strSymbol & vbTab & Str(Bars.TickMove) & vbTab & Str(Bars.Prop(eBARS_MinMoveInTicks)) _
                            & vbTab & "-> " & Str(dTickMove) & vbTab & Str(d)
                    End If
                End If
            End If
        End If
    Next
    AddList "done"

End Sub

Private Function LoadYahooStockHist(ByVal strSymbol$) As cGdBars

    Dim s$, i&, iBar&, dDate#, dPrevDate#
    Dim a As New cGdArray, F As New cGdArray
    Dim b As New cGdBars
            
    If InStr(strSymbol, "@") = 0 And SecurityType(strSymbol) = "S" Then
        s = Replace(strSymbol, "_", "-")
        s = "c:\StockData\" & s & ".txt"
        If FileExist(s) Then
            'AA 19620102 65.37 65.75 65.37 65.37 134400
            SetBarProperties b, strSymbol
            a.FromFile s
            b.Size = a.Size
            F.Create eGDARRAY_Doubles, 7
            For i = 0 To a.Size - 1
                s = a(i)
                F.SplitFields s, " "
                dDate = JulFromLong(F.Num(1))
                If dDate > dPrevDate Then
                    ' on 2nd bar, check for bogus 1st bar
                    If iBar = 1 Then
                        If b(eBARS_Close, 0) > 1000000 Or dDate > dPrevDate + 30 Then
                            iBar = 0
                        End If
                    End If
                    dPrevDate = dDate
                    b(eBARS_DateTime, iBar) = dDate
                    b(eBARS_Open, iBar) = F.Num(2)
                    b(eBARS_High, iBar) = F.Num(3)
                    b(eBARS_Low, iBar) = F.Num(4)
                    b(eBARS_Close, iBar) = F.Num(5)
                    b(eBARS_Vol, iBar) = F.Num(6)
                    If b(eBARS_Close, 0) > 1000000 And iBar > 0 Then
                        iBar = iBar
                    End If
                    iBar = iBar + 1
                Else
                    iBar = iBar
                End If
            Next
            b.Size = iBar
            Set F = Nothing
            Set a = Nothing
        End If
    End If
    
    Set LoadYahooStockHist = b

End Function

'Check old stock history:
'- compare beginning of yahoo data with ours
'- look for their date that matches our first date
'- if they have older data, we should prepend it
'- if their close is significantly different (+/- 5%), need to unsplit ours?
Private Sub CheckOldStockHist()

    Dim i&, iRec&, nSymbolID&, dTickTime#, nDate&, iBar1&, iBar2&, d1#, d2#, dDiff#, nDate1&, nDate2&
    Dim strSymbol$, s$
    Dim b1 As New cGdBars, b2 As New cGdBars
    Dim aOut As New cGdArray, aFixData As New cGdArray

    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            Set b2 = LoadYahooStockHist(strSymbol)
            If b2.Size > 0 Then
                ' load our unsplit data to compare
                b1.Size = 0
                DM_GetBars b1, nSymbolID, , , , , , True
                If b1.Size > 0 Then
                    ' check for price gap of first matching date
                    d1 = 0
                    d2 = 0
                    nDate = DateSerial(2002, 1, 1)
                    iBar2 = b2.FindDateTime(nDate)
                    iBar1 = b1.FindDateTime(nDate)
                    For iBar1 = iBar1 To b1.Size - 1
                        nDate = b1(eBARS_DateTime, iBar1)
                        Do While b2(eBARS_DateTime, iBar2) < nDate
                            iBar2 = iBar2 + 1
                            If iBar2 >= b2.Size Then
                                Exit For
                            End If
                        Loop
                        d2 = b2(eBARS_Close, iBar2)
                        If d2 = kNullData Then
                            Exit For
                        End If
                        d1 = b1(eBARS_Close, iBar1)
                        If d1 > 0 And d2 > 0 Then
                            dDiff = Abs(d1 - d2) / d2
                            If dDiff < 0.02 Then ' within 2%
                                Exit For ' this is the first date where the data basically matches
                            End If
                        End If
                    Next
                    If iBar2 > 0 Then
                        ' before this bar is the Yahoo data we should be using?
                        nDate1 = JulToLong(b1(eBARS_DateTime, 0), True)
                        nDate2 = JulToLong(b2(eBARS_DateTime, 0), True)
                        nDate = JulToLong(b2(eBARS_DateTime, iBar2), True)
                        i = Year(b1(eBARS_DateTime, b1.Size - 1))
                        s = strSymbol & vbTab & Str(i) & vbTab & Str(nDate2) & vbTab & Str(nDate1) & vbTab & Str(nDate)
                        'If iBar1 > 0 And iBar2 > 0 Then
                            d1 = b1(eBARS_Close, iBar1)
                            d2 = b2(eBARS_Close, iBar2)
                            If d1 > 0 And d2 > 0 Then
                                dDiff = Abs(d1 - d2) / d2 * 100
                            Else
                                dDiff = -1
                            End If
                            s = s & vbTab & b1.PriceDisplay(d1) & vbTab & b1.PriceDisplay(d2) & vbTab & Format(dDiff, "#0.0")
                        'End If
                        ' check for price gap if our data starts before Yahoo's
                        If nDate1 < nDate2 Then
                            i = b1.FindDateTime(nDate2)
                            d1 = b1(eBARS_Close, i)
                            d2 = b2(eBARS_Close, 0)
                            If d1 > 0 And d2 > 0 Then
                                dDiff = Abs(d1 - d2) / d2 * 100
                            Else
                                dDiff = -1
                            End If
                            If dDiff <> 0 Then
                                s = s & vbTab & b1.PriceDisplay(d1) & vbTab & b1.PriceDisplay(d2) & vbTab & Format(dDiff, "#0.0")
                            End If
                        End If
                        AddList s
                        aOut.Add s
                        
                        ' append to the fix file
                        ' !ABC/9912
                        ' 130923 51.910000 52.240000 51.380000 51.670000 2716900 0 2716900 0 0
                        aFixData.Add "!" & strSymbol & "/9912"
                        For i = 0 To iBar2 - 1
                            nDate = JulToLong(b2(eBARS_DateTime, i), True)
                            If nDate > 0 Then
                                s = Right(Str(nDate), 6) & " " & Str(b2(eBARS_Open, i)) & " " & Str(b2(eBARS_High, i)) _
                                    & " " & Str(b2(eBARS_Low, i)) & " " & Str(b2(eBARS_Close, i)) & " " & Str(b2(eBARS_Vol, i)) _
                                    & " 0 " & Str(b2(eBARS_Vol, i)) & " 0 0"
                                aFixData.Add s
                            End If
                        Next
                    End If
                End If
                
                If gdTickCount > dTickTime + 500 Then
                    dTickTime = gdTickCount
                    DoEvents
                End If
            End If
        End If
    Next
    aOut.ToFile "c:\dump.txt"
    aFixData.ToFile "c:\fix.txt"
    AddList "done"

End Sub

' new method: just use 1/1/2002 as the splice point
Private Sub CheckOldStockHist2()

    Dim i&, iRec&, nSymbolID&, dTickTime#, nDate&, iBar1&, iBar2&, d1#, d2#
    Dim nSpliceDate&, dGap#
    Dim strSymbol$, s$
    Dim b1 As New cGdBars, b2 As New cGdBars
    Dim aOut As New cGdArray, aFixData As New cGdArray

    nSpliceDate = DateSerial(2002, 1, 1)
    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            Set b2 = LoadYahooStockHist(strSymbol)
            iBar2 = b2.FindDateTime(nSpliceDate, False)
            If iBar2 > 0 Then
                ' load our unsplit data to compare
                b1.Size = 0
                DM_GetBars b1, nSymbolID, , , , , , True
                iBar1 = b1.FindDateTime(nSpliceDate, False)
                If iBar1 > 0 Then
                    ' check price gap at splice point
                    dGap = 0
                    d1 = b1(eBARS_Close, iBar1)
                    d2 = b2(eBARS_Close, iBar2)
                    If d1 > 0 And d2 > 0 Then
                        dGap = Abs(d1 - d2) / d2 * 100
                    Else
                        dGap = 0
                    End If
                    ' before this bar is the Yahoo data we should be using?
                    s = strSymbol & vbTab & b1.PriceDisplay(d1) & vbTab & b1.PriceDisplay(d2) & vbTab & Format(dGap, "#0.0")
                    If dGap > 5 Then
                        s = s & vbTab & "###"
                        AddList s
                    End If
                    aOut.Add s
                        
                    ' append to the fix file
                    ' !ABC/9912
                    ' 130923 51.910000 52.240000 51.380000 51.670000 2716900 0 2716900 0 0
                    aFixData.Add "!" & strSymbol & "/9912"
                    For i = 0 To iBar2 - 1
                        nDate = JulToLong(b2(eBARS_DateTime, i), True)
                        If nDate > 0 Then
                            s = Right(Str(nDate), 6) & " " & Str(b2(eBARS_Open, i)) & " " & Str(b2(eBARS_High, i)) _
                                & " " & Str(b2(eBARS_Low, i)) & " " & Str(b2(eBARS_Close, i)) & " " & Str(b2(eBARS_Vol, i)) _
                                & " 0 " & Str(b2(eBARS_Vol, i)) & " 0 0"
                            aFixData.Add s
                        End If
                    Next
                End If
                
                If gdTickCount > dTickTime + 500 Then
                    dTickTime = gdTickCount
                    DoEvents
                End If
            End If
        End If
    Next
    aOut.ToFile "c:\dump.txt"
    aFixData.ToFile "c:\fix.txt"
    AddList "done"

End Sub


Private Sub CheckSectors()

    Dim nSymbolID&, nSectorID&, nSubsectorID&, dValue#, lDate&
    Dim strSymbol$, strSector$, strSubsector$, s$
    Dim aOut As New cGdArray
    
    For nSymbolID = 1 To 300000
        strSymbol = GetSymbol(nSymbolID)
        If Len(strSymbol) > 0 Then
            nSectorID = 0
            nSubsectorID = 0
            ' get symbol for sector
            If DM_GetSnap1(g.DMS, nSymbolID, 162, dValue, lDate) Then
                nSectorID = dValue
            End If
            If DM_GetSnap1(g.DMS, nSymbolID, 163, dValue, lDate) Then
                nSubsectorID = dValue
            End If
            If nSectorID > 0 Or nSubsectorID > 0 Then
                strSector = GetSymbol(nSectorID)
                strSubsector = GetSymbol(nSubsectorID)
                s = strSymbol & vbTab & Str(nSymbolID) & vbTab & strSector & vbTab & Str(nSectorID) & vbTab & strSubsector & vbTab & Str(nSubsectorID)
                If InStr(strSymbol, "@") > 0 Or Left(strSector, 3) <> "$--" Or Left(strSubsector, 2) <> "$-" Or SecurityType(strSymbol) <> "S" Then
                    s = s & vbTab & "###"
                    AddList s
                End If
                aOut.Add s
            End If
        End If
        If nSymbolID Mod 10000 = 0 Then
            AddList Str(nSymbolID)
            DoEvents
        End If
    Next
    aOut.ToFile "c:\Sectors.txt"
    AddList "done"

End Sub

Public Sub CheckEtfTP()

    Dim s$, i&, nDate&, strSymbol$, nSymbol&, dBuyVol#, dSellVol#, iRec&
    Dim b As New cGdBars
    Dim aSymbols As New cGdArray
    Dim aTable As New cGdTable
    
    'Dow30 from 6/8/2009 - 9/23/2012:
    'AA,AXP,BA,BAC,CAT,CSCO,CVX,DD,DIS,GE,HD,HPQ,IBM,INTC,JNJ,JPM,KO,MCD,MDLZ,MMM,MRK,MSFT,PFE,PG,T,TRV,UTX,VZ,WMT,XOM
    s = "AA,AXP,BA,BAC,CAT,CSCO,CVX,DD,DIS,GE,HD,HPQ,IBM,INTC,JNJ,JPM,KO,MCD,MDLZ,MMM,MRK,MSFT,PFE,PG,T,TRV,UTX,VZ,WMT,XOM"
    aSymbols.SplitFields s, ","
    aSymbols.Sort
    
    aTable.CreateField eGDARRAY_Longs, 0, "Date"
    For nSymbol = 0 To aSymbols.Size - 1
        strSymbol = aSymbols(nSymbol)
        aTable.CreateField eGDARRAY_Doubles, nSymbol + 1, strSymbol
        iRec = 0
        For nDate = DateSerial(2009, 6, 8) To DateSerial(2012, 9, 23)
        'For nDate = DateSerial(2013, 1, 1) To Date - 1
            If IsWeekday(nDate) Then
                DM_GetBars b, strSymbol, "1440", nDate, nDate
                If b.Size > 0 Then
                    dBuyVol = b(eBARS_AskVol, 0)
                    dSellVol = b(eBARS_BidVol, 0)
                    If iRec Mod 100 = 0 Then
                        s = strSymbol & vbTab & DateFormat(nDate) & vbTab & Str(dBuyVol - dSellVol)
                        AddList s
                        DoEvents
                    End If
                    
                    If nSymbol = 0 Then
                        aTable.Num(0, iRec) = nDate
                    ElseIf nDate <> aTable.Num(0, iRec) Then
                        s = "*** BAD date: " & strSymbol & " " & DateFormat(nDate)
                        AddList s
                        Exit Sub
                    End If
                    aTable.Num(nSymbol + 1, iRec) = dBuyVol - dSellVol
                    iRec = iRec + 1
                End If
            End If
        Next
        DoEvents
    Next
    
    s = aTable.ToString(, , True)
    FileFromString "c:\test.txt", s
    AddList "done"

End Sub

Private Sub ChkPitRolls()

    Dim i&, iRec&, nSymbolID&, s$, strSymbol$, s1$, s2$
    Dim aOut As New cGdArray

    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            If InStr(strSymbol, "-05") > 0 Then
                s = ConvertFutureSymbol(strSymbol, ePitSymbol)
                If s = strSymbol Then
                    s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                    If Len(s) > 0 And s <> strSymbol Then
                        s1 = RollSymbolForDate(strSymbol)
                        s2 = RollSymbolForDate(s)
                        If Parse(s1, "-", 2) <> Parse(s2, "-", 2) And InStr(s1, "-201") > 0 Then
                            AddList strSymbol & "      " & vbTab & s1 & "    " & vbTab & s2
                            aOut.Add strSymbol & vbTab & s1 & vbTab & s2
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    aOut.ToFile "c:\test.txt"
    AddList "done"

End Sub

Private Sub ChkFormIteration()

    Dim i&, s$
    Dim frm As Form, frms As New cForms
    Dim frmC As frmChart

    i = 0
    AddList Str(Forms.Count)
    Do
        Set frm = frms.NextForm
        If frm Is Nothing Then Exit Do
        
        s = frm.Caption
        AddList Str(i) & vbTab & Str(frm.hWnd) & vbTab & frm.Name & vbTab & s
        i = i + 1
        Sleep 0.25
    Loop
    Set frm = Nothing
    AddList "done " & Str(i)

    i = 0
    AddList Str(Forms.Count)
    frms.Init frmChart
    Do
        Set frmC = frms.NextForm
        If frmC Is Nothing Then Exit Do
        
        s = frmC.Caption
        AddList Str(i) & vbTab & Str(frmC.hWnd) & vbTab & frmC.Name & vbTab & s
        i = i + 1
        Sleep 0.25
    Loop
    Set frmC = Nothing
    AddList "done " & Str(i)

Exit Sub

    AddList Str(Forms.Count)
    For i = Forms.Count - 1 To 10 Step -1
        If i >= Forms.Count Then Exit For
        Set frm = Forms(i)
        s = frm.Caption
        AddList Str(i) & vbTab & Str(frm.hWnd) & vbTab & frm.Name & vbTab & s
        Sleep 0.25
        If s <> Forms(i).Caption Then
            s = ""
        End If
        s = frm.Caption
        'AddList Str(i) & vbTab & frm.Name & vbTab & s
    Next
    Set frm = Nothing
    AddList "done 1"
    
Exit Sub
    
    AddList Str(Forms.Count)
    i = 0
    For Each frm In Forms
        s = frm.Caption
        AddList Str(i) & vbTab & frm.Name & vbTab & s
        i = i + 1
        Sleep 0.25
    Next
    Set frm = Nothing
    AddList "done 2"

End Sub

Private Sub GetCont67Group()

    Dim i&, iRec&, nSymbolID&, dAvgDPD#, dMaxDPD#
    Dim s$, strSymbol$
    Dim Bars As New cGdBars
    Dim aRanges As New cGdArray, aTrades As New cGdArray, aVolumes As New cGdArray
    Dim aOut As New cGdArray, aOut2 As New cGdArray
    
    aRanges.Create eGDARRAY_Doubles, 0, 0
    aTrades.Create eGDARRAY_Doubles, 0, 0
    aVolumes.Create eGDARRAY_Doubles, 0, 0
    
    ' look for all -067 symbols
    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            If Right(strSymbol, 4) = "-067" Then
                s = ConvertFutureSymbol(strSymbol, eElectronicSymbol)
                If Len(s) = 0 Or s = strSymbol Then
                    ' see if we have recent data
                    DM_GetBars Bars, strSymbol, "daily", LastDailyDownload - 30, 0
                    If Bars.Size > 0 Then
                        aOut.Add strSymbol
                        
                        ' and get intraday data (to calc "Daily Medians")
                        DM_GetBars Bars, strSymbol, "1440min", 20150301, 20160301
                        aRanges.Size = Bars.Size
                        aTrades.Size = Bars.Size
                        aVolumes.Size = Bars.Size
                        For i = 0 To Bars.Size - 1
                            aRanges.Num(i) = (Bars(eBARS_High, i) - Bars(eBARS_Low, i))
                            aTrades.Num(i) = Bars(eBARS_UpTicks, i) + Bars(eBARS_DownTicks, i)
                            aVolumes.Num(i) = Bars(eBARS_Vol, i)
                        Next
                        ' get data for middle day
                        aRanges.Sort
                        aTrades.Sort
                        aVolumes.Sort
                        i = Bars.Size / 2 - 1
                        If i >= 0 Then
                            ' avg $ range per day
                            dAvgDPD = Round(aRanges.Num(i) / Bars.TickMove * Bars.TickValue)
                            dMaxDPD = Round(aRanges.Num(Bars.Size - 1) / Bars.TickMove * Bars.TickValue)
                            s = strSymbol & vbTab & Str(aTrades.Num(i)) & vbTab & Str(aVolumes.Num(i)) & vbTab _
                                & Str(dAvgDPD) & vbTab & Str(dMaxDPD) & vbTab & Str(Bars.Prop(eBARS_Margin)) & vbTab _
                                & Bars.Prop(eBARS_Desc) & vbTab & Bars.Prop(eBARS_Exchange)
                            aOut2.Add s
                            AddList s
                            DoEvents
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    aOut.Sort eGdSort_DeleteDuplicates Or eGdSort_DeleteNullValues
    aOut.ToFile "c:\Cont067.csv"
    aOut2.Sort eGdSort_DeleteNullValues
    aOut2.ToFile "c:\Cont067.txt"
    s = "Symbol" & vbTab & "#Trades" & vbTab & "Volume" & vbTab & "Avg $Daily" & vbTab & "Max $Daily" & vbTab & "Margin" & vbTab & "Description" & vbTab & "Exchange"
    aOut2.Add s, 0
    frmTest.AddList Str(aOut.Size) & " symbols written to C:\Cont067.csv"

End Sub

Private Sub CheckProfileBars()

    Dim i&, s$, n&
    Dim Bars As New cGdBars
    
    lst.Clear
    s = "YC2-067"
    n = DateSerial(2014, 6, 5)
    If 1 Then
        BuildProfileBars Bars, GetSymbolID(s), n, n
    Else
        'DM_GetBars Bars, s, "5", n, n
        DM_GetBars Bars, s, "100t", n, n
    End If
        
    For i = 0 To Bars.Size - 1
        AddList DateFormat(Bars(eBARS_DateTime, i), MM_DD_YYYY, HH_MM, NO_AMPM) & vbTab & Bars.PriceDisplay(Bars(eBARS_Close, i))
    Next
    AddList "done"

End Sub

Private Sub CheckFunctions()

    Dim i&, s$
    Dim F As cFunction
    
    For i = 1 To g.Functions.Count
        Set F = g.Functions.Item(i)
        s = UCase(F.CodedName)
        'If s = "LNLCLOSELONGTERM" Or s = "LNLACCUMDIST" Then
            s = UCase(StripStr(F.CodedText, " "))
            If InStr(s, "~15001)~10001/~01009MOVINGAVG~16001(") > 0 Or InStr(s, "~15001)~10001/~01006MOVAVG~16001(") > 0 Then
                AddList F.CodedName
            End If
        'End If
    Next
    AddList "done"

End Sub

Private Sub CheckCRC()

    Dim i&, iCRC&, j&
    Dim buf As New cMemBuffer
    
    buf.FromFile App.Path & "\NavSuite.exe"
    i = gdCalcMemCRC32(buf.MemPtr, buf.Length)
    iCRC = gdCalcCumulativeCRC32(buf.MemPtr, buf.Length, 0)
    
    iCRC = 0
    For i = 0 To buf.Length Step 1000000
        j = 1000000
        If j >= buf.Length - i Then
            j = buf.Length - i
        End If
        iCRC = gdCalcCumulativeCRC32(buf.MemPtr + i, j, iCRC)
    Next
    
    If iCRC <> gdCalcMemCRC32(buf.MemPtr, buf.Length) Then
        AddList "CRC did NOT match!!"
    Else
        AddList "CRC matched"
    End If

End Sub

#If 0 Then
Private Sub CheckFractZen()

    Dim i&, j&, nDate&
    Dim s$, strSymbol$
    
    
    strSymbol = "SB-067"
    'strSymbol = "$EUR-JPY"
    
    nDate = 0
    i = CalcAutoBreakoutRange(strSymbol)
    j = g.FractZen.GetFractZenRange(strSymbol)
    AddList strSymbol & " " & Str(nDate) & ": " & Str(i) & " " & Str(j)
    
    For nDate = DateSerial(2014, 1, 1) To Date + 1
    'For nDate = LastDailyDownload - 4 To Date + 1
        i = CalcAutoBreakoutRange(strSymbol, nDate - 1)
        j = g.FractZen.GetFractZenRange(strSymbol, nDate)
        AddList strSymbol & " " & Str(nDate) & ": " & Str(i) & " " & Str(j)
        If i <> j Then
            AddList "***** " & DateFormat(nDate) & " " & strSymbol
        End If
    Next
    j = g.FractZen.GetFractZenRange("DX-067")
    AddList "done"

End Sub
#End If


Private Sub TestTAS()

    Dim i&, s$, d#, n&
    Dim nFuncID&, nRangeBars&, nAvgBars&, nMABars&
    Dim nBar&, nYYYYMMDD&, nHHMM&, dOpen#, dHigh#, dLow#, dClose#, dVol#, dOI#
    Dim Bars As New cGdBars
    
    nRangeBars = 10
    nAvgBars = 3
    nMABars = 3
    
    s = "BAC"
    nFuncID = TAS_IndicatorInit("eSignal", "RATIO", s)
    AddList "Init = " & Str(nFuncID)
    If nFuncID > 0 Then
        i = TAS_IndicatorSetParameter(nFuncID, 0, nRangeBars)
        i = TAS_IndicatorSetParameter(nFuncID, 1, nAvgBars)
        i = TAS_IndicatorSetParameter(nFuncID, 2, nMABars)
        
        DM_GetBars Bars, s, "", 20140101, 0
        Bars.AddForecastBars 5
        
        For n = 0 To Bars.Size - 1
            ' this indicator uses values from the previous bar
            nBar = n - 1
            dClose = Bars(eBARS_Close, nBar)
            If dClose = kNullData Then
                d = kNullData
            Else
                d = Bars(eBARS_DateTime, nBar)
                nYYYYMMDD = JulToLong(Int(d), 1)
                nHHMM = Hour(d) * 100 + Minute(d)
                dOpen = Bars(eBARS_Open, nBar)
                dHigh = Bars(eBARS_High, nBar)
                dLow = Bars(eBARS_Low, nBar)
                dVol = Bars(eBARS_Vol, nBar)
                dOI = Bars(eBARS_OI, nBar)
                i = TAS_IndicatorSetBar(nFuncID, nBar, nYYYYMMDD, nHHMM, dOpen, dHigh, dLow, dClose, dVol, dOI)
                d = TAS_IndicatorValue(nFuncID, 0)
            End If
            AddList DateFormat(Bars(eBARS_DateTime, n)) & vbTab & Str(d)
        Next
    End If

End Sub

Private Sub GetVisionMargins()

    Dim s$, i&, iLine&, iCount&, strSymbol$, dMargin#
    Dim aLines As New cGdArray
    Dim aMargins As New cGdArray
    
    s = "http://www.visionfinancialmarkets.com/futures/tools/margin/"
    s = Trim(GetWebPageData(s))
    If Len(s) > 0 Then
        s = Replace(s, Chr(13), Chr(27))
        s = Replace(s, Chr(10), Chr(27))
        Do
            i = Len(s)
            s = Replace(s, Chr(27) & Chr(27), Chr(27))
        Loop While Len(s) < i
        aLines.SplitFields s, Chr(27)
        i = aLines.Size
        'aLines.ToFile "c:\test.txt"
        
        iCount = 999
        For iLine = 0 To aLines.Size - 1
            s = UCase(Trim(aLines(iLine)))
            If Left(s, 10) = "<TR CLASS=" Then
                ' new row
                iCount = 0 ' reset count
                strSymbol = ""
            ElseIf Left(s, 4) = "<TD>" Then
                iCount = iCount + 1
                If iCount = 2 Then
                    ' get our symbol
                    s = Mid(s, 5)
                    s = Parse(s, "<", 1)
                    If Left(s, 1) = "6" Then
                        s = "G" & s
                    ElseIf Left(s, 2) = "M6" Or s = "XK" Or s = "XY" Or Left(s, 1) = "Q" Or s = "E7" Then
                        s = ""
                    ElseIf Left(s, 1) = "X" And Right(s, 1) = "3" Then
                        s = ""
                    ElseIf s = "YG" Then
                        s = "XK"
                    ElseIf s = "YI" Then
                        s = "XY"
                    End If
                    s = ConvertFutureSymbol(s, eElectronicSymbol)
                    If Len(s) > 0 And Len(s) <= 3 Then
                        strSymbol = s & "-067"
                        If GetSymbolID(strSymbol) <= 0 Then
                            strSymbol = ""
                        End If
                    End If
                ElseIf iCount = 3 And Len(strSymbol) > 0 Then
                    s = Mid(s, 5)
                    s = Parse(s, "<", 1)
                    s = StripStr(s, "$,")
                    dMargin = Val(s)
                    If dMargin > 0 Then
                        s = strSymbol & vbTab & Str(dMargin)
                        aMargins.Add s
                    End If
                End If
            End If
        Next
    End If
    
    aMargins.Sort
    aMargins.ToFile "c:\Margins.txt"

End Sub

Private Sub CreateTickFiles()

    Dim i&, iSymbol&, nDate&, nStartDate&, nEndDate&, dTime#, dPrice#, dVol#, dTotalTicks#
    Dim s$, strSymbol$
    Dim bAppend As Boolean
    Dim Bars As New cGdBars
    Dim aSymbols As New cGdArray
    Dim aOut As New cGdArray
    Static bInProgress As Boolean
    
    If Not IsIDE Then Exit Sub ' make sure this isn't allowed for normal clients
    
    If bInProgress Then
        ' if called again while in progress, then set flag to quit
        bInProgress = False
        Exit Sub
    End If
    bInProgress = True
    
    nEndDate = LastDailyDownload
    nStartDate = nEndDate - 6
    'nStartDate = DateSerial(2014, 8, 1)
    
    s = "ZW,ZC,ZS,ZB,ZN,ES,YM,NQ,GC3,SI3,CL3,RB3,G6S,G6C,G6E,G6B,G6J,G6A"
    
    aSymbols.SplitFields UCase(s), ","
    aSymbols.Sort eGdSort_DeleteNullValues Or eGdSort_DeleteDuplicates
    For iSymbol = 0 To aSymbols.Size - 1
        If Not bInProgress Then Exit For
        
        ' all dates for this symbol
        bAppend = False
        For nDate = nStartDate To nEndDate
            aOut.Size = 0
            If Not bInProgress Then Exit For
            
            strSymbol = Trim(aSymbols(iSymbol))
            If Len(strSymbol) > 0 And IsWeekday(nDate) Then
                ' get ticks for this date
                strSymbol = RollSymbolForDate(strSymbol & "-057", nDate)
                DM_GetBars Bars, strSymbol, "each tick", nDate, nDate
                If Bars.Size > 0 Then
                    For i = 0 To Bars.Size - 1
                        ' Symbol, Date Time, Price, Volume
                        dTime = Bars(eBARS_DateTime, i)
                        dPrice = Bars(eBARS_Close, i)
                        dVol = Bars(eBARS_Vol, i)
                        If dTime > 0 And dPrice > 0 Then
                            If dVol < 0 Then dVol = 0
                            s = strSymbol & ", " & DateFormat(dTime, MM_DD_YYYY, HH_MM_SS) _
                                & ", " & Bars.PriceDisplay(dPrice, False, nDate) _
                                & ", " & Str(dVol)
                            aOut.Add s
                        End If
                    Next
                    ' store output for this symbol
                    If aOut.Size > 0 Then
                        dTotalTicks = dTotalTicks + aOut.Size
                        strSymbol = Trim(aSymbols(iSymbol))
                        If DirExist("F:\Ticks") Then
                            s = "F:\Ticks\" & strSymbol & ".txt"
                        Else
                            s = "C:\Ticks\" & strSymbol & ".txt"
                        End If
                        aOut.ToFile s, bAppend
                        bAppend = True
                    End If
                    AddList strSymbol & " " & DateFormat(nDate) & " = " & Str(aOut.Size) & " ticks"
                    DoEvents
                End If
            End If
        Next
        AddList "Total #Ticks = " & Str(dTotalTicks)
    Next
    
    AddList "Done"
    bInProgress = False

End Sub

Private Sub ChkSetBarProps()

    Dim i&, s$, d#
    Dim b As New cGdBars
    Dim aSymbols As New cGdArray

    'Set aSymbols = frmSymbolSelector.ShowMe(s)
    s = "CC-067,CL3-067,CT-067,DX-067,G6A-067,G6B-067,G6C-067,G6E-067,G6J-067,G6S-067,GC3-067,HE-067,HG3-067,HO3-067,KC-067,NG3-067,PA3-067,PL-067,RB3-067,SB-067,ZB-067,ZC-067,ZL-067,ZM-067,ZN-067,ZS-067,ZW-067"
    aSymbols.SplitFields s, ","
    d = gdTickCount
    For i = 0 To aSymbols.Size - 1
        s = aSymbols(i)
        SetBarProperties b, s
    Next
    d = gdTickCount - d
    s = Str(aSymbols.Size) & " symbols in " & Str(Round(d)) & " ms"
    AddList s

End Sub

Private Sub SetAppColors()

    Dim n&, bWhiteForeColor As Boolean
        
    'bWhiteForeColor = True
    If bWhiteForeColor Then
    
        'n = RGB(64, 176, 192) ' aqua
        
        'n = RGB(112, 128, 144)
        'n = 8421504
        'n = RGB(120, 124, 136) ' dark gray
        
        'n = RGB(96, 176, 216)
        'n = RGB(64, 176, 192) ' blue-green
        'n = RGB(90, 135, 190) ' blue
        'n = &H806040 ' bluer
        'n = RGB(96, 128, 160) ' darker blue
        n = RGB(90, 105, 120)
        'n = RGB(64, 72, 80)
        'n = RGB(50, 60, 70) ' quite dark
    
    Else
    
        'n = RGB(214, 218, 222)
        'n = RGB(224, 224, 216) ' light-gray
        
        ' BLUES
        'n = RGB(216, 224, 238) ' blue-gray
        'n = RGB(224, 232, 240) ' blue-gray
        'n = RGB(216, 228, 242)
        'n = RGB(212, 224, 236) ' light-blue
        'n = RGB(220, 228, 240) ' light-blue
        n = RGB(200, 216, 236) ' light-blue
        n = RGB(180, 214, 224)
            n = &HE0D0C0
        n = RGB(170, 200, 224)
        'n = RGB(125, 158, 192) ' mid-blue
        'n = RGB(120, 155, 180) ' darker-blue
    
    End If
    
    SetAppBackColor n, bWhiteForeColor

End Sub

Private Sub BuildDataFix()

    Dim i&, iRec&, nSymbolID&, nDate&, nBar&, dClose#
    Dim s$, strSymbol$
    Dim bFixed As Boolean
    Dim Bars As New cGdBars
    Dim TickBars As New cGdBars
    Dim aOut As New cGdArray
    
    ' look for specific symbols
    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            If Left(strSymbol, 3) = "UB-" And Len(strSymbol) = 9 Then
                AddList strSymbol
                bFixed = False
                ' get full daily history for this contract
                DM_GetBars Bars, strSymbol, "daily", 0, 20150214
                ' for each daily bar, get the last trade for that day
                For nBar = 0 To Bars.Size - 1
                    nDate = Bars(eBARS_DateTime, nBar)
                    TickBars.Size = 0
                    DM_GetBars TickBars, strSymbol, "1440 min", nDate, nDate
                    dClose = TickBars(eBARS_Close, 0)
                    If dClose <> kNullData And Bars(eBARS_Close, nBar) <> dClose Then
                        bFixed = True
'dClose = Bars(eBARS_Close, nBar)
                        Bars(eBARS_Close, nBar) = dClose
                        If dClose > Bars(eBARS_High, nBar) Then
                            Bars(eBARS_High, nBar) = dClose
                        End If
                        If dClose < Bars(eBARS_Low, nBar) Then
                            Bars(eBARS_Low, nBar) = dClose
                        End If
                    End If
                Next
                If bFixed Then
                    ' get fix file output for the entire contract
                    '#UB-201503
                    '20150213 169.343750 169.718750 167.437500 167.593750 79447 517616 83527 564136
                    s = "#" & strSymbol
                    aOut.Add s
                    For nBar = 0 To Bars.Size - 1
                        nDate = Bars(eBARS_DateTime, nBar)
                        s = Format(nDate, "YYYYMMDD") _
                            & " " & Bars.PriceDisplay(Bars(eBARS_Open, nBar), False) _
                            & " " & Bars.PriceDisplay(Bars(eBARS_High, nBar), False) _
                            & " " & Bars.PriceDisplay(Bars(eBARS_Low, nBar), False) _
                            & " " & Bars.PriceDisplay(Bars(eBARS_Close, nBar), False) _
                            & " " & Format(Bars(eBARS_ContVol, nBar), "#0") _
                            & " " & Format(Bars(eBARS_ContOI, nBar), "#0") _
                            & " " & Format(Bars(eBARS_Vol, nBar), "#0") _
                            & " " & Format(Bars(eBARS_OI, nBar), "#0")
                        aOut.Add s
                    Next
                End If
            End If
        End If
    Next
    
    aOut.ToFile "c:\DataFix.txt"

End Sub

Private Sub ChkQM()

    Dim i&, d#, dDate&, iCount&, s$
    Dim Bars As New cGdBars
    
    For dDate = Date - 1 To DateSerial(2014, 7, 1) Step -1
        If IsWeekday(dDate) Then
            DM_GetBars Bars, "QM-067", "each tick", dDate, dDate
            If Bars.Size > 0 Then
                d = dDate + (17 * 60 + 10) / 1440
                For i = Bars.Size - 1 To 0 Step -1
                    If Bars(eBARS_DateTime, i) < d Then
                        If i = Bars.Size - 1 Then
                            d = Bars(eBARS_DateTime, i)
                            s = DateFormat(d, MM_DD_YYYY, HH_MM_SS, NO_AMPM) & vbTab & "???????????????"
                            AddList s
                        End If
                        Exit For
                    End If
                Next
                d = Bars(eBARS_DateTime, i + 1)
                s = DateFormat(d, MM_DD_YYYY, HH_MM_SS, NO_AMPM)
                If d > dDate + (17 * 60 + 14.85) / 1440 Then
                    iCount = iCount + 1
                    s = s & vbTab & "#####"
                    AddList s
                End If
                'AddList s
            End If
        End If
    Next
    AddList Str(iCount)

End Sub

Private Sub FindEhlersData()

    Dim i&, iRec&, nSymbolID&, strSymbol$, dLow#, dHigh#
    Dim Bars As cGdBars
    
    Set Bars = New cGdBars
    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
            If InStr(strSymbol, "-") = 0 Or Right(strSymbol, 4) = "-067" Then
                'DM_GetBars Bars, strSymbol, "monthly", 19900101, 20041231, , , True
                DM_GetBars Bars, strSymbol, "monthly", 19950801, 19960301, , , True
                If Bars.Size > 1 Then
                    For i = 0 To Bars.Size - 1
                        If Month(Bars(eBARS_DateTime, i)) = 10 Then
                            If Bars(eBARS_Low, i - 1) < Bars(eBARS_Low, i) _
                                And Bars(eBARS_Low, i) < Bars(eBARS_Low, i + 1) _
                                And Bars(eBARS_Low, i + 1) < Bars(eBARS_Low, i + 2) _
                                And Bars(eBARS_Low, i + 2) < Bars(eBARS_Low, i + 3) _
                                And Bars(eBARS_Low, i + 3) > Bars(eBARS_Low, i + 4) Then
                                
                                    If Bars(eBARS_High, i + 2) > Bars(eBARS_High, i + 4) _
                                        And Bars(eBARS_High, i + 2) < Bars(eBARS_High, i + 3) Then
                                
                                        If Bars(eBARS_High, i - 1) < Bars(eBARS_High, i) _
                                            And Bars(eBARS_High, i) < Bars(eBARS_High, i + 1) _
                                            And Bars(eBARS_High, i + 1) < Bars(eBARS_High, i + 2) _
                                            And Bars(eBARS_High, i + 2) < Bars(eBARS_High, i + 3) _
                                            And Bars(eBARS_High, i + 3) > Bars(eBARS_High, i + 4) Then
                                
                                            If Bars(eBARS_Low, i + 4) > Bars(eBARS_High, i - 1) _
                                                And Bars(eBARS_Low, i + 2) > Bars(eBARS_High, i - 1) Then
                                        
                                                dLow = Bars(eBARS_Low, i)
                                                dHigh = Bars(eBARS_High, i + 3)
                                                'If dHigh - dLow > 7 And dHigh - dLow < 10 Then
                                                    AddList strSymbol & vbTab & Str(Year(Bars(eBARS_DateTime, i)))
                                                'End If
                                            End If
                                        End If
                                    End If
                            End If
#If 0 Then
                            dLow = Bars(eBARS_Low, i)
                            dHigh = Bars(eBARS_High, i)
                            If dLow > 112 And dLow < 114 And dHigh < 119 Then
                                AddList strSymbol & " ?"
                                dLow = Bars(eBARS_Low, i + 2)
                                dHigh = Bars(eBARS_High, i + 2)
                                If dLow > 117 And dLow < 118 And dHigh > 121 Then
                                    dLow = Bars(eBARS_Low, i + 3)
                                    dHigh = Bars(eBARS_High, i + 3)
                                    If dLow > 118 And dLow < 119 And dHigh > 122 And dHigh < 123 Then
                                        dLow = Bars(eBARS_Low, i + 3)
                                        dHigh = Bars(eBARS_High, i + 3)
                                        If dHigh > 120 And dLow < 115 Then
                                            AddList strSymbol
                                            DoEvents
                                        End If
                                    End If
                                End If
                            End If
#End If
                        End If
                    Next
                End If
            End If
        End If
        If iRec Mod 5000 = 0 Then
            AddList Str(iRec)
            DoEvents
        End If
    Next
    AddList "END"

End Sub

Private Sub ChkDJIA()

    Dim i&, d#, s$, dDate#
    Dim Bars As New cGdBars
    Dim Ticks As New cGdBars
    Dim aDiffs As New cGdArray
    
    dDate = 20150824
    
    lst.Clear
    aDiffs.Create eGDARRAY_Doubles
    
    DM_GetBars Ticks, "$DJIA", "each tick", dDate, dDate
    For i = 0 To Ticks.Size - 1
        d = Ticks(eBARS_DateTime, i)
        d = d - Int(d)
        If d < (9 * 60 + 31) / 1440# Then
            'AddList Format(d, "hh:mm:ss") & vbTab & Str(Ticks(eBARS_Close, i))
        End If
        If i > 0 Then
            d = Abs(Ticks(eBARS_Close, i) - Ticks(eBARS_Close, i - 1))
            If d > 50 Then
                aDiffs.Add d
                AddList Format(Ticks(eBARS_DateTime, i), "hh:mm:ss") & vbTab & Str(Ticks(eBARS_Close, i)) & vbTab & Str(d)
            End If
        End If
    Next
    
    AddList "From ticks:"
    Bars.BuildBars "1 min", Ticks.BarsHandle
    AddList BarDisplay(Bars, 0)
    AddList BarDisplay(Bars, 1)
    
    AddList "From DM:"
    DM_GetBars Bars, "$DJIA", "1 min", dDate, dDate
    AddList BarDisplay(Bars, 0)
    AddList BarDisplay(Bars, 1)
    
    AddList "Diffs:"
    aDiffs.Sort
    For i = 0 To aDiffs.Size - 1
        AddList Str(aDiffs(i))
    Next

End Sub

Private Sub CheckForexSpikes()

    Dim i&, iRec&, nSymbolID&, dSum#, dCount#, dRange#
    Dim s$, strSymbol$
    Dim Bars As New cGdBars
    Dim aFile As New cGdArray

    For iRec = 0 To g.SymbolPool.NumRecords - 1
        nSymbolID = g.SymbolPool.SymbolID(iRec)
        If nSymbolID > 0 Then
            strSymbol = g.SymbolPool.Symbol(iRec)
If strSymbol <> "$CAD-CHF" Then
    'strSymbol = ""
End If
            If Len(strSymbol) = 8 Then
                If IsForex(strSymbol) Then
                    AddList strSymbol
                    DM_GetBars Bars, strSymbol, "daily"
                    dCount = 1
                    dSum = 0
                    For i = Bars.Size - 1 To 0 Step -1
                        If Bars(eBARS_Close, i) <> kNullData Then
                            dRange = Bars(eBARS_High, i) - Bars(eBARS_Low, i)
                            If dRange > 10 * dSum / dCount And dCount > 50 And dSum > 0 Then
                                s = strSymbol & vbTab & DateFormat(Bars(eBARS_DateTime, i)) _
                                    & vbTab & Format(dRange / (dSum / dCount), "#0.0")
                                AddList s
                                aFile.Add s
                            Else
                                dCount = dCount + 1
                                dSum = dSum + dRange
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
    aFile.ToFile "c:\ForexSpikes.txt"
    AddList "Done"

End Sub

' Writes/appends a 4-byte Single into a string
Private Sub PutNumInString(ByRef s$, ByVal Number As Single, Optional nAtOffset As Long = 0)
    
    Dim i&, b(3) As Byte

    If nAtOffset <= 0 Then
        nAtOffset = Len(s) + 1 ' just append
    End If
    
    ' see if need to first make the string bigger
    If nAtOffset + 4 > Len(s) Then
        s = s & String(nAtOffset + 3 - Len(s), 0) ' pad with Nulls for now
    End If
    
    ' copy the bytes into a byte array
    CopyMemory b(0), ByVal GetAddress(Number), 4
    
    ' write the bytes into the string
    For i = 0 To 3
        Mid(s, nAtOffset + i, 1) = Chr(b(i))
    Next

End Sub

Private Function GetNumFromString(ByRef s$, ByRef nAtOffset As Long) As Double
    
    Dim i&, F As Single, b(3) As Byte

    ' make sure this is a valid location
    If nAtOffset >= 1 And nAtOffset + 3 <= Len(s) Then
        ' get the bytes from the string
        For i = 0 To 3
            b(i) = Asc(Mid(s, nAtOffset + i, 1))
        Next
        ' copy the bytes into the number
        CopyMemory ByVal GetAddress(F), b(0), 4
        nAtOffset = nAtOffset + 4
    End If
    GetNumFromString = RoundToSigDigits(F, 8)

End Function


