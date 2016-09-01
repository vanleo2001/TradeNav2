VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAddToChart 
   Caption         =   "Add to Chart"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   Icon            =   "frmAddToChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4425
   Visible         =   0   'False
   Begin HexUniControls.ctlUniTextBoxXP txtDesc 
      Height          =   915
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
      Width           =   4095
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmAddToChart.frx":0442
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
      Tip             =   "frmAddToChart.frx":0462
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAddToChart.frx":0482
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1935
      Left            =   3120
      TabIndex        =   1
      Top             =   660
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
      Caption         =   "frmAddToChart.frx":049E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAddToChart.frx":04BE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAddToChart.frx":04DE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   375
         Left            =   60
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   915
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
         Caption         =   "frmAddToChart.frx":04FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddToChart.frx":0528
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddToChart.frx":0548
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdProp 
         Height          =   375
         Left            =   60
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   915
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
         Caption         =   "frmAddToChart.frx":0564
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddToChart.frx":059A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddToChart.frx":05BA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Default         =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   915
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
         Caption         =   "frmAddToChart.frx":05D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddToChart.frx":05FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddToChart.frx":061E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   480
         Width           =   915
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
         Caption         =   "frmAddToChart.frx":063A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddToChart.frx":0668
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddToChart.frx":0688
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   915
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
         Caption         =   "frmAddToChart.frx":06A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddToChart.frx":06CC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddToChart.frx":06EC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraCommon 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   1080
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
         Caption         =   "frmAddToChart.frx":0708
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAddToChart.frx":0738
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAddToChart.frx":0758
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkFavorites 
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   0
            Width           =   975
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
            Caption         =   "frmAddToChart.frx":0774
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAddToChart.frx":07A8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAddToChart.frx":07C8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgList 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   2595
      _cx             =   4577
      _cy             =   5953
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
      ScrollBars      =   2
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
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   5235
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   9234
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
      Caption         =   "&Studies|&Indicators|&HighlightBars|S&trategies"
      Align           =   0
      Appearance      =   1
      CurrTab         =   3
      FirstTab        =   1
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
      Begin vsOcx6LibCtl.vsElastic vsElastic3 
         Height          =   4860
         Left            =   -4950
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   330
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   8573
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
         Caption         =   " To color the price bars where the condition is true:"
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   4
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   0
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
      Begin vsOcx6LibCtl.vsElastic vseIndicLabel 
         Height          =   4860
         Left            =   -5250
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   330
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   8573
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
         Caption         =   " Select indicator(s) to add to a new pane:"
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   4
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   0
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
      Begin vsOcx6LibCtl.vsElastic vsElastic1 
         Height          =   4860
         Left            =   -5550
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   8573
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
         Caption         =   " Prebuilt panes with indicators and default settings:"
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   4
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   0
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
      Begin vsOcx6LibCtl.vsElastic vsElastic4 
         Height          =   4860
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   330
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   8573
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
         Caption         =   " To run a trading strategy on the chart:"
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   4
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   0
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
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic2 
      Height          =   4860
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   8573
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
      Caption         =   " To color the price bars where the condition is true:"
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   4
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   0
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
End
Attribute VB_Name = "frmAddToChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAddToChart.frm
'' Description: Allow the user to add stuff to the chart
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/01/2013   DAJ         Changed code for loading strategies
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Const kCondBuilderPrompt = "Would you like to build the condition using indicators from the active chart?"

Private Enum eGDCols
    eGDCol_ID = 0
    eGDCol_Name = 1
    eGDCol_Desc = 2
    eGDCol_Favorite = 3
End Enum

Public Enum eChartAddListType
    eAdd_Undefined = -1
    eAdd_Previous = 0
    eAdd_Study = 1
    eAdd_Indicator = 2
    eAdd_AttachedInd = 3
    eAdd_HighlightBars = 4
    eAdd_System = 5
    eAdd_Any = 6
End Enum

Public Enum eChartAddListMode
    eAddMode0_NewPane = 0
    eAddMode1_SelectedPane
    eAddMode2_AttchToInd
    eAddMode3_CondBuilder
    eAddMode4_InstReplay
    eAddMode5_NewFunction
End Enum

Private Type mPrivate
    bCommon As Boolean
    eListType As eChartAddListType
    strSelected As String
    bMoveFocus As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

Private Sub chkFavorites_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        DisplayRows
        If chkFavorites.Value = vbChecked Then
            Call SetIniFileProperty("ShowCommon", True, "AddToChart", g.strIniFile)
        Else
            Call SetIniFileProperty("ShowCommon", False, "AddToChart", g.strIniFile)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddToChart.chkFavorites.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrSection:
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdAdd
    Add

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim strName$, strFile$, iRow&, i&
    Dim aRows As New cGdArray
    
    If fgList.SelectedRows > 1 Then
        If InfBox("Delete the " & CStr(fgList.SelectedRows) & " selected studies?", "?", "+Delete|-Cancel", "Confirm Delete") = "C" Then
            Exit Sub
        End If
    ElseIf fgList.SelectedRows <> 1 Then
        Beep
        Exit Sub
    End If
    With fgList
        For i = 0 To .SelectedRows - 1
            iRow = .SelectedRow(i)
            If iRow >= .FixedRows And iRow < .Rows Then
                If Not .RowHidden(iRow) Then
                    strName = Trim(.TextMatrix(iRow, GDCol(eGDCol_Name)))
                    strFile = App.Path & "\Charts\Templates\" & strName & ".STU"
                    If Not FileExist(strFile) Then
                        Beep
                    ElseIf .SelectedRows < 2 Then
                        If InfBox("Do you wish to permanently delete this| from the list of available studies:||" & strName, "?", "+Delete|-Cancel", "Confirm Delete") = "D" Then
                            KillFile strFile
                            aRows.Add iRow
                        End If
                    Else
                        KillFile strFile
                        aRows.Add iRow
                    End If
                End If
            End If
        Next
    End With
    If aRows.Size > 0 Then
        aRows.Sort eGdSort_Descending
        For i = 0 To aRows.Size - 1
            fgList.RemoveItem aRows(i)
        Next
    End If
    
ErrExit:
    Set aRows = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    Dim strAnswer$
    
    If vsTab.CurrTab = 2 Then
        strAnswer = InfBox(kCondBuilderPrompt, "?", "+Yes|No", "Add to chart")
        If strAnswer = "Y" Then
            m.strSelected = "ShowCondBuilder"
        Else
            m.strSelected = "ShowFunctionEditor"
        End If
    Else
        m.strSelected = "ShowFunctionEditor"
    End If
    
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.cmdNew_Click"

End Sub

Private Sub cmdProp_Click()
On Error GoTo ErrSection:

    Dim strName$, strDesc$, strFile$, i&, iRow&
    Dim aStudy As New cGdArray
    
    With fgList
        ' load selected study
        iRow = .Row
        If iRow >= .FixedRows And iRow < .Rows Then
            strName = Trim(.TextMatrix(iRow, GDCol(eGDCol_Name)))
            strFile = App.Path & "\Charts\Templates\" & strName & ".STU"
            aStudy.FromFile strFile '(need whole file so we can resave it)
        End If
        If Len(strName) = 0 Or aStudy.Size = 0 Then
            Beep
        Else
            ' get name and desc
            For i = 0 To aStudy.Size - 1
                Select Case UCase(Left(aStudy(i), 5))
                Case "[IND;"
                    Exit For
                Case "NAME="
                    strName = Trim(Mid(aStudy(i), 6))
                Case "DESC="
                    strDesc = Trim(Mid(aStudy(i), 6))
                End Select
            Next
            ' edit name and desc
            If frmNameDesc.ShowMe(strName, strDesc, False) Then
                ' replace with new name and desc
                For i = 0 To aStudy.Size - 1
                    Select Case UCase(Left(aStudy(i), 5))
                    Case "[IND;"
                        Exit For
                    Case "[SC; "
                        If Trim(Parse(UCase(aStudy(i)), ";", 2)) <> "PRICE PANE" Then
                            aStudy(i) = "[SC; " & strName & "; 0]"
                        End If
                    Case "NAME="
                        aStudy(i) = "NAME=" & strName
                    Case "DESC="
                        aStudy(i) = "DESC=" & strDesc
                    End Select
                Next
                ' replace file
                KillFile strFile
                strFile = App.Path & "\Charts\Templates\" & strName & ".STU"
                aStudy.ToFile strFile
                
                'ShowType eAdd_Study
                .TextMatrix(iRow, GDCol(eGDCol_Name)) = strName
                .TextMatrix(iRow, GDCol(eGDCol_Desc)) = strDesc
                DisplayDesc
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.cmdProp.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    DisplayDesc

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.fgList.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgList_DblClick()
On Error GoTo ErrSection:
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdAdd
    Add

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.fgList.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgList_KeyUp(KeyCode As Integer, Shift As Integer)

    ' to select row as type letters
    If KeyCode >= 32 And Shift = 0 Then
        On Error Resume Next
        fgList.Row = fgList.Row
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddToChart.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16("kToolsAdd")    'JM:03-30-2009 - this icon does not exist; form has no icon
    
    g.Styler.StyleForm Me
    
    If ExtremeCharts = 1 Then
        vsTab.TabVisible(3) = False
    Else
        vsTab.TabVisible(3) = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next
        
    Dim w&, h&
    
    'check minimum size
    w = fraButtons.Width * 2
    h = fraButtons.Top + fraButtons.Height + txtDesc.Height + 120
    If LimitFormSize(Me, w, h) Then Exit Sub

    With vsTab
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top
    End With

    fraButtons.Left = Me.ScaleWidth - fraButtons.Width - vsTab.Left - 60
    With txtDesc
        .Move fgList.Left, Me.ScaleHeight - .Height - 90, Me.ScaleWidth - .Left * 2, .Height
    End With
    With fgList
        .Move .Left, .Top, fraButtons.Left - .Left - 60, txtDesc.Top - .Top - 60
    End With
    
    #If 0 Then
    With fgList
        If chkCommon.Enabled Then
            chkCommon.Top = Me.ScaleHeight - chkCommon.Height
            h = chkCommon.Top - .Top
        Else
            chkCommon.Top = -2 * chkCommon.Height
            h = Me.ScaleHeight - .Top * 2
        End If
        .Move .Left, .Top, fraButtons.Left - .Left * 2, h
    End With
    #End If
    
    'Me.Refresh

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgList
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 0 '1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShowAndMove
        .ScrollTrack = True
        '.SelectionMode = flexSelectionFree
        .SelectionMode = flexSelectionListBox
        .AllowSelection = True 'False
        '.AllowUserFreezing = flexFreezeColumns
        .SheetBorder = RGB(128, 128, 128)
        .ExtendLastCol = True
        .Editable = flexEDNone ' flexEDKbdMouse
        .AutoSearch = flexSearchFromTop
        
        .Rows = .FixedRows
        .Cols = 4
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_Desc)) = True
        .ColHidden(GDCol(eGDCol_Favorite)) = True
        If .FixedRows > 0 Then
            Select Case m.eListType
            Case eAdd_HighlightBars
                .TextMatrix(0, GDCol(eGDCol_Name)) = "HighlightBars"
            Case eAdd_Study
                .TextMatrix(0, GDCol(eGDCol_Name)) = "Studies"
            Case Else
                .TextMatrix(0, GDCol(eGDCol_Name)) = "Indicators"
            End Select
            .TextMatrix(0, GDCol(eGDCol_Desc)) = "Description"
            .FillStyle = flexFillRepeat
            .Select 0, 0, 0, .Cols - 1
            .CellFontBold = True
            '.CellForeColor = fraAppearance.ForeColor
        End If
        
        '.AutoSize 1
        '.ColWidth(0) = .Width - .ColWidth(1) - 4 * Screen.TwipsPerPixelX
        .ExtendLastCol = True
        '.ColAlignment(1) = flexAlignCenterCenter
        
        '.TextMatrix(1, 0) = "Testing"
        '.Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub LoadSystems()
On Error GoTo ErrSection:
    
    Dim i&, strName$, strDesc$
    Dim rsSystems As Recordset
    Dim lScreenPointer As Long
    
    lScreenPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    With fgList
        .Redraw = flexRDNone
        .Rows = .FixedRows
                
        Set rsSystems = mSysNav.LoadStrategiesRecordset
        If Not (rsSystems.BOF And rsSystems.EOF) Then
            rsSystems.MoveFirst
            
            Do While Not rsSystems.EOF
                If mSysNav.IncludeStrategiesFromRecordset(rsSystems) Then
                    strName = rsSystems!SystemName
                    strDesc = "Library:  " & rsSystems!LibraryName & vbCrLf & "Developer:  " & rsSystems!Developer
                    
                    .AddItem CStr(rsSystems!SystemNumber) & vbTab & strName & vbTab & strDesc
                End If
                
                rsSystems.MoveNext
            Loop
        End If
               
        '.Col = GDCol(eGDCol_Name)
        '.Sort = flexSortGenericAscending
        '.AutoSize 0
        '.Redraw = flexRDBuffered
    End With
    
ErrExit:
    Screen.MousePointer = lScreenPointer
    Set rsSystems = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.LoadSystems", eGDRaiseError_Raise
End Sub

Private Sub LoadStudies()
On Error GoTo ErrSection:
    
    Dim i&, strName$, strDesc$
    Dim aStudies As cGdArray
    
    Set aStudies = GetAllowedList("S")
    With fgList
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        For i = 0 To aStudies.Size - 1
            strName = Parse(aStudies(i), vbTab, 1)
            strDesc = Parse(aStudies(i), vbTab, 4)
            If Len(strName) > 0 Then
                .AddItem vbTab & strName & vbTab & strDesc
            End If
        Next
        
        '.Col = GDCol(eGDCol_Name)
        '.Sort = flexSortGenericAscending
        '.AutoSize 0
        '.Redraw = flexRDBuffered
    End With
    Set aStudies = Nothing
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.LoadStudies", eGDRaiseError_Raise

End Sub

Private Sub LoadFunctions()
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim rsParms As Recordset
    Dim i As Long
    Dim iRow As Long
    Dim iPrevFuncID As Long
    Dim strPrevName As String
    Dim bShow As Boolean
    Dim astrFavorites As New cGdArray
    Dim strText As String
    
    ' load common functions
    'aCommon.FromFile App.Path & "\Info\Common.IND"
    'aCommon.Sort eGdSort_IgnoreCase
    Set astrFavorites = GetFunctionFavorites
    
    'Load all system functions
    'Set rs = g.dbNav.OpenRecordset("qryAllfunctions", dbOpenSnapshot)
    ''Set rs = g.dbNav.OpenRecordset("qryFunctionsAndParms", dbOpenSnapshot)
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblLibrarys.* " & _
                "FROM tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
                "WHERE ((tblLibrarys.Ignore)=0);", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    ValidateCheckSums rs, "tblLibrarys"
    
    iPrevFuncID = -1
    With fgList
        .Redraw = flexRDNone
        .Rows = .FixedRows
        iRow = .FixedRows
        ' add "custom" row
        If m.eListType <> eAdd_AttachedInd Then
            .Rows = .Rows + 1
            .TextMatrix(iRow, GDCol(eGDCol_ID)) = "..."
            If m.eListType = eAdd_HighlightBars Then
                .TextMatrix(iRow, GDCol(eGDCol_Name)) = " CUSTOM HighlightBars"
            Else
                .TextMatrix(iRow, GDCol(eGDCol_Name)) = " CUSTOM Indicator"
            End If
            .TextMatrix(iRow, GDCol(eGDCol_Desc)) = "Can define a custom expression"
            iRow = iRow + 1
        End If
        If Not rs.EOF Then
            rs.MoveLast
            .Rows = rs.RecordCount + .Rows '(for now)
            rs.MoveFirst
            Do While Not rs.EOF
                If rs![tblFunctions.CheckSum] <> 0.5 And rs![tblLibrarys.CheckSum] <> 0.5 And HasModule(NullChk(rs![tblFunctions.RequiredMod])) Then
                    bShow = False
                    
                    'only want to look at first record (parm) of each function
                    If rs!FunctionID <> iPrevFuncID And rs![tblFunctions.SecurityLevel] < 3 _
                            And (rs!Usage And 4) <> 0 Then
                        iPrevFuncID = rs!FunctionID
                        strText = UCase(rs!CodedName)
                        If m.eListType = eAdd_HighlightBars Then
                            'only functions which return an array of booleans
                            If rs!ReturnTypeID = 3 Then
                                'Select Case UCase(rs!FunctionCategory)
                                '    Case "INDICATOR", "DATA"
                                        bShow = True
                                'End Select
                            End If
                        Else
                            'only functions which return an array of numbers
                            If rs!ReturnTypeID = 4 Then
                                'Select Case UCase(rs!FunctionCategory)
                                '    Case "INDICATOR", "DATA"
                                        bShow = True
                                'End Select
                            End If
                            'if attached, only want functions where first parm takes an array
                            If m.eListType = eAdd_AttachedInd And bShow = True Then
                                Set rsParms = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionParms] " & _
                                            "WHERE [FunctionID] = " & rs!FunctionID & " AND [ParmNbr] = 1;", dbOpenDynaset)
                                If Not (rsParms.BOF And rsParms.EOF) Then
                                    If NullChk(rsParms!ParmTypeID, 0) <> 4 Then
                                        bShow = False
                                    End If
                                Else
                                    bShow = False
                                End If
                            ElseIf bShow Then
                                If strText = "SPREAD" Or strText = "RELATIVESTRENGTHRATIO" Then
                                    bShow = False
                                End If
                            End If
                        End If
                        If InStr(strText, "GREENLIGHT") > 0 Then
                            'If InStr(strMods, ",ETA,") <= 0 Then
                            If InStr(g.strAuthorizationString, ",ETA,") <= 0 Then
                                bShow = False
                            End If
                        End If
                    End If
                    If bShow Then
                    
                        .TextMatrix(iRow, GDCol(eGDCol_ID)) = rs!CodedName
                        .TextMatrix(iRow, GDCol(eGDCol_Name)) = rs!FunctionName
                        .TextMatrix(iRow, GDCol(eGDCol_Desc)) = rs!Description
                        '.TextMatrix(X, C_LIBRARYNAME) = rs!LibraryName
                        '.TextMatrix(X, C_CATEGORY) = rs!FunctionCategory
                        '.TextMatrix(X, C_LASTMODIFIED) = Format(rs!LastModified, "mmm dd, yyyy  hh:mm:ss am/pm")
                        'If rs!Reverify Then .TextMatrix(X, C_VERIFIED) = "No" Else .TextMatrix(X, C_VERIFIED) = "Yes"
                        '.TextMatrix(X, C_PREVIEW) = "Usage: " & rs!TradeSenseUsage & Chr(13) & Chr(10) & _
                        '    "Description: " & rs!Description
                        '.TextMatrix(X, C_FUNCTIONID) = rs!FunctionID
                        '.TextMatrix(X, C_IMPLTYPE) = rs!ImplementationTypeID
                        '.TextMatrix(X, C_SECURITYLEVEL) = rs!SecurityLevel
                        'If IsNull(rs!password) Then .TextMatrix(X, C_PASSWORD) = "" Else .TextMatrix(X, C_PASSWORD) = rs!password
                        '.TextMatrix(X, C_CANNOTDELETE) = rs!CannotDelete
                        If astrFavorites.BinarySearch(rs!CodedName) Then
                            .TextMatrix(iRow, GDCol(eGDCol_Favorite)) = "1"
                        End If
                        iRow = iRow + 1
                    End If
                End If
                
                rs.MoveNext
            Loop
        End If
        
        .Rows = iRow

        '.AutoSize 0
        '.Redraw = flexRDBuffered
    End With
    
ErrExit:
    Set rsParms = Nothing
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rsParms = Nothing
    Set rs = Nothing
    RaiseError "frmAddToChart.LoadFunctions", eGDRaiseError_Raise

End Sub

Private Sub DisplayRows()
On Error GoTo ErrSection:

    Dim iRow&, iShowRow&, bJustCommon As Boolean

    If chkFavorites.Value = vbChecked And (m.eListType <> eAdd_Study And m.eListType <> eAdd_System) Then
        bJustCommon = True
    End If
    With fgList
        .Redraw = flexRDNone
        iShowRow = .Row
        For iRow = .FixedRows To .Rows - 1
            If Not bJustCommon Then
                .RowHidden(iRow) = False
            ElseIf ValOfText(.TextMatrix(iRow, GDCol(eGDCol_Favorite))) <> 0 Then
                .RowHidden(iRow) = False
            Else
                .RowHidden(iRow) = True
            End If
        Next
        .Redraw = flexRDBuffered
        If iShowRow >= .FixedRows And iShowRow < .Rows Then
            .Row = iShowRow
            .Select iShowRow, 0, iShowRow, .Cols - 1
            .ShowCell iShowRow, 0
        End If
        .Col = GDCol(eGDCol_Name)
    End With
    
    DisplayDesc

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.DisplayRows", eGDRaiseError_Raise

End Sub

Private Function ShowType(eListType As eChartAddListType)
On Error GoTo ErrSection:

    Dim i&, nID&

    m.eListType = eListType
    m.strSelected = ""

    'load list
    fgList.Redraw = flexRDNone
    InitGrid
    Select Case m.eListType
        Case eAdd_System
            fgList.AllowSelection = False
            LoadSystems
            fraCommon.Visible = False
            cmdProp.Visible = False
            cmdDelete.Visible = False
            cmdNew.Visible = False
        Case eAdd_Study
            LoadStudies
            fraCommon.Visible = False
            cmdProp.Visible = True
            cmdDelete.Visible = True
            cmdNew.Visible = False
        Case Else
            LoadFunctions
            fraCommon.Visible = True
            cmdProp.Visible = False
            cmdDelete.Visible = False
            cmdNew.Visible = True
    End Select
    If Not HasGold(False) Then
        chkFavorites.Value = vbChecked
        fraCommon.Visible = False
    Else
        i = GetIniFileProperty("ShowCommon", 0, "AddToChart", g.strIniFile)
        If i = 0 Then
            chkFavorites.Value = vbUnchecked
        Else
            chkFavorites.Value = vbChecked
        End If
    End If
    
    'sort list
    With fgList
        .Redraw = flexRDNone
        If .Rows > .FixedRows Then
            .Select .FixedRows, GDCol(eGDCol_Name)
            .Sort = flexSortGenericAscending
            .Row = .FixedRows
            .AutoSize GDCol(eGDCol_Name)
        End If
        
        If m.eListType = eAdd_System Then
            If Not ActiveChart Is Nothing Then
                nID = ActiveChart.Chart.SystemID
                If nID > 0 Then
                    For i = .FixedRows To .Rows - 1
                        If ValOfText(.TextMatrix(i, 0)) = nID Then
                            .Row = i
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        
        '.Redraw = flexRDBuffered
    End With
    
    DisplayRows

    If m.bMoveFocus Then
        MoveFocus fgList
        m.bMoveFocus = False
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmAddToChart.ShowType", eGDRaiseError_Raise

End Function

Public Function ShowMe(eListType As eChartAddListType, _
        Optional ByVal strIndicatorCaption$ = "")
On Error GoTo ErrSection:

    Dim i&, iTab&
    Dim Chart As cChart
    
    If Not ActiveChart Is Nothing Then
        Set Chart = ActiveChart.Chart
        If Not Chart Is Nothing Then
            If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then
                InfBox kSeasonalUnavail, "I", , "Add to chart"
                Exit Function
            End If
        End If
    End If
    
    If Len(Trim(strIndicatorCaption)) = 0 Then
        strIndicatorCaption = "Select indicator(s) to add to a new pane:"
    End If
    
    If strIndicatorCaption = "New Chart" Then
        Me.Caption = "Select strategy ..."
        cmdAdd.Caption = "&OK"
    ElseIf strIndicatorCaption <> "ChartIsInGameMode" Then
        vseIndicLabel.Caption = " " & strIndicatorCaption
    End If
    
    If eListType = eAdd_Previous Then
        If m.eListType = eAdd_Previous Then
            eListType = eAdd_Study '(default)
        ElseIf m.eListType = eAdd_AttachedInd Then
            eListType = eAdd_Study
        Else
            eListType = m.eListType
        End If
    End If
    
    Select Case eListType
    Case eAdd_System
        iTab = 3
    Case eAdd_Study, eAdd_Any
        iTab = 0
    Case eAdd_HighlightBars
        iTab = 2
    Case Else
        iTab = 1
    End Select
    
    For i = 0 To vsTab.NumTabs - 1
        If strIndicatorCaption = "New Chart" Then
            vsTab.TabEnabled(i) = False
        Else
            If vsTab.TabCaption(i) = "S&trategies" And strIndicatorCaption = "ChartIsInGameMode" Then
                vsTab.TabEnabled(i) = False
            ElseIf eListType = eAdd_HighlightBars Then
                If i = 2 Then
                    vsTab.TabEnabled(i) = True
                Else
                    vsTab.TabEnabled(i) = False
                End If
            Else
                If i = iTab Or eListType <> eAdd_AttachedInd Then
                    If i = 2 Then
                        If eListType = eAdd_Any Then    'this is true when using 'A' hot key
                            vsTab.TabEnabled(i) = True
                        Else
                            vsTab.TabEnabled(i) = False
                        End If
                    Else
                        vsTab.TabEnabled(i) = True
                    End If
                Else
                    vsTab.TabEnabled(i) = False
                End If
            End If
        End If
    Next
    
    If eListType = eAdd_Any Then
        eListType = eAdd_Study '(default, but all are enabled)
    End If
    
    m.eListType = 0 ' (so "Switch" will not do "ShowType")
    vsTab.FirstTab = 0
    vsTab.CurrTab = iTab
    ShowType eListType

    CenterFormOnChart Me, Chart         '6499
    ShowForm Me, True
    
    eListType = m.eListType
    ShowMe = m.strSelected
    Unload Me

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmAddToChart.ShowMe", eGDRaiseError_Raise

End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    'Call SetIniFileProperty("ShowCommon", optCommon, "AddToChart", g.strIniFile)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub DisplayDesc()
On Error GoTo ErrSection:
    
    Dim iRow&, strDesc$
    
    With fgList
        iRow = .Row
        If iRow >= .FixedRows And iRow < .Rows Then
            If .RowHidden(iRow) = False Then
                'strDesc = Trim(.TextMatrix(iRow, GDCol(eGDCol_Name))) _
                '    & ":" & vbCrLf & Trim(.TextMatrix(iRow, GDCol(eGDCol_Desc)))
                strDesc = Trim(.TextMatrix(iRow, GDCol(eGDCol_Desc)))
            End If
        End If
        txtDesc = strDesc
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.DisplayDesc", eGDRaiseError_Raise

End Sub

Private Sub vsTab_Click()
    m.bMoveFocus = True
End Sub

Private Sub vsTab_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    If OldTab <> NewTab And m.eListType <> 0 Then
        m.bMoveFocus = True
        Select Case NewTab
        Case 0:
            ShowType eAdd_Study
        Case 1:
            ShowType eAdd_Indicator
        Case 2:
            ShowType eAdd_HighlightBars
        Case 3:
            ShowType eAdd_System
        End Select
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.vsTab.Switch", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Add()
On Error GoTo ErrSection:

    Dim i&, r&, strAdd$
    Dim nSelRows&
    
    
    If Me.vsTab.CurrTab = 3 Then
        If Not HasGold(True, "Running a strategy on a chart", False) Then
            Exit Sub
        End If
    End If
    
    'calculate number selected rows that are not hidden (aardvark 934 fix)
    nSelRows = 0
    For i = 0 To fgList.SelectedRows - 1
        r = fgList.SelectedRow(i)
        If Not fgList.RowHidden(r) Then
            nSelRows = nSelRows + 1
        End If
    Next
    
    If nSelRows > 10 Then
        InfBox "You cannot add more than 10 items| at a time to the chart.", "i", , "Add to Chart"
        Exit Sub
    End If
    
    m.strSelected = ""
    With fgList
        For i = 0 To .SelectedRows - 1
            r = .SelectedRow(i)
            strAdd = ""
            If Not .RowHidden(r) Then
                strAdd = .TextMatrix(r, GDCol(eGDCol_ID))
                If Len(strAdd) = 0 Then
                    strAdd = .TextMatrix(r, GDCol(eGDCol_Name))
                End If
                If strAdd = "..." Then
                    If Not HasGold(True, "Custom Indicator") Then
                        strAdd = ""
                    End If
                End If
            End If
            If Len(strAdd) > 0 Then
                If m.strSelected = "" Then
                    m.strSelected = strAdd
                Else
                    m.strSelected = m.strSelected & vbTab & strAdd
                End If
            End If
        Next
    End With
    
    Dim strAnswer$
    
    If strAdd = "..." And vsTab.CurrTab = 2 Then
        strAnswer = InfBox(kCondBuilderPrompt, "?", "+Yes|No", "Add to chart")
        If strAnswer = "Y" Then m.strSelected = "ShowCondBuilder..."
    End If
    
    Me.Hide

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAddToChart.Add", eGDRaiseError_Raise

End Sub

