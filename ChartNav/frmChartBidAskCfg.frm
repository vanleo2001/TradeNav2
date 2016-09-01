VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmChartBidAskCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Bid/Ask Settings"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraBidAskOptions 
      Height          =   2445
      Left            =   218
      TabIndex        =   3
      Top             =   728
      Width           =   3480
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
      Caption         =   "frmChartBidAskCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmChartBidAskCfg.frx":0040
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartBidAskCfg.frx":0060
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboSize 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   1170
         Width           =   915
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
         Tip             =   "frmChartBidAskCfg.frx":007C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":009C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optShowNone 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   810
         Width           =   1890
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
         Caption         =   "frmChartBidAskCfg.frx":00B8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmChartBidAskCfg.frx":00EE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":010E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optShowByPrice 
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   495
         Width           =   2460
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
         Caption         =   "frmChartBidAskCfg.frx":012A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmChartBidAskCfg.frx":016C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":018C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optShowBySize 
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2760
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
         Caption         =   "frmChartBidAskCfg.frx":01A8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmChartBidAskCfg.frx":0204
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":0224
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fg 
         Height          =   795
         Left            =   90
         TabIndex        =   4
         Top             =   1575
         Width           =   1320
         _cx             =   2328
         _cy             =   1402
         _ConvInfo       =   1
         Appearance      =   0
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
         BackColorSel    =   8421504
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483648
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483648
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            Picture         =   "frmChartBidAskCfg.frx":0240
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   19
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   0
            Picture         =   "frmChartBidAskCfg.frx":054A
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   18
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   0
            Picture         =   "frmChartBidAskCfg.frx":0854
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   17
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   255
            Picture         =   "frmChartBidAskCfg.frx":0B5E
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   16
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   255
            Picture         =   "frmChartBidAskCfg.frx":0E68
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   15
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   255
            Picture         =   "frmChartBidAskCfg.frx":1172
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   14
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   510
            Picture         =   "frmChartBidAskCfg.frx":147C
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   510
            Picture         =   "frmChartBidAskCfg.frx":1786
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   12
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   510
            Picture         =   "frmChartBidAskCfg.frx":1A90
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   11
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   765
            Picture         =   "frmChartBidAskCfg.frx":1D9A
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   10
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   765
            Picture         =   "frmChartBidAskCfg.frx":20A4
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   9
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   765
            Picture         =   "frmChartBidAskCfg.frx":23AE
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   8
            Top             =   510
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1020
            Picture         =   "frmChartBidAskCfg.frx":26B8
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   7
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   1020
            Picture         =   "frmChartBidAskCfg.frx":29C2
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   6
            Top             =   255
            Width           =   255
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   1020
            Picture         =   "frmChartBidAskCfg.frx":2CCC
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   5
            Top             =   510
            Width           =   255
         End
      End
      Begin gdOCX.gdSelectColor gdColorBid 
         Height          =   315
         Left            =   2355
         TabIndex        =   20
         Top             =   2070
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdColorAsk 
         Height          =   315
         Left            =   2355
         TabIndex        =   21
         Top             =   1620
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   225
         Left            =   75
         Top             =   1215
         Width           =   1035
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
         Caption         =   "frmChartBidAskCfg.frx":2FD6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmChartBidAskCfg.frx":300A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":302A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   225
         Left            =   1605
         Top             =   2115
         Width           =   1035
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
         Caption         =   "frmChartBidAskCfg.frx":3046
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmChartBidAskCfg.frx":307A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":309A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   225
         Left            =   1605
         Top             =   1665
         Width           =   1035
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
         Caption         =   "frmChartBidAskCfg.frx":30B6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmChartBidAskCfg.frx":30EA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmChartBidAskCfg.frx":310A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   390
      Left            =   1943
      TabIndex        =   2
      Top             =   3323
      Width           =   795
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
      Caption         =   "frmChartBidAskCfg.frx":3126
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmChartBidAskCfg.frx":3154
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmChartBidAskCfg.frx":31B8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   390
      Left            =   1043
      TabIndex        =   1
      Top             =   3323
      Width           =   795
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
      Caption         =   "frmChartBidAskCfg.frx":31D4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmChartBidAskCfg.frx":31FA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmChartBidAskCfg.frx":321A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   405
      Left            =   173
      Top             =   158
      Width           =   3480
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
      Caption         =   "frmChartBidAskCfg.frx":3236
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmChartBidAskCfg.frx":32C8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmChartBidAskCfg.frx":32E8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmChartBidAskCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ePicIdx
    ePicIdx_ArrowNorth = 0
    ePicIdx_ArrowSouth
    ePicIdx_ArrowEast
    ePicIdx_ArrowWest
    ePicIdx_Cross
    ePicIdx_TriUpSolid
    ePicIdx_TriDownSolid
    ePicIdx_CircleSolid
    ePicIdx_SquareSolid
    ePicIdx_DiamondSolid
    ePicIdx_TriUp
    ePicIdx_TriDown
    ePicIdx_Circle
    ePicIdx_Square
    ePicIdx_Diamond
End Enum

Private Type mPrivate
    Chart As cChart

    eMode As eBidAskColorMode
    eImage As eStockImage
    eDir As eImageDir
    eIndex As ePicIdx
    
    nFill As Long
    nSize As Long
    nBidColor As Long
    nAskColor As Long

End Type

Private m As mPrivate

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    If optShowByPrice.Value = True Then
        m.eMode = eBidAskColorMode_ByPrice
    ElseIf optShowBySize.Value = True Then
        m.eMode = eBidAskColorMode_BySize
    Else
        m.eMode = eBidAskColorMode_None
    End If
       
    m.nSize = ValOfText(cboSize.Text)
    m.nBidColor = gdColorBid.Color
    m.nAskColor = gdColorAsk.Color
    
    m.Chart.BidAskPropLet m.eMode, m.eImage, m.eDir, m.nFill, m.nSize, m.nBidColor, m.nAskColor
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmChartBidAskCfg.cmdOK_Click"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartBidAskCfg.Form_Load"

End Sub

Public Sub ShowMe(Chart As cChart)
On Error GoTo ErrSection:

    Set m.Chart = Chart
    If m.Chart Is Nothing Then
        Unload Me
        Exit Sub
    End If

    m.Chart.BidAskPropGet m.eMode, m.eImage, m.eDir, m.nFill, m.nSize, m.nBidColor, m.nAskColor
    
    gdColorBid.Color = m.nBidColor
    gdColorAsk.Color = m.nAskColor
    
    CboSizeInit
    OptButtonsInit m.eMode
    ImgPicInit
    
    CenterFormOnChart Me, Chart
    ShowForm Me, eForm_Modal
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartBidAskCfg.ShowMe"

End Sub

Private Sub CboSizeInit()
On Error GoTo ErrSection:
    
    cboSize.Clear
    cboSize.AddItem ("2")
    cboSize.AddItem ("4")
    cboSize.AddItem ("6")
    cboSize.AddItem ("8")
    cboSize.AddItem ("10")
    cboSize.AddItem ("12")
    cboSize.AddItem ("14")

    If m.nSize < 2 Or m.nSize > 14 Then m.nSize = 2
    
    Select Case m.nSize
        Case 4:
            cboSize.ListIndex = 1
        Case 6:
            cboSize.ListIndex = 2
        Case 8:
            cboSize.ListIndex = 3
        Case 10:
            cboSize.ListIndex = 4
        Case 12:
            cboSize.ListIndex = 5
        Case 14:
            cboSize.ListIndex = 6
        Case Else:
            cboSize.ListIndex = 0
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartBidAskCfg.CboSizeInit"

End Sub

Private Sub ImgPicInit()

    Dim eIndex As ePicIdx
    Dim i As Integer
    
    Select Case m.eImage
        Case eCNI_Cross
            eIndex = ePicIdx_Cross
        Case eCNI_Arrow
            Select Case m.eDir
                Case eCNI_East
                    eIndex = ePicIdx_ArrowEast
                Case eCNI_West
                    eIndex = ePicIdx_ArrowWest
                Case eCNI_South
                    eIndex = ePicIdx_ArrowSouth
                Case Else
                    eIndex = ePicIdx_ArrowNorth
            End Select
        Case eCNI_Triangle
            If m.eDir = eCNI_North Then
                If m.nFill = 0 Then
                    eIndex = ePicIdx_TriUp
                Else
                    eIndex = ePicIdx_TriUpSolid
                End If
            Else
                If m.nFill = 0 Then
                    eIndex = ePicIdx_TriDown
                Else
                    eIndex = ePicIdx_TriDownSolid
                End If
            End If
        Case eCNI_Diamond
            If m.nFill = 0 Then
                eIndex = ePicIdx_Diamond
            Else
                eIndex = ePicIdx_DiamondSolid
            End If
        Case eCNI_Circle
            If m.nFill = 0 Then
                eIndex = ePicIdx_Circle
            Else
                eIndex = ePicIdx_CircleSolid
            End If
        Case eCNI_Square
                If m.nFill = 0 Then
                    eIndex = ePicIdx_Square
                Else
                    eIndex = ePicIdx_SquareSolid
                End If
        Case Else
            eIndex = ePicIdx_DiamondSolid
    End Select
    
    i = eIndex
    pic_Click i

End Sub

Private Sub OptButtonsInit(ByVal eMode As eBidAskColorMode)
On Error GoTo ErrSection:

    Select Case eMode
        Case eBidAskColorMode_ByPrice
            optShowNone.Value = False
            optShowBySize.Value = False
            optShowByPrice.Value = True
        Case eBidAskColorMode_BySize
            optShowNone.Value = False
            optShowBySize.Value = True
            optShowByPrice.Value = False
        Case Default
            optShowNone.Value = True
            optShowBySize.Value = False
            optShowByPrice.Value = False
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartBidAskCfg.OptButtonsInit"

End Sub

Private Sub pic_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim i&
    
    For i = 0 To 14
        If i = Index Then
            pic(i).BorderStyle = 1
        Else
            pic(i).BorderStyle = 0
        End If
    Next

    Select Case Index
        Case ePicIdx_ArrowNorth
            m.eImage = eCNI_Arrow
            m.eDir = eCNI_North
        Case ePicIdx_ArrowSouth
            m.eImage = eCNI_Arrow
            m.eDir = eCNI_South
        Case ePicIdx_ArrowEast
            m.eImage = eCNI_Arrow
            m.eDir = eCNI_East
        Case ePicIdx_ArrowWest
            m.eImage = eCNI_Arrow
            m.eDir = eCNI_West
        Case ePicIdx_Cross
            m.eImage = eCNI_Cross
        Case ePicIdx_TriUpSolid
            m.eImage = eCNI_Triangle
            m.eDir = eCNI_North
            m.nFill = 1
        Case ePicIdx_TriDownSolid
            m.eImage = eCNI_Triangle
            m.eDir = eCNI_South
            m.nFill = 1
        Case ePicIdx_CircleSolid
            m.eImage = eCNI_Circle
            m.nFill = 1
        Case ePicIdx_SquareSolid
            m.eImage = eCNI_Square
            m.nFill = 1
        Case ePicIdx_DiamondSolid
            m.eImage = eCNI_Diamond
            m.nFill = 1
        Case ePicIdx_TriUp
            m.eImage = eCNI_Triangle
            m.eDir = eCNI_North
            m.nFill = 0
        Case ePicIdx_TriDown
            m.eImage = eCNI_Triangle
            m.eDir = eCNI_South
            m.nFill = 0
        Case ePicIdx_Circle
            m.eImage = eCNI_Circle
            m.nFill = 0
        Case ePicIdx_Square
            m.eImage = eCNI_Square
            m.nFill = 0
        Case ePicIdx_Diamond
            m.eImage = eCNI_Diamond
            m.nFill = 0
    End Select
    
    m.eIndex = Index
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartBidAskCfg.pic_Click"

End Sub

Private Sub pic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    Dim i As Integer
    
    i = -1
    Select Case KeyCode
        Case vbKeyLeft
            If m.eIndex = 0 Then
                i = 4
            Else
                i = m.eIndex - 1
                If i = 4 Then
                    i = 9
                ElseIf i = 9 Then
                    i = 14
                End If
            End If
        Case vbKeyRight
            If m.eIndex = 14 Then
                i = 10
            Else
                i = m.eIndex + 1
                If i = 5 Then
                    i = 0
                ElseIf i = 10 Then
                    i = 5
                End If
            End If
        Case vbKeyUp
            If m.eIndex - 5 >= 0 Then
                i = m.eIndex - 5
            Else
                i = m.eIndex + 10
            End If
        Case vbKeyDown
            If m.eIndex + 5 <= 14 Then
                i = m.eIndex + 5
            Else
                i = m.eIndex - 10
            End If
    End Select
    
    If i >= 0 And i <= 14 Then
        m.eIndex = i
        pic_Click i
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmChartBidAskCfg.pic_KeyDown"
End Sub

