VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTTSummaryCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vsOcx6LibCtl.vsIndexTab tabSettings 
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   12621
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
      Caption         =   "&Trading|Co&nsole|&Web Status Report"
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
      Begin HexUniControls.ctlUniFrameWL fraWeb 
         Height          =   6780
         Left            =   7260
         TabIndex        =   39
         Top             =   330
         Width           =   6225
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
         Caption         =   "frmTTSummaryCfg.frx":0000
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTSummaryCfg.frx":002C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTSummaryCfg.frx":004C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdWebSample 
            Height          =   435
            Left            =   1680
            TabIndex        =   6
            Top             =   5700
            Width           =   2835
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
            Caption         =   "frmTTSummaryCfg.frx":0068
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":00C4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":00E4
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgAccounts 
            Height          =   2295
            Left            =   120
            TabIndex        =   21
            Top             =   1920
            Width           =   5955
            _cx             =   10504
            _cy             =   4048
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
         Begin MSComctlLib.Slider sliFontSize 
            Height          =   315
            Left            =   3060
            TabIndex        =   25
            Top             =   4440
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   7
            SelStart        =   4
            Value           =   4
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   1275
            Left            =   360
            Top             =   180
            Width           =   5415
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
            Caption         =   "frmTTSummaryCfg.frx":0100
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":0480
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":04A0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   255
            Left            =   180
            Top             =   5040
            Width           =   5835
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
            Caption         =   "frmTTSummaryCfg.frx":04BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":0574
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":0594
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblURL 
            Height          =   255
            Left            =   180
            Top             =   5340
            Width           =   5835
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
            Caption         =   "frmTTSummaryCfg.frx":05B0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":062C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":064C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   4680
            Top             =   4440
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
            Caption         =   "frmTTSummaryCfg.frx":0668
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":0692
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":06B2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblFontSize 
            Height          =   255
            Left            =   660
            Top             =   4440
            Width           =   2355
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
            Caption         =   "frmTTSummaryCfg.frx":06CE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":0730
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":0750
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAccounts 
            Height          =   255
            Left            =   120
            Top             =   1620
            Width           =   2295
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
            Caption         =   "frmTTSummaryCfg.frx":076C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":07C6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":07E6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraTrading 
         Height          =   6780
         Left            =   45
         TabIndex        =   1
         Top             =   330
         Width           =   6225
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
         Caption         =   "frmTTSummaryCfg.frx":0802
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTSummaryCfg.frx":0836
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTSummaryCfg.frx":0856
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraOptionTrading 
            Height          =   1455
            Left            =   60
            TabIndex        =   20
            Top             =   5220
            Width           =   6075
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
            Caption         =   "frmTTSummaryCfg.frx":0872
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTSummaryCfg.frx":08AE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":08CE
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraOptionFillOptions 
               Height          =   495
               Left            =   555
               TabIndex        =   26
               Top             =   900
               Width           =   5115
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
               Caption         =   "frmTTSummaryCfg.frx":08EA
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmTTSummaryCfg.frx":0916
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":0936
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optBidAsk 
                  Height          =   375
                  Left            =   0
                  TabIndex        =   27
                  Top             =   0
                  Width           =   1515
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
                  Caption         =   "frmTTSummaryCfg.frx":0952
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmTTSummaryCfg.frx":0990
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTSummaryCfg.frx":09B0
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optMidpoint 
                  Height          =   375
                  Left            =   1860
                  TabIndex        =   28
                  Top             =   0
                  Width           =   1515
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
                  Caption         =   "frmTTSummaryCfg.frx":09CC
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmTTSummaryCfg.frx":0A2A
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTSummaryCfg.frx":0A4A
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optTwoThirds 
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1515
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
                  Caption         =   "frmTTSummaryCfg.frx":0A66
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmTTSummaryCfg.frx":0AC2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTSummaryCfg.frx":0AE2
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraOptionOpenEquity 
               Height          =   375
               Left            =   2220
               TabIndex        =   22
               Top             =   240
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
               Caption         =   "frmTTSummaryCfg.frx":0AFE
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmTTSummaryCfg.frx":0B2A
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":0B4A
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optOptionOpenEquityLast 
                  Height          =   375
                  Left            =   2640
                  TabIndex        =   24
                  Top             =   0
                  Width           =   615
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
                  Caption         =   "frmTTSummaryCfg.frx":0B66
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmTTSummaryCfg.frx":0B90
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTSummaryCfg.frx":0BB0
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optOptionOpenEquityBidAsk 
                  Height          =   375
                  Left            =   180
                  TabIndex        =   23
                  Top             =   0
                  Width           =   2355
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
                  Caption         =   "frmTTSummaryCfg.frx":0BCC
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmTTSummaryCfg.frx":0C22
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTSummaryCfg.frx":0C42
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniLabelXP lblOptionOpenEquity 
               Height          =   255
               Left            =   120
               Top             =   300
               Width           =   2115
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
               Caption         =   "frmTTSummaryCfg.frx":0C5E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":0CB6
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":0CD6
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblOptionFill 
               Height          =   255
               Left            =   120
               Top             =   660
               Width           =   5475
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
               Caption         =   "frmTTSummaryCfg.frx":0CF2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":0D76
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":0D96
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraOrderOptions 
            Height          =   2355
            Left            =   60
            TabIndex        =   2
            Top             =   120
            Width           =   6075
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
            Caption         =   "frmTTSummaryCfg.frx":0DB2
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTSummaryCfg.frx":0DEC
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":0E0C
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkConfirmTradeSense 
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   1470
               Width           =   5655
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
               Caption         =   "frmTTSummaryCfg.frx":0E28
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":0EBA
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":0EDA
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkConfirmTriggered 
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Top             =   1185
               Width           =   5655
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
               Caption         =   "frmTTSummaryCfg.frx":0EF6
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":0F86
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":0FA6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkConfirmManual 
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   900
               Width           =   5655
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
               Caption         =   "frmTTSummaryCfg.frx":0FC2
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":104C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":106C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkWarnLimit 
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   2040
               Width           =   5655
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
               Caption         =   "frmTTSummaryCfg.frx":1088
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1140
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1160
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkWarnStop 
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   1755
               Width           =   5655
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
               Caption         =   "frmTTSummaryCfg.frx":117C
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1232
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1252
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkStopBuffer 
               Height          =   195
               Left            =   120
               TabIndex        =   3
               Top             =   270
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
               Caption         =   "frmTTSummaryCfg.frx":126E
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1310
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1330
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtStopBuffer 
               Height          =   285
               Left            =   4920
               TabIndex        =   4
               Top             =   225
               Width           =   420
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTSummaryCfg.frx":134C
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
               Alignment       =   2
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmTTSummaryCfg.frx":136E
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":138E
            End
            Begin gdOCX.gdScrollBar sbStopBuffer 
               Height          =   360
               Left            =   5340
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   180
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   635
            End
            Begin HexUniControls.ctlUniLabelXP lblStopBuffer 
               Height          =   375
               Left            =   360
               Top             =   480
               Width           =   5655
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
               Caption         =   "frmTTSummaryCfg.frx":13AA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":146A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":148A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraJournal 
            Height          =   1215
            Left            =   60
            TabIndex        =   17
            Top             =   3920
            Width           =   6075
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
            Caption         =   "frmTTSummaryCfg.frx":14A6
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTSummaryCfg.frx":14E4
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":1504
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkAutoJournalAutomated 
               Height          =   495
               Left            =   120
               TabIndex        =   19
               Top             =   660
               Width           =   5595
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
               Caption         =   "frmTTSummaryCfg.frx":1520
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1604
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1624
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkAutoJournal 
               Height          =   375
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   5595
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
               Caption         =   "frmTTSummaryCfg.frx":1640
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1720
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1740
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraRejectOptions 
            Height          =   1275
            Left            =   60
            TabIndex        =   12
            Top             =   2560
            Width           =   6075
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
            Caption         =   "frmTTSummaryCfg.frx":175C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTSummaryCfg.frx":17E2
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":1802
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtRejectTimeout 
               Height          =   285
               Left            =   4260
               TabIndex        =   15
               Top             =   585
               Width           =   495
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTSummaryCfg.frx":181E
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
               Tip             =   "frmTTSummaryCfg.frx":1846
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1866
            End
            Begin HexUniControls.ctlUniRadioXP optFlatten 
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   300
               Width           =   3915
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
               Caption         =   "frmTTSummaryCfg.frx":1882
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmTTSummaryCfg.frx":18FE
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":191E
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optNothing 
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   900
               Width           =   3915
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
               Caption         =   "frmTTSummaryCfg.frx":193A
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1986
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":19A6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optAsk 
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   600
               Width           =   5715
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
               Caption         =   "frmTTSummaryCfg.frx":19C2
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1A8C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1AAC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraDisplay 
         Height          =   6780
         Left            =   6960
         TabIndex        =   30
         Top             =   330
         Width           =   6225
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
         Caption         =   "frmTTSummaryCfg.frx":1AC8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTSummaryCfg.frx":1AFC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTSummaryCfg.frx":1B1C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraFonts 
            Height          =   915
            Left            =   60
            TabIndex        =   31
            Top             =   120
            Width           =   4275
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
            Caption         =   "frmTTSummaryCfg.frx":1B38
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTSummaryCfg.frx":1B62
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":1B82
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniButtonImageXP cmdGridFont 
               Height          =   435
               Left            =   180
               TabIndex        =   32
               Top             =   300
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
               Caption         =   "frmTTSummaryCfg.frx":1B9E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1BD2
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1BF2
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblGridFontSample 
               Height          =   495
               Left            =   1380
               Top             =   240
               Width           =   2715
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
               Caption         =   "frmTTSummaryCfg.frx":1C0E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTSummaryCfg.frx":1C4E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTSummaryCfg.frx":1C6E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkShowCents 
            Height          =   255
            Left            =   60
            TabIndex        =   34
            Top             =   1140
            Width           =   4275
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
            Caption         =   "frmTTSummaryCfg.frx":1C8A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":1D12
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":1D32
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgDisplay 
            Height          =   2595
            Left            =   60
            TabIndex        =   36
            Top             =   1680
            Width           =   4335
            _cx             =   7646
            _cy             =   4577
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
         Begin VSFlex7LCtl.VSFlexGrid fgToolbarButtons 
            Height          =   1635
            Left            =   60
            TabIndex        =   38
            Top             =   4560
            Width           =   4335
            _cx             =   7646
            _cy             =   2884
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
         Begin HexUniControls.ctlUniLabelXP lblCustomizeToolbar 
            Height          =   255
            Left            =   60
            Top             =   4320
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
            Caption         =   "frmTTSummaryCfg.frx":1D4E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":1DA4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":1DC4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSummarySettings 
            Height          =   255
            Left            =   60
            Top             =   1440
            Width           =   2415
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
            Caption         =   "frmTTSummaryCfg.frx":1DE0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTSummaryCfg.frx":1E3A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTSummaryCfg.frx":1E5A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   2040
      TabIndex        =   33
      Top             =   7440
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
      Caption         =   "frmTTSummaryCfg.frx":1E76
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTSummaryCfg.frx":1EA2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTSummaryCfg.frx":1EC2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1035
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
         Caption         =   "frmTTSummaryCfg.frx":1EDE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTSummaryCfg.frx":1F04
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTSummaryCfg.frx":1F24
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   1440
         TabIndex        =   37
         Top             =   0
         Width           =   1035
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
         Caption         =   "frmTTSummaryCfg.frx":1F40
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTSummaryCfg.frx":1F6E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTSummaryCfg.frx":1F8E
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmTTSummaryCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTSummaryCfg.frm
'' Description: Allow the user to customize the Trade Console form or set
''              trading options
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/21/2009   DAJ         Only give the wrong side of the market warning if the
''                          trade settings say to give the warning
'' 10/07/2009   DAJ         Implement price shaving for filling option orders
'' 03/11/2010   DAJ         Use global collections, broker properties
'' 03/15/2010   DAJ         Fixed the grid column information persistence
'' 09/29/2010   DAJ         Split out global order confirmation flag and added here
'' 07/27/2011   DAJ         Added customization for Trade Console toolbar buttons
'' 01/30/2012   DAJ         User configure timeout on auto exit reject
'' 06/28/2013   DAJ         Added the web export tab
'' 08/11/2014   DAJ         New flag for how to calculate open equity on options
'' 08/11/2014   DAJ         Warn user if they choose to use last for open equity on options
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Show = 0
    eGDCol_Title
    eGDCol_Level
    eGDCol_ColNumber
    eGDCol_ColWidth
    eGDCol_NumCols
End Enum

Private Enum eGDTabs
    eGDTab_Trading = 0
    eGDTab_Console
    eGDTab_WebExport
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click OK or Cancel?
    
    StopBuffer As cPriceEditor          ' Price editor control for the stop buffer ticks
End Type
Private m As mPrivate

Private Function GDCol(Col As eGDCols) As Long
    GDCol = Col
End Function

Private Function Tabs(nTab As eGDTabs) As Long
    Tabs = nTab
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     True if OK clicked, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    LoadTradingTab
    LoadConsoleTab
    LoadWebTab

    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        ' Save the options from the Trading tab...
        g.Broker.NumTicksStopBuffer = m.StopBuffer.Price
        g.Broker.DontAllowStopMove = CheckBoxValue(chkStopBuffer)
        g.Broker.ConfirmManual = CheckBoxValue(chkConfirmManual)
        g.Broker.ConfirmTriggered = CheckBoxValue(chkConfirmTriggered)
        g.Broker.ConfirmTradeSense = CheckBoxValue(chkConfirmTradeSense)
        g.Broker.WarnStopWrongSide = CheckBoxValue(chkWarnStop)
        g.Broker.WarnLimitWrongSide = CheckBoxValue(chkWarnLimit)
        Select Case True
            Case optAsk
                g.OrderStrategies.RejectOption = eGDOrderRejectOption_Ask
            Case optFlatten
                g.OrderStrategies.RejectOption = eGDOrderRejectOption_Flatten
            Case optNothing
                g.OrderStrategies.RejectOption = eGDOrderRejectOption_Nothing
        End Select
        g.OrderStrategies.RejectTimeout = CLng(ValOfText(txtRejectTimeout.Text))
        g.Broker.AutoJournalPopUp = CheckBoxValue(chkAutoJournal)
        g.Broker.AutoJournalAutomated = CheckBoxValue(chkAutoJournalAutomated)
        Select Case True
            Case optBidAsk
                g.Broker.OptionFillMethod = eGDOptionFill_BidOrAsk
            Case optMidpoint
                g.Broker.OptionFillMethod = eGDOptionFill_Midpoint
            Case optTwoThirds
                g.Broker.OptionFillMethod = eGDOptionFill_TwoThirds
        End Select
        Select Case True
            Case optOptionOpenEquityBidAsk
                g.Broker.OptionOpenEquity = eGDOptionOpenEquity_UseBidAsk
            Case optOptionOpenEquityLast
                g.Broker.OptionOpenEquity = eGDOptionOpenEquity_UseLast
        End Select
        
        ' Save the options from the Console tab...
        g.Broker.ShowCents = CheckBoxValue(chkShowCents)
        g.Broker.GridFont = FontToString(lblGridFontSample.Font)
        
        SaveDisplay
        SaveToolbarSettings
        SaveWebSettings
        
        frmTTSummary.UpdateConsoleSettings
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTSummaryCfg.ShowMe"
    Resume 'RH
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Hide the form and let ShowMe unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdGridFont_Click
'' Description: Allow the user to change the font on the grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdGridFont_Click()
On Error GoTo ErrSection:

    CommonDialogFont frmMain.CommonDialog1, lblGridFontSample.Font

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.cmdGridFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Hide the form and let ShowMe save the information and unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim lRejectTimeout As Long          ' Reject timeout value

    MoveFocus cmdOK
    
    lRejectTimeout = CLng(ValOfText(txtRejectTimeout.Text))
    If (lRejectTimeout > 9999&) Or (lRejectTimeout < 0&) Then
        MoveFocus txtRejectTimeout
        InfBox "Timeout must be between 0 and 9999 seconds", "!", , "Error"
    Else
        m.bOK = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.cmdOK_Click"
    
End Sub

Private Sub cmdWebSample_Click()
On Error GoTo ErrSection:
       
    If Not CanDoAcctStatusWebPage Then
        InfBox "NOTE: You are not currently enabled for this feature.", "!", , "Web Status Report"
    Else
        cmdWebSample.Enabled = False
        SaveWebSettings
        If frmQuotes.AcctStatusCheck(True) Then
            Sleep 1
            RunProcess InternetBrowser, Chr(34) & Trim(lblUrl.Caption) & Chr(34)
        ElseIf Len(g.Broker.WebAccounts) = 0 Then
            InfBox "No accounts have been selected.", "e", , "Web Status Report"
        Else
            InfBox "The report cannot be run.", "e", , "Web Status Report"
        End If
        cmdWebSample.Enabled = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.cmdWebSample_Click"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDisplay_AfterEdit
'' Description: Display subitems if appropriate
'' Inputs:      Row and Column of the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDisplay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.fgDisplay_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDisplay_BeforeEdit
'' Description: Only allow the user to edit the Show column
'' Inputs:      Row and Column of the Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDisplay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Show) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.fgDisplay_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgToolbarButtons_BeforeEdit
'' Description: Only allow the user to edit the Show column
'' Inputs:      Row and Column of the Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgToolbarButtons_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Show) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.fgToolbarButtons_BeforeEdit"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgToolbarButtons_ValidateEdit
'' Description: Validate whether the user can hide the button
'' Inputs:      Row and Column of the Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgToolbarButtons_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nButton As eGDTcButtons         ' Toolbar button in question

    With fgToolbarButtons
        nButton = .RowData(Row)
        
        If CheckedCell(fgToolbarButtons, Row, Col) = True Then
            If (nButton = eGDTcButtons_AutoTrading) Then
                If g.ConsoleForms.NumVisible(eGDConsoleForm_AutoTrading) > 0 Then
                    InfBox "You cannot hide the Automated Trading Items button because you have Automated Trading Items", "!", , "Trade Console"
                    Cancel = True
                End If
            ElseIf (nButton = eGDTcButtons_TradeSenseOrders) Then
                If g.ConsoleForms.NumVisible(eGDConsoleForm_TradeSenseOrders) > 0 Then
                    InfBox "You cannot hide the TradeSense Orders button because you have TradeSense Orders", "!", , "Trade Console"
                    Cancel = True
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.fgDisplay_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Do some initialization when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    
    g.Styler.StyleForm Me
    
    Caption = "Trade Settings"
    CenterTheForm Me
    cmdCancel.Cancel = True
    
    lblUrl.Caption = "www.TradeNavigator.com/Clients/" & RI_GetMachineID & ".htm"
    tabSettings.TabVisible(Tabs(eGDTab_WebExport)) = CanDoAcctStatusWebPage
    
    tabSettings.CurrTab = Tabs(eGDTab_Trading)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Hide the form and let ShowMe save the information and unload
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.Form_QueryUnload"
    
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

    With fraButtons
        .Move (ScaleWidth - .Width) / 2
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optOptionOpenEquityLast_Click
'' Description: Warn the user when they select this option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optOptionOpenEquityLast_Click()
On Error GoTo ErrSection:

    If Visible Then
        InfBox "We recommend that you use the bid or ask|to calculate open equity for options because|the bid and ask are usually much more|recent than the last trade", "i", , "Options Open Equity"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.optOptionOpenEquityLast_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtRejectTimeout_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRejectTimeout_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtRejectTimeout

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.txtRejectTimeout_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitFieldsGrid
'' Description: Initialize the fields grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitFieldsGrid()
On Error GoTo ErrSection:

    With fgDisplay
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Show)) = "Show"
        .TextMatrix(0, GDCol(eGDCol_Title)) = "Header"
        .TextMatrix(0, GDCol(eGDCol_Level)) = "Level"
        .TextMatrix(0, GDCol(eGDCol_ColNumber)) = "Column Number"
        .TextMatrix(0, GDCol(eGDCol_ColWidth)) = "Column Width"
        
        .ColHidden(GDCol(eGDCol_Level)) = True
        .ColHidden(GDCol(eGDCol_ColNumber)) = True
        .ColHidden(GDCol(eGDCol_ColWidth)) = True
                
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.InitFieldsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFieldsGrid
'' Description: Load the display grid
'' Inputs:      String with display information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFieldsGrid()
On Error GoTo ErrSection:

    With fgDisplay
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        LoadDisplayForGrid "Open Orders", GetIniFileProperty("SummaryOrdersDisplay", "", "TTSummary", g.strIniFile)
        LoadDisplayForGrid "Open Positions", GetIniFileProperty("SummaryPositionsDisplay", "", "TTSummary", g.strIniFile)
        LoadDisplayForGrid "Accounts", GetIniFileProperty("SummaryAccountsDisplay", "", "TTSummary", g.strIniFile)
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadFieldsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadConsoleTab
'' Description: Load the controls for the console tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadDisplayForGrid(ByVal strGridName As String, ByVal strDisplay As String)
On Error GoTo ErrSection:

    Dim GridColumns As New cGridColumns ' Collection of grid column objects
    Dim GridColumn As cGridColumn       ' Grid column object
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As Long                 ' State of the grids redraw

    GridColumns.FromString strDisplay

    With fgDisplay
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Header row for the grid...
        .Rows = .Rows + 1
        .RowData(.Rows - 1) = ""
        .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Show)) = flexNoCheckbox
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, GDCol(eGDCol_NumCols) - 1) = ALT_GRID_ROW_COLOR
        .Cell(flexcpFontBold, .Rows - 1, GDCol(eGDCol_Title)) = True
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Title)) = strGridName
        
        For lIndex = 1 To GridColumns.Count
            .Rows = .Rows + 1
            .RowData(.Rows - 1) = strGridName
            
            Set GridColumn = GridColumns.Item(lIndex)
            CheckedCell(fgDisplay, .Rows - 1, GDCol(eGDCol_Show)) = GridColumn.Visible
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Title)) = "     " & GridColumn.Name
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Level)) = Str(GridColumn.Level)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ColNumber)) = Str(GridColumn.Position)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ColWidth)) = Str(GridColumn.Width)
            .RowHidden(.Rows - 1) = Not GridColumn.ShowInSettings
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadDisplayForGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveDisplay
'' Description: Save the grid display strings from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveDisplay()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOrders As New cGdArray      ' Array of grid display information for orders
    Dim astrPositions As New cGdArray   ' Array of grid display information for positions
    Dim astrAccounts As New cGdArray    ' Array of grid display information for accounts
    Dim GridColumn As New cGridColumn   ' Grid column object
    
    astrOrders.Create eGDARRAY_Strings
    astrPositions.Create eGDARRAY_Strings
    astrAccounts.Create eGDARRAY_Strings
    
    With fgDisplay
        For lIndex = .FixedRows To .Rows - 1
            If Len(.RowData(lIndex)) > 0 Then
                GridColumn.Visible = CheckedCell(fgDisplay, lIndex, GDCol(eGDCol_Show))
                GridColumn.Name = Trim(.TextMatrix(lIndex, GDCol(eGDCol_Title)))
                GridColumn.Level = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Level))))
                GridColumn.Position = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_ColNumber))))
                GridColumn.Width = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_ColWidth))))
                GridColumn.ShowInSettings = Not .RowHidden(lIndex)
                
                Select Case UCase(.RowData(lIndex))
                    Case "OPEN ORDERS"
                        astrOrders.Add GridColumn.ToString
                    Case "OPEN POSITIONS"
                        astrPositions.Add GridColumn.ToString
                    Case "ACCOUNTS"
                        astrAccounts.Add GridColumn.ToString
                End Select
            End If
        Next lIndex
    End With
    
    SetIniFileProperty "SummaryOrdersDisplay", "1|" & astrOrders.JoinFields(","), "TTSummary", g.strIniFile
    SetIniFileProperty "SummaryPositionsDisplay", "1|" & astrPositions.JoinFields(","), "TTSummary", g.strIniFile
    SetIniFileProperty "SummaryAccountsDisplay", "1|" & astrAccounts.JoinFields(","), "TTSummary", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.SaveDisplay"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTradingTab
'' Description: Load the controls for the trading tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTradingTab()
On Error GoTo ErrSection:

    Dim nRejectOption As eGDOrderRejectOption   ' Reject option from the INI file

    nRejectOption = g.OrderStrategies.RejectOption
    Select Case nRejectOption
        Case eGDOrderRejectOption_Flatten
            optFlatten.Value = True
            optAsk.Value = False
            optNothing.Value = False
            
        Case eGDOrderRejectOption_Ask
            optFlatten.Value = False
            optAsk.Value = True
            optNothing.Value = False
            
        Case eGDOrderRejectOption_Nothing
            optFlatten.Value = False
            optAsk.Value = False
            optNothing.Value = True
    
    End Select
    txtRejectTimeout.Text = Str(g.OrderStrategies.RejectTimeout)
    
    Set m.StopBuffer = New cPriceEditor
    m.StopBuffer.Init sbStopBuffer, txtStopBuffer, Nothing, g.Broker.NumTicksStopBuffer, 1
    
    CheckBoxValue(chkStopBuffer) = g.Broker.DontAllowStopMove
    CheckBoxValue(chkConfirmManual) = g.Broker.ConfirmManual
    CheckBoxValue(chkConfirmTriggered) = g.Broker.ConfirmTriggered
    CheckBoxValue(chkConfirmTradeSense) = g.Broker.ConfirmTradeSense
    CheckBoxValue(chkWarnStop) = g.Broker.WarnStopWrongSide
    CheckBoxValue(chkWarnLimit) = g.Broker.WarnLimitWrongSide
    CheckBoxValue(chkAutoJournal) = g.Broker.AutoJournalPopUp
    CheckBoxValue(chkAutoJournalAutomated) = g.Broker.AutoJournalAutomated
    
    Select Case g.Broker.OptionFillMethod
        Case eGDOptionFill_BidOrAsk
            optBidAsk.Value = True
            optMidpoint.Value = False
            optTwoThirds.Value = False
        
        Case eGDOptionFill_Midpoint
            optBidAsk.Value = False
            optMidpoint.Value = True
            optTwoThirds.Value = False
        
        Case eGDOptionFill_TwoThirds
            optBidAsk.Value = False
            optMidpoint.Value = False
            optTwoThirds.Value = True
    
    End Select
    
    Select Case g.Broker.OptionOpenEquity
        Case eGDOptionOpenEquity_UseBidAsk
            optOptionOpenEquityBidAsk.Value = True
            optOptionOpenEquityLast.Value = False
            
        Case eGDOptionOpenEquity_UseLast
            optOptionOpenEquityBidAsk.Value = False
            optOptionOpenEquityLast.Value = True
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadTradingTab"
Resume 'RH
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadConsoleTab
'' Description: Load the controls for the console tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadConsoleTab()
On Error GoTo ErrSection:

    FontFromString lblGridFontSample.Font, g.Broker.GridFont
    CheckBoxValue(chkShowCents) = g.Broker.ShowCents
    
    InitFieldsGrid
    LoadFieldsGrid
    
    InitToolbarGrid
    LoadToolbarGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadConsoleTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadWebTab
'' Description: Load the controls for the web tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadWebTab()
On Error GoTo ErrSection:

    InitAccountsGrid
    LoadAccountsGrid
    
    sliFontSize.Min = 1
    sliFontSize.Max = 7
    sliFontSize.Value = g.Broker.WebFontSize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadWebTab"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitToolbarGrid
'' Description: Initialize the toolbar grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitToolbarGrid()
On Error GoTo ErrSection:

    With fgToolbarButtons
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_Title) + 1
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Show)) = "Show"
        .TextMatrix(0, GDCol(eGDCol_Title)) = "Header"
                
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.InitToolbarGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadToolbarGrid
'' Description: Load the toolbar grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadToolbarGrid()
On Error GoTo ErrSection:

    Dim TcBtns As cTradeConsoleButtons  ' Trade console button information
    Dim lIndex As Long                  ' Index into a for loop
    
    Set TcBtns = New cTradeConsoleButtons
    TcBtns.Load
    
    With fgToolbarButtons
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 0 To TcBtns.Size - 1
            If TcBtns.ShowInLists(lIndex) = True Then
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = lIndex
                CheckedCell(fgToolbarButtons, .Rows - 1, GDCol(eGDCol_Show)) = TcBtns.Show(lIndex)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Title)) = TcBtns.Name(lIndex)
            End If
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadToolbarGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveToolbarSettings
'' Description: Save the toolbar button settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveToolbarSettings()
On Error GoTo ErrSection:

    Dim TcBtns As cTradeConsoleButtons  ' Trade console button information
    Dim lIndex As Long                  ' Index into a for loop
    
    Set TcBtns = New cTradeConsoleButtons
    TcBtns.Load
    
    With fgToolbarButtons
        For lIndex = .FixedRows To .Rows - 1
            TcBtns.Show(.RowData(lIndex)) = CheckedCell(fgToolbarButtons, lIndex, GDCol(eGDCol_Show))
        Next lIndex
    End With
    
    TcBtns.Save

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.SaveToolbarSettings"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitAccountsGrid
'' Description: Initialize the accounts grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitAccountsGrid()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If g.nColorTheme = kDarkThemeColor Then .BackColor = g.nColorTheme
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .FixedRows = 1
        .Rows = .FixedRows
        .FixedCols = 0
        .Cols = 3
        
        .ColDataType(0) = flexDTBoolean
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        
        .TextMatrix(0, 0) = "Export"
        .TextMatrix(0, 1) = "Account"
        .TextMatrix(0, 2) = "Type"
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.InitAccountsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsGrid
'' Description: Load the accounts grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAccountsGrid()
On Error GoTo ErrSection:

    Dim Accounts As cPtAccounts         ' Accounts collection
    Dim Account As cPtAccount           ' Account object
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strAccounts As String           ' String of accounts to turn on
    
    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        strAccounts = "," & g.Broker.WebAccounts & ","
        
        .Rows = .FixedRows
        If CanDoAcctStatusWebPage Then
            Set Accounts = g.Broker.SnapshotAccounts
            For lIndex = 1 To Accounts.Count
                Set Account = Accounts(lIndex)
                If Account.AccountType <> eTT_AccountType_SimReplay Then
                    .Rows = .Rows + 1
                    
                    .RowData(.Rows - 1) = Account
                    .TextMatrix(.Rows - 1, 1) = Account.Name
                    .TextMatrix(.Rows - 1, 2) = g.Broker.BrokerName(Account.AccountType)
                    
                    CheckedCell(fgAccounts, .Rows - 1, 0) = (InStr(strAccounts, "," & Account.AccountNumber & ",") <> 0)
                End If
            Next lIndex
        End If
    
        SetBackColors fgAccounts
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.LoadAccountsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lTotal As Long                  ' Total of the column widths
    Dim lExtCol As Long                 ' Column to extend
    
    lExtCol = 1
    
    With fgAccounts
        ' size the custom extended column in order to fill the client width
        .ColHidden(lExtCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        
        lTotal = 0
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal <> .ColWidth(lExtCol) Then
            .ColWidth(lExtCol) = lTotal
        End If
        .ColHidden(lExtCol) = False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
       
ErrSection:
    RaiseError "frmTTSummaryCfg.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveWebSettings
'' Description: Save the settings off of the web tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveWebSettings()
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' List of accounts that are turned on
    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cPtAccount           ' Account object
    
    Set astrAccounts = New cGdArray
    astrAccounts.Create eGDARRAY_Strings
    
    With fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            If CheckedCell(fgAccounts, lIndex, 0) = True Then
                If TypeOf .RowData(lIndex) Is cPtAccount Then
                    Set Account = .RowData(lIndex)
                    astrAccounts.Add Account.AccountNumber
                End If
            End If
        Next lIndex
    End With
    
    g.Broker.WebAccounts = astrAccounts.JoinFields(",")
    g.Broker.WebFontSize = sliFontSize.Value

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTSummaryCfg.SaveWebSettings"
    
End Sub

