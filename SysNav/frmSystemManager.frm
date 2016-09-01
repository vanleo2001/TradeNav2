VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSystemManager 
   Caption         =   "Strategy Manager"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   8700
   Begin vsOcx6LibCtl.vsElastic vsLinked 
      Height          =   315
      Left            =   3840
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
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
      ForeColor       =   12582912
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   "(using data from active chart)"
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
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
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfTest 
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmSystemManager.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Tip             =   "frmSystemManager.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSystemManager.frx":0040
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   6600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   2
      ToolsCount      =   17
      DisplayContextMenu=   0   'False
      Tools           =   "frmSystemManager.frx":005C
      ToolBars        =   "frmSystemManager.frx":2ADF
   End
   Begin VB.Timer tmrOptimizing 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7500
      Top             =   60
   End
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   6735
      Left            =   105
      TabIndex        =   53
      Top             =   120
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   11880
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
      Caption         =   "&Rules|&Inputs|&Data|Sett&ings"
      Align           =   0
      Appearance      =   1
      CurrTab         =   3
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
      Begin HexUniControls.ctlUniFrameWL fraRules 
         Height          =   6360
         Left            =   -9615
         TabIndex        =   31
         Top             =   330
         Width           =   8370
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
         Caption         =   "frmSystemManager.frx":2EB8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSystemManager.frx":2EE4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSystemManager.frx":2F04
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraSignal 
            Height          =   525
            Left            =   180
            TabIndex        =   40
            Top             =   120
            Width           =   6555
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
            Caption         =   "frmSystemManager.frx":2F20
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":2F52
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":2F72
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optSignals 
               Height          =   210
               Index           =   4
               Left            =   240
               TabIndex        =   0
               Tag             =   "5"
               Top             =   210
               Width           =   1005
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
               Caption         =   "frmSystemManager.frx":2F8E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmSystemManager.frx":2FC0
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":2FE0
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSignals 
               Height          =   210
               Index           =   2
               Left            =   3840
               TabIndex        =   3
               Tag             =   "3"
               Top             =   210
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
               Caption         =   "frmSystemManager.frx":2FFC
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3036
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3056
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSignals 
               Height          =   210
               Index           =   1
               Left            =   2640
               TabIndex        =   2
               Tag             =   "2"
               Top             =   210
               Width           =   1065
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
               Caption         =   "frmSystemManager.frx":3072
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":30A6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":30C6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSignals 
               Height          =   210
               Index           =   0
               Left            =   1320
               TabIndex        =   1
               Tag             =   "1"
               Top             =   210
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
               Caption         =   "frmSystemManager.frx":30E2
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":311A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":313A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSignals 
               Height          =   210
               Index           =   3
               Left            =   5160
               TabIndex        =   4
               Tag             =   "4"
               Top             =   210
               Width           =   1110
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
               Caption         =   "frmSystemManager.frx":3156
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":318C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":31AC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraRuleButtons 
            Height          =   3495
            Left            =   6900
            TabIndex        =   42
            Top             =   360
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
            Caption         =   "frmSystemManager.frx":31C8
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":31F4
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":3214
            RightToLeft     =   0   'False
            Begin vsOcx6LibCtl.vsElastic lblNewRule 
               Height          =   195
               Left            =   180
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   120
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   344
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
               ForeColor       =   -2147483640
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "Ne&w Rule"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
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
            Begin HexUniControls.ctlUniButtonImageXP cmdNewRule 
               Height          =   435
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Width           =   1185
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
               Caption         =   "frmSystemManager.frx":3230
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3254
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3274
               RightToLeft     =   0   'False
            End
            Begin vsOcx6LibCtl.vsElastic lblQuickStops 
               Height          =   180
               Left            =   120
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   2220
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   318
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
               ForeColor       =   -2147483640
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "&Quick Stops"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
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
            Begin vsOcx6LibCtl.vsElastic lblAddRule 
               Height          =   225
               Left            =   120
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   600
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   397
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
               ForeColor       =   -2147483640
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "&Add Rule"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
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
            Begin HexUniControls.ctlUniButtonImageXP cmdLinkToEntry 
               Height          =   435
               Left            =   0
               TabIndex        =   14
               Top             =   2580
               Width           =   1185
               _ExtentX        =   0
               _ExtentY        =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmSystemManager.frx":3290
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":32CC
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":334C
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdPyramidInfo 
               Height          =   435
               Left            =   0
               TabIndex        =   16
               Top             =   3060
               Width           =   1185
               _ExtentX        =   0
               _ExtentY        =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmSystemManager.frx":3368
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":33A2
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":33C2
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdRemoveRule 
               Height          =   435
               Left            =   0
               TabIndex        =   11
               Top             =   1440
               Width           =   1185
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
               Caption         =   "frmSystemManager.frx":33DE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3416
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3466
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdEditRule 
               Height          =   435
               Left            =   0
               TabIndex        =   10
               Top             =   960
               Width           =   1185
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
               Caption         =   "frmSystemManager.frx":3482
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":34B6
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3502
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdAddRule 
               Height          =   435
               Left            =   0
               TabIndex        =   9
               Top             =   480
               Width           =   1185
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
               Caption         =   "frmSystemManager.frx":351E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3542
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3562
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdQuickStops 
               Height          =   435
               Left            =   0
               TabIndex        =   13
               Top             =   2100
               Width           =   1185
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
               Caption         =   "frmSystemManager.frx":357E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":35A2
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":35C2
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdTestEntry 
               Height          =   435
               Left            =   0
               TabIndex        =   15
               Top             =   1920
               Width           =   1185
               _ExtentX        =   0
               _ExtentY        =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmSystemManager.frx":35DE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3614
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":36A8
               RightToLeft     =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid vsRules 
            Height          =   3105
            Left            =   180
            TabIndex        =   5
            Top             =   780
            Width           =   6570
            _cx             =   11589
            _cy             =   5477
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
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
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
         Begin HexUniControls.ctlUniRichTextBoxXP txtPreview 
            Height          =   1230
            Left            =   180
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   4020
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   2170
            BackColor       =   12632256
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frmSystemManager.frx":36C4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Tip             =   "frmSystemManager.frx":36E4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":3704
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
      End
      Begin HexUniControls.ctlUniFrameWL fraInputs 
         Height          =   6360
         Left            =   -9315
         TabIndex        =   43
         Top             =   330
         Width           =   8370
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
         Caption         =   "frmSystemManager.frx":3720
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSystemManager.frx":374C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSystemManager.frx":376C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraLinkInputs 
            Height          =   435
            Left            =   60
            TabIndex        =   45
            Top             =   0
            Width           =   9735
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
            Caption         =   "frmSystemManager.frx":3788
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":37B4
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":37D4
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkLinkInputs 
               Height          =   255
               Left            =   180
               TabIndex        =   18
               Top             =   120
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
               Caption         =   "frmSystemManager.frx":37F0
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3878
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3898
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid vsInputs 
            Height          =   4575
            Left            =   0
            TabIndex        =   19
            Top             =   480
            Width           =   9840
            _cx             =   17357
            _cy             =   8070
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
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
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
      Begin HexUniControls.ctlUniFrameWL fraData 
         Height          =   6360
         Left            =   -9015
         TabIndex        =   55
         Top             =   330
         Width           =   8370
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
         Caption         =   "frmSystemManager.frx":38B4
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSystemManager.frx":38E0
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSystemManager.frx":3900
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraDataSettings 
            Height          =   4935
            Left            =   5580
            TabIndex        =   56
            Top             =   120
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
            Caption         =   "frmSystemManager.frx":391C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":3948
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":3968
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkLinkToChart 
               Height          =   375
               Left            =   60
               TabIndex        =   46
               Top             =   60
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
               Caption         =   "frmSystemManager.frx":3984
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3A14
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3A34
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
               Height          =   435
               Left            =   1500
               TabIndex        =   28
               Top             =   660
               Width           =   1065
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
               Caption         =   "frmSystemManager.frx":3A50
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3A88
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3AA8
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniComboImageXP cboBarType 
               Height          =   315
               Left            =   345
               TabIndex        =   30
               Top             =   4560
               Visible         =   0   'False
               Width           =   1845
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
               Tip             =   "frmSystemManager.frx":3AC4
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3AE4
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniFrameWL fraToDate 
               Height          =   2070
               Left            =   0
               TabIndex        =   57
               Top             =   1380
               Width           =   2610
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
               Caption         =   "frmSystemManager.frx":3B00
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmSystemManager.frx":3B44
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3B64
               RightToLeft     =   0   'False
               Begin gdOCX.gdSelectDate dtpToDate 
                  Height          =   330
                  Left            =   420
                  TabIndex        =   24
                  Top             =   1245
                  Width           =   2115
                  _ExtentX        =   3731
                  _ExtentY        =   582
               End
               Begin HexUniControls.ctlUniRadioXP optToDate 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   23
                  Top             =   1260
                  Width           =   2265
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
                  Caption         =   "frmSystemManager.frx":3B80
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmSystemManager.frx":3BB2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmSystemManager.frx":3BD2
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optToEndOfData 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   25
                  Top             =   1620
                  Width           =   1635
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
                  Caption         =   "frmSystemManager.frx":3BEE
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmSystemManager.frx":3C24
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmSystemManager.frx":3C44
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin gdOCX.gdSelectDate dtpFromDate 
                  Height          =   330
                  Left            =   420
                  TabIndex        =   22
                  Top             =   540
                  Width           =   2115
                  _ExtentX        =   3731
                  _ExtentY        =   582
               End
               Begin HexUniControls.ctlUniLabelXP lblFromDate 
                  Height          =   255
                  Left            =   120
                  Top             =   300
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
                  Caption         =   "frmSystemManager.frx":3C60
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmSystemManager.frx":3C8A
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmSystemManager.frx":3CAA
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblToDate 
                  Height          =   255
                  Left            =   180
                  Top             =   975
                  Width           =   555
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
                  Caption         =   "frmSystemManager.frx":3CC6
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmSystemManager.frx":3CEC
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmSystemManager.frx":3D0C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin vsOcx6LibCtl.vsElastic lblBrowse 
               Height          =   225
               Left            =   120
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   780
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   397
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
               ForeColor       =   -2147483640
               FloodColor      =   192
               ForeColorDisabled=   -2147483631
               Caption         =   "Symbol &LookUp"
               Align           =   0
               Appearance      =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
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
            Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
               Height          =   435
               Left            =   0
               TabIndex        =   27
               Top             =   660
               Width           =   1365
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
               Caption         =   "frmSystemManager.frx":3D28
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":3D4C
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3D6C
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblBarType 
               Height          =   210
               Left            =   360
               Top             =   4320
               Visible         =   0   'False
               Width           =   1245
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
               Caption         =   "frmSystemManager.frx":3D88
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":3DC8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3DE8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid vsMarkets 
            Height          =   4950
            Left            =   180
            TabIndex        =   20
            Top             =   120
            Width           =   5190
            _cx             =   9155
            _cy             =   8731
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
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
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
      Begin HexUniControls.ctlUniFrameWL fraSettings 
         Height          =   6360
         Left            =   45
         TabIndex        =   54
         Top             =   330
         Width           =   8370
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
         Caption         =   "frmSystemManager.frx":3E04
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSystemManager.frx":3E30
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSystemManager.frx":3E50
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL Frame2 
            Height          =   1455
            Left            =   4380
            TabIndex        =   50
            Top             =   180
            Width           =   3615
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
            Caption         =   "frmSystemManager.frx":3E6C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":3EB6
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":3ED6
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtDollarsPerTrade 
               Height          =   315
               Left            =   1800
               TabIndex        =   37
               Top             =   600
               Width           =   1095
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmSystemManager.frx":3EF2
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
               Tip             =   "frmSystemManager.frx":3F20
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3F40
            End
            Begin HexUniControls.ctlUniTextBoxXP txtNumShares 
               Height          =   315
               Left            =   1800
               TabIndex        =   38
               Top             =   990
               Width           =   735
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmSystemManager.frx":3F5C
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
               Tip             =   "frmSystemManager.frx":3F82
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":3FA2
            End
            Begin HexUniControls.ctlUniRadioXP optDollarsPerTrade 
               Height          =   220
               Left            =   240
               TabIndex        =   35
               Top             =   660
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   397
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmSystemManager.frx":3FBE
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmSystemManager.frx":4000
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4020
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optSharesPerTrade 
               Height          =   220
               Left            =   240
               TabIndex        =   36
               Top             =   1020
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   397
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmSystemManager.frx":403C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmSystemManager.frx":407C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":409C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblNumShares 
               Height          =   255
               Left            =   180
               Top             =   300
               Width           =   3270
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
               Caption         =   "frmSystemManager.frx":40B8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":4132
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4152
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL Frame1 
            Height          =   1455
            Left            =   120
            TabIndex        =   51
            Top             =   180
            Width           =   3975
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
            Caption         =   "frmSystemManager.frx":416E
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":41DE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":41FE
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtForexCommission 
               Height          =   315
               Left            =   1140
               TabIndex        =   33
               Top             =   660
               Width           =   795
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmSystemManager.frx":421A
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
               Tip             =   "frmSystemManager.frx":423E
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":425E
            End
            Begin HexUniControls.ctlUniTextBoxXP txtCommission 
               Height          =   315
               Left            =   1140
               TabIndex        =   32
               Top             =   300
               Width           =   795
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmSystemManager.frx":427A
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
               Tip             =   "frmSystemManager.frx":42A0
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":42C0
            End
            Begin HexUniControls.ctlUniTextBoxXP txtStockCommission 
               Height          =   315
               Left            =   1140
               TabIndex        =   34
               Top             =   1020
               Width           =   795
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmSystemManager.frx":42DC
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
               Tip             =   "frmSystemManager.frx":4302
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4322
            End
            Begin HexUniControls.ctlUniLabelXP Label9 
               Height          =   255
               Left            =   2040
               Top             =   720
               Width           =   1605
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
               Caption         =   "frmSystemManager.frx":433E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":4386
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":43A6
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label8 
               Height          =   255
               Left            =   240
               Top             =   720
               Width           =   870
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
               Caption         =   "frmSystemManager.frx":43C2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":43EE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":440E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label2 
               Height          =   255
               Left            =   2040
               Top             =   1080
               Width           =   1305
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
               Caption         =   "frmSystemManager.frx":442A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":4468
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4488
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label1 
               Height          =   255
               Left            =   240
               Top             =   1080
               Width           =   870
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
               Caption         =   "frmSystemManager.frx":44A4
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":44D2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":44F2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblFutureFees 
               Height          =   255
               Left            =   240
               Top             =   360
               Width           =   870
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
               Caption         =   "frmSystemManager.frx":450E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":453E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":455E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblStockFees 
               Height          =   255
               Left            =   2040
               Top             =   360
               Width           =   1830
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
               Caption         =   "frmSystemManager.frx":457A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":45C8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":45E8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkForceLimitThrough 
            Height          =   220
            Left            =   120
            TabIndex        =   41
            Top             =   2640
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmSystemManager.frx":4604
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSystemManager.frx":46F2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4712
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAllowReverse 
            Height          =   220
            Left            =   120
            TabIndex        =   39
            Top             =   2040
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmSystemManager.frx":472E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmSystemManager.frx":47A0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":47C0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraPyramiding 
            Height          =   1455
            Left            =   120
            TabIndex        =   48
            Top             =   4800
            Width           =   7905
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
            Caption         =   "frmSystemManager.frx":47DC
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmSystemManager.frx":4810
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4830
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtTradeDepth 
               Height          =   315
               Left            =   3660
               TabIndex        =   52
               Top             =   960
               Width           =   795
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frmSystemManager.frx":484C
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
               Tip             =   "frmSystemManager.frx":4870
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4890
            End
            Begin HexUniControls.ctlUniCheckXP chkPyramid 
               Height          =   255
               Left            =   300
               TabIndex        =   49
               Top             =   360
               Width           =   6090
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
               Caption         =   "frmSystemManager.frx":48AC
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmSystemManager.frx":4938
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4958
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblAllowPyramiding2 
               Height          =   270
               Left            =   600
               Top             =   600
               Width           =   6345
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
               Caption         =   "frmSystemManager.frx":4974
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":4A38
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4A58
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblMaxOpenPositions 
               Height          =   270
               Left            =   300
               Top             =   990
               Width           =   6345
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
               Caption         =   "frmSystemManager.frx":4A74
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmSystemManager.frx":4AEC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmSystemManager.frx":4B0C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBarsTradedBeforeOrders 
            Height          =   315
            Left            =   7065
            TabIndex        =   47
            Top             =   4125
            Width           =   795
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSystemManager.frx":4B28
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
            Tip             =   "frmSystemManager.frx":4B4C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4B6C
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBarsLoadedBeforeTrading 
            Height          =   315
            Left            =   7065
            TabIndex        =   44
            Top             =   3285
            Width           =   795
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmSystemManager.frx":4B88
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
            Tip             =   "frmSystemManager.frx":4BAC
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4BCC
         End
         Begin HexUniControls.ctlUniLabelXP lblForceLimitThrough 
            Height          =   240
            Left            =   120
            Top             =   2400
            Width           =   3090
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
            Caption         =   "frmSystemManager.frx":4BE8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSystemManager.frx":4C2A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4C4A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblReverse 
            Height          =   240
            Left            =   120
            Top             =   1800
            Width           =   2250
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
            Caption         =   "frmSystemManager.frx":4C66
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSystemManager.frx":4CB6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4CD6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblEstimatedLength 
            Height          =   600
            Left            =   120
            Top             =   4080
            Width           =   6345
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
            Caption         =   "frmSystemManager.frx":4CF2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSystemManager.frx":4EE2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4F02
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblEstimatedLengthTitle 
            Height          =   240
            Left            =   120
            Top             =   3840
            Width           =   3270
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
            Caption         =   "frmSystemManager.frx":4F1E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSystemManager.frx":4F80
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":4FA0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBarsRequiredTitle 
            Height          =   240
            Left            =   120
            Top             =   3015
            Width           =   2250
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
            Caption         =   "frmSystemManager.frx":4FBC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSystemManager.frx":5006
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":5026
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBarsRequired 
            Height          =   540
            Left            =   120
            Top             =   3240
            Width           =   6390
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
            Caption         =   "frmSystemManager.frx":5042
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSystemManager.frx":51B0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSystemManager.frx":51D0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin VB.Menu mnuRule 
      Caption         =   "Rule"
      Begin VB.Menu mnuNewRule 
         Caption         =   "Ne&w Rule"
      End
      Begin VB.Menu mnuAddRule 
         Caption         =   "&Add Rule"
      End
      Begin VB.Menu mnuEditRule 
         Caption         =   "&Edit Rule"
      End
      Begin VB.Menu mnuRemoveRule 
         Caption         =   "Remo&ve Rule"
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   "Add Rule to &Favorites"
      End
      Begin VB.Menu mnuQuickStops 
         Caption         =   "Add &Quick Stops"
      End
      Begin VB.Menu mnuLinkEntry 
         Caption         =   "&Link to Entry"
      End
      Begin VB.Menu mnuTestEntry 
         Caption         =   "&Test Entry"
      End
      Begin VB.Menu mnuEditPyramid 
         Caption         =   "Edit P&yramiding Info"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFontRules 
         Caption         =   "&Change Font"
      End
   End
   Begin VB.Menu mnuInputs 
      Caption         =   "Inputs"
      Begin VB.Menu mnuChangeFontInputs 
         Caption         =   "&Change Font"
      End
   End
   Begin VB.Menu mnuMarkets 
      Caption         =   "Markets"
      Begin VB.Menu mnuChangeSymbol 
         Caption         =   "Change &Symbol"
      End
      Begin VB.Menu mnuMarketInfo 
         Caption         =   "&Market Information"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFontMarkets 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmSystemManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSystemManager.frm
'' Description: Allow the user to modify a strategy
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 07/13/2009   DAJ         Change ExitOnEntryBar column to a boolean
'' 03/11/2010   DAJ         Use global Trading Items collection
'' 11/03/2011   DAJ         Start temporary Rule ID at -3 instead of -1
'' 02/29/2012   DAJ         Warn user if turning off the ForceLimitThrough flag
'' 05/01/2013   DAJ         Shadow Trading
'' 05/07/2014   DAJ         Allow FractZen bars for backtesting
'' 06/12/2015   DAJ         Don't add symbols to stream during run if more than kSN_BASKETLIMIT symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Text
Option Explicit

Private Const kErrorCaption = "Strategy Error"

Private Type mPrivate
    System As cSystem
    bShowPassword As Boolean
    bHideDuplicates As Boolean
    bHasDuplicates As Boolean
    bOptimizing As Boolean
    bModal As Boolean
    strName As String
    
    lNextRuleID As Long
    lNextParmID As Long
    lNumEntries As Long
    
    strSaveSymbol As String
    strCondBuilderExpr As String
    bStop As Boolean
    bNewStrategy As Boolean
    bLinkedToChart As Boolean
    bSaveLinkFlag As Boolean
    nWizardStep As Long
    nWizardStart As Long        'step number that back button cannot go past
    nButtonTextColor As Long    'original color save to restore
    strSaveMarket1 As String    'string to hold original market1 info when linking to chart
    
    'values to use when saving system
    dFromDate As Double
    dToDate As Double
    bToEndOfData As Boolean
End Type
Private m As mPrivate

Private Enum eSMTabs
    eSMTab_Rules = 0
    eSMTab_Inputs = 1
    eSMTab_Markets = 2
    eSMTab_Settings = 3
End Enum

Private Enum eRGCols
    eRGCol_Selected = 0
    eRGCol_Alt
    eRGCol_RuleName
    eRGCol_Action
    eRGCol_BuySell
    eRGCol_Late
    eRGCol_RuleID
    eRGCol_RuleType
    eRGCol_Preview
    eRGCol_Sort
    eRGCol_SecurityLevel
    eRGCol_Password
    eRGCol_Linked
    eRGCol_LinkedRules
    eRGCol_Sequence
    eRGCol_LastMod
    eRGCol_LastModKnown
    eRGCol_RuleUse
    eRGCol_Reverify
    eRGCol_ExitOnEntryBar
    eRGCol_ExitBasedOnTrade
    eRGCol_NumContracts
    eRGCol_AsPercent
    eRGCol_PyramidInfo
    eRGCol_SystemNumber
    'eRGCol_Shared
    eRGCol_CondCoded
    eRGCol_PriceCoded
    eRGCol_LimitCoded
    eRGCol_OrderPlacement
    eRGCol_NumCols '(keep this at end)
End Enum

Private Enum eIGCols
    eIGCol_RuleName = 0
    eIGCol_InputName
    eIGCol_InputValue
    eIGCol_FromVal
    eIGCol_ToVal
    eIGCol_IfOptimize
    eIGCol_OptFromValue
    eIGCol_OptToValue
    eIGCol_OptStepValue
    eIGCol_ParmTypeID
    eIGCol_ParmID
    eIGCol_ParmDesc
    eIGCol_RuleID
    eIGCol_Sort
    eIGCol_Req
    eIGCol_Hide
    eIGCol_NumCols
End Enum

Private Enum eMGCols
    eMGCol_ParmName = 0
    eMGCol_Security
    eMGCol_Period
    eMGCol_SecType
    eMGCol_SymbolPath
    eMGCol_Symbol
    eMGCol_MarketSymbol
    eMGCol_Format
    eMGCol_ParmID
    eMGCol_RuleID
    eMGCol_Sort
    eMGCol_GroupID
    eMGCol_SymbolID
    eMGCOl_NumCols
End Enum

'Signal sorting
Private Const optLong = 0
Private Const optLongExit = 1
Private Const optShort = 2
Private Const optShortExit = 3
Private Const optAll = 4

' Rule ID's for QuickStops
Private Const kProfitTargetLongRuleID = 51
Private Const kProfitTargetShortRuleID = 54
Private Const kStopLossLongRuleID = 52
Private Const kStopLossShortRuleID = 55
Private Const kTrailingStopLongRuleID = 53
Private Const kTrailingStopShortRuleID = 56

Public Property Get ID() As Long
    ID = m.System.SystemNumber
End Property

Private Function Tabs(ByVal lTab As eSMTabs) As Long
    Tabs = lTab
End Function
Private Function RGCol(ByVal lColumn As eRGCols) As Long
    RGCol = lColumn
End Function
Private Function IGCol(ByVal lColumn As eIGCols) As Long
    IGCol = lColumn
End Function
Private Function MGCol(ByVal lColumn As eMGCols) As Long
    MGCol = lColumn
End Function

Public Property Let System(pData As cSystem)
    Set m.System = pData
    EnableToolbar True
End Property
Public Property Get System() As cSystem
    Set System = m.System
End Property

Public Property Let Optimizing(ByVal bOptimizing As Boolean)
    m.bOptimizing = bOptimizing
    If m.bOptimizing Then
        Me.Hide
        DoEvents
    Else
        tmrOptimizing.Enabled = True
    End If
End Property

Private Sub GridsToRules()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim lRuleID As Long
    Dim lParmID As Long
        
    For X = vsRules.FixedRows To vsRules.Rows - 1
        lRuleID = CLng(vsRules.TextMatrix(X, RGCol(eRGCol_RuleID)))
        With m.System.Rules.Item(CStr(lRuleID))
            ''.Seq = X
            .Selected = CheckedCell(vsRules, X, RGCol(eRGCol_Selected))
            .Alternate = CheckedCell(vsRules, X, RGCol(eRGCol_Alt))
            .RuleUse = vsRules.TextMatrix(X, RGCol(eRGCol_RuleUse))
            .LastModKnown = Val(vsRules.TextMatrix(X, RGCol(eRGCol_LastModKnown)))
            .LinkedRules = NullChk(vsRules.TextMatrix(X, RGCol(eRGCol_LinkedRules)))
            .SysExitBasedOnEachTrade = CheckedCell(vsRules, X, RGCol(eRGCol_ExitBasedOnTrade))
            .SysNumContracts = vsRules.TextMatrix(X, RGCol(eRGCol_NumContracts))
            .SysAsPercentOfPosition = CheckedCell(vsRules, X, RGCol(eRGCol_AsPercent))
        End With
    Next X
    
    For X = vsInputs.FixedRows To vsInputs.Rows - 1
        lRuleID = CLng(vsInputs.TextMatrix(X, IGCol(eIGCol_RuleID)))
        lParmID = CLng(vsInputs.TextMatrix(X, IGCol(eIGCol_ParmID)))
        
        With m.System.Rules.Item(CStr(lRuleID)).Inputs.Item(CStr(lParmID))
            .Value = ConvertInputValue(vsInputs.TextMatrix(X, IGCol(eIGCol_InputValue)), vsInputs.TextMatrix(X, IGCol(eIGCol_ParmTypeID)))
            .IfOptimize = CheckedCell(vsInputs, X, IGCol(eIGCol_IfOptimize))
            If .ParmTypeID = kSN_RetTrueFalseConstant Then
                .OptFromValue = -1
                .OptToValue = 0
                .OptStepValue = 1
            Else
                .OptFromValue = ValOfText(vsInputs.TextMatrix(X, IGCol(eIGCol_OptFromValue)))
                .OptToValue = ValOfText(vsInputs.TextMatrix(X, IGCol(eIGCol_OptToValue)))
                .OptStepValue = ValOfText(vsInputs.TextMatrix(X, IGCol(eIGCol_OptStepValue)))
            End If
        End With
    Next X
    
    For X = vsMarkets.FixedRows To vsMarkets.Rows - 1
        lRuleID = CLng(vsMarkets.TextMatrix(X, MGCol(eMGCol_RuleID)))
        lParmID = CLng(vsMarkets.TextMatrix(X, MGCol(eMGCol_ParmID)))
        
        With m.System.Rules.Item(CStr(lRuleID)).Inputs.Item(CStr(lParmID))
            .Path = vsMarkets.TextMatrix(X, MGCol(eMGCol_SymbolPath))
            .Symbol = vsMarkets.TextMatrix(X, MGCol(eMGCol_Symbol))
            .MarketSymbol = vsMarkets.TextMatrix(X, MGCol(eMGCol_MarketSymbol))
            .Periodicity = FixPeriod(vsMarkets.TextMatrix(X, MGCol(eMGCol_Period)))
            .Format = vsMarkets.TextMatrix(X, MGCol(eMGCol_Format))
            .SecurityType = vsMarkets.TextMatrix(X, MGCol(eMGCol_SecType))
            .SecurityName = vsMarkets.TextMatrix(X, MGCol(eMGCol_Security))
            .GroupID = vsMarkets.TextMatrix(X, MGCol(eMGCol_GroupID))
            .SymbolID = CLng(Val(vsMarkets.TextMatrix(X, MGCol(eMGCol_SymbolID))))
            
            m.System.Markets.Item(.ParmName) = m.System.Rules.Item(CStr(lRuleID)).Inputs.Item(CStr(lParmID)).MakeCopy
        End With
        
    Next X

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.GridsToRules", eGDRaiseError_Raise

End Sub

Public Property Get NextRuleID() As Long
    m.lNextRuleID = m.lNextRuleID - 1
    NextRuleID = m.lNextRuleID
End Property

Public Property Get NextParmID() As Long
    m.lNextParmID = m.lNextParmID - 1
    NextParmID = m.lNextParmID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAllowReverse_Click
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAllowReverse_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.chkAllowReverse.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkForceLimitThrough_Click
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkForceLimitThrough_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    
    If (Visible = True) And (CheckBoxValue(chkForceLimitThrough) = False) Then
        If InfBox("Expecting a limit order to fill when the market hits your limit price and reverses is unrealistic.  This can cause overly optimistic results in your testing.", "i", "+Turn Off|Keep On", "Warning") = "K" Then
            CheckBoxValue(chkForceLimitThrough) = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.chkForceLimitThrough.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkLinkToChart_Click
'' Description: Allow the strategy to link to a chart
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkLinkToChart_Click()
On Error GoTo ErrSection:

    UseChartSystem chkLinkToChart.Value
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.chkLinkToChart.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewRule_Click
'' Description: Allow the user to create a new rule that will automatically be
''              added to the system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewRule_Click()
On Error GoTo ErrSection:
    
    NewRule

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.cmdNewRule.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPyramidInfo_Click
'' Description: Allow the user to edit the pyramid information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPyramidInfo_Click()
On Error GoTo ErrSection:

    EditPyramidInfo

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdPyramidInfo.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveRule_Click
'' Description: Allow the user to remove rules from the system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveRule_Click()
On Error GoTo ErrSection:

    RemoveRulesFromSystem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdRemoveRule.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRule
'' Description: Adds a rule to the system
'' Inputs:      Rule ID of the rule to add, Whether or not to Refresh Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddRule(ByVal Rule As cRule, Optional ByVal bRefreshRulesGrid As Boolean = True)
On Error GoTo ErrSection:

    Dim bInGrid As Boolean
    Dim X As Long
    Dim lRulePos As Long
    Dim strNewName As String
    Dim lRuleNum As Long
    
    'Quit if rule already exists...
    With vsRules
        For X = .FixedRows To .Rows - 1
            If CLng(.TextMatrix(X, RGCol(eRGCol_RuleID))) = Rule.RuleID Then
                lRulePos = X
                bInGrid = True
                Exit For
            End If
        Next X
    End With
    
    'If Rule no longer exists in table (deleted outside System Mgr) then quit...
    If Rule.RuleID = 0 Then Exit Sub
    
    'Load new rule to rules collection
    With Rule
        If Not bInGrid Then
            If .RuleID < 0 Then
                .Selected = True 'Default new rule to selected
                .RuleUse = .RuleType
                .SecurityLevel = m.System.SecurityLevel
                .Password = m.System.Password
                .CannotDelete = False
                
                .SysAsPercentOfPosition = .AsPercentOfPosition
                .SysExitBasedOnEachTrade = .ExitBasedOnEachTrade
                .SysNumContracts = .NumberContracts
            End If
                
            lRuleNum = 2
            strNewName = .Name
            Do While RuleNameExists(strNewName)
                strNewName = Trim(.Name) & " #" & Format(lRuleNum, "00")
                lRuleNum = lRuleNum + 1
            Loop
            
            If strNewName <> "" And strNewName <> .Name Then .Name = strNewName
            
            If strNewName <> "" Then
                m.System.Rules.Add Rule.RuleID, Rule
                AddRuleToGrid Rule
            Else
                .Delete
            End If
        Else
            'Get updated system number from frmRule...
            m.System.Rules.Item(CStr(Rule.RuleID)) = Rule
            AddRuleToGrid Rule, lRulePos
        End If
    End With
    
    AddInput Rule, bRefreshRulesGrid
    AddMarket Rule, bRefreshRulesGrid
    
    EnableToolbar True
    If bRefreshRulesGrid Then RefreshRulesGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.AddRule", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkLinkInputs_Click
'' Description: Hide/Show the Linked Inputs as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkLinkInputs_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    If chkLinkInputs = vbChecked Then
        HideDuplicateInputs
    Else
        ShowDuplicateInputs
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.chkLinkInputs.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkPyramid_Click
'' Description: Show/Hide Pyramid information on Rules grid as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkPyramid_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    EnableToolbar True
    With vsRules
        If chkPyramid.Value = vbChecked Then
            '.ColHidden(24) = False
            .ColHidden(RGCol(eRGCol_PyramidInfo)) = False
            .ColHidden(.Cols - 1) = True
            Enable txtTradeDepth
            
            For lIndex = vsRules.FixedRows To vsRules.Rows - 1
                DisplayPyramidInfo lIndex
            Next lIndex
            
            ' If the trade depth is less than 2, default it to 10 since a value
            ' less than 2 does not allow pyramiding to happen...
            If ValOfText(txtTradeDepth.Text) < 2 Then
                txtTradeDepth.Text = 10
            End If
            
            cmdPyramidInfo.Enabled = True
        Else
            '.ColHidden(24) = True
            .ColHidden(RGCol(eRGCol_PyramidInfo)) = True
            .ColHidden(.Cols - 1) = False
            Disable txtTradeDepth
        
            cmdPyramidInfo.Enabled = False
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.chkPyramid.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddRule_Click
'' Description: Allow the user to add rule(s) to the system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddRule_Click()
On Error GoTo ErrSection

    AddRules
    If tbToolbar.ToolBars("Wizard").Visible Then SetWizardBack

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdAddRule.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to choose a different symbol from the pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:
    
    BrowseMarkets
    
ErrExit:
    cmdBrowse.Enabled = True
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdBrowse.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdQuickStops_Click
'' Description: Allow the user to add "Quick Stop Rules" to the system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdQuickStops_Click()
On Error GoTo ErrSection:

    QuickStops

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdQuickStops.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLinkToEntry_Click
'' Description: Allow the user to link specific exits to specific entries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLinkToEntry_Click()
On Error GoTo ErrSection:
    
    LinkToEntry
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.cmdLinkToEntry.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditRule_Click
'' Description: Allow the user to edit a rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditRule_Click()
On Error GoTo ErrSection:

    EditRule

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.cmdEditRule.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the current system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "SNV System", Me, 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.PrintMe", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSNvsTS_Click
'' Description: Show the differences in trades between System Navigator and
''              Trade Station
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSNvsTS_Click()
On Error GoTo ErrSection:
    
    With frmTradeDiff
        .lst.Clear
        .txtSysNav = m.System.TradesFile
        ShowForm frmTradeDiff, True
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdSNvsTS.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTestEntry_Click
'' Description: Allow the user to only Test the Selected Entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTestEntry_Click()
On Error GoTo ErrSection:

    RunTest False, vsRules.TextMatrix(vsRules.RowSel, RGCol(eRGCol_RuleID))

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdTestEntry.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        g.Help.ShowF1Help Me
    Else
        frmMain.DockPro_ShortcutKeyDown KeyCode, Shift, Me.Name
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the 'X', ask if they want to save the system
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim strReturn As String

    If UnloadMode <> vbFormCode Then
        Cancel = AskToSave
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form gets resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long               ' Minimum form width allowed
    Dim lMinHeight As Long              ' Minimum form height allowed
    
    ' Figure out the minimum size for the form and limit to that
    lMinWidth = 8820
    lMinHeight = 6990
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
                           
    ' Resize the tabs and the buttons
    With vsIndexTab1
'        .Move .Left, tbToolbar.GetDockHeight(ssDockedTop) + 100, _
            ScaleWidth - (.Left * 2), _
            ScaleHeight - tbToolbar.GetDockHeight(ssDockedTop) - 200
        .Move .Left, .Top, ScaleWidth - (.Left * 2), ScaleHeight - (.Top * 2)
        .Refresh
    End With
        
    ' Resize the contents of the Rules tab
    With fraRuleButtons
        .Move vsIndexTab1.ClientWidth - .Width - fraSignal.Left
    End With
    With txtPreview
        .Move .Left, vsIndexTab1.ClientHeight - .Height - fraSignal.Top, _
            vsIndexTab1.ClientWidth - (.Left * 2)
    End With
    With vsRules
        .Move .Left, .Top, vsIndexTab1.ClientWidth - fraRuleButtons.Width - (.Left * 3), _
            vsIndexTab1.ClientHeight - fraSignal.Height - txtPreview.Height - (fraSignal.Top * 4)
    End With
    
    ' Resize the contents of the Inputs tab
    With vsInputs
        If chkLinkInputs.Visible Then
            .Move .Left, fraLinkInputs.Top + fraLinkInputs.Height, vsIndexTab1.ClientWidth, _
                vsIndexTab1.ClientHeight - (chkLinkInputs.Top + chkLinkInputs.Height)
        Else
            .Move .Left, fraLinkInputs.Top, vsIndexTab1.ClientWidth, _
                vsIndexTab1.ClientHeight
        End If
    End With

    ' Resize the contents of the Data tab
    With fraDataSettings
        .Move vsIndexTab1.ClientWidth - .Width - vsMarkets.Left
    End With
    With vsMarkets
        .Move .Left, .Top, vsIndexTab1.ClientWidth - fraDataSettings.Width - (.Left * 3), _
            vsIndexTab1.ClientHeight - (.Top * 2)
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Reload the grids if necessary upon re-entering the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim i&
    
    ' make backcolors match for labels which go directly on the buttons
    i = cmdNewRule.BackColor
    lblNewRule.BackColor = i
    lblAddRule.BackColor = i
    lblQuickStops.BackColor = i
    lblBrowse.BackColor = i
    
    ' Quickly check the Reverify flag.  If on then force a reverify...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "WHERE [SystemNumber]=" & m.System.SystemNumber & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblSystems"
    If Not rs.EOF Then
        If rs!CheckSum = 0.5 Then
            EnableToolbar False
            Unload Me
            Err.Raise vbObjectError + 1000, , "This Strategy is no longer Valid"
        End If
        
        If rs!Reverify Then EnableToolbar True
    End If
    rs.Close
    
    ' Update any rules that may have changed or been deleted on us
    UpdateRules
    
    If GetActiveWindow = Me.hWnd Then
        Select Case vsIndexTab1.CurrTab
            Case Tabs(eSMTab_Rules)
                MoveFocus vsRules
            Case Tabs(eSMTab_Inputs)
                MoveFocus vsInputs
            Case Tabs(eSMTab_Markets)
                MoveFocus vsMarkets
        End Select
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.Form.Activate", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the controls on the form and set the form size
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText As String
    Dim strFont As String
    
    g.Styler.StyleForm Me
    
    ' NOTE: Don't do this here (security issues)
    ''vsIndexTab1.CurrTab = Tabs(eSMTab_Rules)
    
    tbToolbar.ToolBars("Wizard").Visible = False
    
    'Bar type combo boxes loaded
    With cboBarType
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Quarterly"
        .AddItem "Yearly"
    End With
    
    'Resize and center form
    Width = 11880
    CenterTheForm Me

    cmdTestEntry.Top = cmdLinkToEntry.Top

    strText = GetIniFileProperty("SysMgr", "", "Placement", g.strIniFile)
    If strText <> "" Then SetFormPlacement Me, strText, "LHT"

    Me.Icon = Picture16(ToolbarIcon("ID_Strategies"), , True)
    With tbToolbar
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Run").Picture = Picture16(ToolbarIcon("ID_Performance"))
        .Tools("ID_RunGroup").Picture = Picture16(ToolbarIcon("ID_Performance"))
        .Tools("ID_Notes").Picture = Picture16(ToolbarIcon("ID_News"))
        .Tools("ID_Orders").Picture = Picture16(ToolbarIcon("ID_Orders"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_RunA").Picture = Picture16(ToolbarIcon("ID_Performance"))   'Wizard button
        .Tools("ID_WizardMsg").ChangeAll ssChangeAllForeColor, vbBlue
    End With
        
    dtpFromDate.AllowWeekends = False
    dtpToDate.AllowWeekends = False
    dtpFromDate.MaxDateIsToday = True
    dtpToDate.MaxDateIsToday = True
        
    ''If Not FileExist(AddSlash(App.Path) & "TRDDIFF.MOD") Then cmdSNvsTS.Visible = False
    
    ' For now, hide the bar type combo box - 11/14/00 DAJ
    lblBarType.Visible = False
    cboBarType.Visible = False
    tmrOptimizing.Enabled = False
    
    ''Set m.System = Nothing
    
    mnuRule.Visible = False
    mnuInputs.Visible = False
    mnuMarkets.Visible = False
    
    strFont = GetIniFileProperty("SystemMgrRules", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsRules.Font, strFont
    strFont = GetIniFileProperty("SystemMgrInputs", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsInputs.Font, strFont
    strFont = GetIniFileProperty("SystemMgrMarkets", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsMarkets.Font, strFont
    
    'save original button text color
    m.nButtonTextColor = lblNewRule.ForeColor
    
    ' DAJ 11/03/2011: Need to start the first temporary rule ID at -3 because -1 and -2 are
    ' reservered in the reports as "Manual Entry" and "Manual Exit" respectively...
    m.lNextRuleID = -2
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Load up a blank system to allow for a new system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Add()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass
    
    Set m.System = New cSystem
    m.System.Load 0
    m.System.LibraryID = kSN_UserLibrary
    
    SetEditorCaption Me, "Strategy", ""
    m.strName = ""
    txtCommission = GetIniFileProperty("Expenses", 0, "Systems", g.strIniFile)
    txtBarsLoadedBeforeTrading = GetIniFileProperty("BarsLoadedBeforeTrading", 60, "Systems", g.strIniFile)
    txtBarsTradedBeforeOrders = GetIniFileProperty("BarsTradedBeforeOrders", 60, "Systems", g.strIniFile)
    txtTradeDepth = GetIniFileProperty("TradeDepth", 0, "Systems", g.strIniFile)
    
    cboBarType.Text = GetIniFileProperty("BarTimeFrame", "Daily", "Systems", g.strIniFile)
    If Len(cboBarType.Text) = 0 Then cboBarType.Text = "Daily"
    dtpFromDate = DateSerial(1950, 1, 1)
    dtpToDate = Date
    optToEndOfData = True
    dtpToDate.Enabled = optToDate
    
    ' Initialize the grids...
    InitRulesGrid
    InitInputsGrid
    InitInputsGrid
    
    ' Load the grids...
    LoadGrids
    
    ' Added "LinkInputs" stuff 4/17/2002 DAJ
    chkLinkInputs.Value = vbChecked
        
    chkPyramid = vbUnchecked
    HideOrShowInputs
    
    chkAllowReverse = vbChecked
    chkForceLimitThrough = vbChecked
    
    EnableToolbar False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.Add", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadRec
'' Description: Load up a sytem into the system manager
'' Inputs:      System ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadRec(ByVal lSystemNumber As Long)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass
    
gdResetProfiles 200, 299
gdStartProfile 200
gdStartProfile 201
    
    If m.System Is Nothing Then
        Set m.System = New cSystem
        m.System.Load lSystemNumber
        If m.System.SecurityLevel > 0 Then m.bShowPassword = True
    End If
    
    ' 6/4/01: clear the "rules display filter" (back to ALL)
    optSignals(4) = True
    
    With m.System
        m.strName = .SystemName
        txtCommission.Text = Format(Val(.Expenses), "$#,##0.00")
        txtBarsLoadedBeforeTrading.Text = Format(Val(.BarsLoadedBeforeTrading), "#,##0")
        txtBarsTradedBeforeOrders.Text = Format(Val(.BarsTradedBeforeOrders), "#,##0")
        
        If .LinkInputs Then chkLinkInputs = vbChecked Else chkLinkInputs = vbUnchecked
        
        If .BarTimeFrame <> "" Then cboBarType.Text = .BarTimeFrame
        
        If Len(cboBarType.Text) = 0 Then cboBarType.Text = "Daily"
        
        'Set test date ranges
        If .FromDate = 0 Then
            dtpFromDate = DateSerial(1950, 1, 1)
        Else
            dtpFromDate = .FromDate
        End If
        If .ToEndOfData Then
            optToEndOfData = True
        Else
            optToDate = True
        End If
        If .ToDate = 0 Then
            dtpToDate = Date
            optToEndOfData = True
        Else
            dtpToDate = .ToDate
        End If
        'set date to use for saving
        m.dFromDate = .FromDate
        m.dToDate = .ToDate
        m.bToEndOfData = .ToEndOfData
        
        dtpToDate.Enabled = optToDate
        If .Pyramid = True And .TradeDepth < 2 Then .TradeDepth = 10
        txtTradeDepth.Text = Format(Val(.TradeDepth), "#,##0")
        If .UseSharesPerTrade Then
            optSharesPerTrade.Value = True
        Else
            optDollarsPerTrade.Value = True
        End If
        txtNumShares.Text = Format(.NumShares, "#,##0")
        txtDollarsPerTrade.Text = Format(.DollarsPerTrade, "$#,##0")
        txtStockCommission.Text = Format(.StockExpenses, "$#,##0.00")
        txtForexCommission.Text = Format(.ForexExpenses, "$#,##0.00")
        If .AllowReverse Then chkAllowReverse = vbChecked Else chkAllowReverse = vbUnchecked
        If .ForceLimitThrough Then chkForceLimitThrough = vbChecked Else chkForceLimitThrough = vbUnchecked
    End With
    
    SetEditorCaption Me, "Strategy", m.strName
gdStopProfile 201
    
    FixSequence False
    
    ' Intialize the grids...
    InitRulesGrid
    InitInputsGrid
    InitMarketsGrid
    
    ' Load the grids...
gdStartProfile 205
    LoadGrids
gdStopProfile 205
    
    ' Hide inputs if necessary...
    'With Rules
    '    For lIndex = 1 To .Count
    '        If .Item(lIndex).Alternate = False And .Item(lIndex).Selected = False Then
    '            HideInputs .Item(lIndex).RuleID, True
    '        Else
    '            HideInputs .Item(lIndex).RuleID, False
    '        End If
    '    Next lIndex
    '    ShowLinkInputs
    'End With
    
    With vsRules
        If m.System.Pyramid Then
            .ColHidden(RGCol(eRGCol_PyramidInfo)) = False ' Pyramid info column
            .ColHidden(.Cols - 1) = True ' Blank column
            Enable txtTradeDepth
            chkPyramid = 1
        Else
            .ColHidden(RGCol(eRGCol_PyramidInfo)) = True 'Pyramid info column
            .ColHidden(.Cols - 1) = False 'Blank column
            Disable txtTradeDepth
            chkPyramid = 0
        End If
    End With

gdStopProfile 200
If IsIDE Then
    FileFromString App.Path & "\Chk\SysLoad.chk", gdGetProfiles(200, 299, vbCrLf), True
End If
    
    DoEvents
    EnableToolbar m.System.Reverify

ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.LoadRec", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings to Ini File upon form unload
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    If g.CurrentSystem Is m.System Then Set g.CurrentSystem = Nothing
    Set m.System = Nothing

    tmrOptimizing.Enabled = False
    SetIniFileProperty "SysMgr", GetFormPlacement(Me), "Placement", g.strIniFile
    
    SetIniFileProperty "BarTimeFrame", cboBarType.Text, "Systems", g.strIniFile
    SetIniFileProperty "Expenses", txtCommission.Text, "Systems", g.strIniFile
    SetIniFileProperty "StockExpenses", txtStockCommission.Text, "Systems", g.strIniFile
    SetIniFileProperty "ForexExpenses", txtForexCommission.Text, "Systems", g.strIniFile
    SetIniFileProperty "UseSharesPerTrade", optSharesPerTrade.Value, "Systems", g.strIniFile
    SetIniFileProperty "NumShares", txtNumShares.Text, "Systems", g.strIniFile
    SetIniFileProperty "DollarsPerTrade", ValOfText(txtDollarsPerTrade.Text), "Systems", g.strIniFile
    SetIniFileProperty "BarsLoadedBeforeTrading", txtBarsLoadedBeforeTrading.Text, "Systems", g.strIniFile
    SetIniFileProperty "BarsTradedBeforeOrders", txtBarsTradedBeforeOrders.Text, "Systems", g.strIniFile
    SetIniFileProperty "TradeDepth", txtTradeDepth.Text, "Systems", g.strIniFile
    
    SetIniFileProperty "SystemMgrRules", FontToString(vsRules.Font), "Fonts", g.strIniFile
    SetIniFileProperty "SystemMgrInputs", FontToString(vsInputs.Font), "Fonts", g.strIniFile
    SetIniFileProperty "SystemMgrMarkets", FontToString(vsMarkets.Font), "Fonts", g.strIniFile
    
    If m.bSaveLinkFlag Then SetIniFileProperty "LinkToChart", Abs(m.bLinkedToChart), "Systems", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.Form.Unload", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the system to the database
'' Inputs:      None
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Save(ByVal strButton As String) As Boolean
On Error GoTo ErrSection:
    
    Dim rs              As Recordset
    Dim bRename         As Boolean
    Dim lIndex          As Long
    Dim bSaveAs As Boolean
    Dim strNewName As String
    Dim strText As String
    Dim frm As Form
    
    ' If this strategy is being automatically traded, warn the user that their
    ' changes will not take effect until they stop and restart the auto trade item...
    If g.TradingItems.IsStrategyAutoTrading(m.System.SystemNumber) = True Then
        If strButton = "ID_Save" Then
            InfBox "In order for your changes to take effect in automated trading, you will need to stop and restart the appropriate automated trading item.", "i", , "Warning"
        End If
    End If
    
    'switch back to original market1 if market1 is linked to chart
    If m.bLinkedToChart Then
        chkLinkToChart.Value = 0
    End If
    
    If ValidateMarkets = False Then Exit Function
    
    FixSequence
    
    ' Handle Rename/Save As
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        strText = "Save the current Strategy as..."
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
        bRename = True
    ElseIf strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Strategy as..."
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & "Copy of " & m.strName & " ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        strText = "Rename the current Strategy as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
        bRename = True
    Else
        strNewName = m.strName
    End If
    
    'If the system name changes, allow user to make a copy
    strNewName = Trim(strNewName)
    Do While Len(strNewName) > 0 And strNewName <> m.System.SystemName
        'Make sure single quotes not in system name
        If InStr(strNewName, "'") > 0 Then
            InfBox "Single quotes not allowed in Strategy Name", "e"
        Else
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                    "WHERE [SystemName]='" & strNewName & "';", dbOpenSnapshot)
            If Not rs.EOF Then
                InfBox "Strategy '" & strNewName & "' already exists", "e"
            Else
                Exit Do
            End If
        End If
        strText = "Rename the Strategy as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
    Loop
    
    If Len(Trim(strNewName)) = 0 Then
        Exit Function 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    m.strName = Trim(strNewName)
    SetEditorCaption Me, "Strategy", m.strName
    
    Screen.MousePointer = vbHourglass
    
    'Recalculate step values in Inputs grid
    VerifyStepValues
           
    With m.System
        If .SystemNumber = 0 Or bSaveAs Then
            If .SecurityLevel < 2 Then
                .SecurityLevel = 0
                .CannotDelete = False
                .LibraryID = kSN_UserLibrary
                .Password = ""
            End If
        ElseIf bRename = True Then
            'User must be authorized to rename (save)
            If m.bShowPassword = True Then
                If Not g.Security.CanSave(.SecurityLevel, .Password) Then
                    GoTo ErrExit:
                Else
                    m.bShowPassword = False
                End If
            End If
        End If
        
        .SystemName = m.strName
        .FromDate = dtpFromDate ' m.dFromDate
        .ToDate = dtpToDate ' m.dToDate
                
        .LinkInputs = (chkLinkInputs = vbChecked)
        If .LinkInputs Then SetDuplicates
        
        .Expenses = ValOfText(txtCommission.Text)
        .BarsLoadedBeforeTrading = ValOfText(txtBarsLoadedBeforeTrading.Text)
        .BarsTradedBeforeOrders = ValOfText(txtBarsTradedBeforeOrders.Text)
        For lIndex = vsMarkets.FixedRows To vsMarkets.Rows - 1
            If vsMarkets.TextMatrix(lIndex, 0) = "Market1" Then
                .BarTimeFrame = vsMarkets.TextMatrix(lIndex, 2)
                Exit For
            End If
        Next lIndex
        '.BarTimeFrame = cboBarType.Text
        .ToEndOfData = optToEndOfData.Value ' = m.bToEndOfData
        .MMid = 0
        .Pyramid = Val(chkPyramid.Value) * -1
        .TradeDepth = ValOfText(txtTradeDepth.Text)
        .UseSharesPerTrade = optSharesPerTrade.Value
        .NumShares = CLng(ValOfText(txtNumShares.Text))
        .DollarsPerTrade = ValOfText(txtDollarsPerTrade.Text)
        .StockExpenses = ValOfText(txtStockCommission.Text)
        .ForexExpenses = ValOfText(txtForexCommission.Text)
        .AllowReverse = (chkAllowReverse = vbChecked)
        .ForceLimitThrough = (chkForceLimitThrough = vbChecked)
    End With
        
        GridsToRules
        If Len(m.System.SystemName) <> 0 Then
            If bSaveAs Then Set m.System = m.System.MakeCopy(NextSystemID)
        End If
            
    With m.System
        .Save
        g.bDirtyLibrariesMDB = True
        
        'Adjust values if changed in validate method
        txtBarsLoadedBeforeTrading = .BarsLoadedBeforeTrading
        txtBarsTradedBeforeOrders = .BarsTradedBeforeOrders
        txtTradeDepth = .TradeDepth
    End With
        
    
    EnableToolbar False
    Save = True
    
    ' update any charts with this system
    For lIndex = 0 To Forms.Count - 1
        If IsFrmChart(Forms(lIndex)) Then
            Set frm = Forms(lIndex)
            With frm.Chart
                If .ShowTrades And .SystemID = m.System.SystemNumber Then
                    .ShowTrades = False
                    .GenerateChart eRedo1_Scrolled
                    .ShowTrades = True
                    .GenerateChart eRedo1_Scrolled
                End If
            End With
        End If
    Next
    Set frm = Nothing
    
ErrExit:
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

ErrSection:
    ''Select Case m.System.ErrNbr
    ''    Case 2: MoveFocus dtpFromDate
    ''    Case Else
    ''End Select
    RaiseError "frmSystemManager.Save", eGDRaiseError_Raise

End Function

Private Sub lblAddRule_Click()

    cmdAddRule_Click
    
End Sub

Private Sub lblBrowse_Click()

    cmdBrowse_Click

End Sub

Private Sub lblNewRule_Click()

    cmdNewRule_Click

End Sub

Private Sub lblQuickStops_Click()

    cmdQuickStops_Click

End Sub

Private Sub mnuAddRule_Click()
On Error GoTo ErrSection:

    AddRules

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuAddRule.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuChangeFontInputs_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsInputs, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuChangeFontInputs.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuChangeFontRules_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsRules, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuChangeFontRules.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuChangeFontMarkets_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsMarkets, True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuChangeFontMarkets.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuChangeSymbol_Click()
On Error GoTo ErrSection:
    
    BrowseMarkets

ErrExit:
    cmdBrowse.Enabled = True
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuChangeSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuEditPyramid_Click()
On Error GoTo ErrSection:

    EditPyramidInfo

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuEditPyramid.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuEditRule_Click()
On Error GoTo ErrSection:

    EditRule

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuEditRule.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuFavorites_Click()
On Error GoTo ErrSection:

    Dim lRuleID As Long

    lRuleID = CLng(vsRules.TextMatrix(vsRules.RowSel, RGCol(eRGCol_RuleID)))
    AddRuleToFavorites m.System.Rules.Item(CStr(lRuleID))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.mnuFavorites.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuLinkEntry_Click()
On Error GoTo ErrSection:
    
    LinkToEntry
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.mnuLinkEntry.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuMarketInfo_Click()
On Error GoTo ErrSection:

    MarketInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.mnuMarketInfo.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub mnuNewRule_Click()
On Error GoTo ErrSection:

    NewRule

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuNewRule.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuQuickStops_Click()
On Error GoTo ErrSection:

    QuickStops

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuQuickStops.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuRemoveRule_Click()
On Error GoTo ErrSection:
    
    RemoveRulesFromSystem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuRemoveRule.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuTestEntry_Click()
On Error GoTo ErrSection:

    RunTest False, vsRules.TextMatrix(vsRules.RowSel, RGCol(eRGCol_RuleID))

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.mnuTestEntry.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optDollarsPerTrade_Click()
    EnableToolbar True
End Sub

Private Sub optSharesPerTrade_Click()
On Error GoTo ErrSection:
    
    Dim strMsg$
    If GetIniFileProperty("SharesPerTrade", 0, "DontAsk", g.strIniFile) = 0 Then
        strMsg = "When backtesting stocks, the 'Dollars per trade' option is recommended since it allows the same percentage increase or decrease in the stock's price to cause the same amount of profit or loss.||Do you still wish to use 'Shares per trade'?"
        strMsg = InfBox(strMsg, "!", "+Yes|-Cancel", "NOTE", , , , , , , , , True)
        If Right(strMsg, 1) = "-" Then
            SetIniFileProperty "SharesPerTrade", 1, "DontAsk", g.strIniFile
        End If
    End If
    If Left(strMsg, 1) = "C" Then
        optDollarsPerTrade.Value = True
    Else
        EnableToolbar True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.optSharesPerTrade_Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSignals_Click
'' Description: Filter the rules grid according to which option the user chose
'' Inputs:      Index of the option button clicked on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSignals_Click(Index As Integer)
    
    Filter

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strID As String
    Dim lSystemID As Long, i&
    Dim bOK As Boolean

    ToggleFocus Me, Me.vsIndexTab1
    
    Select Case Tool.ID
        Case "ID_Save"
            If HasGold(True) Then
                tbToolbar.Tools("ID_Toolbox").Enabled = False
                tbToolbar.Tools("ID_Close").Enabled = False
                
                If Save(Tool.ID) Then
                    lSystemID = m.System.SystemNumber
                    ' TLB 4/13/2015: commented out the following line so as not to lose the
                    ' intraday data that was already loaded (don't know if we really need it?)
                    'Set m.System = Nothing
                    LoadRec lSystemID
                End If
            
                tbToolbar.Tools("ID_Toolbox").Enabled = Not m.bModal
                tbToolbar.Tools("ID_Close").Enabled = True
            End If
        
        Case "ID_SaveAs"
            If HasPlatinum(True) Then
                tbToolbar.Tools("ID_Toolbox").Enabled = False
                tbToolbar.Tools("ID_Close").Enabled = False
                
                If Save(Tool.ID) Then
                    lSystemID = m.System.SystemNumber
                    Set m.System = Nothing
                    LoadRec lSystemID
                End If
            
                tbToolbar.Tools("ID_Toolbox").Enabled = Not m.bModal
                tbToolbar.Tools("ID_Close").Enabled = True
            End If
        
        Case "ID_Rename"
            If HasPlatinum(True) Then
                tbToolbar.Tools("ID_Toolbox").Enabled = False
                tbToolbar.Tools("ID_Close").Enabled = False
                
                If Save(Tool.ID) Then
                    lSystemID = m.System.SystemNumber
                    Set m.System = Nothing
                    LoadRec lSystemID
                End If
            
                tbToolbar.Tools("ID_Toolbox").Enabled = Not m.bModal
                tbToolbar.Tools("ID_Close").Enabled = True
            End If
        
        Case "ID_Run", "ID_RunA"
            If Tool.ID = "ID_RunA" Then ExitWizard
            If HasPlatinum(True) Then
                'Make sure the System's rules have not changed since the last time this
                'system was saved...
                CheckReverifyFlags
                
                'If tbToolbar.Tools("ID_Save").Enabled Or Len(m.System.SystemName) = 0 Then
                '    If Not Save("ID_Save") Then Exit Sub
                'End If
                
                'm.System.Test
                InfBox "Please wait while back-testing your strategy ...", "t", , "Processing", True
                RunTest False
                InfBox
            End If
            
#If 0 Then
        ' TLB 7/1/2014: I'm not sure we really need this button anymore?
        ' (since it's actually better to run a symbol group using the optimizer form)
        Case "ID_RunGroup"
            If HasPlatinum(True) Then
                CheckReverifyFlags
                
                If Tool.State = ssChecked Then
                    Tool.Name = "St&op"
                    m.bStop = False
                    
                    RunTest False, 0&, True
                Else
                    Tool.Name = "Run &Group"
                    m.bStop = True
                End If
            End If
#End If
                
        Case "ID_Orders"
            bOK = False
            If HasModule("JDMP") Then ' first check for non-gold allowances (to skip upgrade message)
                bOK = True
            ElseIf HasGold(True) Then ' then give upgrade message if still not allowed
                bOK = True
            End If
            If bOK Then
                'Make sure the System's rules have not changed since the last time this
                'system was saved...
                CheckReverifyFlags
                
                'If tbToolbar.Tools("ID_Save").Enabled Or Len(m.System.SystemName) = 0 Then
                '    If Not Save("ID_Save") Then Exit Sub
                'End If
                
                'm.System.NextBarReport
                'RunTest True
                
                If Tool.Name = "&Orders" Then
                    If UCase(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_SecType))) = "GROUP" Then
                        Tool.Name = "St&op"
                        tbToolbar.Tools("ID_Orders").Picture = Picture16(ToolbarIcon("kRedLight"))
                        m.bStop = False
                    End If
                    RunTest True
                    If UCase(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_SecType))) = "GROUP" Then
                        Tool.Name = "&Orders"
                        tbToolbar.Tools("ID_Orders").Picture = Picture16(ToolbarIcon("ID_Orders"))
                    End If
                Else
                    m.bStop = True
                    Tool.Name = "&Orders"
                    tbToolbar.Tools("ID_Orders").Picture = Picture16(ToolbarIcon("ID_Orders"))
                End If
            End If

        Case "ID_Notes"
            If HasPlatinum(True) Then
                If ShowPassword Then
                    m.System.Notes = frmNotes.ShowMe(m.System.Notes)
                    EnableToolbar True
                End If
            End If
            
        Case "ID_Print"
            If HasPlatinum(True) Then
                PrintMe
            End If
            
        Case "ID_Toolbox"
            If Not AskToSave Then
                strID = CStr(m.System.SystemNumber)
                Unload Me
                frmToolbox.ShowMe eTab_Systems, strID
            End If
            
        Case "ID_Close"
            If Not AskToSave Then
                RemoveLocalRules
                If m.bModal Then
                    Me.Hide
                Else
                    Unload Me
                End If
            End If
            
        Case "ID_Wizard"
            If SetWizardStart Then
                tbToolbar.Redraw = False
                tbToolbar.ToolBars("General").Visible = False
                tbToolbar.ToolBars("Wizard").Visible = True
                m.nWizardStep = m.nWizardStart
                SetWizardBack
                tbToolbar.Redraw = True
                Call SetIniFileProperty("StrategyWizard", "W", "DontAsk", g.strIniFile)
            End If
        
        Case "ID_ExitWizard"
            ExitWizard
            If vsRules.Rows = vsRules.FixedRows Then
                Call SetIniFileProperty("StrategyWizard", "N", "DontAsk", g.strIniFile)
            End If
            
        Case "ID_Next"
            IncWizardStep
            SetWizardBack
        
        Case "ID_Back"
            DecWizardStep
            SetWizardBack

    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    tbToolbar.Tools("ID_Toolbox").Enabled = Not m.bModal
    tbToolbar.Tools("ID_Close").Enabled = True
    RaiseError "frmSystemMangager.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrOptimizing_Timer
'' Description: Wait for the frmOptimizer to unload before re-showing this form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrOptimizing_Timer()

    If Not FormIsLoaded("frmOptimizer") Then
        tmrOptimizing.Enabled = False
        ShowMe m.System.SystemNumber, True
    End If

End Sub

'========================== System field events ===========================
Private Sub txtBarsLoadedBeforeTrading_Change()
    EnableToolbar True
End Sub
Private Sub txtBarsTradedBeforeOrders_Change()
    EnableToolbar True
End Sub
Private Sub txtCommission_Change()
    EnableToolbar True
End Sub
Private Sub txtCommission_LostFocus()

    On Error Resume Next
    Dim d#, s$
    d = ValOfText(txtCommission.Text)
    s = Format(d, "$#,##0.00")
    If s <> txtCommission.Text Then
        txtCommission.Text = s
    End If
    
End Sub
Private Sub cboBarType_Click()
    EnableToolbar True
End Sub
Private Sub dtpFromDate_Changed()
    EnableToolbar Not m.bLinkedToChart
End Sub
Private Sub dtpToDate_Changed()
    EnableToolbar Not m.bLinkedToChart
End Sub
Private Sub optToDate_Click()
    EnableToolbar Not m.bLinkedToChart
    dtpToDate.Enabled = optToDate
End Sub
Private Sub optToEndOfData_Click()
    EnableToolbar Not m.bLinkedToChart
    dtpToDate.Enabled = optToDate
End Sub

Private Sub txtDollarsPerTrade_Change()
    EnableToolbar True
End Sub
Private Sub txtDollarsPerTrade_LostFocus()

    On Error Resume Next
    Dim d#, s$
    d = ValOfText(txtDollarsPerTrade.Text)
    s = Format(d, "$#,##0")
    If s <> txtDollarsPerTrade.Text Then
        txtDollarsPerTrade.Text = s
    End If
    
End Sub

Private Sub txtForexCommission_Change()
    EnableToolbar True
End Sub
Private Sub txtForexCommission_LostFocus()

    On Error Resume Next
    Dim d#, s$
    d = ValOfText(txtForexCommission.Text)
    s = Format(d, "$#,##0.00")
    If s <> txtForexCommission.Text Then
        txtForexCommission.Text = s
    End If
    
End Sub

Private Sub txtNumShares_Change()
    EnableToolbar True
End Sub

Private Sub txtNumShares_LostFocus()
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    For lIndex = vsRules.FixedRows To vsRules.Rows - 1
        DisplayPyramidInfo lIndex
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.txtNumShares.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtStockCommission_Change()
    EnableToolbar True
End Sub
Private Sub txtStockCommission_LostFocus()

    On Error Resume Next
    Dim d#, s$
    d = ValOfText(txtStockCommission.Text)
    s = Format(d, "$#,##0.00")
    If s <> txtStockCommission.Text Then
        txtStockCommission.Text = s
    End If
    
End Sub

Private Sub txtTradeDepth_Change()
    EnableToolbar True
End Sub

Private Sub txtTradeDepth_LostFocus()
On Error GoTo ErrSection:

    If ValOfText(txtTradeDepth.Text) < 2 Then
        InfBox "Trade Depth must be at least 2 when pyramiding", "i", , "Error"
        txtTradeDepth.Text = "2"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.txtTradeDepth.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsIndexTab1_Click
'' Description: When the user changes tabs, move the focus to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsIndexTab1_Click()

    Select Case vsIndexTab1.CurrTab
        Case Tabs(eSMTab_Rules)
            MoveFocus vsRules
        Case Tabs(eSMTab_Inputs)
            MoveFocus vsInputs
        Case Tabs(eSMTab_Markets)
            MoveFocus vsMarkets
        Case Tabs(eSMTab_Settings)
            MoveFocus txtCommission
    End Select

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsIndexTab1_Switch
'' Description: Don't allow the user to view Rules tab unless permitted
'' Inputs:      Old tab, New tab, Whether to Cancel the switch
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsIndexTab1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)

    If NewTab = Tabs(eSMTab_Inputs) Then
        HideOrShowInputs
    ElseIf NewTab = Tabs(eSMTab_Rules) Then
        If Not HasPlatinum(True, "Editing the Rules") Then
            Cancel = True
        ElseIf m.bShowPassword Then
            If Not g.Security.CanEdit(m.System.SecurityLevel, m.System.Password) Then
                Cancel = True
            Else
                If m.System.SecurityLevel >= 2 Then
                    m.bShowPassword = False
                End If
            End If
        End If
    End If
    
    ' When the user leaves the inputs tab, if they have chosen to link
    ' inputs, set the duplicates to the same value
    If OldTab = Tabs(eSMTab_Inputs) And Not Cancel Then
        If chkLinkInputs = vbChecked Then SetDuplicates
    End If
    
End Sub

Private Sub vsInputs_ChangeEdit()
On Error GoTo ErrSection:
    
    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.vsInputs.ChangeEdit", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub vsInputs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsInputs
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            PopupMenu mnuInputs
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsInputs.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsMarkets_ComboCloseUp
'' Description: When the user closes the combo, set the FinishEdit to True so
''              that the AfterEdit event will occur.
'' Inputs:      Row and Column of the edit, Whether to FinishEdit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsMarkets_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.vsMarkets.ComboCloseUp", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsRules_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    With vsRules
        If .SelectedRows > 0 Then
            .Row = .SelectedRow(0)
            .Col = Col
        End If
    End With
    SetBackColors vsRules

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.AfterSort", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsRules_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)

    'If Col = RGCol(eRGCol_Sequence) Then
        FinishEdit = True
    'End If

End Sub

Private Sub vsRules_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        RemoveRulesFromSystem
    ElseIf KeyCode = vbKeyInsert Then
        AddRules
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsRules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsRules
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            If Button = vbRightButton Then
                .RowSel = lMouseRow
                If .SelectedRows <= 1 Then .Row = lMouseRow
                
                Enable mnuLinkEntry, .TextMatrix(lMouseRow, RGCol(eRGCol_RuleUse)) = "1"
                Enable mnuTestEntry, .TextMatrix(lMouseRow, RGCol(eRGCol_RuleUse)) = "0"
                Enable mnuEditPyramid, chkPyramid = vbChecked
                
                PopupMenu mnuRule
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsMarkets_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With vsMarkets
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .Row = lMouseRow
                .RowSel = lMouseRow
            End If
            
            mnuChangeSymbol.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuMarketInfo.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuMarkets
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsMarkets.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsMarkets_AfterEdit
'' Description: After the user edits the Markets grid, resync as necessary
'' Inputs:      Row and Column of edited cell
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsMarkets_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strName As String
    Dim strPeriod As String
    Dim lIndex As Long
    Dim lGrpIndex As Long
    Dim strRecord As String
    
    Select Case Col
        Case MGCol(eMGCol_Period)
            With vsMarkets
                strName = .TextMatrix(Row, MGCol(eMGCol_ParmName))
                'strPeriod = GetPeriodStr(.TextMatrix(Row, MGCol(eMGCol_Period)))
                strPeriod = FixPeriod(.TextMatrix(Row, MGCol(eMGCol_Period)))
                
                .TextMatrix(Row, MGCol(eMGCol_Period)) = strPeriod
                
                SyncMarkets
            End With
            If tbToolbar.ToolBars("Wizard").Visible Then WizardDataTab
        
        Case MGCol(eMGCol_Symbol)
            With vsMarkets
                lGrpIndex = .ComboData(.ComboIndex)
                If lGrpIndex = 0& Then
                    If .TextMatrix(Row, MGCol(eMGCol_Symbol)) = "<New Symbol>" Then
                        .TextMatrix(Row, MGCol(eMGCol_Symbol)) = m.strSaveSymbol
                        BrowseMarkets
                    End If
                    .TextMatrix(Row, MGCol(eMGCol_GroupID)) = ""
                Else
                    ' Fill in any Market1, Daily, Weekly, etc. with the new info...
                    Select Case lGrpIndex Mod 1000
                        Case Asc("S")
                            .TextMatrix(Row, MGCol(eMGCol_GroupID)) = g.SymbolPool.SymbolGroups(lGrpIndex / 1000&).ID
                            .TextMatrix(Row, MGCol(eMGCol_Security)) = g.SymbolPool.SymbolGroups(lGrpIndex / 1000&).Name
                        Case Asc("F")
                            .TextMatrix(Row, MGCol(eMGCol_GroupID)) = g.SymbolPool.Filters(lGrpIndex / 1000&).ID
                            .TextMatrix(Row, MGCol(eMGCol_Security)) = g.SymbolPool.Filters(lGrpIndex / 1000&).Name
                        Case Asc("C")
                            .TextMatrix(Row, MGCol(eMGCol_GroupID)) = g.SymbolPool.Criterias(lGrpIndex / 1000&).ID
                            .TextMatrix(Row, MGCol(eMGCol_Security)) = g.SymbolPool.Criterias(lGrpIndex / 1000&).Name
                    End Select
                    .TextMatrix(Row, MGCol(eMGCol_MarketSymbol)) = ""
                    .TextMatrix(Row, MGCol(eMGCol_SecType)) = "Group"
                    .TextMatrix(Row, MGCol(eMGCol_Format)) = "CN"
                    .TextMatrix(Row, MGCol(eMGCol_SymbolPath)) = App.Path
                    .TextMatrix(Row, MGCol(eMGCol_Period)) = .TextMatrix(Row, MGCol(eMGCol_Period))
                    .TextMatrix(Row, MGCol(eMGCol_SymbolID)) = "0"
                    SyncMarkets
                                    
                    ' Update changes to INI securities section (most recent changes)...
                    If Not m.bLinkedToChart Then
                        strRecord = .TextMatrix(Row, MGCol(eMGCol_Symbol)) & "|" & AddSlash(App.Path) & "Data" & "|" & .TextMatrix(Row, MGCol(eMGCol_Security))
                        strRecord = strRecord & "|Group|CN|"
                        strRecord = strRecord & .TextMatrix(Row, MGCol(eMGCol_Period)) & "|" & .TextMatrix(Row, MGCol(eMGCol_GroupID)) & "|" & .TextMatrix(Row, MGCol(eMGCol_SymbolID))
                        SetIniFileProperty .TextMatrix(Row, MGCol(eMGCol_ParmName)), strRecord, "Securities", g.strIniFile
                    End If
                End If
            End With
            If tbToolbar.ToolBars("Wizard").Visible Then WizardDataTab
    End Select
    EnableToolbar Not m.bLinkedToChart
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.vsMarkets.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsMarkets_AfterSort
'' Description: After sorting the grid, make sure the back color is set right
'' Inputs:      Column sorted, Order sorted in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsMarkets_AfterSort(ByVal Col As Long, Order As Integer)

    SetBackColors vsMarkets

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsMarkets_BeforeEdit
'' Description: Set up the Period drop-down combo
'' Inputs:      Row and Column of cell being edited, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsMarkets_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Dim strParmName As String           ' Name of this parameter
    Dim strComboList As String          ' Combo list
    
    vsMarkets.ComboList = ""
    Select Case Col
        Case MGCol(eMGCol_Period)
            strParmName = UCase(vsMarkets.TextMatrix(Row, MGCol(eMGCol_ParmName)))
            Select Case strParmName
                Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY", "UNSPLIT"
                    Cancel = True
                Case Else
                    If Left(strParmName, 1) = Chr(34) And Right(strParmName, 1) = Chr(34) Then
                        Cancel = True
                    Else
                        With vsMarkets
                            If .TextMatrix(Row, MGCol(eMGCol_Format)) <> "GT" And .TextMatrix(Row, MGCol(eMGCol_Format)) <> "CN" Then
                                ' e.g. CSI and MS files
                                If .TextMatrix(Row, MGCol(eMGCol_ParmName)) <> "Market1" Then
                                    strComboList = "(Default)|Daily|Weekly|Monthly|Quarterly|Yearly"
                                Else
                                    strComboList = "Daily|Weekly|Monthly|Quarterly|Yearly"
                                End If
                            Else
                                ' TradeNav (and GT?)
                                If .TextMatrix(Row, MGCol(eMGCol_ParmName)) <> "Market1" Then
                                    strComboList = "|(Default)|5 minute|10 minute|15 minute|30 minute|60 minute|Daily|Weekly|Monthly|Quarterly|Yearly"
                                Else
                                    strComboList = "|5 minute|10 minute|15 minute|30 minute|60 minute|Daily|Weekly|Monthly|Quarterly|Yearly"
                                End If
                            
                                If g.FractZen.Allowed Then
                                    strComboList = strComboList & "|FractZen"
                                End If
                            End If
                            
                            .ComboList = strComboList
                        End With
                        EnableToolbar Not m.bLinkedToChart
                    End If
            End Select
        
        Case MGCol(eMGCol_Symbol)
            If UCase(vsMarkets.TextMatrix(Row, MGCol(eMGCol_ParmName))) = "MARKET1" Then
                If HasPlatinum(False) Then
                    m.strSaveSymbol = vsMarkets.TextMatrix(Row, MGCol(eMGCol_Symbol))
                    If Len(vsMarkets.TextMatrix(Row, MGCol(eMGCol_Symbol))) = 0 Then
                        vsMarkets.ComboList = "#0;<New Symbol>|" & SymbolGroups
                    ElseIf vsMarkets.TextMatrix(Row, MGCol(eMGCol_SecType)) = "Group" Then
                        vsMarkets.ComboList = "#0;<New Symbol>|" & SymbolGroups
                    Else
                        vsMarkets.ComboList = "#0;" & vsMarkets.TextMatrix(Row, MGCol(eMGCol_Symbol)) & "|#0;<New Symbol>|" & SymbolGroups
                    End If
                Else
                    Cancel = True
                End If
            Else
                Cancel = True
            End If
            
        Case Else
            Cancel = True
    
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.vsMarkets.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsMarkets_DblClick
'' Description: Allow the user to change the market they are running
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsMarkets_DblClick()
On Error GoTo ErrSection:
    
    BrowseMarkets
    EnableToolbar True
    
ErrExit:
    cmdBrowse.Enabled = True
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsMarkets.DblClick", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsMarkets_ValidateEdit
'' Description: Validate what the user entered in for the period
'' Inputs:      Row and Column edited, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsMarkets_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Period the user has chosen

    Select Case Col
        Case MGCol(eMGCol_Period)
            With vsMarkets
                If .TextMatrix(Row, MGCol(eMGCol_Format)) <> "GT" And .TextMatrix(Row, MGCol(eMGCol_Format)) <> "CN" Then
                    ' e.g. CSI and MS files
                    If IsIntraday(GetPeriodicity(.EditText)) Then
                        Cancel = True
                    End If
                End If
                
                If Cancel = False Then
                    strPeriod = .EditText
                    
                    strPeriod = FixPeriod(strPeriod)
                    If UCase(strPeriod) = "FRACTZEN" And Not g.FractZen.Allowed Then
                        Cancel = True
                        InfBox "You are not authorized to use FractZen bars", "!", , kErrorCaption
                    End If
                    
                    If strPeriod <> .EditText Then
                        .EditText = strPeriod
                    End If
                End If
            End With
    End Select
    If Cancel = False Then EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.vsMarkets.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_AfterEdit
'' Description: After the user edits the grid, resync everything
'' Inputs:      Row and Column of the cell being edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim i&, iRow&, iRuleID&, iSeq&
    Dim aRows As New cGdArray
    Dim strTemp As String
    
    EnableToolbar True
            
    vsRules.Redraw = flexRDNone
    If Col = RGCol(eRGCol_Linked) Then
        ShowLinkInputs
    ElseIf Col = RGCol(eRGCol_Sequence) Then
        iRuleID = Val(vsRules.TextMatrix(Row, RGCol(eRGCol_RuleID)))
        i = Val(vsRules.TextMatrix(Row, Col))
        iSeq = m.System.Rules.Item(Str(iRuleID)).Seq
        If i < iSeq Then
            i = -i
        ElseIf i > iSeq Then
            i = -(i + 1)
        End If
        If i <> iSeq Then
            m.System.Rules.Item(Str(iRuleID)).Seq = i
            FixSequence
            ' resort by priority (sequence)
            vsRules.Sort = flexSortNumericAscending
            SetBackColors vsRules
            For iRow = vsRules.FixedRows To vsRules.Rows - 1
                If Val(vsRules.TextMatrix(iRow, RGCol(eRGCol_RuleID))) = iRuleID Then
                    vsRules.Row = iRow
                    Exit For
                End If
            Next
        End If
    ElseIf Col = RGCol(eRGCol_Alt) Or Col = RGCol(eRGCol_Selected) Then
        ' 6/4/01: do all selected rows for USE and ALT columns
        aRows.Create eGDARRAY_Longs
        If vsRules.SelectedRows > 1 And (Col = RGCol(eRGCol_Selected) Or Col = RGCol(eRGCol_Alt)) Then
            For i = 0 To vsRules.SelectedRows - 1
                If Not vsRules.RowHidden(vsRules.SelectedRow(i)) Then
                    aRows.Add vsRules.SelectedRow(i)
                End If
            Next
        Else
            aRows.Add Row
        End If
        
        For i = 0 To aRows.Size - 1
            iRow = aRows(i)
            iRuleID = Val(vsRules.TextMatrix(iRow, RGCol(eRGCol_RuleID)))
            
            ' If selected is turned on, make sure to turn alternating off
            If Col = RGCol(eRGCol_Selected) Then
                If iRow <> Row Then
                    CheckedCell(vsRules, iRow, RGCol(eRGCol_Selected)) = CheckedCell(vsRules, Row, RGCol(eRGCol_Selected))
                End If
                If CheckedCell(vsRules, iRow, RGCol(eRGCol_Selected)) Then
                    CheckedCell(vsRules, iRow, RGCol(eRGCol_Alt)) = False
                End If
            End If
            
            ' If alternating is turned on, make sure to turn selected off
            If Col = RGCol(eRGCol_Alt) Then
                If iRow <> Row Then
                    CheckedCell(vsRules, iRow, RGCol(eRGCol_Alt)) = CheckedCell(vsRules, Row, RGCol(eRGCol_Alt))
                End If
                If CheckedCell(vsRules, iRow, RGCol(eRGCol_Alt)) Then
                    CheckedCell(vsRules, iRow, RGCol(eRGCol_Selected)) = False
                End If
            End If
            
            ' set properties of the rule
            m.System.Rules.Item(Str(iRuleID)).Selected = CheckedCell(vsRules, iRow, RGCol(eRGCol_Selected))
            m.System.Rules.Item(Str(iRuleID)).Alternate = CheckedCell(vsRules, iRow, RGCol(eRGCol_Alt))
            
            ' If both selected and alternating are off, hide the inputs and turn the
            ' optimization off on them
            If Not CheckedCell(vsRules, iRow, RGCol(eRGCol_Alt)) And _
                Not CheckedCell(vsRules, iRow, RGCol(eRGCol_Selected)) Then
                HideInputs iRuleID, True
            Else
                HideInputs iRuleID, False
            End If
        Next
        FixInputDisplay
    End If
        
ErrExit:
    vsRules.Redraw = flexRDBuffered
    Exit Sub

ErrSection:
    vsRules.Redraw = flexRDBuffered
    RaiseError "frmSystemManager.vsRules.AfterEdit", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_BeforeEdit
'' Description: Don't allow editing of the rulename or shared columns
'' Inputs:      Row and Column of cell being edited, Whether to cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
       
    Dim s$, i&

    'If Col = RGCol(eRGCol_RuleName) Or Col = RGCol(eRGCol_Action) Or Col = RGCol(eRGCol_Linked) Then
    '    Cancel = True
    If Col = RGCol(eRGCol_Sequence) Then
        If vsRules.TextMatrix(Row, RGCol(eRGCol_RuleUse)) = 0 Then
            For i = 1 To m.lNumEntries
                s = s & "|" & Str(i)
            Next
        Else
            For i = m.lNumEntries + 1 To m.System.Rules.Count
                s = s & "|" & Str(i)
            Next
        End If
        vsRules.ComboList = Mid(s, 2)
    ElseIf Col = RGCol(eRGCol_Selected) Or Col = RGCol(eRGCol_Alt) Then
        vsRules.ComboList = ""
    Else
        Cancel = True
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_Click
'' Description: If the user clicks on the pyramid column, bring up the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_Click()
On Error GoTo ErrSection:
    
    With vsRules
        If .MouseRow >= .FixedRows Then
            If .MouseCol = RGCol(eRGCol_PyramidInfo) Then
                .Row = .MouseRow
                EditPyramidInfo
                EnableToolbar True
            ElseIf .MouseCol = RGCol(eRGCol_Linked) Then
                If .TextMatrix(.MouseRow, RGCol(eRGCol_RuleUse)) = "1" Then
                    .Row = .MouseRow
                    LinkToEntry
                End If
            End If
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_DblClick
'' Description: Allow the user to edit the rule (or pyramid information)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_DblClick()
On Error GoTo ErrSection:
    
    With vsRules
        If .MouseRow >= .FixedRows Then
            .Row = .MouseRow
            If .MouseCol = RGCol(eRGCol_RuleName) Then
                EditRule
            ElseIf .MouseCol = RGCol(eRGCol_PyramidInfo) Then
                EditPyramidInfo
                EnableToolbar True
            End If
        End If
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.DblClick", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_MouseMove
'' Description: Show an appropriate tooltip as the user moves over the grid
'' Inputs:      Mouse Button Pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseCol As Long
    Dim lMouseRow As Long
    
    lMouseCol = vsRules.MouseCol
    lMouseRow = vsRules.MouseRow
    
    If lMouseCol = RGCol(eRGCol_RuleName) And lMouseRow >= vsRules.FixedRows Then
        GridTooltip vsRules, RGCol(eRGCol_RuleName)
    Else
        GridTooltip vsRules
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_ValidateEdit
'' Description: Validate what the user entered
'' Inputs:      Row and Column of cell being edited, Whether to cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If m.bShowPassword = True Then
        If Not g.Security.CanSave(m.System.SecurityLevel, m.System.Password) Then
            Cancel = True
            ' need to get it out of edit mode (if an edit window exists)
            If vsRules.EditWindow <> 0 Then
                SendKeys "{Esc}"
            End If
        Else
            m.bShowPassword = False
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsRules.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsRules_AfterRowColChange
'' Description: As the user changes rows in the grid, update the preview
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsRules_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:
    
    Preview NewRow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.vsRules.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSettings_Click
'' Description: Allow the user to change the Market information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSettings_Click()
On Error GoTo ErrSection:
    
    MarketInfo
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.cmdSettings.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterEdit
'' Description: Show/Hide the optimization columns appropriately
'' Inputs:      Row and Column of cell being edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim X As Long
    
    'Show optimization From/To/Step columns
    If Col = IGCol(eIGCol_IfOptimize) Then
        With vsInputs
            ' If inputs linked, need to turn the optimization on/off for
            ' all linked inputs now
            If m.bHideDuplicates Then
                For X = .FixedRows To .Rows - 1
                    If X <> Row Then
                        If SameInput(X, Row) Then
                            .TextMatrix(X, IGCol(eIGCol_IfOptimize)) = .TextMatrix(Row, IGCol(eIGCol_IfOptimize))
                        End If
                    End If
                Next X
            End If
        End With
        HideOptColumns
    End If
    
    If Col = IGCol(eIGCol_OptToValue) Then
        With vsInputs
            If Val(.TextMatrix(Row, IGCol(eIGCol_OptToValue))) < Val(.TextMatrix(Row, IGCol(eIGCol_OptFromValue))) Then
                .TextMatrix(Row, IGCol(eIGCol_OptToValue)) = .TextMatrix(Row, IGCol(eIGCol_OptFromValue))
            End If
        End With
    End If
    
    ColorInputCell Row, Col
    vsInputs.TextMatrix(Row, Col) = FormatNum(ValOfText(vsInputs.TextMatrix(Row, Col)))
    EnableToolbar True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsInputs.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_ValidateEdit
'' Description: Validate what the user entered
'' Inputs:      Row and Column of cell being edited, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Dim dInputValue As Double
    Dim dFromVal As Double
    Dim dToVal As Double
    Dim X As Long
    
    'Get input values
    dInputValue = ValOfText(vsInputs.EditText)
    dFromVal = ValOfText(vsInputs.TextMatrix(Row, IGCol(eIGCol_FromVal)))
    dToVal = ValOfText(vsInputs.TextMatrix(Row, IGCol(eIGCol_ToVal)))
    
    Select Case Col
        Case IGCol(eIGCol_IfOptimize)
            ' must have PLATINUM in order to optimize systems
            If dInputValue Then ' Not CheckedCell(vsInputs, Row, Col) Then
                If Not HasPlatinum(True, "Optimizing strategies") Then
                    Cancel = True
                End If
            End If
        
        Case IGCol(eIGCol_InputValue), IGCol(eIGCol_OptFromValue), IGCol(eIGCol_OptToValue) ', IGCol(eIGCol_OptStepValue)
            If IsNumeric(dInputValue) Then
                If dFromVal <> 0 Or dToVal <> 0 Then
                    If dInputValue < dFromVal Or dInputValue > dToVal Then
                        Cancel = True
                        Err.Raise vbObjectError + 1000, , _
                            "Please enter a value between " & _
                            Format(dFromVal, "general number") & " and " & _
                            Format(dToVal, "general number")
                    End If
                Else
                    If (dInputValue < -100000000000# Or _
                        dInputValue > 100000000000#) Then
                        Cancel = True
                        Err.Raise vbObjectError + 1000, , _
                            "Please enter a value between " & _
                            "-100,000,000,000 and 100,000,000,000"
                    End If
                End If
            End If
            
        Case IGCol(eIGCol_OptStepValue)
            If (dInputValue = 0# Or dInputValue > 100000000000#) Then
                Cancel = True
                Err.Raise vbObjectError + 1000, , _
                    "Please enter a value between " & _
                    "0 and 100,000,000,000"
            End If
        
        Case Else
            Cancel = True
            Exit Sub
    End Select
    
    If Not Cancel Then
        EnableToolbar True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsInputs.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeEdit
'' Description: Only allow the user to edit certain columns
'' Inputs:      Row and Column of cell being edited, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    'No changes allowed for bar structure inputs
    If Val(vsInputs.TextMatrix(Row, IGCol(eIGCol_ParmTypeID))) = kSN_RetBars Then
        Cancel = True
    End If
    
    'Only allow changes to "Value" and optimization columns
    '(Cols-1 is the extra extended column which the focus can move into)
    If Col <> IGCol(eIGCol_InputValue) And Col <> IGCol(eIGCol_IfOptimize) And _
       Col <> IGCol(eIGCol_OptFromValue) And Col <> IGCol(eIGCol_OptToValue) And _
       Col <> IGCol(eIGCol_OptStepValue) And Col <> IGCol(eIGCol_FromVal) And _
       Col <> IGCol(eIGCol_ToVal) Then
        Cancel = True
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.vsInputs.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OptimizationUpdate
'' Description: Takes the current values chosen from the optimizer
'' Inputs:      Rule Names, Input Names, Input Values, Alt Rule Names,
''              Alt Rules Use
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OptimizationUpdate(pRuleNames() As String, _
    pInputNames() As String, pInputValues() As Double, _
    pAltRuleNames() As String, pAltRulesUse() As Boolean)
On Error GoTo ErrSection:

    Dim X As Long
    Dim Y As Long
    Dim lRedraw As Long
    
    With vsRules
        lRedraw = .Redraw
        vsRules.Redraw = flexRDNone
    
        ' Set the Selected flags on the Rules grid
        For Y = 1 To UBound(pInputValues)
            For X = 1 To .Rows - 1
                If .TextMatrix(X, RGCol(eRGCol_RuleName)) = pRuleNames(Y) Then
                    If pInputValues(Y) = -999999999999# Then
                        .TextMatrix(X, RGCol(eRGCol_Selected)) = False
                    End If
                    .TextMatrix(X, RGCol(eRGCol_Alt)) = False
                    Exit For
                End If
            Next X
        Next Y
        
        ' Set the Alternating flags on the Rules grid
        For Y = 1 To UBound(pAltRuleNames)
            For X = 1 To .Rows - 1
                If .TextMatrix(X, RGCol(eRGCol_RuleName)) = pAltRuleNames(Y) Then
                    .TextMatrix(X, RGCol(eRGCol_Selected)) = pAltRulesUse(Y)
                    .TextMatrix(X, RGCol(eRGCol_Alt)) = False
                    Exit For
                End If
            Next X
        Next Y
    
        RefreshRulesGrid
        
        .Redraw = lRedraw
    End With

    For Y = 1 To UBound(pInputNames)
        With vsInputs
            .Redraw = flexRDNone
            For X = .FixedRows To .Rows - 1
                If .TextMatrix(X, IGCol(eIGCol_RuleName)) = pRuleNames(Y) Or pRuleNames(Y) = "" Then
                    If .TextMatrix(X, IGCol(eIGCol_InputName)) = pInputNames(Y) Then
                        If pInputValues(Y) <> -999999999999# Then
                            .TextMatrix(X, IGCol(eIGCol_InputValue)) = FormatNum(pInputValues(Y))
                        End If
                        .TextMatrix(X, IGCol(eIGCol_IfOptimize)) = False
                        If pRuleNames(Y) <> "" Then Exit For
                    End If
                End If
            Next X
            .Redraw = flexRDBuffered
        End With
    Next Y
    
    HideOptColumns
    RefreshInputsGrid
    EnableToolbar True

    HideOrShowInputs

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.OptimizationUpdate", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRules
'' Description: Add rules to the system
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRules()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFilter As Long                 'filter flag for Wizard mode
    Dim RulesToAdd As New cRules        ' Rules to add to the system
    Dim NewRule As New cRule            ' Copy of the Rule to Add
    
    If Not HasPlatinum(True) Then Exit Sub
    If Not ShowPassword Then Exit Sub
        
    lFilter = -1
    If tbToolbar.ToolBars("Wizard").Visible Then
        If optSignals(0).Value = True Then
            lFilter = 0
        ElseIf optSignals(1).Value = True Then
            lFilter = 1
        ElseIf optSignals(2).Value = True Then
            lFilter = 2
        ElseIf optSignals(3).Value = True Then
            lFilter = 3
        End If
    End If
        
    Set RulesToAdd = frmToolbox.ShowAddRules(m.System.SystemNumber, m.System.LibraryID, m.System.Rules, lFilter)
    If Not RulesToAdd Is Nothing Then
        For lIndex = 1 To RulesToAdd.Count
            With RulesToAdd.Item(lIndex)
                Set NewRule = RulesToAdd.Item(lIndex).MakeCopy(NextRuleID, m.System.SystemNumber)
                NewRule.LibraryID = m.System.LibraryID
                If NewRule.SecurityLevel < 2 Then
                    NewRule.CategoryID = 0
                    NewRule.SecurityLevel = m.System.SecurityLevel
                    NewRule.Password = m.System.Password
                    NewRule.CannotDelete = False
                End If
            End With
                
            AddRule NewRule
        Next lIndex
    
        ShowLinkInputs
        RefreshRulesGrid
        
        EnableToolbar True
    End If

ErrExit:
    Set RulesToAdd = Nothing
    Set NewRule = Nothing
    Exit Sub
    
ErrSection:
    Set RulesToAdd = Nothing
    Set NewRule = Nothing
    RaiseError "frmSystemManager.AddRules", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideOrShowInputs
'' Description: Hide or show inputs as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideOrShowInputs()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    ' Hide inputs if necessary...
    With m.System.Rules
        For lIndex = 1 To .Count
            If .Item(lIndex).Selected Or .Item(lIndex).Alternate Then
                HideInputs .Item(lIndex).RuleID, False
            Else
                HideInputs .Item(lIndex).RuleID, True
            End If
        Next lIndex
        FixInputDisplay
        
        If m.bHasDuplicates Then
            If chkLinkInputs = vbChecked Then
                HideDuplicateInputs
            Else
                ShowDuplicateInputs
            End If
        End If
        
        ShowLinkInputs
        HideOptColumns
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.HideOrShowInputs", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback for the print preview form
'' Inputs:      Arguments
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)

    Dim lIndex As Long, lRow As Long, lCol As Long, lTemp As Long
    Dim bEntry As Boolean
    Dim Rule As New cRule
    Dim astrLinked As cGdArray
    Dim lLink As Long
    Dim lRule As Long
    Dim lNum As Long
    Dim lIndex2 As Long
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        '.Font.Name = "Times New Roman"
        .Font.Name = "Microsoft Sans Serif"
        .Font.Name = "Arial"
        .Font.Size = 13
        .Font.Bold = True
        .FontUnderline = True
        .Text = "Strategy:"
        .FontUnderline = False
        .Text = "    " & Trim(m.strName) & vbLf & vbLf
        .Font.Bold = False
        .Font.Size = 9
        
        .Text = "Commission/Slippage: "
        .Font.Bold = True
        .Text = Format(txtCommission.Text, "$#,##0.00") & vbLf
        .Font.Bold = False
        .Text = "Bars Required to Test: " & txtBarsLoadedBeforeTrading.Text & vbLf
        .Text = "Estimated Length of Longest Trade: " & txtBarsTradedBeforeOrders.Text & vbLf
        
        If chkPyramid.Value = vbChecked Then
            .Text = "Pyramiding On (Max Number of Pyramiding Signals=" & txtTradeDepth.Text & ")" & vbLf
        Else
            .Text = "Pyramiding Off" & vbLf
        End If
        
        .Text = "Date Range:  "
        .Font.Bold = True
        .Text = Format(dtpFromDate.Value, "mm/dd/yyyy") & " - " & Format(dtpToDate.Value, "mm/dd/yyyy") & vbLf
        .Font.Bold = False
        
        .Text = vbLf
        
        .Font.Size = 10
        .Font.Bold = True
        .FontUnderline = True
        .Text = "Data Info:" & vbLf & vbLf
        .FontUnderline = False
        .Font.Bold = False
        .Font.Size = 9
        
        .LineSpacing = 100
        .StartTable
        .TableBorder = tbNone
        .TableCell(tcCols) = 5
        
        lIndex = 0&
        For lRow = 0 To vsMarkets.Rows - 1
            If vsMarkets.RowHidden(lRow) = False Then lIndex = lIndex + 1
        Next lRow
        .TableCell(tcRows) = lIndex
        
        .TableCell(tcFontUnderline, 1, 1) = True
        .TableCell(tcFontUnderline, 1, 2) = True
        .TableCell(tcFontUnderline, 1, 3) = True
        .TableCell(tcFontUnderline, 1, 4) = True
        .TableCell(tcFontUnderline, 1, 5) = True
        .TableCell(tcColWidth, , 1) = 1440
        .TableCell(tcColWidth, , 2) = 2880
        .TableCell(tcColWidth, , 3) = 1440
        .TableCell(tcColWidth, , 4) = 1440
        .TableCell(tcColWidth, , 5) = 1440
        
        lIndex = 1&
        For lRow = 0 To vsMarkets.Rows - 1
            If vsMarkets.RowHidden(lRow) = False Then
                .TableCell(tcText, lIndex, 1) = vsMarkets.Cell(flexcpText, lRow, 0)
                .TableCell(tcText, lIndex, 2) = vsMarkets.Cell(flexcpText, lRow, 1)
                .TableCell(tcText, lIndex, 3) = vsMarkets.Cell(flexcpText, lRow, 2)
                .TableCell(tcText, lIndex, 4) = vsMarkets.Cell(flexcpText, lRow, 3)
                .TableCell(tcText, lIndex, 5) = vsMarkets.Cell(flexcpText, lRow, 5)
                If lRow > 0 Then
                    .TableCell(tcFontBold, lIndex, 3) = True
                    .TableCell(tcFontBold, lIndex, 5) = True
                End If
            
                lIndex = lIndex + 1
            End If
        Next lRow
        .EndTable
        
        '.Text = vbCrLf
        .Text = vbLf
        
        .Font.Size = 10
        .Font.Bold = True
        .FontUnderline = True
        .Text = "Rules:" & vbLf '& vbCrLf
        .FontUnderline = False
        .Font.Bold = False
        .Font.Size = 9
        
        If m.System.SecurityLevel < 2 Then
            '.DrawLine .MarginLeft, .CurrentY, .PageWidth - .MarginRight, .CurrentY
            '.Text = vbCrLf
            lNum = (.PageWidth - .MarginLeft - .MarginRight) / .TextWidth("_")
            .CurrentX = .MarginLeft
            For lIndex2 = 0 To lNum - 1
                .Text = "_"
            Next lIndex2
            .Text = vbLf & vbLf
            
            For lIndex = 1 To vsRules.Rows - 1
                ' TLB 4/17/2015: don't bother printing rules which are completely unused
                ' (i.e. must be either set to USE or to Alternating)
                If CheckedCell(vsRules, lIndex, RGCol(eRGCol_Selected)) = True _
                    Or CheckedCell(vsRules, lIndex, RGCol(eRGCol_Alt)) = True Then
                
                    .StartTable
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 2
                    bEntry = (InStr(vsRules.TextMatrix(lIndex, RGCol(eRGCol_Action)), "Entry") > 0)
                    'If chkPyramid.Value = vbChecked Then
                    '    If bEntry Then
                    '        .TableCell(tcRows) = 6
                    '    Else
                    '        .TableCell(tcRows) = 7
                    '    End If
                    'Else
                    '    If bEntry Then
                    '        .TableCell(tcRows) = 5
                    '    Else
                    '        .TableCell(tcRows) = 6
                    '    End If
                    'End If
                    .TableCell(tcRows) = 2
                    .TableCell(tcColWidth, , 1) = "1.2in"
                    .TableCell(tcColWidth, , 2) = "5.3in"
                    
                    .TableCell(tcText, 1, 1) = "RuleName:"
                    .TableCell(tcText, 1, 2) = vsRules.Cell(flexcpText, lIndex, RGCol(eRGCol_RuleName))
                    .TableCell(tcFontBold, 1, 2) = True
                    .TableCell(tcText, 2, 1) = "Rule Text: "
                    If g.Security.CanPreview(vsRules.Cell(flexcpText, lIndex, RGCol(eRGCol_SecurityLevel))) = True Then
                        '.TableCell(tcText, 2, 2) = Rule.GetRTF(vsRules.Cell(flexcpText, lIndex, RGCol(eRGCol_Preview)))
                        rtfTest.TextRTF = Rule.GetRTF(vsRules.Cell(flexcpText, lIndex, RGCol(eRGCol_Preview)))
                        .TableCell(tcText, 2, 2) = Replace(rtfTest.Text, vbCrLf, vbLf)
                    Else
                        .TableCell(tcText, 2, 2) = "Not authorized to view"
                    End If
                    If CheckedCell(vsRules, lIndex, RGCol(eRGCol_Selected)) = False Then
                        .TableCell(tcRows) = .TableCell(tcRows) + 1
                        .TableCell(tcText, .TableCell(tcRows), 1) = "Use:"
                        .TableCell(tcText, .TableCell(tcRows), 2) = "OFF"
                    End If
                    
                    If CheckedCell(vsRules, lIndex, RGCol(eRGCol_Alt)) = True Then
                        .TableCell(tcRows) = .TableCell(tcRows) + 1
                        .TableCell(tcText, .TableCell(tcRows), 1) = "Alternating:"
                        .TableCell(tcText, .TableCell(tcRows), 2) = "ON"
                    End If
                    
                    .TableCell(tcRows) = .TableCell(tcRows) + 1
                    .TableCell(tcText, .TableCell(tcRows), 1) = "Action:"
                    .TableCell(tcText, .TableCell(tcRows), 2) = vsRules.TextMatrix(lIndex, RGCol(eRGCol_Action))
                    If InStr(vsRules.TextMatrix(lIndex, RGCol(eRGCol_Action)), "Exit") > 0 Then
                        .TableCell(tcRows) = .TableCell(tcRows) + 1
                        .TableCell(tcText, .TableCell(tcRows), 1) = "Exit on Entry Bar:"
                        If CheckedCell(vsRules, lIndex, RGCol(eRGCol_ExitOnEntryBar)) Then
                            .TableCell(tcText, .TableCell(tcRows), 2) = "True"
                        Else
                            .TableCell(tcText, .TableCell(tcRows), 2) = "False"
                        End If
                    End If
                    If chkPyramid = vbChecked Then
                        .TableCell(tcRows) = .TableCell(tcRows) + 1
                        .TableCell(tcText, .TableCell(tcRows), 1) = "Pyramiding:"
                        .TableCell(tcText, .TableCell(tcRows), 2) = vsRules.Cell(flexcpText, lIndex, RGCol(eRGCol_PyramidInfo))
                    End If
                    
                    If Not bEntry Then
                        If CheckedCell(vsRules, lIndex, RGCol(eRGCol_Linked)) = True Then
                            Set astrLinked = New cGdArray
                            astrLinked.Create eGDARRAY_Strings
                            astrLinked.SplitFields vsRules.TextMatrix(lIndex, RGCol(eRGCol_LinkedRules)), ","
                            If astrLinked(0) = "" Then astrLinked.Remove 0
                            
                            lRow = .TableCell(tcRows) + 1
                            If astrLinked.Size = 1 Then
                                .TableCell(tcRows) = lRow
                                .TableCell(tcText, lRow, 1) = "Linked To Entry:"
                            ElseIf astrLinked.Size > 0 Then
                                .TableCell(tcRows) = lRow + astrLinked.Size - 1
                                .TableCell(tcText, lRow, 1) = "Linked To Entries:"
                            End If
                            
                            For lLink = 0 To astrLinked.Size - 1
                                If Len(astrLinked(lLink)) > 0 Then
                                    For lRule = vsRules.FixedRows To vsRules.Rows - 1
                                        If vsRules.TextMatrix(lRule, RGCol(eRGCol_RuleID)) = astrLinked(lLink) Then
                                            .TableCell(tcText, lRow, 2) = vsRules.TextMatrix(lRule, RGCol(eRGCol_RuleName))
                                            lRow = lRow + 1
                                            Exit For
                                        End If
                                    Next lRule
                                End If
                            Next lLink
                        End If
                    End If
                    .EndTable
                    
                    '.Text = vbCrLf
                    '.Text = vbLf
                    
                    ' Print INPUTS
                    lTemp = 0&
                    For lRow = 1 To vsInputs.Rows - 1
                        If vsInputs.TextMatrix(lRow, IGCol(eIGCol_RuleID)) = vsRules.TextMatrix(lIndex, RGCol(eRGCol_RuleID)) Then
                            lTemp = lTemp + 1
                        End If
                    Next lRow
                    
                    If lTemp > 0 Then
                        .Text = vbLf
                    
                        .StartTable
                        .TableBorder = tbNone
                        
                        .TableCell(tcCols) = 6
                        .TableCell(tcRows) = lTemp + 1
                        
                        .TableCell(tcColWidth, , 1) = 2880
                        .TableCell(tcColWidth, , 2) = 1080
                        .TableCell(tcColWidth, , 3) = 1080
                        .TableCell(tcColWidth, , 4) = 1080
                        .TableCell(tcColWidth, , 5) = 1080
                        .TableCell(tcColWidth, , 6) = 1080
                        
                        .TableCell(tcFontUnderline, 1, 1) = True
                        .TableCell(tcFontUnderline, 1, 2) = True
                        .TableCell(tcFontUnderline, 1, 3) = True
                        .TableCell(tcFontUnderline, 1, 4) = True
                        .TableCell(tcFontUnderline, 1, 5) = True
                        .TableCell(tcFontUnderline, 1, 6) = True
                        .TableCell(tcText, 1, 1) = "Input"
                        .TableCell(tcText, 1, 2) = "Value"
                        .TableCell(tcText, 1, 3) = "Optimize"
                        .TableCell(tcText, 1, 4) = "From Val"
                        .TableCell(tcText, 1, 5) = "To Val"
                        .TableCell(tcText, 1, 6) = "Step Val"
                        
                        lTemp = 2
                        For lRow = 1 To vsInputs.Rows - 1
                            If vsInputs.TextMatrix(lRow, IGCol(eIGCol_RuleID)) = vsRules.TextMatrix(lIndex, RGCol(eRGCol_RuleID)) Then
                                .TableCell(tcText, lTemp, 1) = vsInputs.TextMatrix(lRow, 1)
                                .TableCell(tcText, lTemp, 2) = vsInputs.TextMatrix(lRow, 2)
                                If CheckedCell(vsInputs, lRow, 5) Then
                                    .TableCell(tcText, lTemp, 3) = "ON"
                                    .TableCell(tcText, lTemp, 4) = vsInputs.TextMatrix(lRow, 3)
                                    .TableCell(tcText, lTemp, 5) = vsInputs.TextMatrix(lRow, 4)
                                    .TableCell(tcText, lTemp, 6) = vsInputs.TextMatrix(lRow, 6)
                                Else
                                    .TableCell(tcText, lTemp, 3) = "OFF"
                                End If
                                
                                .TableCell(tcFontBold, lTemp, 1) = True
                                .TableCell(tcFontBold, lTemp, 2) = True
                                
                                lTemp = lTemp + 1
                            End If
                        Next lRow
                        
                        .EndTable
                        '.Text = vbLf
                    End If
                    
                    '.DrawLine .MarginLeft, .CurrentY, .PageWidth - .MarginRight, .CurrentY
                    '.Text = vbLf
                    .CurrentX = .MarginLeft
                    For lIndex2 = 0 To lNum - 1
                        .Text = "_"
                    Next lIndex2
                    .Text = vbLf & vbLf
                End If
            Next lIndex
        Else
            .Text = "Not authorized to see rules"
        End If
        
        .EndDoc
    End With

    Set Rule = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLinkInputs
'' Description: Show/Hide the linked inputs check box as appropriate
'' Inputs:      Whether to Show or Hide
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLinkInputs()
On Error Resume Next

    Dim lFrameBottom As Long
    
    lFrameBottom = fraLinkInputs.Top + fraLinkInputs.Height
    fraLinkInputs.Visible = m.bHasDuplicates

    With vsInputs
        If m.bHasDuplicates Then
            .Move .Left, lFrameBottom, vsIndexTab1.ClientWidth, _
                    vsIndexTab1.ClientHeight - lFrameBottom
        Else
            .Move .Left, fraLinkInputs.Top, vsIndexTab1.ClientWidth, _
                    vsIndexTab1.ClientHeight
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowPassword
'' Description: Determines if a user has permission or not
'' Inputs:      None
'' Returns:     True if has permission, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ShowPassword() As Boolean
On Error GoTo ErrSection:

    If m.bShowPassword = True Then
        If g.Security.CanSave(m.System.SecurityLevel, m.System.Password) Then
            m.bShowPassword = False
            ShowPassword = True
        End If
    Else
        ShowPassword = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.ShowPassword", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      System ID to load, Whether to load or just show form
'' Returns:     True if changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lSystemID As Long, Optional ByVal bJustShowForm As Boolean = False, _
    Optional ByVal bModal As Boolean = True, Optional ByVal strTextToPaste = "", _
    Optional ByVal bNew As Boolean = False, _
    Optional ByVal bSaveLinkFlag As Boolean = False, _
    Optional ByRef Chart As cChart = Nothing) As Boolean
On Error GoTo ErrSection:
        
    Dim dLastModified As Double         ' Date to compare to see if system was changed
    Dim strUseWizard$
    Dim i&, s$
    Dim bIsOwner As Boolean             ' Is the current user an owner of this object?
        
    m.bModal = bModal
    m.bNewStrategy = bNew
    If Not bJustShowForm Then
        If lSystemID = 0 Then
            Add
        Else
            LoadRec lSystemID
        End If
    End If
    
    If m.System.IsGuru = True Then
        bIsOwner = IsOwnerOfGuruObject(m.System.LibraryID)
    Else
        bIsOwner = True
    End If
    
    If bIsOwner Then
        m.bSaveLinkFlag = bSaveLinkFlag
        'Default to 1st tab (if can view and if has Platinum)
        If m.System.SecurityLevel < 2 And HasPlatinum(False) Then
            vsIndexTab1.CurrTab = Tabs(eSMTab_Rules)
        Else
            vsIndexTab1.CurrTab = Tabs(eSMTab_Inputs)
        End If
        
        Screen.MousePointer = vbDefault
        dLastModified = m.System.LastModified
           
        If Not Chart Is Nothing Then CenterFormOnChart Me, Chart            '6499
        ShowForm Me, bModal, frmMain, , ALT_GRID_ROW_COLOR
        If bModal Then
            If m.System.LastModified > dLastModified Then ShowMe = True
        ElseIf m.bNewStrategy Then
            ' see if wizard setting is stored (i.e. don't ask anymore)
            strUseWizard = GetIniFileProperty("StrategyWizard", "", "DontAsk", g.strIniFile)
            If Len(strUseWizard) = 0 Then
                strUseWizard = "Would you like to use the Strategy Wizard| to help create the new strategy?"
                strUseWizard = InfBox(strUseWizard, "?", "+Wizard|-No", "Strategy Wizard", , , , , , , , , True)
                If InStr(strUseWizard, "-") > 0 Then
                    ' don't ask anymore, so store this wizard setting for future use
                    Call SetIniFileProperty("StrategyWizard", Left(strUseWizard, 1), "DontAsk", g.strIniFile)
                End If
            End If
            If UCase(Left(strUseWizard, 1)) = "W" Then
                ' use the wizard to start the new strategy
                tbToolbar.Redraw = False
                tbToolbar.ToolBars("General").Visible = False
                tbToolbar.ToolBars("Wizard").Visible = True
                SetWizardBack
                tbToolbar.Redraw = True
            End If
        ElseIf 0 Then ' IsIDE Then
            On Error Resume Next
            s = ""
            i = m.System.Markets.Index("Market1")
            If i > 0 Then
                s = m.System.Markets(i).Symbol
                If Len(s) > 0 Then
                    If IsForex(s) Or SecurityType(s) = "S" Then
                        vsIndexTab1.CurrTab = Tabs(eSMTab_Settings)
                        s = "Some new settings have recently been added related to backtesting Stocks and Forex symbols."
                        InfBox s, "i", , "Please Note ..."
                    End If
                End If
            End If
            On Error GoTo ErrSection:
        End If
           
        If Len(strTextToPaste) > 0 Then
            If Not NewRule(strTextToPaste) Then Unload Me
        End If
        
        ShowMe = True
    Else
        InfBox "You are not authorized to view this strategy", "!", , "Strategy Error"
        ShowMe = False
    End If

ErrExit:
    If (bModal And Not m.bOptimizing) Or (bIsOwner = False) Then
        Unload Me
    End If
    Exit Function

ErrSection:
    Unload Me
    RaiseError "frmSystemManager.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Filter
'' Description: Filter the rules grid according to the option chosen
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Filter()
On Error GoTo ErrSection:
    
    Dim lIndex As Long               ' Index into a for loop
    Dim bVisibleRows As Boolean         ' Are there visible rows?
    Dim lRedraw As Long                 ' Current status of the grid redraw
    
    With vsRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Hide all of the rows
        For lIndex = .FixedRows To .Rows - 1
            .RowHidden(lIndex) = True
        Next lIndex
        
        ' Show the rows based on the option selected
        For lIndex = 1 To .Rows - 1
            Select Case True
                Case optSignals(optAll)
                    .RowHidden(lIndex) = False
                Case optSignals(optLong)
                    If .TextMatrix(lIndex, RGCol(eRGCol_Action)) = "Long Entry" Then .RowHidden(lIndex) = False
                Case optSignals(optLongExit)
                    If .TextMatrix(lIndex, RGCol(eRGCol_Action)) = "Long Exit" Then .RowHidden(lIndex) = False
                Case optSignals(optShort)
                    If .TextMatrix(lIndex, RGCol(eRGCol_Action)) = "Short Entry" Then .RowHidden(lIndex) = False
                Case optSignals(optShortExit)
                    If .TextMatrix(lIndex, RGCol(eRGCol_Action)) = "Short Exit" Then .RowHidden(lIndex) = False
            End Select
        Next lIndex
        
        ' Figure out if there are visible rows or not
        For lIndex = .FixedRows To .Rows - 1
            If Not .RowHidden(lIndex) Then
                bVisibleRows = True
                Exit For
            End If
        Next lIndex
        
        RefreshRulesGrid
        SetBackColors vsRules
        
        .Redraw = lRedraw
    End With
    
    
    'Preview the first row in the grid...
    If bVisibleRows Then
        Preview 1
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.Filter", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshRulesGrid
'' Description: Refresh the rules grid
'' Inputs:      Whether to force the first row to be selected
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshRulesGrid(Optional ByVal bForceFirstRow As Boolean = False)
On Error GoTo ErrSection:
    
    Dim lRuleID As Long                 ' ID of the currently selected row
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the redraw
    
    With vsRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        FixSequence
        
        If .Rows > .FixedRows Then
            If .RowSel < .FixedRows Or .RowSel >= .Rows Then
                .Row = .FixedRows
                .RowSel = .FixedRows
            End If
            lRuleID = CLng(.TextMatrix(.RowSel, RGCol(eRGCol_RuleID)))

            .AutoSize RGCol(eRGCol_RuleName)
            .Col = RGCol(eRGCol_RuleName)
            .Sort = flexSortGenericAscending
            SetBackColors vsRules
            
            If bForceFirstRow Then
                .Row = .FixedRows
                .RowSel = .FixedRows
            Else
                For lIndex = .FixedRows To .Rows - 1
                    If CLng(.TextMatrix(lIndex, RGCol(eRGCol_RuleID))) = lRuleID Then
                        .Row = lIndex
                        .RowSel = lIndex
                        Exit For
                    End If
                Next lIndex
            End If
            
            Preview .RowSel
        Else
            .AutoSize RGCol(eRGCol_RuleName)
            cmdLinkToEntry.Enabled = False
            cmdEditRule.Enabled = False
            cmdRemoveRule.Enabled = False
            cmdTestEntry.Enabled = False
            cmdPyramidInfo.Enabled = False
        End If
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RefreshRulesGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Preview
'' Description: Update the preview box with the current rule
'' Inputs:      Row selected in grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Preview(ByVal lRow As Long)
On Error GoTo ErrSection:
    
    Dim lRuleID As Long                 ' Rule ID of the currently selected rule
    Dim Rule As New cRule               ' Temporary variable for getting RTF
    
    ' Make sure authorized to view rule (TLB 7/23/2014: can preview if password has already been given)
    If g.Security Is Nothing Then Set g.Security = New cSecurity
    If g.Security.CanPreview(vsRules.Cell(flexcpValue, lRow, RGCol(eRGCol_SecurityLevel))) Or m.bShowPassword = False Then
        txtPreview.TextRTF = Rule.GetRTF(vsRules.TextMatrix(lRow, RGCol(eRGCol_Preview)))
    Else
        txtPreview.SelColor = vbBlack
        txtPreview.Text = "Not authorized to view"
    End If
    
    ' If Rule is Reserved for Stops then disable Edit button
    lRuleID = CLng(vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleID)))
    If vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleType)) = kSN_RESERVED_STOP_RULE Then
        Enable cmdEditRule, False
    Else
        Enable cmdEditRule, True
    End If
    
    ' If Rule is an exit, enable "Linked Entries" button
    If vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleUse)) = 1 Then
        cmdLinkToEntry.Visible = True
        cmdTestEntry.Visible = False
    Else
        cmdLinkToEntry.Visible = False
        cmdTestEntry.Visible = True
    End If
    
    If vsRules.Rows > vsRules.FixedRows Then
        If tbToolbar.ToolBars("General").Visible Then
            'do not want these two buttons to be enabled in wizard mode
            cmdLinkToEntry.Enabled = True
            cmdTestEntry.Enabled = True
        End If
        cmdEditRule.Enabled = True
        cmdRemoveRule.Enabled = True
        Enable cmdPyramidInfo, (chkPyramid.Value = vbChecked)
    End If
    
ErrExit:
    Set Rule = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.Preview", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitRulesGrid
'' Description: Set up the rules grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitRulesGrid()
On Error GoTo ErrSection:

    Dim lWidth As Long                  ' Saved column width
    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With vsRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Clear
        
        If HasPlatinum(False) Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionListBox
        .AllowSelection = True      'Allow multiple selection
        .ExtendLastCol = True
        .GridLines = flexGridFlat
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .Ellipsis = flexEllipsisEnd
        .ScrollTrack = True
        .ScrollTips = False
        .Cols = RGCol(eRGCol_NumCols) + 1
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1
        '.FormatString = "Use|Alt|Signal Name|Action|Long|LExit|Short|SExit"
        .TextMatrix(0, RGCol(eRGCol_Selected)) = "Use"
        .TextMatrix(0, RGCol(eRGCol_Alt)) = "Alt"
        .TextMatrix(0, RGCol(eRGCol_RuleName)) = "Signal Name"
        .TextMatrix(0, RGCol(eRGCol_Action)) = "Action"
        .TextMatrix(0, RGCol(eRGCol_Linked)) = "Linked"
        .TextMatrix(0, RGCol(eRGCol_Sequence)) = "Priority"
        .TextMatrix(0, RGCol(eRGCol_PyramidInfo)) = "Pyramid Info"
        '.TextMatrix(0, RGCol(eRGCol_Shared)) = "Shared"
        
        'Set heading properties
        .ColDataType(RGCol(eRGCol_Selected)) = flexDTBoolean
        .ColDataType(RGCol(eRGCol_Alt)) = flexDTBoolean
        .ColDataType(RGCol(eRGCol_Late)) = flexDTBoolean
        .ColDataType(RGCol(eRGCol_Reverify)) = flexDTBoolean
        .ColDataType(RGCol(eRGCol_ExitOnEntryBar)) = flexDTBoolean
        .ColDataType(RGCol(eRGCol_ExitBasedOnTrade)) = flexDTBoolean
        .ColDataType(RGCol(eRGCol_AsPercent)) = flexDTBoolean
'        .ColDataType(RGCol(eRGCol_Linked)) = flexDTBoolean
        
        .ColAlignment(RGCol(eRGCol_Selected)) = flexAlignCenterTop
        .ColAlignment(RGCol(eRGCol_Alt)) = flexAlignCenterTop
        .ColAlignment(RGCol(eRGCol_Linked)) = flexAlignCenterTop
        
        .ColHidden(RGCol(eRGCol_RuleID)) = True
        .ColHidden(RGCol(eRGCol_BuySell)) = True
        .ColHidden(RGCol(eRGCol_RuleType)) = True
        .ColHidden(RGCol(eRGCol_Late)) = True
        .ColHidden(RGCol(eRGCol_Preview)) = True
        .ColHidden(RGCol(eRGCol_Sort)) = True
        .ColHidden(RGCol(eRGCol_SecurityLevel)) = True
        .ColHidden(RGCol(eRGCol_Password)) = True
        .ColHidden(RGCol(eRGCol_LinkedRules)) = True
        .ColHidden(RGCol(eRGCol_LastMod)) = True
        .ColHidden(RGCol(eRGCol_LastModKnown)) = True
        .ColHidden(RGCol(eRGCol_RuleUse)) = True
        .ColHidden(RGCol(eRGCol_Reverify)) = True
        .ColHidden(RGCol(eRGCol_ExitOnEntryBar)) = True
        .ColHidden(RGCol(eRGCol_ExitBasedOnTrade)) = True
        .ColHidden(RGCol(eRGCol_NumContracts)) = True
        .ColHidden(RGCol(eRGCol_AsPercent)) = True
        .ColHidden(RGCol(eRGCol_SystemNumber)) = True
        .ColHidden(RGCol(eRGCol_CondCoded)) = True
        .ColHidden(RGCol(eRGCol_PriceCoded)) = True
        .ColHidden(RGCol(eRGCol_LimitCoded)) = True
        .ColHidden(RGCol(eRGCol_OrderPlacement)) = True
        '.ColHidden(RGCol(eRGCol_Shared)) = True
        
        .ColWidthMax = 3500
        .TextMatrix(0, RGCol(eRGCol_Action)) = "Short Entry"
        .AutoSize 0, .Cols - 1, False, 75
        .TextMatrix(0, RGCol(eRGCol_Action)) = "Action"
        
        lWidth = .ColWidth(RGCol(eRGCol_Selected))
        .ColWidth(RGCol(eRGCol_Selected)) = lWidth
        .ColWidth(RGCol(eRGCol_Alt)) = lWidth
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.InitRulesGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrids
'' Description: Load the Rule, Input, and Market grids from the Rules collection
''              of the System object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrids()
On Error GoTo ErrSection:

    Dim lRule As Long                   ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim Rule As cRule                   ' Temporary Rule object
    Dim lInput As Long                  ' Index into a for loop
        
    vsRules.Redraw = flexRDNone
    vsInputs.Redraw = flexRDNone
    vsMarkets.Redraw = flexRDNone
    
    For lRule = 1 To m.System.Rules.Count
        Set Rule = m.System.Rules.Item(lRule)
gdStartProfile 210
        AddRuleToGrid Rule
gdStopProfile 210

gdStartProfile 211
        For lInput = 1 To Rule.Inputs.Count
            Rule.Inputs.Item(lInput).RuleName = Rule.Name
            If Rule.Inputs.Item(lInput).ParmTypeID = 5 Then
                AddMarketToGrid Rule.Inputs.Item(lInput)
            Else
                AddInputToGrid Rule.Inputs.Item(lInput)
            End If
        Next lInput
gdStopProfile 211

        DisplayPyramidInfo vsRules.Rows - 1
gdStartProfile 213
        HideInputs Rule.RuleID, (Rule.Selected = False) And (Rule.Alternate = False)
gdStopProfile 213
    Next lRule
    
gdStartProfile 214
    FixInputDisplay
gdStopProfile 214
    
gdStartProfile 215
    RefreshRulesGrid True
gdStopProfile 215
gdStartProfile 216
    RefreshInputsGrid
gdStopProfile 216
    RefreshMarketsGrid
        
    vsRules.Redraw = flexRDBuffered
    vsInputs.Redraw = flexRDBuffered
    vsMarkets.Redraw = flexRDBuffered

ErrExit:
    Set Rule = Nothing
    Exit Sub
    
ErrSection:
    Set Rule = Nothing
    RaiseError "frmSystemManager.LoadGrids", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditPyramidInfo
'' Description: Allow the user to edit the pyramiding information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditPyramidInfo()
On Error GoTo ErrSection:

    Dim bEnter As Boolean
    Dim bPercent As Boolean
    Dim bPosition As Boolean
    Dim lNumContracts As Long
    Dim lRow As Long
    Dim lMarket1Row As Long
    Dim strUnits As String
    Dim lNumUnits As Long

    If Not ShowPassword Then Exit Sub

    lMarket1Row = Market1Row
    If lMarket1Row > -1 Then
        Select Case vsMarkets.TextMatrix(lMarket1Row, MGCol(eMGCol_SecType))
            Case "S", "I"
                strUnits = "Share(s)"
                lNumUnits = CLng(ValOfText(txtNumShares.Text))
            
            Case Else
                strUnits = "Contract(s)"
                lNumUnits = 1&
        End Select
    Else
        strUnits = "Contract(s)"
        lNumUnits = 1&
    End If
    
    With vsRules
        lRow = .Row
        bEnter = (.TextMatrix(lRow, RGCol(eRGCol_RuleUse)) = "0")
        bPercent = CheckedCell(vsRules, lRow, RGCol(eRGCol_AsPercent))
        bPosition = Not CheckedCell(vsRules, lRow, RGCol(eRGCol_ExitBasedOnTrade))
        lNumContracts = CLng(.TextMatrix(lRow, RGCol(eRGCol_NumContracts)))
        
        If frmPyramidOptions.ShowMe(bEnter, bPosition, bPercent, lNumContracts, .TextMatrix(lRow, RGCol(eRGCol_RuleName)), lNumUnits, strUnits) = True Then
            .TextMatrix(lRow, RGCol(eRGCol_NumContracts)) = Trim(CStr(lNumContracts))
            
            If bEnter Then
                CheckedCell(vsRules, lRow, RGCol(eRGCol_AsPercent)) = False
                CheckedCell(vsRules, lRow, RGCol(eRGCol_ExitBasedOnTrade)) = False
            Else
                If bPosition = True Then
                    CheckedCell(vsRules, lRow, RGCol(eRGCol_ExitBasedOnTrade)) = False
                    CheckedCell(vsRules, lRow, RGCol(eRGCol_AsPercent)) = bPercent
                Else
                    CheckedCell(vsRules, lRow, RGCol(eRGCol_ExitBasedOnTrade)) = True
                End If
            End If
            
            DisplayPyramidInfo lRow
        End If
    End With
    
    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.EditPyramidInfo", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditRule
'' Description: Allow the user to edit a rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditRule(Optional ByVal lRuleID As Long = 0&)
On Error GoTo ErrSection:
    
    Dim lRow As Long
    Dim lLE&, lLX&, lSE&, lSX&
    Dim frm As frmRule
    Dim Rule As cRule
    
    If Not HasPlatinum(True) Then Exit Sub
    'If Not ShowPassword Then Exit Sub
    'If m.bShowPassword Then
    '    If Not g.Security.CanEdit(m.System.SecurityLevel, m.System.Password) Then Exit Sub
    '    m.bShowPassword = False
    'End If
        
    If vsRules.Rows = vsRules.FixedRows Then Exit Sub
    If lRuleID = 0& Then
        lRow = vsRules.RowSel
        lRuleID = CLng(vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleID)))
    Else
        For lRow = vsRules.FixedRows To vsRules.Rows - 1
            If CLng(vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleID))) = lRuleID Then
                vsRules.Row = lRow
                vsRules.RowSel = lRow
                Exit For
            End If
        Next lRow
    End If
    GetRuleNumbers lLE, lLX, lSE, lSX
    Set frm = New frmRule
    Set Rule = m.System.Rules.Item(CStr(lRuleID)).MakeCopy(lRuleID, m.System.SystemNumber)
    Rule.RuleType = Rule.RuleUse '(make sure we use the setting from tblSystemRules)
    frm.ShowFromSysMgr Rule, m.System.SystemNumber, m.System.LibraryID, lLE, lLX, lSE, lSX, Me
    
    ShowLinkInputs

ErrExit:
    Set frm = Nothing
    Exit Sub

ErrSection:
    Set frm = Nothing
    RaiseError "frmSystemManager.EditRule", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRuleToGrid
'' Description: Add the rule to the grid
'' Inputs:      Rule to add to the grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRuleToGrid(Rule As cRule, Optional ByVal lRowToAdd As Long = -1)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lRedraw As Long
    Dim strPyramid As String

    With Rule
        ' Right now, the exit based on trade is disabled, so don't allow
        ' those values to go through.  11/02/2001 DAJ
        If .ExitBasedOnEachTrade = True Then
            .ExitBasedOnEachTrade = False
            .AsPercentOfPosition = False
            .NumberContracts = 1
        End If
        
        lRedraw = vsRules.Redraw
        vsRules.Redraw = flexRDNone
        
        If lRowToAdd = -1 Then
            vsRules.Rows = vsRules.Rows + 1
            lRow = vsRules.Rows - 1
            CheckedCell(vsRules, lRow, RGCol(eRGCol_Selected)) = .Selected
            CheckedCell(vsRules, lRow, RGCol(eRGCol_Alt)) = .Alternate
            vsRules.TextMatrix(lRow, RGCol(eRGCol_LinkedRules)) = NullChk(.LinkedRules)
            'CheckedCell(vsRules, lRow, RGCol(eRGCol_Linked)) = (Len(NullChk(.LinkedRules)) > 0)
            If .LastModKnown = 0 Then
                vsRules.TextMatrix(lRow, RGCol(eRGCol_LastModKnown)) = Str(.LastModified)
            Else
                vsRules.TextMatrix(lRow, RGCol(eRGCol_LastModKnown)) = Str(.LastModKnown)
            End If
            CheckedCell(vsRules, lRow, RGCol(eRGCol_ExitBasedOnTrade)) = .SysExitBasedOnEachTrade
            vsRules.TextMatrix(lRow, RGCol(eRGCol_NumContracts)) = .SysNumContracts
            CheckedCell(vsRules, lRow, RGCol(eRGCol_AsPercent)) = .SysAsPercentOfPosition
        Else
            lRow = lRowToAdd
            vsRules.TextMatrix(lRow, RGCol(eRGCol_LastModKnown)) = Str(.LastModified)
        End If
        vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleName)) = .Name
        vsRules.TextMatrix(lRow, RGCol(eRGCol_Action)) = Action(.BuySell, .RuleUse)
        vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleID)) = .RuleID
        vsRules.TextMatrix(lRow, RGCol(eRGCol_BuySell)) = .BuySell
        vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleType)) = .RuleType
        vsRules.TextMatrix(lRow, RGCol(eRGCol_RuleUse)) = .RuleUse
        If .LateAction Or .LateCondition Then
            CheckedCell(vsRules, lRow, RGCol(eRGCol_Late)) = True
        Else
            CheckedCell(vsRules, lRow, RGCol(eRGCol_Late)) = False
        End If
        vsRules.TextMatrix(lRow, RGCol(eRGCol_Sort)) = .Name
        vsRules.TextMatrix(lRow, RGCol(eRGCol_LastMod)) = .LastModified
        vsRules.TextMatrix(lRow, RGCol(eRGCol_Preview)) = .CondFillWords
        vsRules.TextMatrix(lRow, RGCol(eRGCol_SecurityLevel)) = .SecurityLevel
        vsRules.TextMatrix(lRow, RGCol(eRGCol_Password)) = .Password
        CheckedCell(vsRules, lRow, RGCol(eRGCol_Reverify)) = .Reverify
        If .Reverify Then
            vsRules.Cell(flexcpForeColor, lRow, RGCol(eRGCol_RuleName)) = vbRed
        Else
            vsRules.Cell(flexcpForeColor, lRow, RGCol(eRGCol_RuleName)) = vsRules.Cell(flexcpForeColor, lRow, RGCol(eRGCol_Alt))
        End If
        CheckedCell(vsRules, lRow, RGCol(eRGCol_ExitOnEntryBar)) = .ExitOnEntryBar
        
        vsRules.TextMatrix(lRow, RGCol(eRGCol_CondCoded)) = .CondCoded
        vsRules.TextMatrix(lRow, RGCol(eRGCol_PriceCoded)) = .PriceCoded
        vsRules.TextMatrix(lRow, RGCol(eRGCol_LimitCoded)) = .Price2Coded
        vsRules.TextMatrix(lRow, RGCol(eRGCol_OrderPlacement)) = .OrderPlacement
        vsRules.TextMatrix(lRow, RGCol(eRGCol_Sequence)) = .Seq
        
        ' Build the pyramiding information string
        DisplayPyramidInfo lRow
        
        'vsRules.ColHidden(RGCol(eRGCol_Linked)) = Not ShowLinked
        ShowLinkedExits
        
        vsRules.TextMatrix(lRow, RGCol(eRGCol_SystemNumber)) = .SystemNumber
        vsRules.Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.AddRuleToGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveLocalRules
'' Description: Remove all temporary local rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLocalRules()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rs As Recordset
    Dim lSysNbr As Long
    
    With vsRules
        'For X = .FixedRows To .Rows - 1
        '    lSysNbr = CLng(.TextMatrix(X, RGCol(eRGCol_SystemNumber)))
        '
        '    '-2 is a New/Local rule (delete it)
        '    If lSysNbr = -2 Or lSysNbr = -4 Then
        '        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        '            "WHERE [RuleID]=" & .TextMatrix(X, RGCol(eRGCol_RuleID)) & ";", dbOpenDynaset)
        '        If Not rs.EOF Then
        '            rs.Delete
        '        End If
        '    End If
        'Next X
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RemoveLocalRules", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveRule
'' Description: Deltes the system rule if it is new or edited and local
'' Inputs:      Row selected in the grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveRule(ByVal lRow As Long)
On Error GoTo ErrSection:
    
    Dim lRuleID As Long
    Dim lRedraw As Long
    
    With vsRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRuleID = CLng(.TextMatrix(lRow, RGCol(eRGCol_RuleID)))
        
        RemoveInput lRuleID
        RemoveMarket lRuleID
        
        m.System.Rules.Remove CStr(lRuleID)
    
        .RemoveItem lRow
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.RemoveRule", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckReverifyFlags
'' Description: Check each rule to see if it needs reverifying
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckReverifyFlags()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rs As Recordset
    Dim strReverify As String
    Dim strLastMod As String
    Dim strRemoved As String
    Dim strInvalid As String
    Dim strMsg As String
    
'Exit Sub
    
    With vsRules
        For X = .FixedRows To .Rows - 1
            'Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                    "WHERE [RuleID]=" & .TextMatrix(X, RGCol(eRGCol_RuleID)) & ";", dbOpenDynaset)
            'ValidateCheckSums rs, "tblRules"
            'If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
            'If rs.EOF Then
            '    strRemoved = strRemoved & UCase(.TextMatrix(X, RGCol(eRGCol_RuleName))) & vbCrLf
            'Else
            '    If rs!CheckSum = 0.5 Then
            '        strInvalid = strInvalid & UCase(.TextMatrix(X, RGCol(eRGCol_RuleName))) & vbCrLf
            '    End If
                
            '    If rs!Reverify Then
                If CheckedCell(vsRules, X, RGCol(eRGCol_Reverify)) Then
                    strReverify = strReverify & UCase(.TextMatrix(X, RGCol(eRGCol_RuleName))) & vbCrLf
                End If
                
            '    If Str(rs!LastModified) <> .TextMatrix(X, RGCol(eRGCol_LastModKnown)) Then
            '        strLastMod = strLastMod & UCase(.TextMatrix(X, RGCol(eRGCol_RuleName))) & vbCrLf
            '    End If
            'End If
            'rs.Close
        Next X
    End With
    
    'If Len(strInvalid) > 0 Then
    '    strMsg = strMsg & "One or more rules are invalid.  " & _
            "Please remove the following rule(s):" & vbCrLf & strInvalid
    'End If
    
    If Len(strReverify) > 0 Then
        strMsg = strMsg & "One or more rules need Reverifying.  " & _
            "Please Reverify the following rule(s):" & vbCrLf & strReverify
    End If
    
    'If Len(strLastMod) > 0 Then
    '    If Len(strMsg) > 0 Then strMsg = strMsg & vbCrLf
    '    strMsg = strMsg & "One or more rules need Reverifying.  " & _
            "Please Reverify the following rule(s):" & vbCrLf & strLastMod
    'End If
    
    'If Len(strRemoved) > 0 Then
    '    If Len(strMsg) > 0 Then strMsg = strMsg & vbCrLf
    '    strMsg = strMsg & _
            "One or more rules have been removed since the strategy " & _
            "was last saved.  You may want to check any inputs for " & _
            "the following rule(s):" & vbCrLf & strRemoved
    'End If
    
    If Len(strMsg) > 0 Then Err.Raise vbObjectError + 1000, , strMsg
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "frmSystemManager.CheckReverifyFlags", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitInputsGrid
'' Description: Initialize the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitInputsGrid()
On Error GoTo ErrSection:
    
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .AllowBigSelection = False
        .AllowSelection = False
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .TabBehavior = flexTabCells
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTips = False
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = IGCol(eIGCol_NumCols)
        .Rows = 1
        .FixedCols = 0 '2
        .FixedRows = 1
        
        .ColHidden(IGCol(eIGCol_RuleID)) = True
        .ColHidden(IGCol(eIGCol_ParmTypeID)) = True
        .ColHidden(IGCol(eIGCol_ParmID)) = True
        .ColHidden(IGCol(eIGCol_FromVal)) = True
        .ColHidden(IGCol(eIGCol_ToVal)) = True
        .ColHidden(IGCol(eIGCol_ParmDesc)) = True
        .ColHidden(IGCol(eIGCol_OptFromValue)) = True
        .ColHidden(IGCol(eIGCol_OptToValue)) = True
        .ColHidden(IGCol(eIGCol_OptStepValue)) = True
        .ColHidden(IGCol(eIGCol_Sort)) = True
        .ColHidden(IGCol(eIGCol_Req)) = True
        .ColHidden(IGCol(eIGCol_Hide)) = True
        
        .TextMatrix(0, IGCol(eIGCol_RuleName)) = "Rule Name"
        .TextMatrix(0, IGCol(eIGCol_InputValue)) = "Value"
        .TextMatrix(0, IGCol(eIGCol_InputName)) = "Input"
        .TextMatrix(0, IGCol(eIGCol_IfOptimize)) = "Optimize"
        .TextMatrix(0, IGCol(eIGCol_OptFromValue)) = "From"
        .TextMatrix(0, IGCol(eIGCol_OptToValue)) = "To"
        .TextMatrix(0, IGCol(eIGCol_OptStepValue)) = "Step"
        
        .ColAlignment(IGCol(eIGCol_RuleName)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_InputName)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_InputValue)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_IfOptimize)) = flexAlignCenterCenter
        .ColAlignment(IGCol(eIGCol_OptFromValue)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_OptToValue)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_OptStepValue)) = flexAlignLeftCenter
        
        .ColDataType(IGCol(eIGCol_IfOptimize)) = flexDTBoolean
        .ColDataType(IGCol(eIGCol_Req)) = flexDTBoolean
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.InitInputsGrid", eGDRaiseError_Raise
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInputsGrid
'' Description: Load up the inputs grid
'' Inputs:      Whether or not to hide duplicates
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadInputsGrid(ByVal bHideDuplicates As Boolean)
On Error GoTo ErrSection:
        
    Dim X As Long
    Dim lRedraw As Long

    'Leave if no inputs exist in collection
    If m.System.Inputs Is Nothing Then Exit Sub
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For X = 1 To m.System.Inputs.Count
            AddInputToGrid m.System.Inputs.Item(X)
        Next X
        
        HideOptColumns
        If bHideDuplicates Then HideDuplicateInputs
        RefreshInputsGrid
        
        SetBackColors vsInputs
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.LoadInputsGrid", eGDRaiseError_Raise

End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddInputToGrid
'' Description: Add an input to the grid
'' Inputs:      Row to fill, Input to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddInputToGrid(Parm As cInput, Optional ByVal lRowToAdd As Long = -1&)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim lRedraw As Long
    Dim lRow As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRowToAdd = -1 Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        Else
            lRow = lRowToAdd
        End If
        
        .TextMatrix(lRow, IGCol(eIGCol_RuleID)) = Parm.RuleID
        .TextMatrix(lRow, IGCol(eIGCol_RuleName)) = Parm.RuleName
        .TextMatrix(lRow, IGCol(eIGCol_InputName)) = Parm.ParmName
        .TextMatrix(lRow, IGCol(eIGCol_Sort)) = Parm.RuleName & Parm.ParmName
        .Cell(flexcpFontBold, lRow, IGCol(eIGCol_InputName)) = True
        .Cell(flexcpForeColor, lRow, IGCol(eIGCol_InputName)) = vbBlack
        .TextMatrix(lRow, IGCol(eIGCol_ParmTypeID)) = Parm.ParmTypeID
        .TextMatrix(lRow, IGCol(eIGCol_ParmID)) = Parm.ParmID
        .TextMatrix(lRow, IGCol(eIGCol_ParmDesc)) = Parm.ParmDesc
        .TextMatrix(lRow, IGCol(eIGCol_FromVal)) = ""
        .TextMatrix(lRow, IGCol(eIGCol_ToVal)) = ""
        .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = ""
        .TextMatrix(lRow, IGCol(eIGCol_Req)) = Parm.Required
        
        'Optimization fields
        .TextMatrix(lRow, IGCol(eIGCol_IfOptimize)) = Parm.IfOptimize
        If Parm.ParmTypeID = kSN_RetTrueFalseConstant Then
            .TextMatrix(lRow, IGCol(eIGCol_OptFromValue)) = "True"
            .TextMatrix(lRow, IGCol(eIGCol_OptToValue)) = "False"
            .TextMatrix(lRow, IGCol(eIGCol_OptStepValue)) = ""
        Else
            .TextMatrix(lRow, IGCol(eIGCol_OptFromValue)) = FormatNum(Parm.OptFromValue)
            .TextMatrix(lRow, IGCol(eIGCol_OptToValue)) = FormatNum(Parm.OptToValue)
            .TextMatrix(lRow, IGCol(eIGCol_OptStepValue)) = FormatNum(Parm.OptStepValue)
        End If
            
        'Set the value (or default if one doesn't exist).  The bars and
        'trades type structure is always "Market1" and "Trades"
        Select Case Parm.ParmTypeID
        
            Case kSN_RetBars
                .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = Parm.ParmName
        
            Case kSN_RetTrades
                .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = Parm.ParmName
                
            Case kSN_RetTrueFalseConstant
                If Parm.Value = "" Or ValOfText(Parm.Value) <> 0 Then
                    .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = "True"
                Else
                    .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = "False"
                End If
                
            Case kSN_RetNumericConstant, kSN_RetTrueFalse, kSN_RetNumeric ', kSN_RetTrueFalseConstant
                If Parm.Value = "" Then
                    .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = FormatNum(Val(Parm.DefaultValue))
                Else
                    .TextMatrix(lRow, IGCol(eIGCol_InputValue)) = FormatNum(Val(Parm.Value))
                End If
                ColorInputCell lRow, IGCol(eIGCol_InputValue)
                
                .TextMatrix(lRow, IGCol(eIGCol_FromVal)) = FormatNum(Parm.FromValue)
                .TextMatrix(lRow, IGCol(eIGCol_ToVal)) = FormatNum(Parm.ToValue)
                ColorInputCell lRow, IGCol(eIGCol_FromVal)
                ColorInputCell lRow, IGCol(eIGCol_ToVal)
                
        End Select
        
        If Parm.ParmTypeID = kSN_RetBars Then .RowHidden(lRow) = True
        
        If m.bHideDuplicates Then
            For lIndex = .FixedRows To .Rows - 1
                If lRow <> lIndex And SameInput(lRow, lIndex) Then
                    If Not HideDuplicateBit(lIndex) Then
                        HideDuplicateBit(lRow) = True
                        Exit For
                    End If
                End If
            Next lIndex
        End If
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.AddInputToGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideOptColumns
'' Description: Show/Hide the optimization columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideOptColumns()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim bHideOpt As Boolean
    Dim lRedraw As Long
    Dim bPlat As Boolean
    
    bPlat = HasPlatinum(False)
    
    ' If at least one optimization field is checked then show optimization fields
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        bHideOpt = True
        For X = .FixedRows To .Rows - 1
            If .RowHidden(X) = False Then
                If CheckedCell(vsInputs, X, IGCol(eIGCol_IfOptimize)) = True Then
                    If Not bPlat Then
                        ' optimizing only allowed with Platinum
                        CheckedCell(vsInputs, X, IGCol(eIGCol_IfOptimize)) = False
                    Else
                        bHideOpt = False
                        Exit For
                    End If
                End If
            End If
        Next X
        .ColHidden(IGCol(eIGCol_OptFromValue)) = bHideOpt
        .ColHidden(IGCol(eIGCol_OptToValue)) = bHideOpt
        .ColHidden(IGCol(eIGCol_OptStepValue)) = bHideOpt
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.HideOptColumns", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideDuplicateInputs
'' Description: Hide all but one occurrance inputs with the same name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideDuplicateInputs()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' State of the grid's redraw property
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop

    m.bHideDuplicates = True

    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Turn all of the 'Hide Duplicate' bits off
        For lIndex = .FixedRows To .Rows - 1
            HideDuplicateBit(lIndex) = False
        Next lIndex
        
        ' Walk through and turn appropriate 'Hide Duplicate' bits on
        For lIndex = .FixedRows To .Rows - 1
            If Not (HideRuleBit(lIndex) Or HideDuplicateBit(lIndex)) Then
                For lIndex2 = .FixedRows To .Rows - 1
                    If lIndex2 <> lIndex Then
                        If SameInput(lIndex, lIndex2) Then
                            HideDuplicateBit(lIndex2) = True
                        End If
                    End If
                Next lIndex2
            End If
        Next lIndex
        
        FixInputDisplay
        SetBackColors vsInputs
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.HideDuplicateInputs", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideInputs
'' Description: Hide or Show the inputs of a specific rule
'' Inputs:      ID of the Rule, Whether or not to Hide the inputs
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideInputs(ByVal lRuleID As Long, ByVal bHide As Boolean)
On Error GoTo ErrSection:
    
    Dim lRedraw As Long                 ' State of the grid's redraw property
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
gdStartProfile 220
        ' Turn the appropriate 'Hide Inputs' bits on or off
        For lIndex = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lIndex, IGCol(eIGCol_RuleID))) = lRuleID Then
                HideRuleBit(lIndex) = bHide
                
                ' If duplicates are hidden and we are now hiding an input that
                ' was shown before, hide it and show the next duplicate instead
                If m.bHideDuplicates Then
                    If bHide And Not HideDuplicateBit(lIndex) Then
                        For lIndex2 = .FixedRows To .Rows - 1
                            If lIndex <> lIndex2 And SameInput(lIndex, lIndex2) Then
                                If Not HideRuleBit(lIndex2) Then
                                    HideDuplicateBit(lIndex) = True
                                    HideDuplicateBit(lIndex2) = False
                                    Exit For
                                End If
                            End If
                        Next lIndex2
                    ElseIf Not bHide Then
                        For lIndex2 = .FixedRows To .Rows - 1
                            If SameInput(lIndex, lIndex2) Then
                                HideDuplicateBit(lIndex2) = (lIndex <> lIndex2)
                            End If
                        Next lIndex2
                    End If
                End If
                
                'Exit For
            End If
        Next lIndex
gdStopProfile 220
       
        ' TLB: Don't do this here anymore -- it's too inefficient to do with every
        ' call to HideInputs. Instead, call this after all the calls to HideInputs.
        'FixInputDisplay

gdStartProfile 222
        SetBackColors vsInputs
gdStopProfile 222
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.HideInputs", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowDuplicateInputs
'' Description: Show the duplicate inputs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowDuplicateInputs()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' State of the grid's redraw property
    Dim lIndex As Long                  ' Index into a for loop
    
    m.bHideDuplicates = False
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            HideDuplicateBit(lIndex) = False
        Next lIndex
        
        FixInputDisplay
        SetBackColors vsInputs
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.ShowDuplicateInputs", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDuplicates
'' Description: Set the duplicate bits accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetDuplicates()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' State of the grid's redraw property
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop

    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If Not HideDuplicateBit(lIndex) Then
                For lIndex2 = .FixedRows To .Rows - 1
                    If lIndex2 <> lIndex Then
                        If SameInput(lIndex, lIndex2) Then
                            .TextMatrix(lIndex2, IGCol(eIGCol_InputValue)) = .TextMatrix(lIndex, IGCol(eIGCol_InputValue))
                            .TextMatrix(lIndex2, IGCol(eIGCol_IfOptimize)) = .TextMatrix(lIndex, IGCol(eIGCol_IfOptimize))
                            .TextMatrix(lIndex2, IGCol(eIGCol_OptFromValue)) = .TextMatrix(lIndex, IGCol(eIGCol_OptFromValue))
                            .TextMatrix(lIndex2, IGCol(eIGCol_OptToValue)) = .TextMatrix(lIndex, IGCol(eIGCol_OptToValue))
                            .TextMatrix(lIndex2, IGCol(eIGCol_OptStepValue)) = .TextMatrix(lIndex, IGCol(eIGCol_OptStepValue))
                        End If
                    End If
                Next lIndex2
            End If
        Next lIndex
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.SetDuplicates", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixInputDisplay
'' Description: Hide an Input accordingly
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixInputDisplay()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Value of the grid's redraw property
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        m.bHasDuplicates = False
        For lIndex = .FixedRows To .Rows - 1
gdStartProfile 230
            ' Hide the row if necessary
            .RowHidden(lIndex) = HideRuleBit(lIndex) Or HideDuplicateBit(lIndex)
            
            ' Set the Rule Name column back to the original value and height
            .TextMatrix(lIndex, IGCol(eIGCol_RuleName)) = Parse(.TextMatrix(lIndex, IGCol(eIGCol_RuleName)), vbCrLf, 1)
            .RowHeight(lIndex) = .RowHeight(0)
gdStopProfile 230
            
            ' If the duplicates are hidden, we need to get all of the
            ' applicable rule names for this input and display them all
            If Not HideRuleBit(lIndex) And Not HideDuplicateBit(lIndex) Then
                For lIndex2 = .FixedRows To .Rows - 1
                    If lIndex <> lIndex2 And SameInput(lIndex, lIndex2) Then
gdStartProfile 235
                        .TextMatrix(lIndex2, IGCol(eIGCol_RuleName)) = Parse(.TextMatrix(lIndex2, IGCol(eIGCol_RuleName)), vbCrLf, 1)
gdStopProfile 235
gdStartProfile 236
                        .RowHeight(lIndex2) = .RowHeight(0)
gdStopProfile 236
gdStartProfile 237
                        If Not HideRuleBit(lIndex2) Then
                            m.bHasDuplicates = True
                            If m.bHideDuplicates Then
                                .TextMatrix(lIndex, IGCol(eIGCol_RuleName)) = .TextMatrix(lIndex, IGCol(eIGCol_RuleName)) & vbCrLf & .TextMatrix(lIndex2, IGCol(eIGCol_RuleName))
                                .RowHeight(lIndex) = .RowHeight(lIndex) + .RowHeight(0)
                            End If
                        End If
gdStopProfile 237
                    End If
                Next lIndex2
            End If
        Next lIndex
        
gdStartProfile 240
        .AutoSize IGCol(eIGCol_RuleName)
gdStopProfile 240
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.FixInputDisplay", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshInputsGrid
'' Description: Refresh the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshInputsGrid()
On Error GoTo ErrSection:
    
    Dim lRedraw As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            .Select 1, IGCol(eIGCol_Sort)
            HideOptColumns
            If m.bHideDuplicates Then HideDuplicateInputs
            .AutoSize 0, .Cols - 1, False, 75
        End If
        
        If .ColWidth(IGCol(eIGCol_RuleName)) > 2000 Then .ColWidth(IGCol(eIGCol_RuleName)) = 2000
        If .ColWidth(IGCol(eIGCol_InputName)) > 2500 Then .ColWidth(IGCol(eIGCol_InputName)) = 2500
        If .ColWidth(IGCol(eIGCol_InputValue)) < 1200 Then .ColWidth(IGCol(eIGCol_InputValue)) = 1200
        If .ColWidth(IGCol(eIGCol_OptFromValue)) < 1000 Then .ColWidth(IGCol(eIGCol_OptFromValue)) = 1000
        If .ColWidth(IGCol(eIGCol_OptToValue)) < 1000 Then .ColWidth(IGCol(eIGCol_OptToValue)) = 1000
        If .ColWidth(IGCol(eIGCol_OptStepValue)) < 1000 Then .ColWidth(IGCol(eIGCol_OptStepValue)) = 1000
        
        SetBackColors vsInputs
        
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RefreshInputsGrid", eGDRaiseError_Raise

End Sub

Private Sub RefreshMarketsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    Dim bDataOtherThanCN As Boolean

    With vsMarkets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            HideDuplicateMarkets
            bDataOtherThanCN = DataOtherThanCN
            vsMarkets.ColHidden(MGCol(eMGCol_Format)) = Not bDataOtherThanCN
            vsMarkets.ColHidden(MGCol(eMGCol_SymbolPath)) = Not bDataOtherThanCN
            SetBackColors vsMarkets
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RefreshMarketsGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorInputCell
'' Description: Color the input cell red for negative, black for positive
'' Inputs:      Row and Column of cell being edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorInputCell(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    If ValOfText(vsInputs.TextMatrix(Row, Col)) < 0 Then
        vsInputs.Cell(flexcpForeColor, Row, Col) = vbRed
    Else
        vsInputs.Cell(flexcpForeColor, Row, Col) = vbBlack
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.ColorInputCell", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SameInput
'' Description: Determine if this is a matched input
'' Inputs:      Two rows to compare
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SameInput(ByVal lRow1 As Long, ByVal lRow2 As Long) As Boolean
On Error GoTo ErrSection:
    
'gdStartProfile 241
    With vsInputs
        SameInput = (UCase(.TextMatrix(lRow1, IGCol(eIGCol_InputName))) = UCase(.TextMatrix(lRow2, IGCol(eIGCol_InputName))))
    End With
'gdStopProfile 241
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.SameInput", eGDRaiseError_Raise
    
End Function

Private Property Get HideRuleBit(ByVal lIndex As Long) As Boolean
    HideRuleBit = (Val(vsInputs.TextMatrix(lIndex, IGCol(eIGCol_Hide))) And 1)
End Property
Private Property Get HideDuplicateBit(ByVal lIndex As Long) As Boolean
    HideDuplicateBit = (Val(vsInputs.TextMatrix(lIndex, IGCol(eIGCol_Hide))) And 2)
End Property
Private Property Let HideRuleBit(ByVal lIndex As Long, ByVal bValue As Boolean)
    With vsInputs
        If bValue Then
            .TextMatrix(lIndex, IGCol(eIGCol_Hide)) = CStr(Val(.TextMatrix(lIndex, IGCol(eIGCol_Hide))) Or 1)
        Else
            .TextMatrix(lIndex, IGCol(eIGCol_Hide)) = CStr(Val(.TextMatrix(lIndex, IGCol(eIGCol_Hide))) And Not 1)
        End If
    End With
End Property
Private Property Let HideDuplicateBit(ByVal lIndex As Long, ByVal bValue As Boolean)
    With vsInputs
        If bValue Then
            .TextMatrix(lIndex, IGCol(eIGCol_Hide)) = CStr(Val(.TextMatrix(lIndex, IGCol(eIGCol_Hide))) Or 2)
        Else
            .TextMatrix(lIndex, IGCol(eIGCol_Hide)) = CStr(Val(.TextMatrix(lIndex, IGCol(eIGCol_Hide))) And Not 2)
        End If
    End With
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyStepValues
'' Description: Verify the step values for optimization
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyStepValues()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim dFromValue As Double
    Dim dToValue As Double
    Dim dStepValue As Double
    
    With vsInputs
        For X = .FixedRows To .Rows - 1
            If .TextMatrix(X, IGCol(eIGCol_IfOptimize)) = True Then
                dFromValue = ValOfText(.TextMatrix(X, IGCol(eIGCol_OptFromValue)))
                dToValue = ValOfText(.TextMatrix(X, IGCol(eIGCol_OptToValue)))
                dStepValue = ValOfText(.TextMatrix(X, IGCol(eIGCol_OptStepValue)))
                If dStepValue > (dToValue - dFromValue) Or dStepValue <= 0 Then
                    dStepValue = dToValue - dStepValue
                End If
            End If
        Next X
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.VerifyStepValues", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddInput
'' Description: Add an input to the system
'' Inputs:      Rule ID of the input, Whether to refresh the grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddInput(ByVal Rule As cRule, Optional ByVal bRefreshGrid As Boolean = True)
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim bFound As Boolean
    Dim Y As Long
    
    If Rule.Inputs Is Nothing Then Exit Sub
    
    ' Remove any parms that no longer exist in the collection...
    With vsInputs
        For X = .Rows - 1 To .FixedRows Step -1
            If CLng(.TextMatrix(X, IGCol(eIGCol_RuleID))) = Rule.RuleID Then
                If Not Rule.Inputs.Found(.TextMatrix(X, IGCol(eIGCol_ParmID))) Then
                    .RemoveItem X
                End If
            End If
        Next X
    End With

    ' Load rule inputs to inputs grid...
    For X = 1 To Rule.Inputs.Count
        With Rule.Inputs.Item(X)
            ' Input cannot be a bar structure type
            If .ParmTypeID <> kSN_RetBars Then
                .RuleName = Rule.Name
            
                ' Make sure input doesn't already exist
                bFound = False
                For Y = 1 To vsInputs.Rows - 1
                    If ValOfText(vsInputs.TextMatrix(Y, IGCol(eIGCol_ParmID))) = .ParmID Then
                        ' TLB & DAJ: need to "re-add" this input so display will be correct for it
                        AddInputToGrid Rule.Inputs.Item(X), Y
                        bFound = True
                        Exit For
                    End If
                Next Y
                
                If Not bFound Then
                    AddInputToGrid Rule.Inputs.Item(X)
                End If
            End If
        End With
    Next X
    
    If bRefreshGrid Then
        ' Hide the inputs if the rule is neither selected nor alternating
        HideInputs Rule.RuleID, (Rule.Selected = False) And (Rule.Alternate = False)
        FixInputDisplay
        RefreshInputsGrid
        SetBackColors vsInputs
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.AddInput", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveInput
'' Description: Remove an input from a system
'' Inputs:      Rule ID of the input
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveInput(ByVal lRuleID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lRedraw As Long
    
    'Remove all inputs for selected rule
    HideInputs lRuleID, True
    FixInputDisplay
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If ValOfText(.TextMatrix(lIndex, IGCol(eIGCol_RuleID))) = lRuleID Then
                .RemoveItem lIndex
            End If
        Next lIndex
        .Redraw = lRedraw
    End With
    
    HideOptColumns
    RefreshInputsGrid
    SetBackColors vsInputs
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RemoveInput", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddMarket
'' Description: Add a security to the system
'' Inputs:      Rule ID of the security, Whether to hide duplicates
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddMarket(ByVal Rule As cRule, Optional ByVal bHideDuplicates As Boolean = True)
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim strSymbol As String
    Dim strSecurity As String
    Dim strSecType As String
    Dim strFormat As String
    Dim strPeriod As String
    Dim strPath As String
    Dim strMarketSymbol As String
    Dim strGroupID As String
    Dim iIndex As Long
    Dim bFound As Boolean
    Dim lRedraw As Long
    Dim lRuleID As Long
    Dim lParmID As Long
    Dim lSymbolID As Long
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    'Quit if no inputs found...
    If Rule.Inputs Is Nothing Then Exit Sub
    
    With vsMarkets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        For X = 1 To Rule.Inputs.Count
            'Only add bar structure inputs...
            If Rule.Inputs.Item(X).ParmTypeID = kSN_RetBars Then
            
                'Make sure it doesn't already exist in grid...
                If Not MarketFound(CStr(Rule.Inputs.Item(X).ParmID)) Then
                    
                    ' If the parm name already exists, take all the data from it
                    bFound = False
                    For iIndex = .FixedRows To .Rows - 1
                        If .TextMatrix(iIndex, MGCol(eMGCol_ParmName)) = Rule.Inputs.Item(X).ParmName Then
                            strSymbol = .TextMatrix(iIndex, MGCol(eMGCol_Symbol))
                            strSecurity = .TextMatrix(iIndex, MGCol(eMGCol_Security))
                            strSecType = .TextMatrix(iIndex, MGCol(eMGCol_SecType))
                            strFormat = .TextMatrix(iIndex, MGCol(eMGCol_Format))
                            strPeriod = .TextMatrix(iIndex, MGCol(eMGCol_Period))
                            strPath = .TextMatrix(iIndex, MGCol(eMGCol_SymbolPath))
                            strMarketSymbol = .TextMatrix(iIndex, MGCol(eMGCol_MarketSymbol))
                            strGroupID = .TextMatrix(iIndex, MGCol(eMGCol_GroupID))
                            lSymbolID = CLng(Val(.TextMatrix(iIndex, MGCol(eMGCol_SymbolID))))
                            bFound = True
                            Exit For
                        End If
                    Next iIndex
                    
                    ' If the parm name does not exist yet, try to get data from the ini file
                    If bFound = False Then
                        'Get default strSecurity info from INI.  It will exist here if the
                        'strSecurity was added to a previous system.
                        Select Case UCase(Rule.Inputs.Item(X).ParmName)
                            Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY"
                                For iIndex = .FixedRows To .Rows - 1
                                    If UCase(.TextMatrix(iIndex, MGCol(eMGCol_ParmName))) = "MARKET1" Then
                                        strSymbol = .TextMatrix(iIndex, MGCol(eMGCol_Symbol))
                                        strSecurity = .TextMatrix(iIndex, MGCol(eMGCol_Security))
                                        strSecType = .TextMatrix(iIndex, MGCol(eMGCol_SecType))
                                        strFormat = .TextMatrix(iIndex, MGCol(eMGCol_Format))
                                        strPeriod = Left(Rule.Inputs.Item(X).ParmName, 1)
                                        strPath = .TextMatrix(iIndex, MGCol(eMGCol_SymbolPath))
                                        strMarketSymbol = .TextMatrix(iIndex, MGCol(eMGCol_MarketSymbol))
                                        strGroupID = .TextMatrix(iIndex, MGCol(eMGCol_GroupID))
                                        lSymbolID = CLng(Val(.TextMatrix(iIndex, MGCol(eMGCol_SymbolID))))
                                        bFound = True
                                        Exit For
                                    End If
                                Next iIndex
                            
                            Case "UNSPLIT"
                                For iIndex = .FixedRows To .Rows - 1
                                    If UCase(.TextMatrix(iIndex, MGCol(eMGCol_ParmName))) = "MARKET1" Then
                                        strSymbol = .TextMatrix(iIndex, MGCol(eMGCol_Symbol))
                                        strSecurity = .TextMatrix(iIndex, MGCol(eMGCol_Security))
                                        strSecType = .TextMatrix(iIndex, MGCol(eMGCol_SecType))
                                        strFormat = .TextMatrix(iIndex, MGCol(eMGCol_Format))
                                        strPath = .TextMatrix(iIndex, MGCol(eMGCol_SymbolPath))
                                        strMarketSymbol = .TextMatrix(iIndex, MGCol(eMGCol_MarketSymbol))
                                        strGroupID = .TextMatrix(iIndex, MGCol(eMGCol_GroupID))
                                        lSymbolID = CLng(Val(.TextMatrix(iIndex, MGCol(eMGCol_SymbolID))))
                                        bFound = True
                                        Exit For
                                    End If
                                Next iIndex
                            
                            Case Else
                                If Left(Rule.Inputs.Item(X).ParmName, 1) = Chr(34) And Right(Rule.Inputs.Item(X).ParmName, 1) = Chr(34) Then
                                    strSymbol = Parse(Replace(Rule.Inputs.Item(X).ParmName, Chr(34), ""), ",", 1)
                                    strPeriod = Parse(Replace(Rule.Inputs.Item(X).ParmName, Chr(34), ""), ",", 2)
                                    If Len(strPeriod) = 0 Then strPeriod = "(Default)"
                                    If Len(strSymbol) = 0 Then
                                        For iIndex = .FixedRows To .Rows - 1
                                            If UCase(.TextMatrix(iIndex, MGCol(eMGCol_ParmName))) = "MARKET1" Then
                                                strSymbol = .TextMatrix(iIndex, MGCol(eMGCol_Symbol))
                                                strSecurity = .TextMatrix(iIndex, MGCol(eMGCol_Security))
                                                strSecType = .TextMatrix(iIndex, MGCol(eMGCol_SecType))
                                                strFormat = .TextMatrix(iIndex, MGCol(eMGCol_Format))
                                                strPath = .TextMatrix(iIndex, MGCol(eMGCol_SymbolPath))
                                                strMarketSymbol = .TextMatrix(iIndex, MGCol(eMGCol_MarketSymbol))
                                                strGroupID = .TextMatrix(iIndex, MGCol(eMGCol_GroupID))
                                                lSymbolID = CLng(Val(.TextMatrix(iIndex, MGCol(eMGCol_SymbolID))))
                                                bFound = True
                                                Exit For
                                            End If
                                        Next iIndex
                                    End If
                                    
                                    If SetBarProperties(Bars, strSymbol) = True Then
                                        strSecurity = Bars.Prop(eBARS_Desc)
                                        strSecType = Chr(Bars.Prop(eBARS_SecurityType))
                                        strFormat = "CN"
                                        strPath = AddSlash(App.Path) & "Data"
                                        strMarketSymbol = Bars.Prop(eBARS_MarketSymbol)
                                        strGroupID = ""
                                        lSymbolID = Bars.Prop(eBARS_SymbolID)
                                    Else
                                        strSecurity = ""
                                        strSecType = ""
                                        strFormat = ""
                                        strPath = ""
                                        strMarketSymbol = ""
                                        strGroupID = ""
                                        lSymbolID = 0&
                                    End If
                                Else
                                    DefaultSecurityInfo Rule.Inputs.Item(X).ParmName, _
                                        strSymbol, strSecurity, strSecType, strFormat, strPeriod, strPath, _
                                        strMarketSymbol, strGroupID, lSymbolID
                                End If
                        End Select
                    End If
                    
                    With Rule.Inputs.Item(X)
                        .Path = strPath
                        .Symbol = strSymbol
                        .MarketSymbol = strMarketSymbol
                        .Periodicity = strPeriod
                        .Format = strFormat
                        .SecurityType = strSecType
                        .SecurityName = strSecurity
                        .GroupID = strGroupID
                        .SymbolID = lSymbolID
                    End With
                    
                    AddMarketToGrid Rule.Inputs.Item(X)
                End If
            End If
        Next X
        
        ' Delete any dead-wood that may be left...
        For X = .Rows - 1 To .FixedRows Step -1
            lRuleID = CLng(.TextMatrix(X, MGCol(eMGCol_RuleID)))
            lParmID = CLng(.TextMatrix(X, MGCol(eMGCol_ParmID)))
            
            If lRuleID = Rule.RuleID Then
                If Not Rule.Inputs.Found(CStr(lParmID)) Then
                    .RemoveItem X
                End If
            End If
        Next X
        
        .Redraw = lRedraw
    End With
        
    If bHideDuplicates Then RefreshMarketsGrid
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.AddMarket", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketFound
'' Description: Is the market in the grid?
'' Inputs:      Parm ID of the market
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MarketFound(ByVal lParmID As Long) As Boolean
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    MarketFound = False
    With vsMarkets
        For lIndex = .FixedRows To .Rows - 1
            If lParmID = ValOfText(.TextMatrix(lIndex, MGCol(eMGCol_ParmID))) Then
                MarketFound = True
                Exit For
            End If
        Next lIndex
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.MarketFound", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveMarket
'' Description: Remove a market from a system
'' Inputs:      Rule ID of the market
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveMarket(ByVal lRuleID As Long)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    With vsMarkets
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If ValOfText(.TextMatrix(lIndex, MGCol(eMGCol_RuleID))) = lRuleID Then
                .RemoveItem lIndex
            End If
        Next lIndex
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RemoveMarket", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitMarketsGrid
'' Description: Initialize the markets grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitMarketsGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With vsMarkets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Clear
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionByRow
        .ExtendLastCol = True
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        .GridLines = flexGridFlat
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ScrollBars = flexScrollBarBoth
        .ScrollTips = False
        .ScrollTrack = True
        .Cols = MGCol(eMGCOl_NumCols)
        
        .TextMatrix(0, MGCol(eMGCol_ParmName)) = "Identifier"
        .TextMatrix(0, MGCol(eMGCol_Security)) = "Security"
        .TextMatrix(0, MGCol(eMGCol_SecType)) = "Sec Type"
        .TextMatrix(0, MGCol(eMGCol_SymbolPath)) = "Symbol Path"
        .TextMatrix(0, MGCol(eMGCol_Symbol)) = "Symbol"
        .TextMatrix(0, MGCol(eMGCol_Period)) = "Bar Period"
        .TextMatrix(0, MGCol(eMGCol_Format)) = "Format"
        
        .ColHidden(MGCol(eMGCol_MarketSymbol)) = True
        .ColHidden(MGCol(eMGCol_ParmID)) = True
        .ColHidden(MGCol(eMGCol_RuleID)) = True
        .ColHidden(MGCol(eMGCol_Sort)) = True
        .ColHidden(MGCol(eMGCol_GroupID)) = True
        .ColHidden(MGCol(eMGCol_SymbolID)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.InitMarketsGrid", eGDRaiseError_Raise
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMarketsGrid
'' Description: Load the markets grid from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMarketsGrid()
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    Dim Security As cSystemSecurity
    Dim lRedraw As Long
    Dim bOtherThanCN As Boolean
    
    With vsMarkets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        'Load ALL securities (including dups by ParmName) into grid
        For lIndex = 1 To m.System.Securities.Count
            Set Security = m.System.Securities.Item(lIndex)
            With Security
                If .Format <> "CN" Then bOtherThanCN = True
                AddMarketToGrid .ParmName, .SecurityName, .SecurityType, _
                    .Format, .Path, .Symbol, .Periodicity, .MarketSymbol, _
                    .ParmID, .RuleID
            End With
        Next lIndex
        
        'Hide duplicate securities...
        HideDuplicateMarkets
        
        ' If only has ChartNav data then hide the Path and Format columns
        .ColHidden(MGCol(eMGCol_Format)) = Not bOtherThanCN
        .ColHidden(MGCol(eMGCol_SymbolPath)) = Not bOtherThanCN
        
        SetBackColors vsMarkets
    
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.LoadMarketsGrid", eGDRaiseError_Raise
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideDuplicateMarkets
'' Description: Hide duplicate markets
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideDuplicateMarkets()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim Y As Long
    Dim strParmName As String
    Dim lRedraw As Long
    
    With vsMarkets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        'Sort grid by Parm Name first...
        If .Rows - 1 > 0 Then
            .Select 1, MGCol(eMGCol_Sort), .Rows - 1
            .Sort = flexSortGenericAscending
        End If
        
        'Reset all lines to unhidden...
        For X = .FixedRows To .Rows - 1
            .RowHidden(X) = False
        Next X
        
        'Only show one security per group.  For example, if "Market1" is
        'used by 10 rules, only show it once...
        For X = .FixedRows To .Rows - 1
            If Not .RowHidden(X) Then
                Select Case UCase(.TextMatrix(X, MGCol(eMGCol_ParmName)))
'                    Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY"
'                        .RowHidden(X) = True
                    Case Else
                        strParmName = .TextMatrix(X, MGCol(eMGCol_ParmName))
                        For Y = X + 1 To .Rows - 1
                            If strParmName = .TextMatrix(Y, MGCol(eMGCol_ParmName)) Then
                                .RowHidden(Y) = True
                            Else
                                Exit For
                            End If
                        Next Y
                        X = Y - 1
                End Select
            End If
        Next X
    
        'Enable/Disable the Browse security buttons...
        If .Rows - 1 <> 0 Then
            .Row = 1
            cmdBrowse.Enabled = True
            cmdSettings.Enabled = True
        Else
            cmdBrowse.Enabled = False
            cmdSettings.Enabled = False
        End If
        
        'Re-sort and display the securities data...
        If .Rows - 1 > 0 Then
            .Select 1, MGCol(eMGCol_Sort)
            .Sort = flexSortGenericAscending
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.HideDuplicateMarkets", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddMarketToGrid
'' Description: Add a market to the grid
'' Inputs:      Market information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddMarketToGrid(Parm As cInput, Optional ByVal lRowToAdd As Long = -1&)
On Error GoTo ErrSection:
    
    Dim lRow As Long
    Dim lRedraw As Long
    
    With vsMarkets
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRowToAdd = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        Else
            lRow = lRowToAdd
        End If
        
        .TextMatrix(lRow, MGCol(eMGCol_ParmName)) = Parm.ParmName
        Select Case UCase(Parm.ParmName)
            Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY"
                Parm.Periodicity = Left(Parm.ParmName, 1)
        End Select
        .TextMatrix(lRow, MGCol(eMGCol_Security)) = Parm.SecurityName
        .TextMatrix(lRow, MGCol(eMGCol_SecType)) = Parm.SecurityType
        .TextMatrix(lRow, MGCol(eMGCol_SymbolPath)) = SetPath(Parm.Format, Parm.Path)
        If Parm.SymbolID = 0 Then
            .TextMatrix(lRow, MGCol(eMGCol_Symbol)) = Parm.Symbol
        Else
            .TextMatrix(lRow, MGCol(eMGCol_Symbol)) = GetSymbol(Parm.SymbolID)
        End If
        .TextMatrix(lRow, MGCol(eMGCol_MarketSymbol)) = Parm.MarketSymbol
        If Len(Parm.Periodicity) = 1 Then
            Select Case UCase(Parm.Periodicity)
                Case "D"
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Daily"
                Case "W"
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Weekly"
                Case "M"
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Monthly"
                Case "Q"
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Quarterly"
                Case "Y"
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Yearly"
                Case "("
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = "(Default)"
            End Select
        Else
            .TextMatrix(lRow, MGCol(eMGCol_Period)) = FixPeriod(Parm.Periodicity)
        End If
        .TextMatrix(lRow, MGCol(eMGCol_Format)) = Parm.Format
        .TextMatrix(lRow, MGCol(eMGCol_ParmID)) = Parm.ParmID
        .TextMatrix(lRow, MGCol(eMGCol_RuleID)) = Parm.RuleID
        If Parm.ParmName = "Market1" Then
            .TextMatrix(lRow, MGCol(eMGCol_Sort)) = " "
        Else
            .TextMatrix(lRow, MGCol(eMGCol_Sort)) = Parm.ParmName
        End If
        .TextMatrix(lRow, MGCol(eMGCol_GroupID)) = Parm.GroupID
        .TextMatrix(lRow, MGCol(eMGCol_SymbolID)) = Parm.SymbolID
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.AddMarketToGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MarketInfo
'' Description: Allow the user to edit the market information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MarketInfo(Optional ByVal strSymbol As String = "") As Boolean
On Error GoTo ErrSection:

    Dim strSym As String
    Dim astrSymbols As New cGdArray

    With vsMarkets
        If Len(strSymbol) > 0 Then
            strSym = strSymbol
        ElseIf .RowSel >= .FixedRows And .RowSel < .Rows Then
            If Len(.TextMatrix(.RowSel, MGCol(eMGCol_GroupID))) > 0 Then
                astrSymbols.Create eGDARRAY_Strings
                Set astrSymbols = frmSymbolSelector.ShowMe("", False, False, "Select a Symbol to view Market Information", False)
                If astrSymbols.Size > 0 Then strSym = astrSymbols(0)
            ElseIf .TextMatrix(.RowSel, MGCol(eMGCol_SymbolID)) = "0" Then
                strSym = .TextMatrix(.RowSel, MGCol(eMGCol_Symbol))
            Else
                strSym = GetSymbol(CLng(Val(.TextMatrix(.RowSel, MGCol(eMGCol_SymbolID)))))
            End If
        End If
        
        If Len(strSym) > 0 Then MarketInfo = frmMarkets.ShowMe(strSym)
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.MarketInfo", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DefaultSecurityInfo
'' Description: See Below
'' Inputs:      Market Information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This returns default security information for symbols added to the system
'with no symbol description (Symbol="").  Each new symbol is stored on the
'INI file as a string.  This routine parses the string and returns the info.
'Market1=US-9967|C:\GD\BACK67\QMASTER.|T. BONDS            |F|CSI|D
'GC=GC-9967|C:\GD\BACK67\QMASTER.|COMEX GOLD          |F|CSI|D
'TQ=TQ-9967|C:\GD\BACK67\QMASTER.|DAY T-BONDS         |F|CSI|D
Private Sub DefaultSecurityInfo(pParmName As String, pSymbol As String, _
    pSecurity As String, pSecType As String, pFormat As String, _
    pPeriod As String, pPath As String, pMarketSymbol As String, pGroupID As String, lSymbolID As Long)
On Error GoTo ErrSection:
    
    Dim strRecord           As String
    Dim strTempPath         As String
    Dim astrRecord As New cGdArray      ' Array broken out from delimited string
    
    strRecord = GetIniFileProperty(pParmName, "", "Securities", g.strIniFile)
    astrRecord.SplitFields strRecord, "|"
        
    If astrRecord.Size > 0 Then
        pSymbol = astrRecord(0)
        strTempPath = astrRecord(1)
        pSecurity = astrRecord(2)
        pSecType = astrRecord(3)
        pFormat = astrRecord(4)
        If pFormat = "" Then pFormat = "CN"
        pPeriod = astrRecord(5)
        pGroupID = astrRecord(6)
        pPath = SetPath(pFormat, strTempPath)
        pMarketSymbol = GetMarketSymbol(pSymbol, pSecType)
        If Len(astrRecord(7)) > 0 Then
            lSymbolID = CLng(Val(astrRecord(7)))
        ElseIf IsAlpha(pGroupID) = False Then
            lSymbolID = CLng(Val(pGroupID))
            pGroupID = ""
            
            ' This block is in here because of a bug in saving the default security information,
            ' so we will also save it correctly right here (10/18/2005 DAJ)...
            astrRecord(6) = ""
            astrRecord(7) = Str(lSymbolID)
            SetIniFileProperty pParmName, astrRecord.JoinFields("|"), "Securities", g.strIniFile
        Else
            lSymbolID = GetSymbolID(pSymbol)
        End If
    Else
        pSymbol = ""
        pSecurity = ""
        pSecType = ""
        pFormat = ""
        pPeriod = ""
        pPath = ""
        pMarketSymbol = ""
        pGroupID = ""
        lSymbolID = 0&
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.DefaultSecurityInfo", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPath
'' Description: Set the path correctly for an external data file
'' Inputs:      Format and Path of the market
'' Returns:     Corrected Path
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetPath(ByVal strFormat As String, ByVal strPath As String) As String
On Error GoTo ErrSection:
    
    Dim iPlace As Long

    Select Case UCase(strFormat)
        Case "CSI"
            iPlace = InStr(strPath, "\Qmaster")
            If iPlace <> 0 Then strPath = Mid(strPath, 1, iPlace - 1)
        Case "MS7"
            iPlace = InStr(strPath, "\Master")
            If iPlace <> 0 Then strPath = Mid(strPath, 1, iPlace - 1)
        Case "CN"
            strPath = AddSlash(App.Path) & "Data"
    End Select
    
    SetPath = strPath
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.SetPath", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetMarketSymbol
'' Description: Get the market symbol for a Genesis symbol
'' Inputs:      Symbol, Security Type
'' Returns:     Market Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMarketSymbol(ByVal strSymbol As String, ByVal strSecType As String) As String
On Error GoTo ErrSection:
    
    Dim lLoc        As Long
    
    If strSecType = "S" Then
        GetMarketSymbol = "!"
    ElseIf strSecType = "I" Then
        GetMarketSymbol = "$"
    Else
        lLoc = InStr(strSymbol, "-")
        If lLoc = 0 Then lLoc = InStr(strSymbol, "_")
        If lLoc = 0 Then
            If strSecType = "F" Then
                GetMarketSymbol = Mid(strSymbol, 1, 3)
            Else
                GetMarketSymbol = strSymbol
            End If
        Else
            GetMarketSymbol = Mid(strSymbol, 1, lLoc - 1)
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.GetMarketSymbol", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetRuleNumbers
'' Description: Get the next rule numbers for the appropriate types
'' Inputs:      Long Entry, Long Exit, Short Entry, Short Exit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetRuleNumbers(lLE As Long, lLX As Long, lSE As Long, lSX As Long)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid
    Dim X As Long                       ' Temporary variable
    Dim strRuleName As String           ' Name of the rule on the current row
    
    ' Initialize the variables
    lLE = 0
    lLX = 0
    lSE = 0
    lSX = 0
    
    With vsRules
        For lRow = .FixedRows To .Rows - 1
            strRuleName = UCase(.TextMatrix(lRow, RGCol(eRGCol_RuleName)))
            If IsDefaultRuleName(strRuleName) Then
                If Left(strRuleName, 12) = "LONG ENTRY #" Then
                    X = DefaultRuleNumber(strRuleName)
                    If X > lLE Then lLE = X
                ElseIf Left(strRuleName, 11) = "LONG EXIT #" Then
                    X = DefaultRuleNumber(strRuleName)
                    If X > lLX Then lLX = X
                ElseIf Left(strRuleName, 13) = "SHORT ENTRY #" Then
                    X = DefaultRuleNumber(strRuleName)
                    If X > lSE Then lSE = X
                ElseIf Left(strRuleName, 12) = "SHORT EXIT #" Then
                    X = DefaultRuleNumber(strRuleName)
                    If X > lSX Then lSX = X
                End If
            End If
        Next lRow
    End With
    
    ' Next available number will be one more than what we have...
    lLE = lLE + 1
    lLX = lLX + 1
    lSE = lSE + 1
    lSX = lSX + 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.GetRuleNumbers", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleNameExists
'' Description: Determine whether a rule name exists in the system
'' Inputs:      Name to search for
'' Returns:     True if Exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RuleNameExists(ByVal strName As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long

    For lIndex = 1 To m.System.Rules.Count
        If m.System.Rules.Item(lIndex).Name = strName Then
            RuleNameExists = True
            Exit For
        End If
    Next lIndex
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.RuleNameExists", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateRules
'' Description: Determine if any rules in the system have changed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateRules()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim lMatch As Long                  ' Grid row for given Rule ID
    Dim alRuleIDs As New cGdArray       ' Array of Rule ID's for the system
    Dim alSorted As New cGdArray        ' Sorted index of the Rule ID's
    Dim rs As Recordset                 ' Recordset from the database
    Dim Rule As New cRule
    
    Screen.MousePointer = vbHourglass
    vsRules.Redraw = flexRDNone
    vsInputs.Redraw = flexRDNone
    vsMarkets.Redraw = flexRDNone
    
    ' Store all Rule ID's from grid
    alRuleIDs.Create eGDARRAY_Longs, vsRules.Rows - 1
    alSorted.Create eGDARRAY_Longs, vsRules.Rows - 1
    For lIndex = 1 To vsRules.Rows - 1
        alRuleIDs(lIndex) = CLng(vsRules.TextMatrix(lIndex, RGCol(eRGCol_RuleID)))
    Next
    gdSortAsIndex alSorted.ArrayHandle, alRuleIDs.ArrayHandle, 1, eGdSort_Default, 1, alRuleIDs.Size - 1
            
    ' Get query of all System Rules from DB
    Set rs = g.dbNav.OpenRecordset("SELECT tblSystemRules.*, tblRules.* " & _
                "FROM tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID " & _
                "WHERE (((tblSystemRules.SystemNumber)=" & m.System.SystemNumber & "));", dbOpenDynaset)
    
    ValidateCheckSums rs, "tblRules"
    ValidateCheckSums rs, "tblSystemRules"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do Until rs.EOF
        If gdBinarySearchAsIndex(alSorted.ArrayHandle, alRuleIDs.ArrayHandle, rs![tblRules.RuleID], lMatch, eGdSort_Default, 1, alRuleIDs.Size - 1) Then
            If rs![tblRules.CheckSum] = 0.5 Or rs![tblSystemRules.CheckSum] = 0.5 Then
                EnableToolbar True
                
            ElseIf Round(rs!LastModified, 6) > Round(Val(vsRules.TextMatrix(alSorted(lMatch), RGCol(eRGCol_LastModKnown))), 6) Then
                EnableToolbar True
                Set Rule = New cRule
                'Rule.RuleID = rs![tblRules.RuleID]
                'Rule.Load
                Rule.LoadWithSystemInfo rs![tblRules.RuleID]
                AddRule Rule 'rs![tblRules.RuleID]
            End If
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
    ' See if any rules have been deleted on us ...
    'Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] ORDER BY [RuleID];", dbOpenDynaset)
    'ValidateCheckSums rs, "tblRules"
    'For lIndex = alRuleIDs.Size - 1 To 1 Step -1
    '    rs.FindFirst "[RuleID]=" & alRuleIDs(lIndex)
    '    If rs.NoMatch Then
    '        EnableToolbar True
    '        RemoveRule lIndex
    '    ElseIf rs!CheckSum = 0.5 Then
    '        EnableToolbar True
    '        vsRules.RemoveItem lIndex
    '    ElseIf rs!Reverify Then
    '        vsRules.Cell(flexcpForeColor, lIndex, RGCol(eRGCol_RuleName)) = vbRed
    '    Else
    '        vsRules.Cell(flexcpForeColor, lIndex, RGCol(eRGCol_RuleName)) = vsRules.Cell(flexcpForeColor, lIndex, RGCol(eRGCol_Alt))
    '    End If
    'Next

ErrExit:
    vsRules.Redraw = flexRDBuffered
    vsInputs.Redraw = flexRDBuffered
    vsMarkets.Redraw = flexRDBuffered
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set Rule = Nothing
    Exit Sub

ErrSection:
    vsRules.Redraw = flexRDBuffered
    vsInputs.Redraw = flexRDBuffered
    vsMarkets.Redraw = flexRDBuffered
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set Rule = Nothing
    RaiseError "frmSystemManager.RuleChanged", eGDRaiseError_Raise

End Sub

Public Sub ShowSystem(SystemToLoad As cSystem)
On Error GoTo ErrSection:

    Set m.System = SystemToLoad
    LoadRec m.System.SystemNumber
    EnableToolbar True
    ShowForm Me, False, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.ShowSystem", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub EnableToolbar(ByVal bDirty As Boolean)
On Error GoTo ErrSection:

    With tbToolbar
        .Tools("ID_Save").Enabled = bDirty
        .Tools("ID_SaveAs").Enabled = (Trim(m.strName) <> "")
        .Tools("ID_Rename").Enabled = (Trim(m.strName) <> "")
        .Tools("ID_Toolbox").Enabled = Not m.bModal
        
        .Tools("ID_Run").Enabled = (vsRules.Rows > vsRules.FixedRows)
        .Tools("ID_RunGroup").Enabled = (vsRules.Rows > vsRules.FixedRows)
        .Tools("ID_Orders").Enabled = (vsRules.Rows > vsRules.FixedRows)
        
        If vsMarkets.Rows > vsMarkets.FixedRows Then
            If Len(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_GroupID))) > 0 Then
                .Tools("ID_Run").Picture = Picture16(ToolbarIcon("ID_StrategyBaskets"))
                
                ' TLB 7/1/2014: I'm not sure we really need this button anymore?
                ' (since it's actually better to run a symbol group using the optimizer form)
                .Tools("ID_RunGroup").Visible = False 'True
            Else
                .Tools("ID_Run").Picture = Picture16(ToolbarIcon("ID_Performance"))
                .Tools("ID_RunGroup").Visible = False
            End If
        Else
            .Tools("ID_Run").Picture = Picture16(ToolbarIcon("ID_Performance"))
            .Tools("ID_RunGroup").Visible = False
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.EnableToolbar", eGDRaiseError_Raise
    
End Sub

Private Function Market1Row() As Long
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    Market1Row = -1&
    With vsMarkets
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, MGCol(eMGCol_ParmName)) = "Market1" Then
                Market1Row = lIndex
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.Market1Row", eGDRaiseError_Raise
    
End Function

Private Sub DisplayPyramidInfo(ByVal Row As Long)
On Error GoTo ErrSection:

    Dim bEnter As Boolean
    Dim bPercent As Boolean
    Dim bPosition As Boolean
    Dim lNumContracts As Long
    Dim lMarket1Row As Long
    Dim strUnits As String
    Dim lNumUnits As Long
    
    lMarket1Row = Market1Row
    If lMarket1Row > -1 Then
        Select Case vsMarkets.TextMatrix(lMarket1Row, MGCol(eMGCol_SecType))
            ' TLB 10/31/2011 #6459: "shares" should only be for stocks, not forex
            Case "S" ', "I"
                strUnits = "Share(s)"
                lNumUnits = CLng(ValOfText(txtNumShares.Text))
            
            Case Else
                strUnits = "Contract(s)"
                lNumUnits = 1&
        End Select
    Else
        strUnits = "Contract(s)"
        lNumUnits = 1&
    End If

    With vsRules
        bEnter = (.TextMatrix(Row, RGCol(eRGCol_RuleUse)) = "0")
        bPercent = CheckedCell(vsRules, Row, RGCol(eRGCol_AsPercent))
        bPosition = Not CheckedCell(vsRules, Row, RGCol(eRGCol_ExitBasedOnTrade))
        lNumContracts = CLng(.TextMatrix(Row, RGCol(eRGCol_NumContracts)))
        
        If bEnter Then
            .TextMatrix(Row, RGCol(eRGCol_PyramidInfo)) = _
                    "Enter " & Format(lNumContracts * lNumUnits, "#,##0") & " " & strUnits
        Else
            If bPosition = True Then
                If bPercent = True Then
                    .TextMatrix(Row, RGCol(eRGCol_PyramidInfo)) = _
                        "Exit " & Trim(CStr(lNumContracts)) & "% of position"
                Else
                    .TextMatrix(Row, RGCol(eRGCol_PyramidInfo)) = _
                        "Exit " & Format(lNumContracts * lNumUnits, "#,##0") & " " & strUnits
                End If
            Else
                .TextMatrix(Row, RGCol(eRGCol_PyramidInfo)) = "Exit based on trade"
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.DisplayPyramidInfo", eGDRaiseError_Raise
    
End Sub

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    
    If AskToSaveRules Then
        AskToSave = True
    Else
        If tbToolbar.Tools("ID_Save").Enabled Then
            If WindowState = vbMinimized Then WindowState = vbNormal
    
            strResponse = InfBox("Do you want to save your changes?||Clicking No will undo any changes you have made to the strategy or any rules in the strategy.", "?", "+Yes|No|-Cancel", Caption)
            Select Case strResponse
                Case "C"
                    AskToSave = True
                Case "Y"
                    Save "ID_Save"
                Case "N"
                    RemoveLocalRules
            End Select
        End If
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError Me.Name & ".AskToSave"

End Function

Private Sub RunTest(ByVal bNextBarReport As Boolean, Optional ByVal lRuleID As Long = 0&)
On Error GoTo ErrSection:

    Dim i&, j&, nMinFZ&, nMaxFZ&, nMinLB&, nMaxLB&
    Dim s$, strPeriod$, strSymbol$, strSpeeds$
    Dim bRunGroup As Boolean
    Dim aSymbols As New cGdArray

    If ValidateMarkets = False Then Exit Sub
    If lRuleID <> 0& Then
        If Not HasPlatinum(True) Then Exit Sub
    End If
    
    If optToEndOfData = False Then
        If dtpToDate.Value < dtpFromDate.Value Then
            Err.Raise vbObjectError + 1000, , "To Date must be after From Date"
        End If
    End If

    FixSequence

    With m.System
        .SystemName = m.strName
        .FromDate = dtpFromDate.Value
        .ToDate = dtpToDate.Value
        .LinkInputs = (chkLinkInputs = vbChecked)
        If .LinkInputs Then SetDuplicates
        
        .Expenses = ValOfText(txtCommission.Text)
        .BarsLoadedBeforeTrading = ValOfText(txtBarsLoadedBeforeTrading.Text)
        .BarsTradedBeforeOrders = ValOfText(txtBarsTradedBeforeOrders.Text)
        For i = vsMarkets.FixedRows To vsMarkets.Rows - 1
            If vsMarkets.TextMatrix(i, 0) = "Market1" Then
                .BarTimeFrame = vsMarkets.TextMatrix(i, 2)
                Exit For
            End If
        Next
        .ToEndOfData = optToEndOfData.Value * -1
        .MMid = 0
        .Pyramid = Val(chkPyramid.Value) * -1
        .TradeDepth = ValOfText(txtTradeDepth.Text)
        .UseSharesPerTrade = optSharesPerTrade.Value
        .NumShares = CLng(ValOfText(txtNumShares.Text))
        .DollarsPerTrade = ValOfText(txtDollarsPerTrade.Text)
        .StockExpenses = ValOfText(txtStockCommission.Text)
        .ForexExpenses = ValOfText(txtForexCommission.Text)
        .AllowReverse = (chkAllowReverse.Value = vbChecked)
        .ForceLimitThrough = (chkForceLimitThrough.Value = vbChecked)
        
        If lRuleID = 0& Then
            Select Case True
                Case optSignals(optLong), optSignals(optLongExit)
                    If InfBox("Run all entries or only long entries?", "?", "+All|-Long", "Strategy Run") = "A" Then
                        .RunMode = eGDRunMode_All
                    Else
                        .RunMode = eGDRunMode_Long
                    End If
                Case optSignals(optShort), optSignals(optShortExit)
                    If InfBox("Run all entries or only short entries?", "?", "+All|-Short", "Strategy Run") = "A" Then
                        .RunMode = eGDRunMode_All
                    Else
                        .RunMode = eGDRunMode_Short
                    End If
                Case Else
                    .RunMode = eGDRunMode_All
            End Select
        Else
            .RunMode = eGDRunMode_All
        End If
        
        GridsToRules
        .ClearBars
                
        bRunGroup = False
        strSymbol = vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_Symbol))
        strPeriod = vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_Period))
        If UCase(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_SecType))) = "GROUP" Then
            bRunGroup = True
        ElseIf UCase(strPeriod) = "FRACTZEN" Then
            ' ONLY for John/Terry, see if want to test different speeds (only if not streaming)
            If Trim(UCase(FileToString(App.Path & "\AutoBreakout.flg", , True))) = "PROJECTX" Then
                ' for now, selecting custom FZ speed doesn't work when optimizing either
                If Not bNextBarReport And Not g.RealTime.Active And Not .Optimized Then
                    s = g.FractZen.GetSpeedInfo(strSymbol)
                    nMinFZ = Val(Parse(s, ",", 1))
                    nMinLB = Val(Parse(s, ",", 2))
                    s = Str(nMinFZ) & "L" & Str(nMinLB)
                    If 0 Then ' .Optimized Then
                        ' when optimizing, can only get a single FractZen speed
                        strSpeeds = InfBox("Enter a single FractZen speed to test:", "?", , "TEST FractZen Speed", , , , , , "s", s)
                        nMaxFZ = 0
                        nMaxLB = 0
                    Else
                        ' when not optimizing, can get a range of FractZen speeds to test
                        s = s & " - " & s
                        strSpeeds = InfBox("Enter range of FractZen speeds to test:", "?", , "TEST FractZen Speed", , , , , , "s", s)
                        nMaxFZ = nMinFZ
                        nMaxLB = nMinLB
                    End If
                    strSpeeds = Trim(UCase(strSpeeds))
                    If Len(strSpeeds) = 0 Then
                        strSymbol = "" ' to ABORT
                    Else
                        s = Parse(strSpeeds, "-", 1)
                        nMinFZ = Val(Parse(s, "L", 1))
                        nMinLB = Val(Parse(s, "L", 2))
                        If nMaxFZ > 0 Then
                            s = Parse(strSpeeds, "-", 2)
                            nMaxFZ = Val(Parse(s, "L", 1))
                            nMaxLB = Val(Parse(s, "L", 2))
                            If nMaxFZ > nMinFZ And nMinFZ > 0 Then
                                bRunGroup = True
                            ElseIf nMaxLB > nMinLB And nMinLB > 0 Then
                                bRunGroup = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If bRunGroup Then
            If nMaxFZ > nMinFZ Or nMaxLB > nMinLB Then
                ' if testing different FractZen speeds
                For i = nMinFZ To nMaxFZ
                    For j = nMinLB To nMaxLB
                        aSymbols.Add strSymbol & vbTab & Str(i) & "L" & Str(j)
                    Next
                Next
            Else
                ' load symbol group
                Set aSymbols = SymbolsInGroup(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_GroupID)))
            End If
            RunMultipleSymbols aSymbols, bNextBarReport, lRuleID
        ElseIf Len(strSymbol) > 0 Then
            ' load market1 data (except for "external" symbols -- e.g. CSI or MS)
            If Left(strSymbol, 1) = "*" Then
                .Bars = Nothing ' just let cSystem.RunEngine load the external data
            Else
                If nMinFZ > 0 Then
                    g.FractZen.SetSpeedInfo strSymbol, nMinFZ, nMinLB
                End If
                If UCase(strPeriod) = "FRACTZEN" Or GetPeriodicity(strPeriod) < ePRD_Days Then
                    InfBox "Please wait while loading data for your strategy ...", "t", , "Processing", True
                End If
                .LoadMarket1Bars strSymbol, strPeriod, bNextBarReport
                If UCase(strPeriod) = "FRACTZEN" Or GetPeriodicity(strPeriod) < ePRD_Days Then
                    InfBox "Please wait while back-testing your strategy ...", "t", , "Processing", True
                End If
            End If
            ' run the backtest
            If bNextBarReport Then
                .NextBarReport
            Else
                .Test False, lRuleID
            End If
            InfBox ""
        End If
        If nMinFZ > 0 Then
            ' clear the FractZen speed back to it's normal default
            g.FractZen.SetSpeedInfo strSymbol
        End If
        
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.Test", eGDRaiseError_Raise
    
End Sub

Private Function ValidateMarkets() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long
    Dim Bars As New cGdBars
    Dim astrSymbols As New cGdArray
    
    ValidateMarkets = True
    
    With vsMarkets
        For lIndex = .FixedRows To .Rows - 1
            ' Validate that the user has a symbol selected for each market...
            If Len(.TextMatrix(lIndex, MGCol(eMGCol_Symbol))) = 0 Then
                .Row = lIndex
                .RowSel = lIndex
                
                BrowseMarkets
                If Len(.TextMatrix(lIndex, MGCol(eMGCol_Symbol))) = 0 Then
                    ValidateMarkets = False
                    Exit For
                End If
                
            ElseIf UCase(.TextMatrix(lIndex, MGCol(eMGCol_Period))) = "FRACTZEN" And Not g.FractZen.Allowed Then
                InfBox "You are not authorized to use FractZen bars", "!", , kErrorCaption
                ValidateMarkets = False
                Exit For
                
            ' Validate that there is market information for each symbol...
            Else
                If UCase(.TextMatrix(lIndex, MGCol(eMGCol_SecType))) = "GROUP" Then
                    Set astrSymbols = SymbolsInGroup(.TextMatrix(lIndex, MGCol(eMGCol_GroupID)))
                    For lIndex2 = 0 To astrSymbols.Size - 1
                        GetMarketInfo astrSymbols(lIndex2), Bars
                        If Bars.Prop(eBARS_TickValue) = 0 Or Bars.Prop(eBARS_TickMove) = 0 Or Bars.Prop(eBARS_MinMoveInTicks) = 0 Then
                            .Row = lIndex
                            .RowSel = lIndex
                        
                            If Not MarketInfo(astrSymbols(lIndex2)) Then
                                ValidateMarkets = False
                                Exit For
                            End If
                        End If
                    Next lIndex2
                    If Not ValidateMarkets Then Exit For
                Else
                    GetMarketInfo .TextMatrix(lIndex, MGCol(eMGCol_Symbol)), Bars
                    If Bars.Prop(eBARS_TickValue) = 0 Or Bars.Prop(eBARS_TickMove) = 0 Or Bars.Prop(eBARS_MinMoveInTicks) = 0 Then
                        .Row = lIndex
                        .RowSel = lIndex
                    
                        If Not MarketInfo Then
                            ValidateMarkets = False
                            Exit For
                        End If
                    End If
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Set Bars = Nothing
    Exit Function
    
ErrSection:
    Set Bars = Nothing
    RaiseError "frmSystemManager.ValidateMarkets", eGDRaiseError_Raise
    
End Function

Private Sub BrowseMarkets(Optional ByVal strNewSym$ = "")
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray
    Dim strSymbol As String
    Dim Bars As New cGdBars
    Dim strInputName As String
    Dim lRedraw As Long
    Dim X As Long
    Dim strPeriod As String
    Dim strPeriodMkt1 As String
    Dim strRecord As String
    Dim strFormat As String
    Dim strPath As String
    Dim strSecType As String
    Dim bRedisplayPyramid As Boolean
    Dim strParmName As String
    Dim lSymbolID As Long
    Dim lRow As Long
    
    cmdBrowse.Enabled = False
    
    With vsMarkets
        If .Row >= .FixedRows And .Row < .Rows Then
            Select Case UCase(.TextMatrix(.Row, MGCol(eMGCol_ParmName)))
                Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY", "UNSPLIT"
                    .Row = .FixedRows
                    .RowSel = .Row
            End Select
            
            strSymbol = .TextMatrix(.Row, MGCol(eMGCol_Symbol))
            strParmName = .TextMatrix(.Row, MGCol(eMGCol_ParmName))
            strSecType = .TextMatrix(.Row, MGCol(eMGCol_SecType))
            lSymbolID = CLng(Val(.TextMatrix(.Row, MGCol(eMGCol_SymbolID))))
            If UCase(strParmName) = "MARKET1" Then
                strPeriodMkt1 = .TextMatrix(.Row, MGCol(eMGCol_Period))
            End If
            lRow = .Row
        End If
    End With
    
    ' Don't allow the user to edit one of the special "Symbol,Period" parms...
    If Left(strParmName, 1) = Chr(34) And Right(strParmName, 1) = Chr(34) Then
        InfBox Chr(34) & "Symbol,Period" & Chr(34) & " parameters cannot be edited", "!", , "Browse Markets"
        Exit Sub
    End If
        
    Bars.Prop(eBARS_Symbol) = strSymbol
    strSymbol = Bars.Prop(eBARS_Symbol)
    astrSymbols.Create eGDARRAY_Strings
    If Len(strNewSym) > 0 Then
        astrSymbols.Add strNewSym
    Else
        If strSecType = "Group" Then
            Set astrSymbols = frmSymbolSelector.ShowMe("", False, True, "Select a Symbol for " & strParmName, True)
        Else
            Set astrSymbols = frmSymbolSelector.ShowMe(strSymbol, False, True, "Select a Symbol for " & strParmName, True)
        End If
    End If
    If astrSymbols.Size > 0 Then
        With vsMarkets
            lRedraw = .Redraw
            .Redraw = flexRDNone
            
            If InStr(astrSymbols(0), "|") = 0 Then
                SetBarProperties Bars, astrSymbols(0)
                strInputName = .TextMatrix(.Row, MGCol(eMGCol_ParmName))
                        
                If UCase(strInputName) = "MARKET1" And .TextMatrix(lRow, MGCol(eMGCol_SecType)) <> Bars.SecurityType Then
                    bRedisplayPyramid = True
                End If
                .TextMatrix(lRow, MGCol(eMGCol_Security)) = Bars.Prop(eBARS_Desc)
                .TextMatrix(lRow, MGCol(eMGCol_SecType)) = Bars.SecurityType
                .TextMatrix(lRow, MGCol(eMGCol_SymbolPath)) = App.Path
                .TextMatrix(lRow, MGCol(eMGCol_Symbol)) = Bars.Prop(eBARS_Symbol)
                If .TextMatrix(lRow, MGCol(eMGCol_Period)) = "" Then
                    If UCase(strInputName) = "MARKET1" Then
                        .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Daily"
                    Else
                        .TextMatrix(lRow, MGCol(eMGCol_Period)) = "(Default)"
                    End If
                End If
                If Len(strPeriodMkt1) > 0 Then
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = strPeriodMkt1
                End If
                strPeriod = .TextMatrix(lRow, MGCol(eMGCol_Period))
                .TextMatrix(lRow, MGCol(eMGCol_Format)) = "CN"
                .TextMatrix(lRow, MGCol(eMGCol_MarketSymbol)) = Bars.Prop(eBARS_MarketSymbol)
                .TextMatrix(lRow, MGCol(eMGCol_GroupID)) = ""
                .TextMatrix(lRow, MGCol(eMGCol_SymbolID)) = Str(Bars.Prop(eBARS_SymbolID))
                
                SyncMarkets
                .ColHidden(MGCol(eMGCol_SymbolPath)) = True
                .ColHidden(MGCol(eMGCol_Format)) = True
                
                ' Update changes to INI securities section (most recent changes)...
                strRecord = Bars.Prop(eBARS_Symbol) & "|" & AddSlash(App.Path) & "Data" & "|" & Bars.Prop(eBARS_Desc)
                strRecord = strRecord & "|" & Bars.SecurityType & "|CN"
                strRecord = strRecord & "|" & strPeriod & "||" & Str(Bars.Prop(eBARS_SymbolID)) & "|"
                SetIniFileProperty strInputName, strRecord, "Securities", g.strIniFile
            Else
                strInputName = .TextMatrix(.Row, MGCol(eMGCol_ParmName))
                        
                If UCase(strInputName) = "MARKET1" And Parse(astrSymbols(0), "|", 4) <> Bars.SecurityType Then
                    bRedisplayPyramid = True
                End If
                .TextMatrix(lRow, MGCol(eMGCol_Security)) = Parse(astrSymbols(0), "|", 3)
                .TextMatrix(lRow, MGCol(eMGCol_SecType)) = Parse(astrSymbols(0), "|", 4)
                .TextMatrix(lRow, MGCol(eMGCol_SymbolPath)) = Parse(astrSymbols(0), "|", 2)
                .TextMatrix(lRow, MGCol(eMGCol_Symbol)) = Parse(astrSymbols(0), "|", 1)
                If .TextMatrix(lRow, MGCol(eMGCol_Period)) = "" Then
                    If UCase(strInputName) = "MARKET1" Then
                        .TextMatrix(lRow, MGCol(eMGCol_Period)) = "Daily"
                    Else
                        .TextMatrix(lRow, MGCol(eMGCol_Period)) = "(Default)"
                    End If
                End If
                If Len(strPeriodMkt1) > 0 Then
                    .TextMatrix(lRow, MGCol(eMGCol_Period)) = strPeriodMkt1
                End If
                strPeriod = .TextMatrix(lRow, MGCol(eMGCol_Period))
                .TextMatrix(lRow, MGCol(eMGCol_Format)) = Parse(astrSymbols(0), "|", 5)
                .TextMatrix(lRow, MGCol(eMGCol_MarketSymbol)) = GetMarketSymbol(Parse(astrSymbols(0), "|", 1), Parse(astrSymbols(0), "|", 4))
                .TextMatrix(lRow, MGCol(eMGCol_GroupID)) = ""
                .TextMatrix(lRow, MGCol(eMGCol_SymbolID)) = "0"
                SyncMarkets
                
                .ColHidden(MGCol(eMGCol_SymbolPath)) = False
                .ColHidden(MGCol(eMGCol_Format)) = False
                
                SetIniFileProperty strInputName, astrSymbols(0), "Securities", g.strIniFile
            End If
            
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = lRedraw
        End With
    End If
    
    If bRedisplayPyramid Then
        With vsRules
            For X = .FixedRows To .Rows - 1
                DisplayPyramidInfo X
            Next X
        End With
    End If
    
    EnableToolbar Not m.bLinkedToChart
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.BrowseMarkets", eGDRaiseError_Raise

End Sub

Private Function NewRule(Optional ByVal strTextToPaste As String = "") As Boolean
On Error GoTo ErrSection:
    
    Dim lLE&, lLX&, lSE&, lSX&          ' Last default rule for each type
    Dim frm As frmRule
    Dim activeFrm As Form
    Dim Rule As New cRule
    Dim strNewRule As String
    Dim lActionIdx As Long              'passed to form rule to set action drop box list index
    Dim strCBPrompt$
    
    If Not HasPlatinum(True) Then Exit Function
    If Not ShowPassword Then Exit Function
        
    m.strCondBuilderExpr = ""
    If Len(strTextToPaste) = 0 Then
        strCBPrompt = GetIniFileProperty("UseCondBuilderNewRule", "", "DontAsk", g.strIniFile)
        If Len(strCBPrompt) = 0 Then
            strCBPrompt = "Would you like to build the rule's condition|using indicators from the active chart?"
            strCBPrompt = InfBox(strCBPrompt, "?", "+Yes|-No", "Condition Builder", , , , , , , , , True)
            If InStr(strCBPrompt, "-") > 0 Then
                ' don't ask anymore, store for future use
                Call SetIniFileProperty("UseCondBuilderNewRule", "N", "DontAsk", g.strIniFile)
            End If
        End If
        If UCase(Left(strCBPrompt, 1)) = "Y" Then
            Set activeFrm = ActiveChart
            If Not activeFrm Is Nothing Then
                frmConditionBuilder.ShowMe activeFrm.Chart, , eType_Rule, Me
                If m.strCondBuilderExpr = "UserCancel" Then Exit Function
            End If
        End If
        strNewRule = m.strCondBuilderExpr
    Else
        strNewRule = strTextToPaste
    End If
        
    NewRule = True
    GetRuleNumbers lLE, lLX, lSE, lSX
    
    lActionIdx = -1
    If tbToolbar.ToolBars("Wizard").Visible Then
        If optSignals(0).Value = True Then
            lActionIdx = 0
        ElseIf optSignals(1).Value = True Then
            lActionIdx = 1
        ElseIf optSignals(2).Value = True Then
            lActionIdx = 2
        ElseIf optSignals(3).Value = True Then
            lActionIdx = 3
        End If
    End If
    
    Set frm = New frmRule
    frm.ShowFromSysMgr Rule, m.System.SystemNumber, m.System.LibraryID, lLE, lLX, lSE, lSX, Me, strNewRule, lActionIdx
    
    EnableToolbar True
    
ErrExit:
    Set Rule = Nothing
    Set frm = Nothing
    Exit Function
    
ErrSection:
    Set Rule = Nothing
    Set frm = Nothing
    RaiseError "frmSystemManager.NewRule", eGDRaiseError_Raise
    
End Function

Private Sub RemoveRulesFromSystem()
On Error GoTo ErrSection:

    Dim X As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid redraw
    Dim lRowSel As Long                 ' Currently selected rowf
    
    If Not HasPlatinum(True) Then Exit Sub
    If Not ShowPassword Then Exit Sub
    
    With vsRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRowSel = .RowSel
        For X = .SelectedRows - 1 To 0 Step -1
            If Not .RowHidden(.SelectedRow(X)) Then RemoveRule .SelectedRow(X)
        Next X
        
        RefreshRulesGrid False
        RefreshInputsGrid
        RefreshMarketsGrid
        
        .Redraw = lRedraw
    End With
            
    MoveFocus vsRules
    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RemoveRulesFromSystem", eGDRaiseError_Raise
    
End Sub

Private Sub QuickStops()
On Error GoTo ErrSection:

    Dim X As Long                    ' Number of rules to add
    Dim lRow As Long                    ' Index into a for loop
    Dim alRuleIDs As cGdArray           ' Rule ID's to add
    Dim adStopValues As cGdArray        ' Values of stop amounts
    Dim lTotStops As Long               ' Total number of rules to add
    Dim dProfitTarget As Double         ' User entered Profit Target
    Dim dStopLoss As Double             ' User entered Stop Loss
    Dim dTrailingStop As Double         ' User entered Trailing Stop
    Dim Rule As cRule
    Dim NewRule As New cRule
    Dim lRuleNum As Long
    Dim lIndex As Long
    Dim lOldParmID As Long
    Dim Parm As cInput
    Dim lLongStopLoss&, lLongTrailingStop&, lLongProfitTarget&
    Dim lShortStopLoss&, lShortTrailingStop&, lShortProfitTarget&
    
    If Not HasPlatinum(True) Then Exit Sub
    If Not ShowPassword Then Exit Sub
    
    lLongStopLoss = kStopLossLongRuleID
    lShortStopLoss = kStopLossShortRuleID
    lLongTrailingStop = kTrailingStopLongRuleID
    lShortTrailingStop = kTrailingStopShortRuleID
    lLongProfitTarget = kProfitTargetLongRuleID
    lShortProfitTarget = kProfitTargetShortRuleID

    ' Get current values if already in the system...
    With vsInputs
        For X = .FixedRows To .Rows - 1
            Select Case UCase(Trim(Parse(.TextMatrix(X, IGCol(eIGCol_RuleName)), vbCrLf, 1)))
                Case UCase("Exit Long: Profit Target")
                    dProfitTarget = ValOfText(.TextMatrix(X, IGCol(eIGCol_InputValue)))
                    lLongProfitTarget = CLng(ValOfText(.TextMatrix(X, IGCol(eIGCol_RuleID))))
                Case UCase("Exit Short: Profit Target")
                    dProfitTarget = ValOfText(.TextMatrix(X, IGCol(eIGCol_InputValue)))
                    lShortProfitTarget = CLng(ValOfText(.TextMatrix(X, IGCol(eIGCol_RuleID))))
                Case UCase("Exit Long: Stop Loss")
                    dStopLoss = ValOfText(.TextMatrix(X, IGCol(eIGCol_InputValue)))
                    lLongStopLoss = CLng(ValOfText(.TextMatrix(X, IGCol(eIGCol_RuleID))))
                Case UCase("Exit Short: Stop Loss")
                    dStopLoss = ValOfText(.TextMatrix(X, IGCol(eIGCol_InputValue)))
                    lShortStopLoss = CLng(ValOfText(.TextMatrix(X, IGCol(eIGCol_RuleID))))
                Case UCase("Exit Long: Trailing Stop")
                    dTrailingStop = ValOfText(.TextMatrix(X, IGCol(eIGCol_InputValue)))
                    lLongTrailingStop = CLng(ValOfText(.TextMatrix(X, IGCol(eIGCol_RuleID))))
                Case UCase("Exit Short: Trailing Stop")
                    dTrailingStop = ValOfText(.TextMatrix(X, IGCol(eIGCol_InputValue)))
                    lShortTrailingStop = CLng(ValOfText(.TextMatrix(X, IGCol(eIGCol_RuleID))))
            End Select
        Next X
    End With
    
    ' Show Quick stops form...
    If frmQuickStops.ShowMe(dProfitTarget, dStopLoss, dTrailingStop) Then
        Set alRuleIDs = New cGdArray
        alRuleIDs.Create eGDARRAY_Longs
        Set adStopValues = New cGdArray
        adStopValues.Create eGDARRAY_Doubles
        
        ' Add Target profit for Long (and short)
        If dProfitTarget > 0 Then
            alRuleIDs.Add lLongProfitTarget
            adStopValues.Add dProfitTarget
            alRuleIDs.Add lShortProfitTarget
            adStopValues.Add dProfitTarget
        End If
    
        ' Add Stop Loss for Long (and short)
        If dStopLoss > 0 Then
            alRuleIDs.Add lLongStopLoss
            adStopValues.Add dStopLoss
            alRuleIDs.Add lShortStopLoss
            adStopValues.Add dStopLoss
        End If
        
        ' Add Trailing Stop for Long (and short)
        If dTrailingStop > 0 Then
            alRuleIDs.Add lLongTrailingStop
            adStopValues.Add dTrailingStop
            alRuleIDs.Add lShortTrailingStop
            adStopValues.Add dTrailingStop
        End If
    
        ' Post Rules, Inputs, Markets to collections...
        For X = 0 To alRuleIDs.Size - 1
            If Not m.System.Rules.Found(CStr(alRuleIDs(X))) Then
                Set Rule = New cRule
                With Rule
                    .LoadWithSystemInfo alRuleIDs(X)
                    
                    If g.Security.CanEdit(.SecurityLevel, .Password, .Name) Then
                        Set NewRule = .MakeCopy(NextRuleID, m.System.SystemNumber)
                        NewRule.LibraryID = m.System.LibraryID
                        NewRule.SecurityLevel = m.System.SecurityLevel
                        NewRule.Password = m.System.Password
                        NewRule.CannotDelete = False
                        alRuleIDs(X) = NewRule.RuleID
                    End If
                End With
                
                AddRule NewRule
            Else
                Set NewRule = m.System.Rules.Item(CStr(alRuleIDs(X)))
            End If
            
            'Update Inputs collection with inputs values entered on Quick Stops form
            For lRow = vsInputs.FixedRows To vsInputs.Rows - 1
                If CLng(vsInputs.TextMatrix(lRow, IGCol(eIGCol_RuleID))) = NewRule.RuleID Then
                    vsInputs.TextMatrix(lRow, IGCol(eIGCol_InputValue)) = FormatNum(adStopValues(X))
                    Exit For
                End If
            Next lRow
        Next X
        
        ShowLinkInputs
    
        EnableToolbar True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.QuickStops", eGDRaiseError_Raise
    
End Sub

Private Sub LinkToEntry()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current row in the grid
    Dim pt As POINTAPI                  ' Current screen point
    Dim lRuleID As Long                 ' Current Rule ID
    Dim strLink As String               ' String returned from form
    
    If Not HasPlatinum(True) Then Exit Sub
    If Not ShowPassword Then Exit Sub
    
    With vsRules
        lRow = .RowSel
    
        ' Get popup window position
        pt.X = .ColPos(RGCol(eRGCol_RuleName)) / Screen.TwipsPerPixelX
        pt.Y = (.RowPos(lRow) + .RowHeight(lRow)) / Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        pt.X = pt.X * Screen.TwipsPerPixelX
        pt.Y = pt.Y * Screen.TwipsPerPixelY
    
        'Get the currently selected Exit rule ID
        lRuleID = CLng(.TextMatrix(lRow, RGCol(eRGCol_RuleID)))
    
        'Show form...
        If frmLinkedRules.ShowMe(pt.X, pt.Y, m.System.Rules, lRuleID, strLink) Then
            m.System.Rules.Item(CStr(lRuleID)).LinkedRules = strLink
            .TextMatrix(lRow, RGCol(eRGCol_LinkedRules)) = strLink
            'CheckedCell(vsRules, lRow, RGCol(eRGCol_Linked)) = (Len(strLink) > 0)
            'vsRules.ColHidden(RGCol(eRGCol_Linked)) = Not ShowLinked
            ShowLinkedExits
        End If
    End With
    
    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.LinkToEntry", eGDRaiseError_Raise
    
End Sub

Private Function DataOtherThanCN() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    DataOtherThanCN = False
    With vsMarkets
        For lIndex = .FixedRows To .Rows - 1
            If UCase(.TextMatrix(lIndex, MGCol(eMGCol_Format))) <> "CN" Then
                DataOtherThanCN = True
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.DataOtherThanCN", eGDRaiseError_Raise
    
End Function

Private Function AskToSaveRules() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim bCancel As Boolean

    bCancel = False
    For lIndex = Forms.Count - 1 To 0 Step -1
        If UCase(Forms(lIndex).Name) = "FRMRULE" Then
            If Forms(lIndex).SystemID = m.System.SystemNumber Then
                ' if editor is not visible, it just didn't get
                ' unloaded all the way so kill it now
                If Not Forms(lIndex).Visible Then
                    Unload Forms(lIndex)
                Else
                    Forms(lIndex).SetFocus
                    If Forms(lIndex).AskToSave Then
                        bCancel = True
                        Exit For
                    Else
                        Unload Forms(lIndex)
                    End If
                End If
            End If
        End If
    Next lIndex
    
    AskToSaveRules = bCancel

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.AskToSaveRules", eGDRaiseError_Raise
    
End Function

Private Function Action(ByVal bBuySell As Boolean, ByVal lRuleUse As Long) As String
On Error GoTo ErrSection:

    If bBuySell Then
        If lRuleUse = 0 Then
            Action = "Long Entry"
        Else
            Action = "Short Exit"
        End If
    Else
        If lRuleUse = 0 Then
            Action = "Short Entry"
        Else
            Action = "Long Exit"
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.Action", eGDRaiseError_Raise
    
End Function

Private Sub ShowLinkedExits()
On Error GoTo ErrSection:

    Dim lRow As Long, bHideCol As Boolean
    
    With vsRules
        bHideCol = True
        For lRow = .FixedRows To .Rows - 1
            If .TextMatrix(lRow, RGCol(eRGCol_RuleUse)) <> 1 Then
                .Cell(flexcpChecked, lRow, RGCol(eRGCol_Linked)) = flexNoCheckbox
            ElseIf Len(.TextMatrix(lRow, RGCol(eRGCol_LinkedRules))) > 0 Then
                .Cell(flexcpChecked, lRow, RGCol(eRGCol_Linked)) = flexChecked
                .Cell(flexcpPictureAlignment, lRow, RGCol(eRGCol_Linked)) = .ColAlignment(RGCol(eRGCol_Linked))
                bHideCol = False
            Else
                .Cell(flexcpChecked, lRow, RGCol(eRGCol_Linked)) = flexUnchecked
                .Cell(flexcpPictureAlignment, lRow, RGCol(eRGCol_Linked)) = .ColAlignment(RGCol(eRGCol_Linked))
            End If
        Next lRow
        '.ColHidden(RGCol(eRGCol_Linked)) = bHideCol
        .ColHidden(RGCol(eRGCol_Sequence)) = bHideCol
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.ShowLinkedExits", eGDRaiseError_Raise
End Sub

Private Sub FixSequence(Optional ByVal bUpdateDisplay As Boolean = True)
On Error GoTo ErrSection:

    Dim bChanged As Boolean, bSaveRedraw As Boolean
    Dim iRule As Long, iSeq As Long, nValue&, i&
    Dim aSeq As New cGdArray
    Dim Rule As cRule
    
    aSeq.Create eGDARRAY_Longs
    
    m.lNumEntries = 0
    For iRule = 1 To m.System.Rules.Count
        'iSeq = m.System.Rules.Item(iRule).Seq
        Set Rule = m.System.Rules.Item(iRule)
        ' Create value to sort by: USSSSNNNN
        '   where U = 0 for entry or 1 for exit
        '   and SSSS = current sequence number
        '   and NNNN = rule number in collection
        If Rule.Seq < 0 Then
            ' negative number is temporarily set in order to insert
            ' (e.g. -3 will get inserted between 2 and 3)
            nValue = Abs(Rule.Seq) * 2 - 1
        Else
            ' double the current sequence to allow for inserts
            nValue = Rule.Seq * 2
        End If
        ' if 0 or too high, put at the end
        If nValue <= 0 Or nValue > 9999 Then
            nValue = 9999
        End If
        ' make all exits go after all entries
        If Rule.RuleUse = 0 Then
            m.lNumEntries = m.lNumEntries + 1
        Else
            nValue = nValue + 10000
        End If
        ' append rule number to the value
        aSeq.Add nValue * 10000 + iRule
    Next
    
    ' sort the values
    aSeq.Sort eGdSort_Stable
    
    ' assign new sequence numbers to each rule
    For i = 1 To aSeq.Size
        iRule = aSeq(i - 1) Mod 10000
        If m.System.Rules.Item(iRule).Seq <> i Then
            m.System.Rules.Item(iRule).Seq = i
            bChanged = True
        End If
    Next
    
    ' update the display in the rules grid
    If bUpdateDisplay And bChanged Then
        With vsRules
            bSaveRedraw = .Redraw
            .Redraw = flexRDNone
            For i = .FixedRows To .Rows - 1
                iRule = Val(.TextMatrix(i, RGCol(eRGCol_RuleID)))
                If iRule <> 0 Then
                    .TextMatrix(i, RGCol(eRGCol_Sequence)) = m.System.Rules.Item(Str(iRule)).Seq
                End If
            Next
            .Redraw = bSaveRedraw
        End With
    End If
    
    Set Rule = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.FixSequence", eGDRaiseError_Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolGroups
'' Description: Build a pipe delimited string of symbol groups to display in
''              the dropdown list for Market1
'' Inputs:      None
'' Returns:     String of Symbol Groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolGroups() As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrGroups As New cGdArray      ' Array of group names
    
    astrGroups.Create eGDARRAY_Strings
    For lIndex = 1 To g.SymbolPool.SymbolGroups.Count
        If g.SymbolPool.SymbolGroups(lIndex).GroupType = eGROUP_Normal And g.SymbolPool.SymbolGroups(lIndex).IsActive = True Then
            astrGroups.Add g.SymbolPool.SymbolGroups(lIndex).Name & ";#" & Str(lIndex * 1000 + Asc("S"))
        End If
    Next lIndex
    
    If ScansEnabled Then
        For lIndex = 1 To g.SymbolPool.Filters.Count
            If g.SymbolPool.Filters(lIndex).IsActive Then
                astrGroups.Add g.SymbolPool.Filters(lIndex).Name & ";#" & Str(lIndex * 1000 + Asc("F"))
            End If
        Next lIndex
        
        For lIndex = 1 To g.SymbolPool.Criterias.Count
            If g.SymbolPool.Criterias(lIndex).IsBoolean And g.SymbolPool.Criterias(lIndex).IsActive Then
                astrGroups.Add g.SymbolPool.Criterias(lIndex).Name & ";#" & Str(lIndex * 1000 + Asc("C"))
            End If
        Next lIndex
    End If
    
    astrGroups.Sort eGdSort_IgnoreCase
    
    For lIndex = 0 To astrGroups.Size - 1
        astrGroups(lIndex) = Parse(astrGroups(lIndex), ";", 2) & ";" & Parse(astrGroups(lIndex), ";", 1)
    Next lIndex
        
    SymbolGroups = astrGroups.JoinFields("|")

ErrExit:
    Set astrGroups = Nothing
    Exit Function
    
ErrSection:
    Set astrGroups = Nothing
    RaiseError "frmSystemManager.SymbolGroups", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolsInGroup
'' Description: Build an array of the symbols in a symbol group.
'' Inputs:      ID of the Symbol Group
'' Returns:     Array of symbols in that group
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolsInGroup(ByVal strGroupID As String) As cGdArray
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Array of symbols to return
    Dim aIndex As New cGdArray          ' Index into the symbol pool
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Create the arrays...
    astrSymbols.Create eGDARRAY_Strings
    aIndex.Create eGDARRAY_Longs
    
    ' Get the index from the symbol pool for this symbol group...
    Select Case Right(strGroupID, 3)
        Case "GRP"
            Set aIndex = g.SymbolPool.ArrayTable.CreateIndex(g.SymbolPool.FieldNumForID("GRP:" & strGroupID))
        Case "FIL"
            Set aIndex = g.SymbolPool.ArrayTable.CreateIndex(g.SymbolPool.FieldNumForID("FIL:" & strGroupID))
        Case "SCN"
            Set aIndex = g.SymbolPool.ArrayTable.CreateIndex(g.SymbolPool.FieldNumForID("DSV:" & strGroupID))
    End Select
    
    ' Walk through the index placing each symbol in the array...
    For lIndex = 0 To aIndex.Size - 1
        astrSymbols(lIndex) = g.SymbolPool.Symbol(aIndex(lIndex))
    Next lIndex
    
    ' Return the array of symbols...
    Set SymbolsInGroup = astrSymbols

ErrExit:
    Set astrSymbols = Nothing
    Set aIndex = Nothing
    Exit Function
    
ErrSection:
    Set astrSymbols = Nothing
    Set aIndex = Nothing
    RaiseError "frmSystemManager.SymbolsInGroup", eGDRaiseError_Raise
    
End Function

Private Sub RunMultipleSymbols(aSymbols As cGdArray, ByVal bNextBarReport As Boolean, Optional ByVal lRuleID As Long = 0&)
On Error GoTo ErrSection:

    Dim astrParms As cGdArray           ' Parameter array
    Dim astrTrades As cGdArray          ' Trades array
    Dim astrFiles As New cGdArray       ' Array of next bar files
    Dim lIndex As Long                  ' Index into a for loop
    Dim rc As Long                      ' Return code from the optimizer form
    Dim strPeriod As String             ' Period to run
    Dim strFileName As String           ' Name of the trades file
    Dim dNextBarDate As Double          ' Date for the next bar report
    Dim bAssumeNoPosition As Boolean    ' Does the user want to assume no current position?
    Dim bIgnoreNextBarData As Boolean   ' Does the user want to ignore next bar data?
    Dim s$, strSymbol$
    
    ' Initialize the necessary arrays...
    Set astrParms = New cGdArray
    astrParms.Create eGDARRAY_Strings
    Set astrTrades = New cGdArray
    astrTrades.Create eGDARRAY_Strings
    
    ' Get the symbols in the symbol group...
    'Set aSymbols = SymbolsInGroup(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_GroupID)))
    If aSymbols.Size = 0 Then
        Err.Raise vbObjectError + 1000, , "No Symbols in the selected Symbol Group"
    End If

    If bNextBarReport Then
        dNextBarDate = GetNextBarDate(bAssumeNoPosition, bIgnoreNextBarData)
        If dNextBarDate = -99999# Then
            m.bStop = False
            StatusMsg ""
            GoTo ErrExit
        End If
    Else
        Set g.CurrentSystem = m.System
        frmOptimizer.Init aSymbols.Size, astrParms, eGDOptMode_MultipleRun
        frmOptimizer.TestRuleID = lRuleID
        frmOptimizer.TestRunMode = m.System.RunMode
    End If
    astrFiles.Create eGDARRAY_Strings
    
    ' Retain the periodicity of Market1...
    strPeriod = m.System.Markets("Market1").Periodicity
    
    ' Run each symbol in the symbol group...
    If bNextBarReport Then Screen.MousePointer = vbHourglass
    For lIndex = 0 To aSymbols.Size - 1
        strSymbol = Parse(aSymbols(lIndex), vbTab, 1)
        If UCase(strPeriod) = "FRACTZEN" Then
            s = Parse(aSymbols(lIndex), vbTab, 2)
            g.FractZen.SetSpeedInfo strSymbol, Val(Parse(s, "L", 1)), Val(Parse(s, "L", 2))
        End If
        m.System.LoadMarket1Bars strSymbol, strPeriod, bNextBarReport, (aSymbols.Size <= kSN_BASKETLIMIT)
        
        If bNextBarReport = False Then
            'If IsIDE And UCase(strPeriod) = "FRACTZEN" Then
            If 0 Then
                m.System.TestFractZen False, astrTrades
            Else
                m.System.Test False, lRuleID, astrTrades, False
            End If
            strFileName = Replace(m.System.NextBarFile(eGDNextBarMode_RunMult), "\NB", "\S")
            
            astrParms(0) = vbTab & "<system>" & vbTab & m.System.SystemName & vbTab & "true"
            astrParms(1) = vbTab & "<symbol>" & vbTab & strSymbol & vbTab & "true"
            If UCase(strPeriod) = "FRACTZEN" Then
                s = g.FractZen.GetSpeedInfo(strSymbol)
                s = Parse(s, ",", 1) & "L" & Parse(s, ",", 2) ' e.g. "8L3"
                strFileName = Left(strFileName, Len(strFileName) - 4) & "-" & s & ".TXT"
            Else
                s = strPeriod
            End If
            astrParms(2) = vbTab & "<period>" & vbTab & s & vbTab & "true"
            astrParms(3) = vbTab & "<fromdate>" & vbTab & Str(CLng(DateOf(m.System.FromDate))) & vbTab & "true"
            astrParms(4) = vbTab & "<todate>" & vbTab & Str(CLng(DateOf(m.System.ToDate))) & vbTab & "true"
            astrParms(5) = vbTab & "<toend>" & vbTab & Str(CLng(m.System.ToEndOfData)) & vbTab & "true"
            astrParms(6) = vbTab & "<overrides>" & vbTab & vbTab & "true"
            astrParms(7) = vbTab & "<filenames>" & vbTab & strFileName & vbTab & "true"
            
            astrTrades.ToFile strFileName
            astrFiles.Add strFileName
            
            rc = frmOptimizer.Add(lIndex + 1, astrTrades, astrParms)
            If rc <> kSN_OPTIMIZATION_IN_PROGRESS Then Exit For
        Else
            StatusMsg Replace(m.System.SystemName, "&", "&&") & " (" & strSymbol & ")"
            m.System.NextBarReport eGDNextBarMode_RunMult, dNextBarDate, bAssumeNoPosition, bIgnoreNextBarData
            astrFiles.Add m.System.NextBarFile(eGDNextBarMode_RunMult)
            DoEvents
            If m.bStop Then Exit For
        End If
    Next lIndex
    
    If bNextBarReport = True Then
        m.bStop = False
        StatusMsg ""
        frmNextBar.ShowMeMult astrFiles, , m.strName
        
        'For lIndex = 0 To astrFiles.Size - 1
        '    KillFile astrFiles(lIndex)
        '    KillFile Replace(astrFiles(lIndex), "\NB", "\RB")
        'Next lIndex
    Else
        If Len(m.System.SystemName) > 0 Then
            frmOptimizer.SetUpMergedRun m.System.SystemName, m.System.Pyramid, astrFiles.ArrayHandle, m.System.RptRulesHandle
        Else
            frmOptimizer.SetUpMergedRun "New Strategy", m.System.Pyramid, astrFiles.ArrayHandle, m.System.RptRulesHandle
        End If
    End If

ErrExit:
    Set astrParms = Nothing
    Set astrFiles = Nothing
    Set astrTrades = Nothing
    Exit Sub
    
ErrSection:
    'StatusMsg ""
    'If FormIsLoaded("frmOptimizer") Then frmOptimizer.StopRun
    Set astrParms = Nothing
    Set astrFiles = Nothing
    Set astrTrades = Nothing
    RaiseError "frmSystemManager.RunMultipleSymbols", eGDRaiseError_Raise
    
End Sub

' TLB: as of 7/8/2014 this routine should be OBSOLETE (since the "Merged Reports" button is now always hidden)
#If 0 Then
Private Sub RunMergedReports(Optional ByVal lRuleID As Long = 0&)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Array of symbols to run strategy on
    Dim astrFiles As New cGdArray       ' Array of next bar files
    Dim astrTrades As cGdArray          ' Trades array
    Dim lIndex As Long                  ' Index into a for loop
    Dim rc As Long                      ' Return code from the optimizer form
    Dim strPeriod As String             ' Period to run
    Dim strFileName As String           ' Name of the trades file
        
    ' Get the symbols in the symbol group...
    Set astrSymbols = SymbolsInGroup(vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_GroupID)))
    
    Set astrFiles = New cGdArray
    astrFiles.Create eGDARRAY_Strings
    Set astrTrades = New cGdArray
    astrTrades.Create eGDARRAY_Strings
    
    ' Retain the periodicity of Market1...
    strPeriod = m.System.Markets("Market1").Periodicity
    
    ' Run each symbol in the symbol group...
    Screen.MousePointer = vbHourglass
    For lIndex = 0 To astrSymbols.Size - 1
        m.System.LoadMarket1Bars astrSymbols(lIndex), strPeriod, False
        
        m.System.Test False, lRuleID, astrTrades, False
        strFileName = Replace(m.System.NextBarFile(eGDNextBarMode_RunMult), "\NB", "\S")
        astrTrades.ToFile strFileName
            
        StatusMsg "Testing " & Replace(m.System.SystemName, "&", "&&") & " on " & astrSymbols(lIndex)
        astrFiles.Add strFileName
        DoEvents
        If m.bStop Then Exit For
    Next lIndex
    
    tbToolbar.ToolBars("General").Tools("ID_RunGroup").State = ssUnchecked
    m.System.ShowMergedReports astrFiles

ErrExit:
    StatusMsg ""
    Screen.MousePointer = vbDefault
    Set astrSymbols = Nothing
    Set astrFiles = Nothing
    Set astrTrades = Nothing
    Exit Sub
    
ErrSection:
    StatusMsg ""
    Screen.MousePointer = vbDefault
    Set astrSymbols = Nothing
    Set astrFiles = Nothing
    Set astrTrades = Nothing
    RaiseError "frmSystemManager.RunMergedReports", eGDRaiseError_Raise
    
End Sub
#End If

Public Property Let CondBuilderExpr(ByVal strExpr As String)
On Error Resume Next
    
    m.strCondBuilderExpr = strExpr

End Property

Private Sub SetWizardText()
On Error GoTo ErrSection:

    Dim strInfo$, i&
    
    'always make sure these buttons are off in wizard mode
    cmdTestEntry.Enabled = False
    cmdLinkToEntry.Enabled = False
    chkLinkToChart.Enabled = False
    
    With tbToolbar
        If m.nWizardStep = 10 Then
            If .Tools("ID_Next").Visible Then .Tools("ID_Next").Visible = False
            If Not .Tools("ID_RunA").Visible Then .Tools("ID_RunA").Visible = True
        Else
            If Not .Tools("ID_Next").Visible Then .Tools("ID_Next").Visible = True
            If .Tools("ID_RunA").Visible Then .Tools("ID_RunA").Visible = False
        End If
    End With

    Select Case m.nWizardStep
        Case 0:
            strInfo = "Use 'New Rule' or 'Add Rule' button to add LONG entries."
        Case 1:
            strInfo = "Use 'New Rule' or 'Add Rule' button to add SHORT entries."
        Case 2:
            strInfo = "Use 'Quick Stops' button to add protective stops and profit targets."
        Case 3:
            strInfo = "Use 'New Rule' or 'Add Rule' button to add LONG exits."
        Case 4:
            strInfo = "Use 'New Rule' or 'Add Rule' button to add SHORT exits."
        Case 5:
            strInfo = "Press 'Next' when done adding/editing rules."
        Case 6:
            If vsInputs.Rows > vsInputs.FixedRows Then
                strInfo = "Change input values as desired."
            Else
                strInfo = "Select symbol/group, bar period and date range."
                m.nWizardStep = 7
            End If
        Case 7:
            strInfo = "Select symbol/group, bar period and date range."
        Case 8:
            'security types: F=futures, I=index, S=stocks, GROUP=group
            strInfo = vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_SecType))
            If strInfo = "F" Then
                strInfo = "Set commission/slippage fees."
            Else
                strInfo = "Set commission/slippage fees and number of shares."
            End If
        Case 9:
            strInfo = "Can adjust other settings (default values recommended)."
        Case 10:
            strInfo = "Your strategy is complete. Hit run to test it!"
    End Select
    
    If strInfo <> tbToolbar.Tools("ID_WizardMsg").Name Then tbToolbar.Tools("ID_WizardMsg").ChangeAll ssChangeAllName, strInfo
    SetAsisstControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.SetWizardText", eGDRaiseError_Raise

End Sub

Private Sub SetWizardBack()
On Error GoTo ErrSection:
   
    If m.nWizardStep = m.nWizardStart Then
        tbToolbar.Tools("ID_Back").Enabled = False
    Else
        tbToolbar.Tools("ID_Back").Enabled = True
    End If
    
    SetWizardText
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.SetWizardBack", eGDRaiseError_Raise

End Sub

Private Sub SetAsisstControls(Optional ByVal bReset As Boolean = False)
On Error GoTo ErrSection:

    Static nHighestStep As Long
    
    If bReset Then
        nHighestStep = 0
        WizardDataTab True
        IncWizardStep True
        Exit Sub
    End If
    
    If nHighestStep < m.nWizardStep Then nHighestStep = m.nWizardStep
    
    WizardTabsIndex
    
    Select Case m.nWizardStep
        Case 0, 1, 2, 3, 4, 5:
            WizardRuleTab nHighestStep
        Case 7:
            WizardDataTab
        Case 8, 9:
            WizardSettingsTab
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.SetWizardControls", eGDRaiseError_Raise

End Sub

Private Function HasEntryRule() As Boolean
On Error GoTo ErrSection:

    Dim i&, strText$
    Dim bRule As Boolean
    
    For i = vsRules.FixedRows To vsRules.Rows - 1
        strText = vsRules.TextMatrix(i, RGCol(eRGCol_Action))
        If InStr(strText, "Entry") Then
            bRule = True
            Exit For
        End If
    Next
    
    HasEntryRule = bRule

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.HasEntryRule", eGDRaiseError_Raise

End Function

Private Function HasExitRule(nLongEntries&, nShortEntries&, _
    Optional ByVal bMsg As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, strText$
    Dim nLE&, nSE&          'count of long & short entry rules
    Dim nLX&, nSX&          'count of long & short exit rules
    Dim bOK As Boolean
    
    For i = vsRules.FixedRows To vsRules.Rows - 1
        strText = vsRules.TextMatrix(i, RGCol(eRGCol_Action))
        If InStr(strText, "Long Entry") Then nLE = nLE + 1
        If InStr(strText, "Short Entry") Then nSE = nSE + 1
        If InStr(strText, "Long Exit") Then nLX = nLX + 1
        If InStr(strText, "Short Exit") Then nSX = nSX + 1
        
        If nLE > 0 And (nSE > 0 Or nLX > 0) Then
            bOK = True
        ElseIf nSE > 0 And nSX > 0 Then
            bOK = True
        End If
        
        If bOK Then Exit For
    Next
    
    nLongEntries = nLE
    nShortEntries = nSE
    
    If bMsg And Not bOK Then
        If nLE > 0 And nLX = 0 And nSE = 0 Then
            InfBox "You have long entries. You need to add a short entry or a long exit."
        ElseIf nSE > 0 And nSX = 0 Then
            InfBox "You have short entries. You need to add a long entry or a short exit."
        ElseIf nSE = 0 And nLE = 0 Then
            InfBox "You do not have any entry rules."
        End If
    End If
    
    HasExitRule = bOK

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.HasExitRule", eGDRaiseError_Raise

End Function

Private Sub WizardRuleTab(ByVal nHighestStep&)
On Error GoTo ErrSection:

   If nHighestStep >= 4 Then
        lblQuickStops.Enabled = True
        cmdQuickStops.Enabled = True
        optSignals(0).Enabled = True   'long entry
        optSignals(1).Enabled = True   'long exit
        optSignals(2).Enabled = True   'short entry
        optSignals(3).Enabled = True   'short exit
        optSignals(4).Enabled = True   'all rules
    Else
        optSignals(0).Enabled = False   'long entry
        optSignals(1).Enabled = False   'long exit
        optSignals(2).Enabled = False   'short entry
        optSignals(3).Enabled = False   'short exit
        optSignals(4).Enabled = False   'all rules
        
        If nHighestStep < 2 Then
            lblQuickStops.Enabled = False
            cmdQuickStops.Enabled = False
        End If
    End If

    'set rule radio button to match current step
    Select Case m.nWizardStep
        Case 0:
            If nHighestStep < 4 Then optSignals(0).Enabled = True
            optSignals(0).Value = True
        Case 1:
            If nHighestStep < 4 Then optSignals(2).Enabled = True
            optSignals(2).Value = True
        Case 2:
            If nHighestStep = 2 Then
                lblQuickStops.Enabled = True
                cmdQuickStops.Enabled = True
            End If
        Case 3:
            If nHighestStep < 4 Then optSignals(1).Enabled = True
            optSignals(1).Value = True
        Case 4:
            optSignals(3).Value = True
        Case Else:
            optSignals(4).Value = True
    End Select
    
    'set button text color to match current step
    If m.nWizardStep = 2 Then
        If lblNewRule.ForeColor <> m.nButtonTextColor Then lblNewRule.ForeColor = m.nButtonTextColor
        If lblAddRule.ForeColor <> m.nButtonTextColor Then lblAddRule.ForeColor = m.nButtonTextColor
        If lblQuickStops.ForeColor <> vbBlue Then lblQuickStops.ForeColor = vbBlue
    Else
        If lblNewRule.ForeColor <> vbBlue Then lblNewRule.ForeColor = vbBlue
        If lblAddRule.ForeColor <> vbBlue Then lblAddRule.ForeColor = vbBlue
        If lblQuickStops.ForeColor <> m.nButtonTextColor Then lblQuickStops.ForeColor = m.nButtonTextColor
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.WizardRuleTab", eGDRaiseError_Raise
End Sub

Private Sub WizardDataTab(Optional ByVal bReset As Boolean = False)
On Error GoTo ErrSection:

    Static bWizardSymbol As Boolean
    Static bWizardBar As Boolean
    Static bWizardCalendar As Boolean
    
    Dim strInfRtn$
        
    If bReset Then
        bWizardSymbol = False
        bWizardBar = False
        bWizardCalendar = False
        Exit Sub
    End If
        
    If cmdSettings.Enabled Then cmdSettings.Enabled = False
    If bWizardSymbol And bWizardBar Then bWizardCalendar = True
    
    fraToDate.Enabled = bWizardCalendar
    lblFromDate.Enabled = bWizardCalendar
    lblToDate.Enabled = bWizardCalendar
    dtpFromDate.Enabled = bWizardCalendar
    dtpToDate.Enabled = bWizardCalendar
    optToDate.Enabled = bWizardCalendar
    optToEndOfData.Enabled = bWizardCalendar
    
    'set button text color for symbol lookup
    lblBrowse.ForeColor = vbBlue
    
    If bWizardCalendar = True Then Exit Sub
                
    If Not bWizardSymbol Then
        bWizardSymbol = True
        strInfRtn = InfBox("icon=? ; buttons=+Symbol|Group|Use Existing ; msg=Would you like to apply this strategy to one symbol or a group of symbols?")
        If strInfRtn = "S" Then
            cmdBrowse_Click
        ElseIf strInfRtn = "G" Then
            With vsMarkets
                .ShowCell .FixedRows + 1, eMGCol_Symbol
                .Col = eMGCol_Symbol
                .EditCell
                SendKeys "{F4}"
            End With
            Exit Sub
        End If
    End If
        
    If bWizardSymbol And Not bWizardBar Then
        bWizardBar = True
        With vsMarkets
            .ShowCell .FixedRows + 1, eMGCol_Period
            .Col = eMGCol_Period
            .EditCell
            SendKeys "{F4}"
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.WizardDataTab", eGDRaiseError_Raise

End Sub

Private Sub WizardSettingsTab()
On Error GoTo ErrSection:
    
    Dim strSecType$
    Dim bEnable As Boolean
    Dim ctlEdit As ctlUniTextBoxXP 'TextBox 'RH changed from Textbox
       
    If m.nWizardStep >= 9 Then
        bEnable = True
    Else
        bEnable = False
    End If
    
    chkAllowReverse.Enabled = bEnable
    chkForceLimitThrough.Enabled = bEnable
    txtBarsLoadedBeforeTrading.Enabled = bEnable
    txtBarsTradedBeforeOrders.Enabled = bEnable
    chkPyramid.Enabled = bEnable
    txtTradeDepth.Enabled = bEnable
            
    'security types: F=futures, I=index, S=stocks, GROUP=group
    strSecType = vsMarkets.TextMatrix(vsMarkets.FixedRows, MGCol(eMGCol_SecType))
    
    'set focus to correct commission/fees control
    Select Case strSecType
        Case "S"
            If txtCommission.Enabled Then txtCommission.Enabled = False
            If Not txtStockCommission.Enabled Then txtStockCommission.Enabled = True
            If Not txtNumShares.Enabled Then txtNumShares.Enabled = True
            If Not txtDollarsPerTrade.Enabled Then txtDollarsPerTrade.Enabled = True
            Set ctlEdit = txtStockCommission
        Case "F"
            If Not txtCommission.Enabled Then txtCommission.Enabled = True
            If txtStockCommission.Enabled Then txtStockCommission.Enabled = False
            If txtNumShares.Enabled Then txtNumShares.Enabled = False
            If txtDollarsPerTrade.Enabled Then txtDollarsPerTrade.Enabled = False
            Set ctlEdit = txtCommission
        Case Else
            If Not txtCommission.Enabled Then txtCommission.Enabled = True
            If Not txtStockCommission.Enabled Then txtStockCommission.Enabled = True
            If Not txtNumShares.Enabled Then txtNumShares.Enabled = True
            If Not txtDollarsPerTrade.Enabled Then txtDollarsPerTrade.Enabled = True
            Set ctlEdit = txtCommission
    End Select
    
    If Not ctlEdit Is Nothing Then
        ctlEdit.SelLength = Len(ctlEdit.Text)
        ctlEdit.SetFocus
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.WizardSettingsTab", eGDRaiseError_Raise

End Sub

Private Sub WizardTabsIndex()
On Error GoTo ErrSection:

    If m.nWizardStep < 6 Then
        vsIndexTab1.TabEnabled(0) = True    'rule
        vsIndexTab1.TabEnabled(1) = False   'inputs
        vsIndexTab1.TabEnabled(2) = False   'data
        vsIndexTab1.TabEnabled(3) = False   'settings
        vsIndexTab1.CurrTab = 0
    ElseIf m.nWizardStep = 6 And vsInputs.Rows > vsInputs.FixedRows Then
        vsIndexTab1.TabEnabled(0) = False
        vsIndexTab1.TabEnabled(1) = True    'inputs
        vsIndexTab1.TabEnabled(2) = False
        vsIndexTab1.TabEnabled(3) = False
        vsIndexTab1.CurrTab = 1
        vsInputs.Row = 1
        vsInputs.Col = 2
        vsInputs.EditCell
    ElseIf m.nWizardStep = 7 Then
        vsIndexTab1.TabEnabled(0) = False
        vsIndexTab1.TabEnabled(1) = False
        vsIndexTab1.TabEnabled(2) = True    'data
        vsIndexTab1.TabEnabled(3) = False
        vsIndexTab1.CurrTab = 2
    ElseIf m.nWizardStep >= 8 Then
        vsIndexTab1.TabEnabled(0) = False
        vsIndexTab1.TabEnabled(1) = False
        vsIndexTab1.TabEnabled(2) = False
        vsIndexTab1.TabEnabled(3) = True    'settings
        vsIndexTab1.CurrTab = 3
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.WizardTabsIndex", eGDRaiseError_Raise

End Sub

Private Function SetWizardStart() As Boolean
On Error GoTo ErrSection:

    Dim strInfRtn$, lSystemID&
    
    SetWizardStart = True
    
    If m.bNewStrategy Then
        m.nWizardStart = 0
        Exit Function
    End If

    'check for strategies that cannot be copied or edited, i.e run-only
    If g.Security.CanPreview(m.System.SecurityLevel) Then
        strInfRtn = InfBox("icon=? ; buttons=+Yes|No ; msg=Do you want to make changes to the rules used in the strategy?")
    Else
       strInfRtn = InfBox("icon=? ; buttons=+Run|Cancel ; msg=You cannot edit or make a copy of this strategy. Would you like to run it?")
       If strInfRtn = "C" Then
            SetWizardStart = False
            Exit Function
       End If
    End If
        
    'start with step 6 to give user opportunity to modify inputs
    If strInfRtn = "R" Or strInfRtn = "N" Then
        m.nWizardStart = 6
        Exit Function
    End If

    'check security for editing
    If g.Security.CanSave(m.System.SecurityLevel, m.System.Password) Then
        m.nWizardStart = 0
    Else
        strInfRtn = InfBox("icon=? ; buttons=+Make Copy|Cancel ; msg=You do not have permission to edit the strategy. Would you like to make a copy and edit the copy?")
        If strInfRtn = "M" Then
            If HasPlatinum(True) Then
                If Save(tbToolbar.Tools("ID_SaveAs").ID) Then
                    lSystemID = m.System.SystemNumber
                    Set m.System = Nothing
                    LoadRec lSystemID
                    m.nWizardStart = 0
                Else
                    SetWizardStart = False
                End If
            Else
                SetWizardStart = False
            End If
        Else
            SetWizardStart = False
        End If
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.SetWizardStart", eGDRaiseError_Raise

End Function

Private Sub RestoreControls()
On Error GoTo ErrSection:

    'Rules tab
    optSignals(0).Enabled = True
    optSignals(1).Enabled = True
    optSignals(2).Enabled = True
    optSignals(3).Enabled = True
    optSignals(4).Enabled = True
    optSignals(4).Value = True          'set 'All' button on
    cmdQuickStops.Enabled = True
    lblQuickStops.Enabled = True
    lblNewRule.ForeColor = m.nButtonTextColor   'restore button text colors
    lblAddRule.ForeColor = m.nButtonTextColor
    lblQuickStops.ForeColor = m.nButtonTextColor
    
    'Data tab
    fraToDate.Enabled = True            'date controls
    lblFromDate.Enabled = True
    lblToDate.Enabled = True
    dtpFromDate.Enabled = True
    dtpToDate.Enabled = True
    optToDate.Enabled = True
    optToEndOfData.Enabled = True
    cmdSettings.Enabled = True          'market button
    cmdBrowse.Enabled = True            'symbol lookup button
    lblBrowse.ForeColor = m.nButtonTextColor
    chkLinkToChart.Enabled = True
    
    'Settings tab
    txtStockCommission.Enabled = True
    txtNumShares.Enabled = True
    txtDollarsPerTrade.Enabled = True
    chkAllowReverse.Enabled = True
    chkForceLimitThrough.Enabled = True
    txtBarsLoadedBeforeTrading.Enabled = True
    txtBarsTradedBeforeOrders.Enabled = True
    chkPyramid.Enabled = True
    txtTradeDepth.Enabled = True
    
    'Tabs index
    vsIndexTab1.TabEnabled(1) = True    'inputs
    vsIndexTab1.TabEnabled(2) = True
    vsIndexTab1.TabEnabled(3) = True
    
    If g.Security.CanPreview(m.System.SecurityLevel) Then
        vsIndexTab1.TabEnabled(0) = True
        vsIndexTab1.CurrTab = 0
    Else
        vsIndexTab1.TabEnabled(0) = False
        vsIndexTab1.CurrTab = 1
    End If
    
    SetAsisstControls True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.RestoreControls", eGDRaiseError_Raise

End Sub

Private Function IncWizardStep(Optional ByVal bReset As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Static bWizardQuickStop As Boolean
    
    Dim nNextStep&, nLE&, nSE&
    Dim bOkay As Boolean
    Dim bRC As Boolean

    If bReset Then
        bWizardQuickStop = False
        Exit Function
    End If
    
    'verify if okay to go to next step
    Select Case m.nWizardStep
        Case 0:
            bOkay = True
        Case 1:
            'next step is quick stops, must have entry rules
            If HasEntryRule Then
                bOkay = True
            Else
                InfBox "You must add an entry rule before going to the next step."
            End If
        Case 2, 3:
            'step=2 --> next step=3=long exits, must have entry rules
            'step=3 --> next step=4=short exits, must have entry rules
            If HasEntryRule Then bOkay = True
        Case 4, 5, 6:
            'step=4 --> next step=5=done with rules, must have valid rules
            'step=5 --> next step=6=inputs, must have valid rules
            'step=6 --> next step=7=data, must have valid rules
            bOkay = HasExitRule(nLE, nSE, True)
        Case 7:
            'step=7 --> next step=8=settings, must have valid rules & valid data
            bOkay = HasExitRule(nLE, nSE, True)
            If bOkay Then WizardMarketData
        Case 8, 9:
            'settings tabs, no need to check since all other tabs are disabled at this point
            bOkay = True
    End Select
    
    If Not bOkay Then Exit Function
   
    'determine what next step should be
    Select Case m.nWizardStep
        Case 0:
            nNextStep = 1   'add short entries
        Case 1:
            nNextStep = 2   'quickstop
            'bring up quickstop form if first time
            If Not bWizardQuickStop Then
                bWizardQuickStop = True
                QuickStops
            End If
        Case 2:
            'go to long exits if strategy has long entries otherwise go to short exits
            HasExitRule nLE, nSE
            If nLE = 0 Then
                nNextStep = 4
            Else
                nNextStep = 3
            End If
        Case 3:
            'go to short exits if strategy has short entries otherwise go to done with rules
            HasExitRule nLE, nSE
            If nSE = 0 Then
                nNextStep = 5
            Else
                nNextStep = 4
            End If
        Case 4, 5, 6, 7, 8, 9:
            nNextStep = m.nWizardStep + 1
    End Select
    

    m.nWizardStep = nNextStep
    
    IncWizardStep = bOkay
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.IncWizardStep", eGDRaiseError_Raise

End Function

Private Function DecWizardStep() As Boolean
On Error GoTo ErrSection:

    Dim nPrevStep&, nLE&, nSE&
       
    Select Case m.nWizardStep
        Case 0:
            nPrevStep = 0
        Case 1, 2, 3:
            'step=1 --> prev step=0=long entries
            'step=2 --> prev step=1=short entries
            'step=3 --> prev step=2=quick stops
            nPrevStep = m.nWizardStep - 1
        Case 4:
            'step=4 --> prev step=3=long exits, if there are long entries
            'step=4 --> prev step=2=quick stops, if there are no long entries
            HasExitRule nLE, nSE
            If nLE > 0 Then
                nPrevStep = 3
            Else
                nPrevStep = 2
            End If
        Case 5:
            'step=5 --> prev step=4=short exits, if there are short entries
            'step=5 --> prev step=3=long exits, if there are long entries, but no short entries
            'step=5 --> prev step=0=add long entries, if there are no entries at all (user may have deleted entries)
            HasExitRule nLE, nSE
            If nSE > 0 Then
                nPrevStep = 4
            ElseIf nLE > 0 Then
                nPrevStep = 3
            Else
                nPrevStep = 0
            End If
        Case 6:
            'step=6 --> prev step=5=finishing add/edit rules
            nPrevStep = 5
        Case 7:
            'step=7 --> prev step=6=inputs, if there are inputs
            'step=7 --> prev step=5=finishing add/edit rules, if there are no inputs
            If vsInputs.Rows > vsInputs.FixedRows Then
                nPrevStep = 6
            Else
                nPrevStep = 5
            End If
        Case 8, 9, 10:
            'step=8  --> prev step=7=data tab
            'step=9  --> prev step=8=settings tab
            'step=10 --> prev step=9=settings advanced settings tab
            nPrevStep = m.nWizardStep - 1
        Case Else
            nPrevStep = m.nWizardStep - 1
    End Select
        
    m.nWizardStep = nPrevStep
    DecWizardStep = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.DecWizardStep", eGDRaiseError_Raise

End Function

Private Sub ExitWizard()
On Error GoTo ErrSection:

    RestoreControls
    tbToolbar.Redraw = False
    tbToolbar.ToolBars("General").Visible = True
    tbToolbar.ToolBars("Wizard").Visible = False
    tbToolbar.Redraw = True
    m.nWizardStep = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.ExitWizard", eGDRaiseError_Raise

End Sub

Private Sub WizardMarketData()
On Error GoTo ErrSection:

    Dim i&
        
    With vsMarkets
        'make sure all rows have valid bar period and security name
        For i = .FixedRows + 1 To .Rows - 1
            If Not .RowHidden(i) Then
                If .TextMatrix(i, MGCol(eMGCol_Period)) = "" Then .TextMatrix(i, MGCol(eMGCol_Period)) = "(Default)"
                If .TextMatrix(i, MGCol(eMGCol_Symbol)) = "" Then
                    .Row = i
                    BrowseMarkets
                End If
            End If
        Next
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.WizardMarketData", eGDRaiseError_Raise

End Sub

Public Property Get StrategyWizard() As Boolean

    StrategyWizard = tbToolbar.ToolBars("Wizard").Visible

End Property

Public Sub UseChartSystem(ByVal nUseChart As Long)
On Error GoTo ErrSection:

    Dim frmActive As Form
    Dim Chart As cChart
    Dim aMarket1 As cGdArray
    Dim strBarPeriod$, i&

    'do not allow system switching in wizard mode
    If tbToolbar.ToolBars("Wizard").Visible Then Exit Sub

    If chkLinkToChart.Value <> nUseChart Then
        If nUseChart = 1 Then m.bLinkedToChart = True
        'setting the check box value will trigger the check box click event
        chkLinkToChart.Value = nUseChart
        Exit Sub
    End If
        
    If nUseChart <> 0 Then
        'get active chart object
        Set frmActive = ActiveChart
        If Not frmActive Is Nothing Then
            Set Chart = frmActive.Chart
            If Chart Is Nothing Then
                Set frmActive = Nothing
                m.bLinkedToChart = False
                Exit Sub
            End If
        End If
        
        m.bLinkedToChart = True
        m.dFromDate = dtpFromDate.Value
        m.dToDate = dtpToDate.Value
        m.bToEndOfData = optToEndOfData.Value * -1
        'vsIndexTab1.TabCaption(2) = "&Data (linked)"
        
        'set test date controls to match chart's date range
        dtpFromDate.Value = Chart.Bars(eBARS_DateTime, 0)
        dtpToDate.Value = Chart.Bars(eBARS_DateTime, Chart.Bars.Size - 1)
        optToEndOfData.Value = Chart.ToEndOfData
            
        'save original market1 data information to string
        m.strSaveMarket1 = ""
        With vsMarkets
            .Redraw = flexRDNone
            If .Rows > .FixedRows Then
                For i = eMGCol_ParmName To eMGCOl_NumCols - 1
                    m.strSaveMarket1 = m.strSaveMarket1 & .TextMatrix(.FixedRows, i) & vbTab
                Next
                If .TextMatrix(.FixedRows, MGCol(eMGCol_SecType)) = "Group" Then
                    m.strSaveMarket1 = m.strSaveMarket1 & Str(.ComboIndex)
                End If
                'set first row values with data information from chart
                .Row = .FixedRows
                .Enabled = False
                .TextMatrix(.FixedRows, MGCol(eMGCol_Period)) = Chart.Bars.Prop(eBARS_PeriodicityStr)
            End If
            BrowseMarkets Chart.Symbol
            .Redraw = flexRDBuffered
        End With
        vsLinked.Visible = True
    Else
        'Set test date controls to match saved date range
        dtpFromDate = m.dFromDate
        dtpToDate = m.dToDate
        optToEndOfData.Value = m.bToEndOfData
        optToDate = Not m.bToEndOfData
        
        'restore original market1 data to top row
        Set aMarket1 = New cGdArray
        aMarket1.SplitFields m.strSaveMarket1, vbTab
        With vsMarkets
            .Redraw = flexRDNone
            If .Rows > .FixedRows Then
                For i = eMGCol_ParmName To eMGCOl_NumCols - 1
                    .TextMatrix(.FixedRows, i) = aMarket1(i)
                Next
                If aMarket1.Size = eMGCOl_NumCols + 1 Then
                    .Row = .FixedRows
                    .Col = eMGCol_Symbol
                    .ComboIndex = Val(aMarket1(eMGCOl_NumCols))     'symbol groups
                    vsMarkets_AfterEdit .FixedRows, eMGCol_Symbol
                Else
                    .Row = .FixedRows
                    BrowseMarkets .TextMatrix(.FixedRows, eMGCol_Symbol)
                End If
            End If
            .Enabled = True
            .Redraw = flexRDBuffered
        End With
        
        Set aMarket1 = Nothing
        m.strSaveMarket1 = ""
        m.bLinkedToChart = False
        'vsIndexTab1.TabCaption(2) = "&Data"
        vsLinked.Visible = False
    End If

    fraToDate.Enabled = Not m.bLinkedToChart
    lblFromDate.Enabled = Not m.bLinkedToChart
    lblToDate.Enabled = Not m.bLinkedToChart
    dtpFromDate.Enabled = Not m.bLinkedToChart
    dtpToDate.Enabled = Not m.bLinkedToChart
    optToDate.Enabled = Not m.bLinkedToChart
    optToEndOfData.Enabled = Not m.bLinkedToChart
    cmdBrowse.Enabled = Not m.bLinkedToChart
    lblBrowse.Enabled = Not m.bLinkedToChart

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmSystemManager.UseChartSystem", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SyncMarkets
'' Description: Synchronize the hidden markets in the markets grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SyncMarkets()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strParmName As String           ' Parameter name of the current row
    Dim strParmName2 As String          ' Parameter name of the current row
    Dim strDone As String               ' Markets that have been done
    Dim strPeriod As String             ' Period to put in the grid
    
    strDone = vbTab
    With vsMarkets
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                strParmName = .TextMatrix(lIndex, MGCol(eMGCol_ParmName))
                If InStr(strDone, vbTab & strParmName & vbTab) = 0 Then
                    For lIndex2 = lIndex + 1 To .Rows - 1
                        strParmName2 = .TextMatrix(lIndex2, MGCol(eMGCol_ParmName))
                        
                        If strParmName2 = strParmName Then
                            .TextMatrix(lIndex2, MGCol(eMGCol_Security)) = .TextMatrix(lIndex, MGCol(eMGCol_Security))
                            .TextMatrix(lIndex2, MGCol(eMGCol_SecType)) = .TextMatrix(lIndex, MGCol(eMGCol_SecType))
                            .TextMatrix(lIndex2, MGCol(eMGCol_SymbolPath)) = .TextMatrix(lIndex, MGCol(eMGCol_SymbolPath))
                            .TextMatrix(lIndex2, MGCol(eMGCol_Symbol)) = .TextMatrix(lIndex, MGCol(eMGCol_Symbol))
                            .TextMatrix(lIndex2, MGCol(eMGCol_Period)) = .TextMatrix(lIndex, MGCol(eMGCol_Period))
                            .TextMatrix(lIndex2, MGCol(eMGCol_Format)) = .TextMatrix(lIndex, MGCol(eMGCol_Format))
                            .TextMatrix(lIndex2, MGCol(eMGCol_MarketSymbol)) = .TextMatrix(lIndex, MGCol(eMGCol_MarketSymbol))
                            .TextMatrix(lIndex2, MGCol(eMGCol_GroupID)) = .TextMatrix(lIndex, MGCol(eMGCol_GroupID))
                            .TextMatrix(lIndex2, MGCol(eMGCol_SymbolID)) = .TextMatrix(lIndex, MGCol(eMGCol_SymbolID))
                            
                            If InStr(strDone, vbTab & strParmName2 & vbTab) = 0 Then
                                strDone = strDone & strParmName2 & vbTab
                            End If
                        ElseIf UCase(strParmName) = "MARKET1" And Market1Equivalent(strParmName2, strPeriod) Then
                            .TextMatrix(lIndex2, MGCol(eMGCol_Security)) = .TextMatrix(lIndex, MGCol(eMGCol_Security))
                            .TextMatrix(lIndex2, MGCol(eMGCol_SecType)) = .TextMatrix(lIndex, MGCol(eMGCol_SecType))
                            .TextMatrix(lIndex2, MGCol(eMGCol_SymbolPath)) = .TextMatrix(lIndex, MGCol(eMGCol_SymbolPath))
                            .TextMatrix(lIndex2, MGCol(eMGCol_Symbol)) = .TextMatrix(lIndex, MGCol(eMGCol_Symbol))
                            .TextMatrix(lIndex2, MGCol(eMGCol_Period)) = strPeriod
                            .TextMatrix(lIndex2, MGCol(eMGCol_Format)) = .TextMatrix(lIndex, MGCol(eMGCol_Format))
                            .TextMatrix(lIndex2, MGCol(eMGCol_MarketSymbol)) = .TextMatrix(lIndex, MGCol(eMGCol_MarketSymbol))
                            .TextMatrix(lIndex2, MGCol(eMGCol_GroupID)) = .TextMatrix(lIndex, MGCol(eMGCol_GroupID))
                            .TextMatrix(lIndex2, MGCol(eMGCol_SymbolID)) = .TextMatrix(lIndex, MGCol(eMGCol_SymbolID))
                        
                            If InStr(strDone, vbTab & strParmName2 & vbTab) = 0 Then
                                strDone = strDone & strParmName2 & vbTab
                            End If
                        End If
                    Next lIndex2
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSystemManager.SyncMarkets", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Market1Equivalent
'' Description: Should the given parameter be given Market1's information?
'' Inputs:      Parameter Name
'' Returns:     True if Equivalent to Market1, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Market1Equivalent(ByVal strParmName As String, Optional strPeriod As String) As Boolean
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol of the parameter

    Market1Equivalent = False
    Select Case UCase(strParmName)
        Case "MARKET1"
            strPeriod = ""
            Market1Equivalent = True
        
        Case "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY"
            strPeriod = strParmName
            Market1Equivalent = True
        
        Case "UNSPLIT"
            strPeriod = "(Default)"
            Market1Equivalent = True
        
        Case Else
            If Left(strParmName, 1) = Chr(34) And Right(strParmName, 1) = Chr(34) Then
                strSymbol = Parse(Replace(strParmName, Chr(34), ""), ",", 1)
                strPeriod = Parse(Replace(strParmName, Chr(34), ""), ",", 2)
                
                If Len(strSymbol) = 0 And Len(strPeriod) > 0 Then
                    Market1Equivalent = True
                End If
            End If
            
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.Market1Equivalent", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPeriodStr
'' Description: Wrapper of mGdDll version that accounts for (Default)
'' Inputs:      Periodicity
'' Returns:     Periodicity String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetPeriodStr(ByVal Periodicity As Variant) As String
On Error GoTo ErrSection:

    If VarType(Periodicity) = vbString And Left(Periodicity, 1) = "(" Then
        GetPeriodStr = "(Default)"
    Else
        GetPeriodStr = mGdDll.GetPeriodStr(Periodicity)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.GetPeriodStr", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNextBarDate
'' Description: Ask the user for the date to use for next bar reports
'' Inputs:      Assume No Position, Ignore Next Bar Data
'' Returns:     Next Bar Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNextBarDate(bAssumeNoPosition As Boolean, bIgnoreNextBarData As Boolean) As Double
On Error GoTo ErrSection:

    Dim dNewYorkTime As Double          ' Current date and time in New York
    Dim dNextBarDate As Double          ' Date (and time) of the next bar report
    Dim lMousePointer As Long           ' Current state of the mouse pointer

    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbDefault

    ' Come up with an educated guess as to the next bar date...
    dNewYorkTime = ConvertTimeZone(Now)
    If Hour(dNewYorkTime) < 14 Then
        dNextBarDate = Int(dNewYorkTime)
    Else
        dNextBarDate = Int(dNewYorkTime) + 1
    End If
    Do While Not IsWeekday(dNextBarDate)
        dNextBarDate = dNextBarDate + 1
    Loop
    
    ' Verify our educated guess with the user...
    If frmNextBarOpt.ShowMe(dNextBarDate, False, False, bAssumeNoPosition, bIgnoreNextBarData) Then
        GetNextBarDate = dNextBarDate
    Else
        GetNextBarDate = -99999#
    End If

ErrExit:
    Screen.MousePointer = lMousePointer
    Exit Function
    
ErrSection:
    Screen.MousePointer = lMousePointer
    RaiseError "frmSystemManager.GetNextBarDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPeriod
'' Description: Fix the bar period
'' Inputs:      Bar Period
'' Returns:     Fixed Bar Period
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixPeriod(ByVal strPeriod As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If UCase(strPeriod) = "AUTO BREAKOUT" Or UCase(strPeriod) = "FRACTZEN" Then
        strReturn = "FractZen" 'strPeriod
    Else
        strReturn = GetPeriodStr(strPeriod)
    End If
    
    FixPeriod = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSystemManager.FixPeriod"
    
End Function


