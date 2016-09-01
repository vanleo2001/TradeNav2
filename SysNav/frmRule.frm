VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmRule 
   Caption         =   "Rule Editor"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9465
   Visible         =   0   'False
   Begin vsOcx6LibCtl.vsElastic vsSystemName 
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   450
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
      Caption         =   "Rule: System name"
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   60
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   11
      DisplayContextMenu=   0   'False
      Tools           =   "frmRule.frx":0000
      ToolBars        =   "frmRule.frx":03E6
   End
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   5130
      Left            =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   100
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   9049
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
      Caption         =   "&Rule|&Rule|&Inputs|&Advanced"
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
      Begin HexUniControls.ctlUniFrameWL fraAdvanced 
         Height          =   4755
         Left            =   45
         TabIndex        =   5
         Top             =   330
         Width           =   9135
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
         Caption         =   "frmRule.frx":062C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmRule.frx":0658
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRule.frx":0678
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraPyramiding 
            Height          =   3735
            Left            =   300
            TabIndex        =   15
            Top             =   720
            Width           =   6495
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
            Caption         =   "frmRule.frx":0694
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmRule.frx":06C0
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":06E0
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraPyramidExit 
               Height          =   2055
               Left            =   0
               TabIndex        =   20
               Top             =   1560
               Width           =   5775
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
               Caption         =   "frmRule.frx":06FC
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmRule.frx":0736
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":0756
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optPosition 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   22
                  Top             =   660
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
                  Caption         =   "frmRule.frx":0772
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmRule.frx":07A2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":07C2
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniFrameWL fraPosition 
                  Height          =   615
                  Left            =   1080
                  TabIndex        =   9
                  Top             =   900
                  Width           =   3375
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
                  Caption         =   "frmRule.frx":07DE
                  Enabled         =   -1  'True
                  ForeColor       =   -2147483642
                  BackColor       =   -2147483633
                  Tip             =   "frmRule.frx":080A
                  VistaStyle      =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":082A
                  RightToLeft     =   0   'False
                  Begin HexUniControls.ctlUniRadioXP optNumContracts 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   23
                     Top             =   0
                     Width           =   1935
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
                     Caption         =   "frmRule.frx":0846
                     Enabled         =   -1  'True
                     Align           =   0
                     RadioBackColor  =   -2147483643
                     RadioForeColor  =   -2147483640
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     Pressed         =   0   'False
                     Tip             =   "frmRule.frx":088C
                     Style           =   -1
                     MousePointer    =   0
                     MouseIcon       =   "frmRule.frx":08AC
                     ShowFocus       =   -1  'True
                     RightToLeft     =   0   'False
                  End
                  Begin HexUniControls.ctlUniTextBoxXP txtContractsToExit 
                     Height          =   285
                     Left            =   2160
                     TabIndex        =   24
                     Top             =   0
                     Width           =   855
                     _ExtentX        =   0
                     _ExtentY        =   0
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     Enabled         =   -1  'True
                     Locked          =   0   'False
                     Text            =   "frmRule.frx":08C8
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
                     Tip             =   "frmRule.frx":08E8
                     HideSelection   =   -1  'True
                     RightToLeft     =   0   'False
                     ManualStart     =   0   'False
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmRule.frx":0908
                  End
                  Begin HexUniControls.ctlUniRadioXP optPercent 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   25
                     Top             =   360
                     Width           =   1935
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
                     Caption         =   "frmRule.frx":0924
                     Enabled         =   -1  'True
                     Align           =   0
                     RadioBackColor  =   -2147483643
                     RadioForeColor  =   -2147483640
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     Pressed         =   0   'False
                     Tip             =   "frmRule.frx":096A
                     Style           =   -1
                     MousePointer    =   0
                     MouseIcon       =   "frmRule.frx":098A
                     ShowFocus       =   -1  'True
                     RightToLeft     =   0   'False
                  End
                  Begin HexUniControls.ctlUniTextBoxXP txtPercent 
                     Height          =   360
                     Left            =   2160
                     TabIndex        =   26
                     Top             =   360
                     Width           =   855
                     _ExtentX        =   0
                     _ExtentY        =   0
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     Enabled         =   -1  'True
                     Locked          =   0   'False
                     Text            =   "frmRule.frx":09A6
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
                     Tip             =   "frmRule.frx":09C6
                     HideSelection   =   -1  'True
                     RightToLeft     =   0   'False
                     ManualStart     =   0   'False
                     RoundedBorders  =   0   'False
                     MousePointer    =   0
                     MouseIcon       =   "frmRule.frx":09E6
                  End
               End
               Begin HexUniControls.ctlUniRadioXP optTrade 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   27
                  Top             =   1620
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
                  Caption         =   "frmRule.frx":0A02
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmRule.frx":0A2C
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":0A4C
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblBasedOn 
                  Height          =   255
                  Left            =   240
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
                  Caption         =   "frmRule.frx":0A68
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmRule.frx":0AA4
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":0AC4
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraPyramidEnter 
               Height          =   855
               Left            =   0
               TabIndex        =   17
               Top             =   480
               Width           =   5775
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
               Caption         =   "frmRule.frx":0AE0
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmRule.frx":0B22
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":0B42
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtContractsToEnter 
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   19
                  Top             =   345
                  Width           =   855
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmRule.frx":0B5E
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
                  Tip             =   "frmRule.frx":0B7E
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":0B9E
               End
               Begin HexUniControls.ctlUniLabelXP lblContractsToEnter 
                  Height          =   255
                  Left            =   240
                  Top             =   360
                  Width           =   2175
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
                  Caption         =   "frmRule.frx":0BBA
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmRule.frx":0C14
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":0C34
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniLabelXP Label7 
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   6135
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
               Caption         =   "frmRule.frx":0C50
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmRule.frx":0CEC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":0D0C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniComboImageXP cboCategory 
            Height          =   315
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   2235
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
            Tip             =   "frmRule.frx":0D28
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":0D48
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCategory 
            Height          =   195
            Left            =   360
            Top             =   300
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
            Caption         =   "frmRule.frx":0D64
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmRule.frx":0DA0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":0DC0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraInputs 
         Height          =   4755
         Left            =   -9780
         TabIndex        =   11
         Top             =   330
         Width           =   9135
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
         Caption         =   "frmRule.frx":0DDC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmRule.frx":0E08
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRule.frx":0E28
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid vsInputs 
            Height          =   4275
            Left            =   90
            TabIndex        =   12
            Top             =   390
            Width           =   8955
            _cx             =   15796
            _cy             =   7541
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
            Height          =   240
            Index           =   1
            Left            =   90
            Top             =   90
            Width           =   8955
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
            Caption         =   "frmRule.frx":0E44
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmRule.frx":0F52
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":0F72
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraAdvRule 
         Height          =   4755
         Left            =   -10080
         TabIndex        =   13
         Top             =   330
         Width           =   9135
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
         Caption         =   "frmRule.frx":0F8E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmRule.frx":0FBA
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRule.frx":0FDA
         RightToLeft     =   0   'False
         Begin NavTradeSenseV3.Editor edAdvanced 
            Height          =   4050
            Left            =   120
            TabIndex        =   10
            Top             =   540
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   7144
         End
         Begin HexUniControls.ctlUniLabelXP Label6 
            Height          =   240
            Left            =   165
            Top             =   195
            Width           =   8115
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
            Caption         =   "frmRule.frx":0FF6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmRule.frx":10E8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":1108
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraRule 
         Height          =   4755
         Left            =   -10380
         TabIndex        =   28
         Top             =   330
         Width           =   9135
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
         Caption         =   "frmRule.frx":1124
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmRule.frx":1150
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRule.frx":1170
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkAdvanced 
            Height          =   220
            Left            =   6180
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
            Width           =   1980
            _ExtentX        =   3493
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
            Caption         =   "frmRule.frx":118C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmRule.frx":11CC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":11EC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkGlobal 
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   3855
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
            Caption         =   "frmRule.frx":1208
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmRule.frx":1296
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":12B6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraOrder 
            Height          =   2145
            Left            =   150
            TabIndex        =   30
            Top             =   2385
            Width           =   7140
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
            Caption         =   "frmRule.frx":12D2
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmRule.frx":1328
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":1348
            RightToLeft     =   0   'False
            Begin NavTradeSenseV3.Editor Editor2 
               Height          =   615
               Left            =   3240
               TabIndex        =   7
               Top             =   450
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   1085
            End
            Begin HexUniControls.ctlUniFrameWL Frame1 
               Height          =   1005
               Left            =   240
               TabIndex        =   21
               Top             =   360
               Width           =   2595
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
               Caption         =   "frmRule.frx":1364
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmRule.frx":1390
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":13B0
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniCheckXP chkCanExitOnEntryBar 
                  Height          =   255
                  Left            =   600
                  TabIndex        =   4
                  Top             =   600
                  Width           =   1815
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
                  Caption         =   "frmRule.frx":13CC
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmRule.frx":1416
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":1436
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniComboImageXP cboAction 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   3
                  Top             =   240
                  Width           =   2175
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
                  Tip             =   "frmRule.frx":1452
                  Sorted          =   0   'False
                  HScroll         =   0   'False
                  RoundedBorders  =   -1  'True
                  IconDim         =   16
                  MousePointer    =   0
                  MouseIcon       =   "frmRule.frx":1472
                  DropDownOnTextClick=   -1  'True
                  DropDownWidth   =   -1
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniComboImageXP cboOrderPlacement 
               Height          =   315
               Left            =   285
               TabIndex        =   6
               Top             =   1725
               Width           =   2595
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
               Tip             =   "frmRule.frx":148E
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":14AE
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin NavTradeSenseV3.Editor edLimit 
               Height          =   630
               Left            =   3240
               TabIndex        =   8
               Top             =   1380
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   1111
            End
            Begin HexUniControls.ctlUniLabelXP lblStop 
               Height          =   240
               Left            =   3255
               Top             =   240
               Width           =   5145
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
               Caption         =   "frmRule.frx":14CA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmRule.frx":1500
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":1520
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label5 
               Height          =   195
               Left            =   240
               Top             =   1485
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
               Caption         =   "frmRule.frx":153C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmRule.frx":1572
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":1592
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblActionDesc 
               Height          =   1140
               Left            =   225
               Top             =   855
               Width           =   2775
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
               Caption         =   "frmRule.frx":15AE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmRule.frx":15CE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":15EE
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblLimit 
               Height          =   255
               Left            =   3240
               Top             =   1140
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
               Caption         =   "frmRule.frx":160A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmRule.frx":1640
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":1660
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraCondition 
            Height          =   1695
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   8865
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
            Caption         =   "frmRule.frx":167C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmRule.frx":16AE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmRule.frx":16CE
            RightToLeft     =   0   'False
            Begin NavTradeSenseV3.Editor Editor1 
               Height          =   1095
               Left            =   180
               TabIndex        =   2
               Top             =   480
               Width           =   8565
               _ExtentX        =   15108
               _ExtentY        =   1931
            End
            Begin HexUniControls.ctlUniLabelXP lblCondition 
               Height          =   240
               Left            =   180
               Top             =   240
               Width           =   7665
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
               Caption         =   "frmRule.frx":16EA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmRule.frx":17DE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmRule.frx":17FE
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin VB.Image picSmall 
            Height          =   2040
            Left            =   3480
            Picture         =   "frmRule.frx":181A
            Top             =   2460
            Width           =   5475
         End
         Begin VB.Image picLarge 
            Height          =   2535
            Left            =   3480
            Picture         =   "frmRule.frx":B5EA
            Top             =   2460
            Visible         =   0   'False
            Width           =   6870
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
Attribute VB_Name = "frmRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmRule.frm
'' Description: Form for the management of rules
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Date         Author      Description
'' 04/17/2009   DAJ         Don't check price or price2 expressions unless need to
'' 10/30/2009   DAJ         Make sure we look for IF right in expression
'' 02/23/2010   DAJ         Make sure to do a ValOfText on Default Value
'' 05/17/2010   DAJ         Moved some functions to a module for global use
'' 06/16/2011   DAJ         Made changes for calling from Highlight Bar Reporter
'' 06/27/2011   DAJ         Added the VerifyHighlightBarReport routine
'' 09/15/2011   DAJ         Allow Daily,Weekly,Monthly for HighlightBar Report
'' 10/03/2011   DAJ         In BuildRule, fixed the insertion of IF
'' 12/17/2015   DAJ         When adding a rule back to the system manager in save, add a copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    lCurrentRuleID As Long
    lReturnValue As Long
    lSystemNumber As Long
    lSystemLibID As Long
    strSystemName As String
    bReloadMenu As Boolean
    bNewRule As Boolean
    bShared As Boolean
    bResavingAll As Boolean
    bSkipAutoIf As Boolean
    strCalledFrom As String
    frmSysMgr As Form
    
    SaveGrid As cInputs
    ListLoading As cListLoading
    Inputs As cInputs
    Rule As cRule
    
    lLongEntry As Long
    lLongExit As Long
    lShortEntry As Long
    lShortExit As Long
    
    strName As String
End Type
Private m As mPrivate

'Input columns
Private Enum eIGCols
    eIGCol_InputName = 0
    eIGCol_DefaultValue
    eIGCol_FromVal
    eIGCol_ToVal
    eIGCol_ParmTypeID
    eIGCol_ParmID
    eIGCol_ParmDesc
    eIGCol_RuleID
    eIGCol_Sort
    eIGCol_Req
    eIGCol_ListID
    eIGCol_NumCols
End Enum

'Action types
Private Enum eRuleAction
    eRA_LongEntry = 0
    eRA_LongExit = 1
    eRA_ShortEntry = 2
    eRA_ShortExit = 3
End Enum

'Order types
Private Enum eOrderType
    eOT_Market = 0
    eOT_Stop = 1
    eOT_Limit = 2
    eOT_StopLimit = 3
    eOT_MarketClose = 4
    eOT_StopClose = 5
    eOT_LimitClose = 6
    eOT_StopLimitClose = 7
End Enum

Private Enum eTabs
    eTab_Rule = 0
    eTab_AdvancedRule = 1
    eTab_Inputs = 2
    eTab_Advanced = 3
End Enum

Private Function Tabs(ByVal lTab As eTabs) As Long
    Tabs = lTab
End Function
Private Function IGCol(ByVal lColumn As eIGCols) As Long
    IGCol = lColumn
End Function
Private Function Action(ByVal lAction As eRuleAction) As Long
    Action = lAction
End Function
Private Function OrderType(ByVal lOrderType As eOrderType) As Long
    OrderType = lOrderType
End Function

Public Property Get ID() As Long
    ID = m.Rule.RuleID
End Property
Public Property Get SystemID() As Long
    SystemID = m.lSystemNumber
End Property
Property Get ReloadMenu() As Boolean
    ReloadMenu = m.bReloadMenu
End Property
Property Let CalledFrom(pData As String)
    m.strCalledFrom = pData
End Property
Property Let SystemNumber(pData As Long)
    m.lSystemNumber = pData
End Property
Property Let LongEntryNum(pData As Long)
    m.lLongEntry = pData
End Property
Property Let LongExitNum(pData As Long)
    m.lLongExit = pData
End Property
Property Let ShortEntryNum(pData As Long)
    m.lShortEntry = pData
End Property
Property Let ShortExitNum(pData As Long)
    m.lShortExit = pData
End Property
Property Let SystemLibID(pData As Long)
    m.lSystemLibID = pData
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadRec
'' Description: Load the rule for the ID given
'' Inputs:      Rule ID to load
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadRec(ByVal lRuleID As Long, Optional ByVal lSystemID As Long = 0&, _
                Optional ByVal bIgnoreSecurity As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    'ShowForm Me
    'm.lReturnValue = LockWindowUpdate(Me.hWnd)
    
    m.lCurrentRuleID = lRuleID
    m.lSystemNumber = lSystemID
    m.bNewRule = False
    
    'If Rule exists in the collection then continue...
    Set m.Rule = New cRule
    With m.Rule
        .RuleID = m.lCurrentRuleID
        .Load
        
        If Not bIgnoreSecurity Then
            If Not g.Security.CanEdit(.SecurityLevel, .Password) Then
                GoTo ErrExit:
            End If
        End If
        
        'Load screen with rule settings
        ClearForm
        m.strName = .Name
        
        'Zero for System number means global
        If .SystemNumber <> 0 Or m.strCalledFrom = "frmSystemManager" Then
            chkGlobal.Value = vbUnchecked
            m.bShared = False
        Else
            chkGlobal.Value = vbChecked
            m.bNewRule = False
            m.bShared = True
        End If
        chkGlobal.Visible = False
        If m.lSystemNumber <> 0 Then
            vsSystemName.Caption = "Local Rule for " & Replace(SystemNameForID(m.lSystemNumber), "&", "&&")
        Else
            vsSystemName.Caption = "Building Block Rule"
        End If
        
        Editor1.TextRTF = .GetRTF(.Cond)
        edLimit.TextRTF = .GetRTF(.Price2RTF)
        
        If .BuySell Then
            If .RuleType = 0 Then
                If IsDefaultRuleName(.Name) Then m.lLongEntry = DefaultRuleNumber(.Name)
                cboAction.ListIndex = Action(eRA_LongEntry)
            Else
                If IsDefaultRuleName(.Name) Then m.lShortExit = DefaultRuleNumber(.Name)
                cboAction.ListIndex = Action(eRA_ShortExit)
            End If
        Else
            If .RuleType = 0 Then
                If IsDefaultRuleName(.Name) Then m.lShortEntry = DefaultRuleNumber(.Name)
                cboAction.ListIndex = Action(eRA_ShortEntry)
            Else
                If IsDefaultRuleName(.Name) Then m.lLongExit = DefaultRuleNumber(.Name)
                cboAction.ListIndex = Action(eRA_LongExit)
            End If
        End If
        
        If .ExitOnEntryBar = True Then chkCanExitOnEntryBar.Value = vbChecked Else chkCanExitOnEntryBar.Value = vbUnchecked
        cboOrderPlacement.Text = .OrderPlacement
        Editor2.TextRTF = .GetRTF(.Price)
        
        'Temporary, remove after converting all rules from rtf stored format into
        'tokened text format 7-17-2000 MT
        If Left(.Cond, 1) = "{" Then
            Editor1.TextRTF = .Cond
            edLimit.TextRTF = .Price2RTF
            Editor2.TextRTF = .Price
        End If
        
        ' Take care of pyramiding stuff 5/18/2001 DAJ
        If .RuleType = 0 Then
            txtContractsToEnter.Text = Format(.NumberContracts, "#,##0")
        Else
            If .ExitBasedOnEachTrade Then
                optTrade.Value = True
            Else
                optPosition.Value = True
                If .AsPercentOfPosition Then
                    optPercent.Value = True
                    txtPercent.Text = CStr(.NumberContracts)
                Else
                    optNumContracts.Value = True
                    txtContractsToExit.Text = Format(.NumberContracts, "#,##0")
                End If
            End If
        End If
        
        ' Select the appropriate item in the rule category combo box...
        If .CategoryID > 0 Then
            cboCategory.Text = RuleCategoryFromID(.CategoryID)
        ElseIf .RuleType = 0 Then
            cboCategory.Text = "Other Entries"
        Else
            cboCategory.Text = "Other Exits"
        End If
    End With
    
    SetEditorCaption Me, "Rule", m.strName
    
    'Initiliaze inputs grid
    Set m.Inputs = m.Rule.Inputs
    LoadInputsGrid
    
    'Build advanced view
    With m.Rule
        edAdvanced.TextRTF = .GetRTF(EZToAdvanced(.BuySell, .OrderPlacement, _
            .Cond, .Price, .Price2RTF))
        chkAdvanced_Click
    End With
    
    m.lReturnValue = LockWindowUpdate(0)
    
    tbToolbar.Tools("ID_Verify").Enabled = False
    tbToolbar.Tools("ID_Diagram").Enabled = True
    EnableToolbar False
    
    LoadRec = True
    
ErrExit:
    m.lReturnValue = LockWindowUpdate(0)
    Exit Function

ErrSection:
    RaiseError "frmRule.LoadRec", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EZToAdvanced
'' Description: Combines the condition and price code text and creates the
''              rich text for editing in advanced mode
'' Inputs:      Buy/Sell, Order Placement, Condition, Stop Price, Limit Price
'' Returns:     Advanced rich text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EZToAdvanced(pBuySell As Boolean, pOrderPlacement As String, _
    pCond As String, pPrice1 As String, pPrice2 As String) As String
On Error GoTo ErrSection:
    
    Dim strTemp         As String

    'Condition, Then followed by <enterkey>
    strTemp = Trim(pCond) & " ~24005THEN ~80003<E> ~44005<Tab> "

    'Order action (Buy/Sell)
    If pBuySell Then
        strTemp = strTemp & "~24003BUY ~16001( "
    Else
        strTemp = strTemp & "~24004SELL ~16001( "
    End If
    
    'Entry Price plus a comma...
    'Price may be blank for existing rules (before the new tradesense merge).  In
    'these cases, default the price coded text for market orders.
    If pOrderPlacement = "Market" Or pOrderPlacement = "Market on close" Then
        strTemp = strTemp & "~01013Next Bar Open ~22001, "
    Else
        strTemp = strTemp & Trim(pPrice1) & " ~22001, "
    End If
    
    'Order type
    strTemp = strTemp & "~20" & Format(Len(pOrderPlacement) + 2, "000") & _
        """" & pOrderPlacement & """"
    
    'Stop/Limit Price
    If Len(pPrice2) > 0 And pPrice2 <> "~00000" Then
        strTemp = strTemp & " ~22001, " & pPrice2
    End If
    
    'Function right paren...
    strTemp = Trim(strTemp) & " ~17001) ~80003<E> ~43005EndIf ~80003<E>"
    
    'Return reformatting rule...
    EZToAdvanced = strTemp
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.EZToAdvanced", eGDRaiseError_Raise
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Set the rule manager up for a new rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add()
On Error GoTo ErrSection:
    
    m.bSkipAutoIf = True
    m.bNewRule = True
    
    'Setup form...
    ClearForm
    
    chkAdvanced_Click
    
    cboAction.ListIndex = Action(eRA_LongEntry)
    chkCanExitOnEntryBar.Value = vbUnchecked
    chkCanExitOnEntryBar.Enabled = False
    cboOrderPlacement.ListIndex = OrderType(eOT_Market)
    
    txtContractsToEnter.Text = "1"
    txtContractsToExit.Text = "1"
    txtPercent.Text = "100"
    
    'If called from System Manager then default to Local rule.  Otherwise
    'it was called from the menu so default to global.
    If m.strCalledFrom = "frmSystemManager" Then
        'chkGlobal.Enabled = True
        chkGlobal.Value = vbUnchecked
        m.bShared = False
    Else
        'chkGlobal.Enabled = False
        chkGlobal.Value = vbChecked
        m.bShared = True
    End If
    chkGlobal.Visible = False
    If m.lSystemNumber <> 0 Then
        vsSystemName.Caption = "Local Rule for " & Replace(SystemNameForID(m.lSystemNumber), "&", "&&")
    Else
        vsSystemName.Caption = "Building Block Rule"
    End If
    
    ' If there are categories in the combo box, select the first one...
    'If cboCategory.ListCount > 0 Then cboCategory.ListIndex = 0
    If cboCategory.ListCount > 0 Then cboCategory.Text = "Other Entries"
    
    ''m.strName = SetDefaultRuleName
    m.strName = ""
    SetEditorCaption Me, "Rule", m.strName
    
    Set m.Rule = New cRule
    m.lCurrentRuleID = 0
    
    tbToolbar.Tools("ID_Verify").Enabled = False
    tbToolbar.Tools("ID_Diagram").Enabled = True
    EnableToolbar False
    
    Set m.Inputs = New cInputs
    LoadInputsGrid
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Add", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the rule
'' Inputs:      Button, Ask Password?
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save(ByVal strButton As String, Optional ByVal bAskPassword As Boolean = True) As Boolean
On Error GoTo ErrSection:
    
    Dim bSuccess As Boolean
    Dim rs As Recordset
    Dim strNewName As String
    Dim strTemp As String
    Dim bAskRename As Boolean
    Dim bSaveAs As Boolean
    Dim strText As String
    Dim strError As String
    Dim strTempName As String
    Dim lIndex As Long
       
    ' Check to make sure that the user has permission to Save/Rename first.  On a
    ' SaveAs situation, make sure if the rule is local that the user can save to
    ' the system.  (DAJ: 05/22/2003)
    If bAskPassword Then
        If strButton <> "ID_SaveAs" Then
            If Not g.Security.CanSave(m.Rule.SecurityLevel, m.Rule.Password) Then GoTo ErrExit:
        ElseIf m.strCalledFrom = "frmSystemManager" Then
            If Not m.frmSysMgr Is Nothing Then
                If Not g.Security.CanSave(m.frmSysMgr.System.SecurityLevel, m.frmSysMgr.System.Password) Then GoTo ErrExit:
            End If
        End If
    End If
       
    ' TLB 8/9/00: must clear unused expressions so engine won't
    ' choke on a possibly invalid yet hidden expression
    If (Not Editor2.Visible And Editor2.Text <> "") _
        Or (Not edLimit.Visible And edLimit.Text <> "") Then
            tbToolbar.Tools("ID_Verify").Enabled = True
    End If

    ' Only reverify if necessary (if a rule has changed)
    If tbToolbar.Tools("ID_Verify").Enabled Then
        Verify
        If tbToolbar.Tools("ID_Verify").Enabled Then Exit Function
    End If
    
    ' Make sure that if the user used Next Bar High, Next Bar Low, or Next Bar Close
    ' either in the rule or in one of the functions used in the rule that we only
    ' allow them to do "On Close" type orders...
    If CloseOnlyOrders Then
        If InStr(UCase(cboOrderPlacement.Text), "CLOSE") = 0 Then
            Err.Raise vbObjectError + 1000, , "Only 'On Close' or 'Close Only' type of orders can reference the high, low or close of the next bar."
        End If
    End If
    
    ' Handle Rename/Save As
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        strText = "Save the current Rule as..."
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & SetDefaultRuleName & " ; " & strText)
    ElseIf strButton = "ID_SaveAs" Then
        strText = "Save a copy of the current Rule as..."
        strTempName = SetDefaultRuleName
        If strTempName = "" Then strTempName = m.strName & " #02"
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & strTempName & " ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        strText = "Rename the current Rule as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    Else
        strNewName = m.strName
    End If
    
    ' Verify that it is a good name
    Do While Len(Trim(strNewName)) > 0
        strError = m.Rule.ValidName(strNewName)
        If strError <> "" Then
            InfBox strError, "e", , "Error"
        Else
            Exit Do
        End If
        strText = "Rename the Rule as..."
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
    Loop
    
    If Len(Trim(strNewName)) = 0 Then
        Exit Function 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    m.strName = Trim(strNewName)
        
        
'???????????????????????
        
    ' Make sure rule name is unique in tblRules if marked global.  If the
    ' flag is checked off (made local), make sure the rule isn't in any
    ' other systems.
    bAskRename = True
    If m.bShared <> (chkGlobal.Value = vbChecked) Then
        If chkGlobal.Value = vbChecked Then
            If IsDefaultRuleName(m.strName) Then
                strNewName = InfBox("Please Enter in a Name for the New Shared Rule", "", , "Shared Rule Name", , , , , , "string")
                If Trim(strNewName) = "" Then GoTo ErrExit
                If InStr(1, strNewName, "'") > 0 Then
                    Err.Raise vbObjectError + 1000, , "Single quotes not allowed in Rule Name"
                End If
                bAskRename = False
                m.strName = Trim(strNewName)
            End If
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                "WHERE (([Name] = '" & m.strName & "') AND " & _
                "([RuleID] <> " & m.Rule.RuleID & "));", dbOpenSnapshot)
            If Not rs.EOF Then
                Err.Raise vbObjectError + 1000, , "Rules made available to all strategies must have a unique name"
            End If
            
            ' Since the user is making this rule global, we cannot assume
            ' that the rule is in the "current" library anymore since they
            ' can use the rule anywhere.  It is up to the user to move the
            ' rule back into an appropriate library
            m.Rule.LibraryID = kSN_UserLibrary
        Else
            strNewName = InfBox("This will make a local copy of this shared rule.  Please enter in a new name for the local copy:", _
                        "?", , "Copy Rule", , , , , , "string", SetDefaultRuleName)
            If Trim(strNewName) = "" Then GoTo ErrExit
            If InStr(1, strNewName, "'") > 0 Then
                Err.Raise vbObjectError + 1000, , "Single quotes not allowed in Rule Name"
            End If
            m.lSystemNumber = -2
            If RuleExists(strNewName) Then
                Err.Raise vbObjectError + 1000, , "Rule Name '" & strNewName & "' already exists in strategy"
            End If
            bAskRename = False
            m.strName = Trim(strNewName)
            m.lCurrentRuleID = 0
            m.bNewRule = True
            
            ' Since the user has made rule a local copy, we need to assign
            ' the Library ID of the System to ensure that this new rule stays
            ' with the System
            m.Rule.LibraryID = m.lSystemLibID
        End If
    ElseIf m.bNewRule = True Then
        bAskRename = False
        If chkGlobal.Value = vbChecked Then
            ' Since the user is making this rule global, we cannot assume
            ' that the rule is in the "current" library anymore since they
            ' can use the rule anywhere.  It is up to the user to move the
            ' rule back into an appropriate library
            m.Rule.LibraryID = kSN_UserLibrary
        Else
            ' Since the user has made rule a local copy, we need to assign
            ' the Library ID of the System to ensure that this new rule stays
            ' with the System
            m.Rule.LibraryID = m.lSystemLibID
        End If
    End If
    
    'If the rule name changes, allow user to make a copy
    m.strName = Trim(m.strName)
    'If (Len(m.Rule.Name) > 0) And (m.strName <> m.Rule.Name) Then
    If m.strName <> m.Rule.Name Then
        If RuleExists(m.strName) Then
            strText = "Rule " & m.strName & " already exists"
            m.strName = ""
            Err.Raise vbObjectError + 1000, , strText
        End If
        If bSaveAs Then
            If m.strCalledFrom = "frmSystemManager" Then
                If Not m.frmSysMgr Is Nothing Then
                    Set m.Rule = m.Rule.MakeCopy(m.frmSysMgr.NextRuleID, m.lSystemNumber)
                    m.lCurrentRuleID = m.Rule.RuleID
                End If
            Else
                Set m.Rule = m.Rule.MakeCopy
                m.lCurrentRuleID = 0
            End If
        End If
    End If
    
    'Validate and save rule
    m.lReturnValue = LockWindowUpdate(Me.hWnd)
    
    SetEditorCaption Me, "Rule", m.strName
    With m.Rule
        
        'If new rule or copy of existing then put into user library with no restrictions
        If m.lCurrentRuleID = 0 Or bSaveAs Then
            If .SecurityLevel < 2 Then
                .SecurityLevel = 0
                .CannotDelete = False
                If bAskRename Then .LibraryID = kSN_UserLibrary
                .Password = ""
            End If
        Else
            'User must be authorized to rename (save)
            'If Not g.Security.CanSave(.SecurityLevel, .Password) Then
            '    GoTo ErrExit:
            'End If
        End If
        
        .RuleID = m.lCurrentRuleID
        
        SetEditorCaption Me, "Rule", m.strName
        
        .Name = m.strName
        .LastModified = Now()
        .ActionCodedName = "N/A"
                
        'If opBuy Then .BuySell = True Else .BuySell = False
        Select Case cboAction.ListIndex
            Case Action(eRA_LongEntry)
                .BuySell = True
                .RuleType = 0
            Case Action(eRA_LongExit)
                .BuySell = False
                .RuleType = 1
            Case Action(eRA_ShortEntry)
                .BuySell = False
                .RuleType = 0
            Case Action(eRA_ShortExit)
                .BuySell = True
                .RuleType = 1
        End Select
        If chkCanExitOnEntryBar.Value = vbChecked Then
            .ExitOnEntryBar = (.RuleType = 1)
        Else
            .ExitOnEntryBar = False
        End If
            
        If .RuleType = 0 Then
            ' Entry
            .ExitBasedOnEachTrade = False
            .AsPercentOfPosition = False
            .NumberContracts = ValOfText(txtContractsToEnter.Text)
        Else
            ' Exit
            If optPosition.Value = True Then
                If optNumContracts.Value = True Then
                    .ExitBasedOnEachTrade = False
                    .AsPercentOfPosition = False
                    .NumberContracts = ValOfText(txtContractsToExit.Text)
                Else
                    .ExitBasedOnEachTrade = False
                    .AsPercentOfPosition = True
                    .NumberContracts = ValOfText(txtPercent.Text)
                End If
            Else
                .ExitBasedOnEachTrade = True
                .AsPercentOfPosition = False
            End If
        End If
        
        If m.bShared Then
            .CategoryID = cboCategory.ItemData(cboCategory.ListIndex)
        Else
            .CategoryID = 0&
        End If
        .OrderPlacement = cboOrderPlacement.Text
        '.RuleType = 0
    
        'Load current grid values to inputs collection which is then saved
        'in the cRule class
        LoadInputsCollection
        .Inputs = m.Inputs.MakeCopy(.RuleID, m.lSystemNumber)
        
        If m.strCalledFrom = "frmSystemManager" Then
            .RuleUse = .RuleType
            .Validate
            .CondFillWords = .PreviewRTF
            .LastModKnown = .LastModified
            If .RuleID = 0 Then
                .SystemNumber = -2&
                If Not m.frmSysMgr Is Nothing Then
                    .RuleID = m.frmSysMgr.NextRuleID
                    .SystemNumber = m.frmSysMgr.ID
                    m.lCurrentRuleID = .RuleID
                End If
            End If
            
            If .RuleID <= 0 Then
                For lIndex = 1 To .Inputs.Count
                    .Inputs.Item(lIndex).ParmID = (.RuleID * 100) + (lIndex * Sgn(.RuleID))
                    .Inputs.Item(lIndex).RuleID = .RuleID
                    .Inputs.Item(lIndex) = .Inputs.Item(lIndex)
                Next lIndex
            End If
        Else
            .Save
            m.lCurrentRuleID = .RuleID
            g.bDirtyLibrariesMDB = True
        
            ' Reload the rule to get the PreviewRTF text...
            m.Rule.Load
            
            ' If the user changed a favorite, ask if they wish to update
            ' any locals with the same name...
            If m.lSystemNumber = 0 Then frmUpdateLocals.ShowMe m.Rule
        End If
    
        ' For a local rule with a default rule name, increase the appropriate
        ' variable so that if the user hits SaveAs, it comes up with the correct
        ' default rule name
        If IsDefaultRuleName(m.strName) And (m.Rule.SystemNumber <> 0) Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry)
                    m.lLongEntry = m.lLongEntry + 1
                Case Action(eRA_LongExit)
                    m.lLongExit = m.lLongExit + 1
                Case Action(eRA_ShortEntry)
                    m.lShortEntry = m.lShortEntry + 1
                Case Action(eRA_ShortExit)
                    m.lShortExit = m.lShortExit + 1
            End Select
        End If

        ' Refresh the Rule in the Global table...
        If m.lCurrentRuleID > 0 Then RefreshRule m.Rule
        
        tbToolbar.Tools("ID_Verify").Enabled = False
        tbToolbar.Tools("ID_Diagram").Enabled = True
        EnableToolbar False
        
        If m.strCalledFrom = "frmSystemManager" Then
            If Not m.frmSysMgr Is Nothing Then
                If IsWindow(m.frmSysMgr.hWnd) Then
                    m.frmSysMgr.AddRule m.Rule.MakeCopy(m.Rule.RuleID, m.Rule.SystemNumber) '.RuleID
                End If
            End If
        End If
    End With
    
    bSuccess = True
    
ErrExit:
    Screen.MousePointer = vbDefault
    m.lReturnValue = LockWindowUpdate(0)
    Save = bSuccess
    Exit Function

ErrSection:
    If Err.Number < 0 Then
        Select Case m.Rule.ErrNbr
            Case 2: MoveFocus Editor1
            Case 3: MoveFocus Editor2
            Case 4: MoveFocus edLimit
        End Select
    End If
    Screen.MousePointer = vbDefault
    m.lReturnValue = LockWindowUpdate(0)
    RaiseError "frmRule.Save", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAction_Change
'' Description: Reset controls based on the new action
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAction_Change()
On Error GoTo ErrSection:
   
    DisplayOrderControls
    EnableToolbar True
    
    If Me.Visible Then
        If m.strCalledFrom = "frmSystemManager" And IsDefaultRuleName(m.strName) Then
            ''m.strName = SetDefaultRuleName
            ''SetEditorCaption Me, "Rule", m.strName
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.cboAction.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAction_Click
'' Description: Reset controls based on the new action
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAction_Click()
On Error GoTo ErrSection:
    
    DisplayOrderControls
    EnableToolbar True
    
    If Me.Visible Then
        If m.strCalledFrom = "frmSystemManager" And IsDefaultRuleName(m.strName) Then
            ''m.strName = SetDefaultRuleName
            ''SetEditorCaption Me, "Rule", m.strName
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.cboAction.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDefaultRuleName
'' Description: Set the default rule name for a local rule
'' Inputs:      None
'' Returns:     Default Rule Name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetDefaultRuleName() As String
On Error GoTo ErrSection:
    
    Dim lNumber As Long
    Dim bBuySell As Boolean
    Dim lRuleType As Long
    Dim strAction As String
        
    If chkGlobal.Value = vbUnchecked Then
        Select Case cboAction.ListIndex
            Case Action(eRA_LongEntry)
                bBuySell = True
                lRuleType = 0
                strAction = "Long Entry"
                lNumber = m.lLongEntry
            Case Action(eRA_LongExit)
                bBuySell = False
                lRuleType = 1
                strAction = "Long Exit"
                lNumber = m.lLongExit
            Case Action(eRA_ShortEntry)
                bBuySell = False
                lRuleType = 0
                strAction = "Short Entry"
                lNumber = m.lShortEntry
            Case Action(eRA_ShortExit)
                bBuySell = True
                lRuleType = 1
                strAction = "Short Exit"
                lNumber = m.lShortExit
        End Select
                
        SetDefaultRuleName = strAction & " #" & Format(lNumber, "000")
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.SetDefaultRuleName", eGDRaiseError_Raise
    Resume ErrExit
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOrderPlacement_Change
'' Description: Reset controls based on the new order placement
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOrderPlacement_Change()
On Error GoTo ErrSection:
    
    DisplayOrderControls
    EnableToolbar True
    
    If Me.Visible Then
        Select Case cboOrderPlacement.ListIndex
            Case OrderType(eOT_Stop), OrderType(eOT_StopClose)
                MoveFocus Editor2
            
            Case OrderType(eOT_Limit), OrderType(eOT_LimitClose)
                MoveFocus Editor2
            
            Case OrderType(eOT_StopLimit), OrderType(eOT_StopLimitClose)
                If Len(Trim(Editor2.Text)) = 0 Then
                    MoveFocus Editor2
                ElseIf Len(Trim(edLimit.Text)) = 0 Then
                    MoveFocus edLimit
                Else
                    MoveFocus Editor2
                End If
                
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.cboOrderPlacement.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOrderPlacement_Change
'' Description: Reset controls based on the new order placement
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOrderPlacement_Click()
On Error GoTo ErrSection:
    
    DisplayOrderControls
    EnableToolbar True

    If Me.Visible Then
        Select Case cboOrderPlacement.ListIndex
            Case OrderType(eOT_Stop), OrderType(eOT_StopClose)
                MoveFocus Editor2
            
            Case OrderType(eOT_Limit), OrderType(eOT_LimitClose)
                MoveFocus Editor2
            
            Case OrderType(eOT_StopLimit), OrderType(eOT_StopLimitClose)
                If Len(Trim(Editor2.Text)) = 0 Then
                    MoveFocus Editor2
                ElseIf Len(Trim(edLimit.Text)) = 0 Then
                    MoveFocus edLimit
                Else
                    MoveFocus Editor2
                End If
                
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.cboOrderPlacement.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCanExitOnEntryBar_Click
'' Description: Enable the Save Button (mark the Rule as Dirty)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCanExitOnEntryBar_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.chkCanExitOnEntryBar.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAdvanced_Click
'' Description: Show/Hide Advanced functionality as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAdvanced_Click()
On Error GoTo ErrSection:

    Dim bBuy As Byte

    'Advanced Editor:  Concatenate EZ editor stuff into Advanced editor
    If chkAdvanced.Value = 1 Then
        vsIndexTab1.TabVisible(0) = False
        vsIndexTab1.TabVisible(1) = True
        vsIndexTab1.CurrTab = Tabs(eTab_AdvancedRule)
        
        'If Buy/Sell changes or Order Placement changes then rebuild
        'Advanced text to guarentee syntax is updated
        If Len(Trim(Editor1.Text)) > 0 And Trim(Editor1.Text) <> "If" Then
            bBuy = cboAction.ListIndex = Action(eRA_LongEntry) Or cboAction.ListIndex = Action(eRA_ShortExit)
            edAdvanced.Text = BuildRule(Editor1.Text, bBuy, Editor2.Text, _
                cboOrderPlacement.ListIndex, edLimit.Text)
            Verify
        End If
    Else
        'EZ Editor: Parse Advanced editor format into EZ editor tab
        vsIndexTab1.TabVisible(Tabs(eTab_Rule)) = True
        vsIndexTab1.TabVisible(Tabs(eTab_AdvancedRule)) = False
        vsIndexTab1.CurrTab = Tabs(eTab_Rule)
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.chkAdvanced.Click", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkGlobal_Click
'' Description: Enable the Save Button (mark the Rule as Dirty)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkGlobal_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.chkGlobal.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "SNV Rule", Me, 0
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.PrintMe", eGDRaiseError_Raise
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edAdvanced_Change
'' Description: Enable/Disable buttons as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edAdvanced_Change()
On Error GoTo ErrSection:

    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = True
    tbToolbar.Tools("ID_Diagram").Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.edAdvanced.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edAdvanced_EditFunction
'' Description: Allow the user to edit the function chosen
'' Inputs:      Function ID, Function Name, Whether it was Found
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edAdvanced_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edAdvanced.EditFunction", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edAdvacned_GotFocus
'' Description: Set the control up appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edAdvanced_GotFocus()
On Error GoTo ErrSection:

    Dim Control As Control
    
    'Disable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Me.Controls
'       Control.TabStop = False
'    Next Control
    
    Set g.ActiveEditor = edAdvanced
    With edAdvanced
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = False
        .Usage = 2             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With
    
    If Len(Trim(edAdvanced.Text)) = 0 And Not m.bSkipAutoIf Then
        edAdvanced.Text = ""
        SendKeys "IF "
    End If
    
    m.bSkipAutoIf = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edAdvanced.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edAdvanced_KeyUp
'' Description: Perform actions based on the user pressing a key
'' Inputs:      KeyCode of Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edAdvanced_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    'Alt-1 to show rule tree
    If KeyCode = 49 And Shift = 4 Then
        VerifyRuleDebug
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edAdvanced.KeyUp", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edAdvanced_LostFocus
'' Description: Turn the Editor off upon losing focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edAdvanced_LostFocus()
On Error GoTo ErrSection:
    
    Dim Control     As Control
    
    'Enable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Controls
'       Control.TabStop = True
'    Next Control
    
    Set g.ActiveEditor = Nothing
    edAdvanced.RemoveTradeSense
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edAdvanced.LostFocus", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_KeyUp
'' Description: Alt-1 shows the rule tree
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    'Alt-1 to show rule tree
    If KeyCode = 49 And Shift = 4 Then
        VerifyRuleDebug
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor1.KeyUp", eGDRaiseError_Show
    Resume ErrExit:

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
    RaiseError "frmRule.Editor1.NewFunction", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor2_KeyUp
'' Description: Alt-1 shows the rule tree
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    'Alt-1 to show rule tree
    If KeyCode = 49 And Shift = 4 Then
        VerifyRuleDebug
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor2.KeyUp", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor2_NewFunction
'' Description: Allow the user to create a new function
'' Inputs:      Category ID the Function List form was currently on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor2_NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim frm As frmFunctionMgrCT         ' New Function Manager form
    
    Set frm = New frmFunctionMgrCT
    frm.ShowMe 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.Editor2.NewFunction", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edLimit_KeyUp
'' Description: Alt-1 shows the rule tree
'' Inputs:      Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edLimit_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    'Alt-1 to show rule tree
    If KeyCode = 49 And Shift = 4 Then
        VerifyRuleDebug
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edLimit.KeyUp", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_Change
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_Change()
On Error GoTo ErrSection:
    
    Dim lPos As Long                    ' Position of the IF in the text
    
    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = (Len(Trim(Editor1.Text)) > 0)
    tbToolbar.Tools("ID_Diagram").Enabled = False
    
    'lPos = InStr(UCase(Editor1.Text), "IF")
    lPos = PositionOfIf
    If lPos <> 0 Then
        ' If the user tries to type in an assignment expression after the IF, put up an error message...
        If InStr(lPos, Editor1.Text, ":=") <> 0 Then
            Tag = "Editor1"
            InfBox "You cannot have an assignment operator after the IF.", "!", , "Rule Error"
            Editor1.Text = Left(Editor1.Text, lPos - 1) & Replace(Editor1.Text, ":=", "", lPos)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.Editor1.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_EditFunction
'' Description: Allow the user to edit the function they clicked on
'' Inputs:      FunctionID, Function Name, Whether Found or not
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor1.EditFunction", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor2_EditFunction
'' Description: Allow the user to edit the function they clicked on
'' Inputs:      FunctionID, Function Name, Whether Found or not
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor2_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor2.EditFunction", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edLimit_EditFunction
'' Description: Allow the user to edit the function they clicked on
'' Inputs:      FunctionID, Function Name, Whether Found or not
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edLimit_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:

    ShowFunctionMgr FunctionID, FunctionName, Found
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edLimit.EditFunction", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_GotFocus
'' Description: When the editor gets focus, refresh it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_GotFocus()
On Error GoTo ErrSection:
    
    Dim Control As Control
    
    'Disable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Me.Controls
'       Control.TabStop = False
'    Next Control
    
    Set g.ActiveEditor = Editor1
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
    
    If Len(Trim(Editor1.Text)) = 0 And Not m.bSkipAutoIf Then
        Editor1.Text = ""
        SendKeys "IF "
    End If
    
    m.bSkipAutoIf = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor1.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor1_LostFocus
'' Description: When the editor loses focus, remove TradeSense
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor1_LostFocus()
On Error GoTo ErrSection:
    
    Dim Control     As Control
    
    'Enable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Controls
'       Control.TabStop = True
'    Next Control
    
    Set g.ActiveEditor = Nothing
    Editor1.RemoveTradeSense
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor1.LostFocus", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor2_Change
'' Description: Disable/Enable the buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor2_Change()
On Error GoTo ErrSection:
    
    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = (Len(Trim(Editor2.Text)) > 0)
    tbToolbar.Tools("ID_Diagram").Enabled = False
    
    ' If the user tries to type in an assignment expression into the Limit price
    ' expression, put up an error message...
    If InStr(Editor2.Text, ":=") <> 0 Then
        Tag = "Editor2"
        InfBox "You cannot have an assignment operator in this expression.", "!", , "Expression Error"
        Editor2.Text = Replace(Editor2.Text, ":=", "")
        If Len(Editor2.Text) > 0 Then
            Editor2.SelStart = Len(Editor2.Text)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.Editor2.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor2_GotFocus
'' Description: When the editor gets focus, refresh it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor2_GotFocus()
On Error GoTo ErrSection:

    Dim Control As Control
    
    'Disable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Me.Controls
'       Control.TabStop = False
'    Next Control

    Set g.ActiveEditor = Editor2
    With Editor2
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .ShowNewFunction = True
        .Usage = 2             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With
    
    If Len(Trim(Editor2.Text)) = 0 And Not m.bSkipAutoIf Then
        Editor2.Text = ""
        SendKeys " "
    End If
    
    m.bSkipAutoIf = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor2.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Editor2_LostFocus
'' Description: When the editor loses focus, remove Trade Sense
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editor2_LostFocus()
On Error GoTo ErrSection:
    
    Dim Control     As Control
    
    'Enable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Controls
'       Control.TabStop = True
'    Next Control
    
    Set g.ActiveEditor = Nothing
    Editor2.RemoveTradeSense
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Editor2.LostFocus", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edLimit_Change
'' Description: Disable/Enable the buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edLimit_Change()
On Error GoTo ErrSection:
    
    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = (Len(Trim(edLimit.Text)) > 0)
    tbToolbar.Tools("ID_Diagram").Enabled = False
    
    ' If the user tries to type in an assignment expression into the Limit price
    ' expression, put up an error message...
    If InStr(edLimit.Text, ":=") <> 0 Then
        Tag = "edLimit"
        InfBox "You cannot have an assignment operator in the Limit price expression.", "!", , "Expression Error"
        edLimit.Text = Replace(edLimit.Text, ":=", "")
        If Len(edLimit.Text) > 0 Then
            edLimit.SelStart = Len(edLimit.Text)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.edLimit.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edLimit_GotFocus
'' Description: When the editor gets focus, refresh it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edLimit_GotFocus()
On Error GoTo ErrSection:

    Dim Control As Control
    
    'Disable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Me.Controls
'       Control.TabStop = False
'    Next Control

    Set g.ActiveEditor = edLimit
    With edLimit
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .ShowNewFunction = True
        .Usage = 2             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With
    
    If Len(Trim(edLimit.Text)) = 0 And Not m.bSkipAutoIf Then
        edLimit.Text = ""
        SendKeys " "
    End If
    
    m.bSkipAutoIf = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edLimit.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edLimit_LostFocus
'' Description: When the editor loses focus, remove Trade Sense
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edLimit_LostFocus()
On Error GoTo ErrSection:
    
    Dim Control     As Control
    
    'Enable tabbing for all controls so formatting can occur in Editor
    'Ignore errors for controls without the TabStop property.
' Turn this off for now because of conflicts with combo boxes 4/22/2002 DAJ
'    On Error Resume Next
'    For Each Control In Controls
'       Control.TabStop = True
'    Next Control
    
    Set g.ActiveEditor = Nothing
    edLimit.RemoveTradeSense
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.edLimit.LostFocus", eGDRaiseError_Show
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    edLimit_NewFunction
'' Description: Allow the user to create a new function
'' Inputs:      Category ID the Function List form was currently on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub edLimit_NewFunction(ByVal lCategoryID As Long)
On Error GoTo ErrSection:

    Dim frm As frmFunctionMgrCT         ' New Function Manager form
    
    Set frm = New frmFunctionMgrCT
    frm.ShowMe 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.edLimit.NewFunction", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form gets focus again, refresh it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
        
    If g.Functions Is Nothing Then
        InitFunctions
    End If
    
    ' Load internally generated TradeSense lists (Symbols, etc.)
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    ' Quickly check the Reverify flag.  If on then force a reverify...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [RuleID]=" & m.Rule.RuleID & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblRules"
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        
        If rs!CheckSum = 0.5 Then
            EnableToolbar False
            Unload Me
            Err.Raise vbObjectError + 1000, , "This Rule is no longer Valid"

        ElseIf rs!Reverify Then
            EnableToolbar True
            tbToolbar.Tools("ID_Verify").Enabled = True
            tbToolbar.Tools("ID_Diagram").Enabled = False
        End If
    End If
    
    ' Hide advanced rules tab
    Select Case UCase(Tag)
        Case "EDLIMIT"
            If GetActiveWindow = hWnd Then
                MoveFocus edLimit
            End If
        
        Case "EDITOR2"
            If GetActiveWindow = hWnd Then
                MoveFocus Editor2
            End If
        
        Case Else
            If chkAdvanced.Value = 1 Then
                vsIndexTab1.TabVisible(0) = False
                vsIndexTab1.TabVisible(1) = True
                vsIndexTab1.CurrTab = Tabs(eTab_AdvancedRule)
                MoveFocus edAdvanced
            Else
                vsIndexTab1.TabVisible(0) = True
                vsIndexTab1.TabVisible(1) = False
                vsIndexTab1.CurrTab = Tabs(eTab_Rule)
                If GetActiveWindow = Me.hWnd Then
                    MoveFocus Editor1
                    ''If Len(Trim(Editor1.Text)) = 0 And Not m.bSkipAutoIf Then SendKeys "IF "
                End If
            End If
    End Select
    Tag = ""
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    Set rs = Nothing
    RaiseError "frmRule.Form.Activate", eGDRaiseError_Show
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
    RaiseError "frmRule.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, initialize the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim strText As String
    Dim X As Long
    Dim strFont As String
    
    g.Styler.StyleForm Me
    
    ' Disable the Exit by Trade option for now - 9/26/2001 DAJ
    optTrade.Visible = False
    optPosition.Visible = False
    fraPosition.Top = optPosition.Top
    fraPosition.Left = optPosition.Left
    chkAdvanced.Visible = False
    
    Me.Icon = Picture16(ToolbarIcon("ID_Rules"), , True)
    
    With tbToolbar
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_SaveFavorite").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Verify").Picture = Picture16(ToolbarIcon("kVerify"))
        .Tools("ID_Diagram").Picture = Picture16(ToolbarIcon("kDiagram"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Functions").Picture = Picture16(ToolbarIcon("ID_Functions"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_CondBuilder").Picture = Picture16(ToolbarIcon("ID_ConditionBuilder"))
    End With
    
    Me.Height = vsIndexTab1.Height + vsIndexTab1.Top * 2 + Me.Height - Me.ScaleHeight
    CenterTheForm Me
    
    With cboAction
        .AddItem "Long Entry (BUY)", Action(eRA_LongEntry)
        .AddItem "Long Exit (SELL)", Action(eRA_LongExit)
        .AddItem "Short Entry (SELL)", Action(eRA_ShortEntry)
        .AddItem "Short Exit (BUY)", Action(eRA_ShortExit)
    End With
    
    With cboOrderPlacement
        .AddItem "Market", OrderType(eOT_Market)
        .AddItem "Stop", OrderType(eOT_Stop)
        .AddItem "Limit", OrderType(eOT_Limit)
        .AddItem "Stop with Limit", OrderType(eOT_StopLimit)
        .AddItem "Market on Close", OrderType(eOT_MarketClose)
        .AddItem "Stop Close Only", OrderType(eOT_StopClose)
        .AddItem "Limit Close Only", OrderType(eOT_LimitClose)
        .AddItem "Stop with Limit Close Only", OrderType(eOT_StopLimitClose)
    End With
    
    ' Load the Category combo box...
    LoadCategories
    
    'Assign default control values...
    vsIndexTab1.CurrTab = Tabs(eTab_Rule)
    
    strText = GetIniFileProperty("RuleMgr", "", "Placement", g.strIniFile)
    SetFormPlacement Me, strText, "LTHW"
    If Me.Width < 9540 Then Me.Width = 9540 '(to show picture)
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("RuleMgr", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString vsInputs.Font, strFont
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the 'X', ask if they wish to save
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode <> vbFormCode Then
        Cancel = AskToSave
    End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Size/Move controls as the Form gets resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim bNarrow As Boolean, nOrderWidth&

    If WindowState = vbMinimized Then
        If TypeOf ActiveControl Is Editor Then
            Set g.ActiveEditor = Nothing
            ActiveControl.RemoveTradeSense
        End If
    End If

    If LimitFormSize(Me, 5000, 4200) Then Exit Sub

    If fraOrder.Width < fraCondition.Width Then
        bNarrow = True
    End If
    
    With vsIndexTab1
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Top - .Left
    End With
    With fraCondition
        .Move .Left, .Top, fraRule.Width - .Left * 2, fraRule.Height - .Top - 2460
    End With
    With Editor1
        .Move .Left, .Top, fraCondition.Width - .Left * 2, fraCondition.Height - .Top - 120
    End With
    lblCondition.Width = fraCondition.Width - lblCondition.Left - 60
    vsSystemName.Width = vsIndexTab1.Width - vsSystemName.Left
    
    If bNarrow Then
        nOrderWidth = cboOrderPlacement.Width + cboOrderPlacement.Left * 2
    Else
        nOrderWidth = fraCondition.Width
    End If
    With fraOrder
        .Move .Left, fraRule.Height - 2310, nOrderWidth, .Height
    End With
    Editor2.Width = Editor1.Width - (Editor2.Left - Editor1.Left)
    edLimit.Width = Editor2.Width
    lblStop.Width = Editor2.Width
    lblLimit.Width = lblStop.Width
    picSmall.Top = fraOrder.Top + 90
    picLarge.Top = picSmall.Top
    
    With vsInputs
        .Move .Left, .Top, fraInputs.Width - .Left * 2, fraInputs.Height - .Top - .Left
    End With
    With edAdvanced
        .Move .Left, .Top, fraAdvRule.Width - .Left * 2, fraAdvRule.Height - .Top - .Left
    End With
    
End Sub

Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont vsInputs, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optNumContracts_Click
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optNumContracts_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    txtPercent.Enabled = False
    txtContractsToExit.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.optNumContracts.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optPercent_Click
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optPercent_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    txtContractsToExit.Enabled = False
    txtPercent.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.optPercent.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optPosition_Click
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optPosition_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    optPercent.Enabled = True
    If optPercent Then txtPercent.Enabled = True
    optNumContracts.Enabled = True
    If optNumContracts Then txtContractsToExit.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.optPosition.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTrade_Click
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTrade_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    optPercent.Enabled = False
    txtPercent.Enabled = False
    optNumContracts.Enabled = False
    txtContractsToExit.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.optTrade.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim strText As String, strID As String
    Dim bCloseOnly As Boolean
    Dim bAutoIf As Boolean
    Dim frm As Form
    
    bAutoIf = m.bSkipAutoIf
    m.bSkipAutoIf = True
    ToggleFocus Me, Me.vsIndexTab1
    m.bSkipAutoIf = bAutoIf
    
    Select Case Tool.ID
        Case "ID_Save", "ID_Rename", "ID_SaveAs"
            Save Tool.ID
            
            'Unload Me
        Case "ID_SaveFavorite"
            AddRuleToFavorites m.Rule
        
        Case "ID_Verify"
            Verify
        
        Case "ID_Diagram"
            bAutoIf = m.bSkipAutoIf
            m.bSkipAutoIf = True
            Diagram
            m.bSkipAutoIf = bAutoIf
            
        Case "ID_Functions"
            bAutoIf = m.bSkipAutoIf
            m.bSkipAutoIf = True
            frmFunctionMgrCT.ShowMe 0, Editor1.SelText
            m.bSkipAutoIf = bAutoIf
            
        Case "ID_Print"
            bAutoIf = m.bSkipAutoIf
            m.bSkipAutoIf = True
            PrintMe
            m.bSkipAutoIf = bAutoIf
            
        Case "ID_Toolbox"
            If Not AskToSave Then
                strID = CStr(m.Rule.RuleID)
                Unload Me
                frmToolbox.ShowMe eTab_Rules, strID
            End If
            
        Case "ID_Close"
            If Not AskToSave Then
                Unload Me
            End If
            
        Case "ID_CondBuilder"
                Set frm = ActiveChart
                If Not frm Is Nothing Then
                    If IsFrmChart(frm) Then
                        frmConditionBuilder.ShowMe frm.Chart, , eType_Rule, Me
                    End If
                End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtContractsToEnter_Change
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContractsToEnter_Change()
On Error GoTo ErrSection:

    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.txtContractsToEnter.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtContractsToExit_Change
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContractsToExit_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.txtContractsToExit.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPercent_Click
'' Description: Enable/Disable buttons appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPercent_Change()
On Error GoTo ErrExit:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.txtPercent.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsIndexTab1_Click
'' Description: If there are no inputs, stay on the Rule Tab
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsIndexTab1_Click()
On Error GoTo ErrSection:
    
    If vsInputs.Rows = 0 Then
        vsIndexTab1.CurrTab = Tabs(eTab_Rule)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.vsIndexTab1.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterEdit
'' Description: After the user edits an input, color and format the cell
'' Inputs:      Row and Column of the cell being edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    ColorCell Row, Col
    vsInputs.TextMatrix(Row, Col) = FormatNum(ValOfText(vsInputs.TextMatrix(Row, Col)))
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.vsInputs.AfterEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub vsInputs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    vsInputs.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.vsInputs.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsInputs_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    
    With vsInputs
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            If Button = vbRightButton Then
                .Row = lMouseRow
                PopupMenu mnuPopUp
            End If
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.vsInputs.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub vsInputs_ChangeEdit()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.vsInputs.ChangeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_ValidateEdit
'' Description: Validate what the user entered
'' Inputs:      Row and Column being edited, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Dim InputValue      As Variant
    'Dim FromVal         As Double
    'Dim ToVal           As Double
    Dim X               As Integer
    
    'Get input values
    InputValue = ValOfText(vsInputs.EditText)
    'FromVal = vsInputs.TextMatrix(Row, IGCol(eIGCol_FromVal))
    'ToVal = vsInputs.Cell(flexcpText, Row, IGCol(eIGCol_ToVal))
    
    If IsNumeric(InputValue) Then
        'If FromVal <> 0 Or ToVal <> 0 Then
        '    If InputValue < FromVal Or InputValue > ToVal Then
        '        Cancel = True
        '        Err.Raise vbObjectError + 1000, , _
        '            "Please enter a value between " & _
        '            Format(FromVal, "general number") & " and " & _
        '            Format(ToVal, "general number")
        '    End If
        'Else
            If (InputValue < -100000000000# Or _
                InputValue > 100000000000#) Then
                Cancel = True
                Err.Raise vbObjectError + 1000, , _
                    "Please enter a value between " & _
                    "-100,000,000,000 and 100,000,000,000"
            End If
        'End If
    End If
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.vsInputs.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeEdit
'' Description: Only allow the user to edit the Default Value column
'' Inputs:      Row and Column being edited, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    If Col <> IGCol(eIGCol_DefaultValue) Then Cancel = True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.vsInputs.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearForm
'' Description: Clear the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearForm()
On Error GoTo ErrSection:
    
    m.strName = ""
    
    Editor1.Text = ""
    Editor2.Text = ""
    edAdvanced.Text = ""
    
    chkCanExitOnEntryBar.Value = vbUnchecked
    chkCanExitOnEntryBar.Enabled = False
    cboAction.ListIndex = Action(eRA_LongEntry)
    cboOrderPlacement.ListIndex = OrderType(eOT_Market)
    If cboCategory.ListCount > 0 Then cboCategory.ListIndex = 0
    
    txtContractsToEnter.Text = "1"
    txtContractsToExit.Text = "1"
    txtPercent.Text = "100"
    optTrade.Value = True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.ClearForm", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up after ourselves when the form unloads
'' Inputs:      Whether to Cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Set m.SaveGrid = Nothing
    Set m.Inputs = Nothing
    Set m.Rule = Nothing
    Set m.ListLoading = Nothing
    Set g.ActiveEditor = Nothing

    EnableToolbar False
    tbToolbar.Tools("ID_Verify").Enabled = False
    tbToolbar.Tools("ID_Diagram").Enabled = True
    tbToolbar.Tools("ID_CondBuilder").Enabled = False

    SetIniFileProperty "RuleMgr", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "RuleMgr", FontToString(vsInputs.Font), "Fonts", g.strIniFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildRule
'' Description: Converts english rules from the EZ editor text boxes
''              into the Advanced editor format before TradeSense validation
'' Inputs:      Condition, Buy/Sell,Entry Price, Order Type, Stop Limit
'' Returns:     Advanced editor text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildRule(pCondition As String, pBuySell As Byte, _
    pEntryPrice As String, pOrderType As Byte, pStopLimit As String) As String
On Error GoTo ErrSection:
    
    Dim strRule As String
    Dim lPos As Long
    Dim astrRule As cGdArray
    Dim lIndex As Long
    
    'IF Condition THEN Action...
    strRule = Trim(pCondition)
    
    Set astrRule = New cGdArray
    astrRule.SplitFields strRule, vbLf
    For lIndex = 0 To astrRule.Size - 1
        If InStr(astrRule(lIndex), ":=") = 0 Then
            If UCase(Left(astrRule(lIndex), 3)) <> "IF " Then
                If InStr(astrRule(lIndex), " IF ") = 0 Then
                    If InStr(astrRule(lIndex), "}IF ") = 0 Then
                        astrRule(lIndex) = "IF " & astrRule(lIndex)
                    End If
                End If
            End If
            
            Exit For
        End If
    Next lIndex
    
    strRule = astrRule.JoinFields(vbCrLf)
    
    If pBuySell = True Then
        strRule = strRule & " THEN " & Chr(13) & Chr(10) & Chr(9) & "BUY ("
    Else
        strRule = strRule & " THEN " & Chr(13) & Chr(10) & Chr(9) & "SELL ("
    End If
    
    'Entry Price...
    If Len(pEntryPrice) > 0 Then
        If pOrderType <> OrderType(eOT_Market) And pOrderType <> OrderType(eOT_MarketClose) Then
            strRule = strRule & pEntryPrice & ", "
        End If
    End If
    
    'Order type...
    Select Case pOrderType
        Case OrderType(eOT_Stop): strRule = strRule & """" & "Stop" & """"
        Case OrderType(eOT_Limit): strRule = strRule & """" & "Limit" & """"
        Case OrderType(eOT_Market): strRule = strRule & "Next Bar Open, " & """" & "Market" & """"
        Case OrderType(eOT_MarketClose): strRule = strRule & "Next Bar Close, " & """" & "Market on Close" & """"
        Case OrderType(eOT_StopClose): strRule = strRule & """" & "Stop Close Only" & """"
        Case OrderType(eOT_LimitClose): strRule = strRule & """" & "Limit Close Only" & """"
        Case OrderType(eOT_StopLimit): strRule = strRule & """" & "Stop with Limit" & """"
        Case OrderType(eOT_StopLimitClose): strRule = strRule & """" & "Stop with Limit Close Only" & """"
    End Select
    
    'Limit Price
    If Len(pStopLimit) > 0 Then
        If pOrderType <> OrderType(eOT_Market) And pOrderType <> OrderType(eOT_MarketClose) Then
            strRule = strRule & ", " & pStopLimit
        End If
    End If
    
    BuildRule = strRule & ")" & Chr(13) & Chr(10) & "ENDIF"
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmRule.BuildRule ", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdvancedToEZ
'' Description: Parse Advanced format coded text and return Condition,
''              Order Type, Entry Price and Stop/Limit price
'' Inputs:      Update Editor, Coded Text, Buy/Sell, Order Type, Condition,
''              Entry Price, Stop Limit Price
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AdvancedToEZ(pUpdateEZEditor As Boolean, pCodedText As String, _
    pBuy As Byte, pOrderType As Byte, _
    pCondition As String, pEntryPrice As String, pStopLimitPrice As String)
On Error GoTo ErrSection:
    
    Dim Parens          As Integer
    Dim Commas          As Integer
    Dim X               As Long
    Dim curPos          As Long
    Dim IfPos           As Long
    Dim EnterPos        As Long
    Dim EntryPos        As Long
    Dim CommaPos1       As Long
    Dim CommaPos2       As Long
    Dim RParenPos       As Long
    Dim ParmStart       As Long
    Dim strOrderType    As String
    Dim wrkText         As String
    Dim ThenPos         As Long
    
    'System Navigator tokens
    Const cIf = 24
    Const cThen = 35
    Const cEnter = 80
    Const cFLParen = 16
    Const cFRParen = 17
    Const cComma = 22
    
    'Condition...
    IfPos = InStr(1, pCodedText, "~" & Format(cIf, "00"))
    ThenPos = InStr(IfPos + 1, pCodedText, "~" & Format(cThen, "00"))
    'EnterPos = InStr(IfPos + 1, pCodedText, "~" & Format(cEnter, "00"))
    EnterPos = InStr(ThenPos + 1, pCodedText, "~" & Format(cEnter, "00"))
    If IfPos = 0 Or EnterPos = 0 Or ThenPos = 0 Then Exit Sub
    pCondition = Left(pCodedText, ThenPos - 1)
    
    'Determine if Buy or Sell...
    If InStr(EnterPos + 1, UCase(pCodedText), "~01003BUY") > 0 Then
        pBuy = 1
    Else
        pBuy = 0
    End If
    
    'Entry Price...
    EntryPos = InStr(EnterPos + 1, pCodedText, "~" & Format(cFLParen, "00"))
    If EntryPos = 0 Then
        edAdvanced.TurnOffEditing
        edAdvanced.ExprIsFormatted = True
        wrkText = pCodedText
        edAdvanced.TextRTF = m.Rule.GetRTF(wrkText)
        Exit Sub
    End If
    
    'Next token is the start of Entry Price (following Function LeftParen)
    EntryPos = InStr(EntryPos + 1, pCodedText, "~")
    
    'If EntryPos is Market1 then ignore hidden parm and adjust entry
    'by skipping market and the comma following it
    If Mid(pCodedText, EntryPos + 6, 7) = "Market1" Then
        EntryPos = InStr(EntryPos + 21, pCodedText, "~")
    End If
    
    'Determine position's of comma's in Buy/Sell action...
    Parens = 1
    Commas = 0
    curPos = EntryPos + 1
    Do Until curPos > Len(pCodedText)
        curPos = InStr(curPos, pCodedText, "~")
        If curPos = 0 Then Exit Do
        Select Case Val(Mid(pCodedText, curPos + 1, 2))
            Case cFLParen: Parens = Parens + 1
            Case cFRParen
                Parens = Parens - 1
                If Parens = 0 Then
                    RParenPos = curPos
                End If
            
            Case cComma
                If Parens = 1 Then
                    Commas = Commas + 1
                    If Commas = 1 Then
                        CommaPos1 = curPos
                    Else
                        CommaPos2 = curPos
                    End If
                End If
        End Select
        curPos = curPos + 1
    Loop
    
    'Entry Price...
    pEntryPrice = Mid(pCodedText, EntryPos, CommaPos1 - EntryPos)
    
    'Order Type...
    ParmStart = InStr(CommaPos1 + 1, pCodedText, "~")
    If CommaPos2 > 0 Then
        strOrderType = Mid(pCodedText, ParmStart, _
            CommaPos2 - ParmStart)
    Else
        strOrderType = Mid(pCodedText, ParmStart, _
            RParenPos - ParmStart)
    End If
    strOrderType = Mid(strOrderType, 8, Val(Mid(strOrderType, 4, 3)) - 2)
    Select Case strOrderType
        Case "Market": pOrderType = OrderType(eOT_Market)
        Case "Stop": pOrderType = OrderType(eOT_Stop)
        Case "Limit": pOrderType = OrderType(eOT_Limit)
        Case "Stop with Limit": pOrderType = OrderType(eOT_StopLimit)
        Case "Market on Close": pOrderType = OrderType(eOT_MarketClose)
        Case "Stop Close Only": pOrderType = OrderType(eOT_StopClose)
        Case "Limit Close Only": pOrderType = OrderType(eOT_LimitClose)
        Case "Stop with Limit Close Only": pOrderType = OrderType(eOT_StopLimitClose)
        Case Else
            pOrderType = OrderType(eOT_Market)
    End Select
    
    'Stop/Limit price
    ParmStart = InStr(CommaPos2 + 1, pCodedText, "~")
    If CommaPos2 > 0 Then
        pStopLimitPrice = Mid(pCodedText, ParmStart, _
            RParenPos - ParmStart)
    End If
    
    '---------------------------------------------------------------------
    If pUpdateEZEditor Then
        
        'Move parsed pieces of Rule back into editor text boxes...
        Editor1.TurnOffEditing
        wrkText = pCondition
        Editor1.TextRTF = m.Rule.GetRTF(wrkText)
        Editor1.ExprIsFormatted = True
        If pEntryPrice <> "" Then
            Editor2.TurnOffEditing
            wrkText = pEntryPrice
            Editor2.TextRTF = m.Rule.GetRTF(wrkText)
            Editor2.ExprIsFormatted = True
        Else
            Editor2.Text = ""
        End If
        If pStopLimitPrice <> "" Then
            edLimit.TurnOffEditing
            wrkText = pStopLimitPrice
            edLimit.TextRTF = m.Rule.GetRTF(wrkText)
            edLimit.ExprIsFormatted = True
        Else
            edLimit.Text = ""
        End If
        
        'Move complete rule into advanced editor text box...
        edAdvanced.TurnOffEditing      'Shut off tradesense validation
        edAdvanced.ExprIsFormatted = True
        wrkText = pCodedText
        edAdvanced.TextRTF = m.Rule.GetRTF(wrkText)
        
        'Set Order type and placement...
        cboOrderPlacement.ListIndex = pOrderType
        DisplayOrderControls
        'If pBuy Then opBuy.value = True Else opSell.value = 1
        If pBuy Then
            If cboAction.ListIndex <> Action(eRA_LongEntry) And cboAction.ListIndex <> Action(eRA_ShortExit) Then
                cboAction.ListIndex = Action(eRA_LongEntry)
            End If
        Else
            If cboAction.ListIndex <> Action(eRA_LongExit) And cboAction.ListIndex <> Action(eRA_ShortEntry) Then
                cboAction.ListIndex = Action(eRA_LongExit)
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.AdvancedToEZ", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Verify
'' Description: Verify the rule text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Verify()
On Error GoTo ErrSection:
    
    Dim svErr           As Long
    Dim svErrDesc       As String
    Dim Expr            As cExpression
    Dim svSource        As String
    Dim Condition       As String
    Dim EntryPrice      As String
    Dim StopLimitPrice  As String
    Dim CondLate        As Boolean
    Dim ActionLate      As Boolean
    Dim OrderTypeInd    As Byte
    Dim Buy             As Byte
    
    Dim bPosition       As Boolean
    Dim bPercent        As Boolean
    Dim strNum1         As String
    Dim strNum2         As String
    Dim strNum3         As String
    
    Dim lIndex As Long                  ' Index into a for loop
    
    If Trim(Editor1.Text) = "" Or Trim(UCase(Editor1.Text)) = "IF" Then
        ' TLB 12/7/2015: as long as it's not a Market order, let's just make it easier for them
        ' and replace a blank condition with "IF True" (for cases when a condition is not needed)
        If cboOrderPlacement.ListIndex <> 0 Then
            Editor1.Text = "IF True"
        Else
            Err.Raise vbObjectError + 1000, , "You must enter a condition"
        End If
    End If
    
    Editor1.Text = FixPeriodInMarkets(Editor1.Text)
    
    'Save current input values entered by user...
    SaveGridValues
    
    ' Save the pyramiding values
    bPosition = optPosition.Value
    bPercent = optPercent.Value
    strNum1 = txtContractsToEnter.Text
    strNum2 = txtContractsToExit.Text
    strNum3 = txtPercent.Text
    
    'Make sure an order price is entered for Systems if the price field
    'is visible (Stop, Stop/Limit, Limit)
    If Editor2.Visible Then
        If Len(Editor2.Text) = 0 Then
            'MoveFocus Editor2
            Err.Raise vbObjectError + 1000, , "You must supply an order price"
        End If
    End If
    If edLimit.Visible Then
        If Len(edLimit.Text) = 0 Then
            'MoveFocus edLimit
            Err.Raise vbObjectError + 1000, , "You must supply an order price"
        End If
    End If
    
    'Shut things off, get ready for verifying rule
    Screen.MousePointer = vbHourglass
    m.lReturnValue = LockWindowUpdate(Me.hWnd)
    edAdvanced.TurnOffEditing
    
    'Verify...
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        
        'Advanced editor
        If chkAdvanced.Value = vbChecked Then
            .ValidateRule edAdvanced.Text
        Else
            'EZ editor
            Buy = cboAction.ListIndex = Action(eRA_LongEntry) Or cboAction.ListIndex = Action(eRA_ShortExit)
            .ValidateRule BuildRule(Editor1.Text, Buy, Editor2.Text, cboOrderPlacement.ListIndex, edLimit.Text)
        End If
        
        ' If these prices are applicable, make sure that neither the Stop price nor the Limit
        ' price are boolean expressions...
        If Editor2.Visible Then
            If IsBooleanExpression(Editor2.Text) Then
                Err.Raise vbObjectError + 1000, , "Stop Price cannot result in a True/False value"
            End If
        End If
        If edLimit.Visible Then
            If IsBooleanExpression(edLimit.Text) Then
                Err.Raise vbObjectError + 1000, , "Limit Price cannot result in a True/False value"
            End If
        End If
        
        ' Verify any "Symbol,Period" market types...
        If Not .Inputs Is Nothing Then
            For lIndex = 1 To .Inputs.Count
                If IsValidMarket(.Inputs.Item(lIndex).ParmName) = False Then
                    Err.Raise vbObjectError + 1000, , "No data can be loaded for " & .Inputs.Item(lIndex).ParmName
                End If
            Next lIndex
        End If
        
        'Save Late calculating flags...
        m.Rule.LateCondition = .LateCondition
        m.Rule.LateAction = .LateAction
        
        'Parse verified Rule back into text boxes...
        AdvancedToEZ True, .EditText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
        m.Rule.Cond = Condition
        m.Rule.Price = EntryPrice
        m.Rule.Price2RTF = StopLimitPrice
        m.Rule.OrderPlacement = cboOrderPlacement.Text 'OrderTypeInd
        
        'Parse Codedtext used by engine and save appropriate pieces...
        AdvancedToEZ False, .CodedText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
        m.Rule.CondCoded = Condition
        If EntryPrice = "" Then
            m.Rule.PriceCoded = "N/A"
        Else
            m.Rule.PriceCoded = EntryPrice
        End If
        If StopLimitPrice = "" Then
            m.Rule.Price2Coded = "N/A"
        Else
            m.Rule.Price2Coded = StopLimitPrice
        End If
        
        '.GetFIDs=Returns the function ID's used in the entire rule...
        m.Rule.CondRefs = .GetFIDs
    
    End With
    
    'Load the inputs from condition and Actions into grid...Restore
    'any changes input values that already exist in the grid
    Set m.Inputs = Expr.Inputs
    LoadInputsGrid
    RestoreGridValues
    
    ' Make sure that the rule does not refrence any function that is not meant
    ' for strategy testing...
    If VerifyUsage(m.Rule.CondRefs) = False Then GoTo ErrExit
    
    ' Restore the pyramiding values
    optPosition.Value = bPosition
    optPercent.Value = bPercent
    txtContractsToEnter.Text = strNum1
    txtContractsToExit.Text = strNum2
    txtPercent.Text = strNum3
    
    m.Rule.Reverify = False
    tbToolbar.Tools("ID_Verify").Enabled = False
    tbToolbar.Tools("ID_Diagram").Enabled = True
    
    Screen.MousePointer = vbDefault
    m.lReturnValue = LockWindowUpdate(0)
    
ErrExit:
    Editor1.TurnOnEditing
    If Editor2.Visible Then Editor2.TurnOnEditing
    If edLimit.Visible Then edLimit.TurnOnEditing
    Set Expr = Nothing
    Exit Sub

ErrSection:
    Screen.MousePointer = vbDefault
    m.lReturnValue = LockWindowUpdate(0)
    
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        svErr = Err.Number
        svSource = Err.Source
        svErrDesc = Err.Description
        
        'Convert codedtext into richtext
        If Not Expr Is Nothing Then
            If Expr.EditText <> "" Then
                edAdvanced.TurnOffEditing
                edAdvanced.ExprIsFormatted = True
                edAdvanced.TextRTF = m.Rule.GetRTF(Expr.EditText)
                AdvancedToEZ True, Expr.EditText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
            End If
        End If
        Set Expr = Nothing
        Err.Raise svErr, svSource, svErrDesc
    Else
        Set Expr = Nothing
        RaiseError "frmRule.Verify ", eGDRaiseError_Raise
    End If
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyRuleDebug
'' Description: Verify the rule text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyRuleDebug()
On Error GoTo ErrSection:
    
    Dim svErr           As Long
    Dim svErrDesc       As String
    Dim Expr            As cExpression
    Dim svSource        As String
    Dim Condition       As String
    Dim EntryPrice      As String
    Dim StopLimitPrice  As String
    Dim CondLate        As Boolean
    Dim ActionLate      As Boolean
    Dim OrderTypeInd    As Byte
    Dim Buy             As Byte
    
    'Save current input values entered by user...
    SaveGridValues
    
    'Make sure an order price is entered for Systems if the price field
    'is visible (Stop, Stop/Limit, Limit)
    If Editor2.Visible Then
        If Len(Editor2.Text) = 0 Then
            Editor2.SetFocus
            Err.Raise vbObjectError + 1000, "You must supply an order price"
        End If
    End If
    
    'Shut things off, get ready for verifying rule
    Screen.MousePointer = vbHourglass
    m.lReturnValue = LockWindowUpdate(Me.hWnd)
    edAdvanced.TurnOffEditing
    
    'Verify...
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        
        'Advanced editor
        If chkAdvanced.Value = 1 Then
            .ValidateRule edAdvanced.Text
        Else
            'EZ editor
            Buy = cboAction.ListIndex = Action(eRA_LongEntry) Or cboAction.ListIndex = Action(eRA_ShortExit)
            .ValidateRule BuildRule(Editor1.Text, Buy, Editor2.Text, _
                cboOrderPlacement.ListIndex, edLimit.Text)
        End If
        
        'Save Late calculating flags...
        m.Rule.LateCondition = .LateCondition
        m.Rule.LateAction = .LateAction
        
        'Parse verified Rule back into text boxes...
        AdvancedToEZ True, .EditText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
        m.Rule.Cond = Condition
        m.Rule.Price = EntryPrice
        m.Rule.Price2RTF = StopLimitPrice
        m.Rule.OrderPlacement = cboOrderPlacement.Text 'OrderTypeInd
        
        'Parse Codedtext used by engine and save appropriate pieces...
        AdvancedToEZ False, .CodedText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
        m.Rule.CondCoded = Condition
        If EntryPrice = "" Then
            m.Rule.PriceCoded = "N/A"
        Else
            m.Rule.PriceCoded = EntryPrice
        End If
        If StopLimitPrice = "" Then
            m.Rule.Price2Coded = "N/A"
        Else
            m.Rule.Price2Coded = StopLimitPrice
        End If
        
        '.GetFIDs=Returns the function ID's used in the entire rule...
        m.Rule.CondRefs = .GetFIDs
    
    End With
    
    'Load the inputs from condition and Actions into grid...Restore
    'any changes input values that already exist in the grid
    Set m.Inputs = Expr.Inputs
    LoadInputsGrid
    RestoreGridValues
    
    tbToolbar.Tools("ID_Verify").Enabled = False
    tbToolbar.Tools("ID_Diagram").Enabled = True
    
    m.lReturnValue = LockWindowUpdate(0)
    
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
    
ErrExit:
    Set Expr = Nothing
    Exit Sub

'=================================================
ErrSection:
    Screen.MousePointer = vbDefault
    m.lReturnValue = LockWindowUpdate(0)
    
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        svErr = Err.Number
        svSource = Err.Source
        svErrDesc = Err.Description
        
        'Convert codedtext into richtext
        edAdvanced.TurnOffEditing
        edAdvanced.ExprIsFormatted = True
        edAdvanced.TextRTF = m.Rule.GetRTF(Expr.EditText)
        
        MsgBox svErrDesc & Chr(13) & Chr(10), vbInformation, "Message"
        Resume ShowTheTree:
    Else
        Set Expr = Nothing
        RaiseError "frmRule.VerifyRuleDebug", eGDRaiseError_Raise
    End If
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveGridValues
'' Description: Save off the grid values from the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveGridValues()
On Error GoTo ErrSection:
    
    Dim X       As Integer
    Dim Y       As Integer
    Dim lParmTypeID As Long             ' Parm Type ID
    
    Set m.SaveGrid = New cInputs
    
    'Search for potential new inputs...
    With vsInputs
        For X = 1 To .Rows - 1
            lParmTypeID = CLng(Val(.TextMatrix(X, IGCol(eIGCol_ParmTypeID))))
            
            m.SaveGrid.Add "", 0, .TextMatrix(X, IGCol(eIGCol_InputName)), "", _
                .TextMatrix(X, IGCol(eIGCol_ParmID)), _
                ConvertInputValue(.TextMatrix(X, IGCol(eIGCol_DefaultValue)), lParmTypeID), _
                0, 0, 0, 0, 0, .TextMatrix(X, IGCol(eIGCol_RuleID)), 0, 0, True, 0, 0, 0, "", ""
        Next X
    End With
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.SaveGridValues", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RestoreGridValues
'' Description: Restore the inputs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RestoreGridValues()
On Error GoTo ErrSection:
    
    Dim lIndex As Integer               ' Index into a for loop
    Dim lIndex2 As Integer              ' Index into a for loop
    Dim Parm As cInput                  ' Input to pull information from
    Dim bFound As Boolean               ' Was the input found in the grid?
    Dim strNewInputs As String          ' New Inputs to display to the user
    Dim strMsg As String                ' Message to display to the user
    
    ' Search for potential new inputs...
    With vsInputs
        For lIndex = 1 To .Rows - 1
            bFound = False
        
            ' Search for input in system inputs by name
            For lIndex2 = 1 To m.SaveGrid.Count
                Set Parm = m.SaveGrid.Item(lIndex2)
                If .Cell(flexcpText, lIndex, IGCol(eIGCol_InputName)) = Parm.ParmName Then
                    .TextMatrix(lIndex, IGCol(eIGCol_ParmID)) = Parm.ParmID
                    .TextMatrix(lIndex, IGCol(eIGCol_RuleID)) = m.Rule.RuleID
                    .TextMatrix(lIndex, IGCol(eIGCol_DefaultValue)) = Parm.Value
                    bFound = True
                    Exit For
                End If
            Next lIndex2
        
            If Not bFound Then
                'TLB 1/22/01: don't ask about markets being added
                If ValOfText(.TextMatrix(lIndex, IGCol(eIGCol_ParmTypeID))) <> 5 Then
                    strNewInputs = strNewInputs & .TextMatrix(lIndex, IGCol(eIGCol_InputName)) & Chr(13) & Chr(10)
                End If
            End If
            
        Next lIndex
        .AutoSize IGCol(eIGCol_InputName)
    End With
    
    If Len(strNewInputs) > 0 Then
        'ask about new inputs
        strMsg = "Unrecognized as existing functions or inputs:|" _
                & strNewInputs & "|Add the above as new INPUTS to this rule?|"
        If AskBox("i=? ; b=+Add|-Cancel ; h=Add Inputs ; " & strMsg) = "C" Then
            'consider it "unverified"
            Err.Raise vbObjectError + 1999, , "Need to fix unrecognized functions."
        Else
            vsIndexTab1.CurrTab = Tabs(eTab_Inputs)
        End If
    End If

ErrExit:
    Set Parm = Nothing
    Set m.SaveGrid = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmRule.RestoreGridValues", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInputsCollection
'' Description: Load the inputs collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadInputsCollection()
On Error GoTo ErrSection:
    
    Dim X       As Integer
    Dim Y       As Integer
    Dim lParmTypeID As Long             ' Parm Type ID
    
    Set m.Inputs = New cInputs
    
    'Search for potential new inputs...
    With vsInputs
        For X = 1 To .Rows - 1
            lParmTypeID = CLng(Val(.TextMatrix(X, IGCol(eIGCol_ParmTypeID))))
            
            m.Inputs.Add "", X, .TextMatrix(X, IGCol(eIGCol_InputName)), _
                .TextMatrix(X, IGCol(eIGCol_ParmDesc)), _
                .TextMatrix(X, IGCol(eIGCol_ParmID)), "", 0, 0, 0, 0, 0, _
                .Cell(flexcpValue, X, IGCol(eIGCol_RuleID)), lParmTypeID, _
                ConvertInputValue(.TextMatrix(X, IGCol(eIGCol_DefaultValue)), lParmTypeID), True, _
                ConvertInputValue(.Cell(flexcpValue, X, IGCol(eIGCol_FromVal)), lParmTypeID), _
                ConvertInputValue(.Cell(flexcpValue, X, IGCol(eIGCol_ToVal)), lParmTypeID), _
                .Cell(flexcpValue, X, IGCol(eIGCol_ListID)), "", ""
        Next X
    End With
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.LoadInputsCollection", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleExists
'' Description: Does the rule name already exist in the database
'' Inputs:      Rule Name
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RuleExists(ByVal strRuleName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    If Not m.frmSysMgr Is Nothing Then
        ' If we were called from the system manager, then we have a local rule
        ' and we need to see if this rule name exists in the current system...
        RuleExists = m.frmSysMgr.RuleNameExists(strRuleName)
        
    ElseIf m.Rule.SystemNumber = 0 Then
        ' If a rule is a "favorite", then we need to check to see if there are
        ' any favorites with that name...
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [Name]='" & strRuleName & "' AND [SystemNumber]=0;", dbOpenSnapshot)
        RuleExists = Not (rs.EOF And rs.BOF)
    
    Else
        ' If the rule is a local rule and we were called from the Toolbox, then
        ' we need to search the database to make sure this rule is unique in the
        ' system...
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [Name]='" & strRuleName & "' AND [SystemNumber]=" & m.lSystemNumber & ";", dbOpenSnapshot)
        RuleExists = Not (rs.EOF And rs.BOF)
    
    End If
        
ErrExit:
    Set rs = Nothing
    Exit Function

ErrSection:
    Set rs = Nothing
    RaiseError "frmRule.RuleExists", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayOrderControls
'' Description: Display the order controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayOrderControls()
On Error GoTo ErrSection:
    
    Dim bNarrow As Boolean
    
    Select Case cboAction.ListIndex
        Case Action(eRA_LongEntry), Action(eRA_ShortEntry)
            chkCanExitOnEntryBar.Enabled = False
            
            lblContractsToEnter.Enabled = True
            txtContractsToEnter.Enabled = True
            lblBasedOn.Enabled = False
            optPosition.Enabled = False
            optTrade.Enabled = False
            optNumContracts.Enabled = False
            txtContractsToExit.Enabled = False
            optPercent.Enabled = False
            txtPercent.Enabled = False
        Case Action(eRA_LongExit), Action(eRA_ShortExit)
            chkCanExitOnEntryBar.Enabled = True
            
            lblContractsToEnter.Enabled = False
            txtContractsToEnter.Enabled = False
            lblBasedOn.Enabled = True
            optPosition.Enabled = True
            optTrade.Enabled = True
            optPosition.Value = True
            optNumContracts.Value = True
            If optPosition Then optPosition_Click Else optTrade_Click
    End Select
    
    Select Case cboOrderPlacement.ListIndex
    
        Case OrderType(eOT_Stop)
            lblStop.Visible = True
            Editor2.Visible = True
            Editor2.Height = 1545
            lblLimit.Visible = False
            edLimit.Visible = False
            
            'If opBuy = True Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                    lblStop.Caption = "Stop Price: Buy if price gets up to..."
                Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                    lblStop.Caption = "Stop Price: Sell if price gets down to..."
            End Select
            
        Case OrderType(eOT_Limit)
            lblStop.Visible = True
            Editor2.Visible = True
            Editor2.Height = 1545
            lblLimit.Visible = False
            edLimit.Visible = False
            
            'If opBuy = True Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                    lblStop.Caption = "Limit Price: Buy if price gets down to..."
                Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                    lblStop.Caption = "Limit Price: Sell if price gets up to..."
            End Select
            
        Case OrderType(eOT_Market)
            bNarrow = True
            lblStop.Visible = False
            Editor2.Visible = False
            Editor2.Height = 1545
            lblLimit.Visible = False
            edLimit.Visible = False
            
        Case OrderType(eOT_MarketClose)
            bNarrow = True
            lblStop.Visible = False
            Editor2.Visible = False
            Editor2.Height = 1545
            lblLimit.Visible = False
            edLimit.Visible = False
            
        Case OrderType(eOT_StopClose)
            lblStop.Visible = True
            Editor2.Visible = True
            Editor2.Height = 1545
            lblLimit.Visible = False
            edLimit.Visible = False
            
            'If opBuy = True Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                    lblStop.Caption = "Stop Price:  Buy if closing price is at or above ..."
                Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                    lblStop.Caption = "Stop Price:  Sell if closing price is at or below ..."
            End Select
            
        Case OrderType(eOT_LimitClose)
            lblStop.Visible = True
            Editor2.Visible = True
            Editor2.Height = 1545
            lblLimit.Visible = False
            edLimit.Visible = False
            
            'If opBuy = True Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                    lblStop.Caption = "Limit Price:  Buy if closing price is at or below ..."
                Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                    lblStop.Caption = "Limit Price:  Sell if closing price is at or above ..."
            End Select
            
        Case OrderType(eOT_StopLimit)
            lblStop.Visible = True
            Editor2.Visible = True
            Editor2.Height = 615 '555
            lblLimit.Visible = True
            edLimit.Visible = True
            
            'If opBuy = True Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                    lblStop.Caption = "Stop Price:  Buy if price is at or above ..."
                    lblLimit.Caption = "Limit Price:  and if price is at or below ..."
                Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                    lblStop.Caption = "Stop Price:  Sell if price is at or below ..."
                    lblLimit.Caption = "Limit Price:  and if price is at or above ..."
            End Select
            
        Case OrderType(eOT_StopLimitClose)
            lblStop.Visible = True
            Editor2.Visible = True
            Editor2.Height = 615 '555
            lblLimit.Visible = True
            edLimit.Visible = True
            
            'If opBuy = True Then
            Select Case cboAction.ListIndex
                Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                    lblStop.Caption = "Stop Price:  Buy if closing price is at or above ..."
                    lblLimit.Caption = "Limit Price:  and if closing price is at or below ..."
                Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                    lblStop.Caption = "Stop Price:  Sell if closing price is at or below ..."
                    lblLimit.Caption = "Limit Price:  and if closing price is at or above ..."
            End Select
            
    End Select
    
    If bNarrow Then
        fraOrder.Width = cboOrderPlacement.Width + cboOrderPlacement.Left * 2
        If Screen.TwipsPerPixelY <= 12 Then
            picLarge.Visible = True
        Else
            picSmall.Visible = True
        End If
    Else
        fraOrder.Width = fraCondition.Width
        picSmall.Visible = False
        picLarge.Visible = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.DisplayOrderControls", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrSave
'' Description: Ask the user if they wish to Save the rule on exit
'' Inputs:      None
'' Returns:     True if Cancelled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelOrSave() As Boolean
On Error GoTo ErrSection:

    Dim strAnswer As String
    Dim bCancelled As Boolean
    Dim bSkipAutoIf As Boolean
    
    ' Save the current SkipAutoIf and turn it off for now so that the Form
    ' Activate won't bring TradeSense back up
    bSkipAutoIf = m.bSkipAutoIf
    m.bSkipAutoIf = True
    
    strAnswer = InfBox("Do you wish to save your changes?", "?", "+Yes|No|-Cancel", "Confirmation")
    If strAnswer = "Y" Then
        Save "ID_Save"
    ElseIf strAnswer = "C" Then
        bCancelled = True
    End If
    
    ' Turn the SkipAutoIf back to its old state
    m.bSkipAutoIf = bSkipAutoIf

    CancelOrSave = bCancelled
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.CancelOrSave", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Diagram
'' Description: Diagram the rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Diagram()
On Error GoTo ErrSection:

    Dim strRule As String               ' Rule to send to the diagram
    Dim astrCondition As cGdArray       ' Array of the condition elements
    Dim astrMacros As cGdArray          ' Array of the macros
    Dim lIndex As Long                  ' Index into a for loop
    Dim strCondition As String          ' Condition
    Dim strVariable As String           ' Variable name
    Dim strExpression As String         ' Expression for variable
    Dim lPos As Long                    ' Position in an array
    
    Set astrCondition = New cGdArray
    Set astrMacros = New cGdArray
    
    strCondition = ""
    astrCondition.SplitFields Editor1.Text, vbLf
    For lIndex = 0 To astrCondition.Size - 1
        If InStr(astrCondition(lIndex), ":=") <> 0 Then
            strVariable = Parse(astrCondition(lIndex), ":=", 1)
            strExpression = Parse(astrCondition(lIndex), ":=", 2)
            If astrMacros.BinarySearch(strVariable & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                astrMacros.Add strVariable & vbTab & strExpression, lPos
            End If
        Else
            strCondition = strCondition & astrCondition(lIndex)
        End If
    Next lIndex
    
    strRule = strCondition & " THEN "
    
    Select Case cboAction.ListIndex
        Case Action(eRA_LongExit), Action(eRA_ShortEntry)
            strRule = strRule & "SELL at "
        Case Action(eRA_LongEntry), Action(eRA_ShortExit)
            strRule = strRule & "BUY at "
    End Select
    
    If Left(cboOrderPlacement, 6) = "Market" Then
        strRule = strRule & UCase(cboOrderPlacement)
    ElseIf Left(cboOrderPlacement, 15) = "Stop with Limit" Then
        strRule = strRule & Editor2.Text & " STOP, " & edLimit.Text & " LIMIT"
        If Right(cboOrderPlacement, 10) = "Close Only" Then
            strRule = strRule & " Close Only"
        End If
    Else
        strRule = strRule & Editor2.Text & " " & UCase(cboOrderPlacement)
    End If
    
    For lIndex = astrMacros.Size - 1 To 0 Step -1
        strVariable = Parse(astrMacros(lIndex), vbTab, 1)
        strExpression = "(" & Parse(astrMacros(lIndex), vbTab, 2) & ")"
        
        strRule = Replace(strRule, strVariable, strExpression)
    Next lIndex

    frmDgm.ShowMe strRule, m.Rule.Name
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.Diagram", eGDRaiseError_Raise

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
    
    Dim lRedraw As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .AllowBigSelection = False
        .AllowSelection = False
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .TabBehavior = flexTabCells
        .BackColorAlternate = g.nAltGridRowColor
        
        .Editable = True
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTips = True
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = IGCol(eIGCol_NumCols) + 1
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1
        
        .ColHidden(IGCol(eIGCol_ParmTypeID)) = True
        .ColHidden(IGCol(eIGCol_ParmID)) = True
        .ColHidden(IGCol(eIGCol_FromVal)) = True
        .ColHidden(IGCol(eIGCol_ToVal)) = True
        .ColHidden(IGCol(eIGCol_ParmDesc)) = True
        .ColHidden(IGCol(eIGCol_Sort)) = True
        .ColHidden(IGCol(eIGCol_Req)) = True
        .ColHidden(IGCol(eIGCol_ListID)) = True
        .ColHidden(IGCol(eIGCol_RuleID)) = True
        
        .TextMatrix(0, IGCol(eIGCol_InputName)) = "Input"
        .TextMatrix(0, IGCol(eIGCol_DefaultValue)) = "Default Value"
        
        .ColAlignment(IGCol(eIGCol_InputName)) = flexAlignLeftCenter
        .ColAlignment(IGCol(eIGCol_DefaultValue)) = flexAlignLeftCenter
        
        .ColDataType(IGCol(eIGCol_Req)) = flexDTBoolean
        
        SetColumnWidths
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.InitInputsGrid", eGDRaiseError_Raise
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadInputsGrid
'' Description: Load the inputs grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadInputsGrid()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim lRedraw As Long

    'Clear and reinitialize grid...
    InitInputsGrid
    
    'Leave if no inputs exist in collection
    If m.Inputs Is Nothing Then Exit Sub
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = m.Inputs.Count + 1
        
        For X = 1 To m.Inputs.Count
            ' For now, make sure that rule input types cannot be arrays
            ' 10/5/2001 - DAJ
            If m.Inputs.Item(X).ParmTypeID = 3 Then m.Inputs.Item(X).ParmTypeID = 6
            If m.Inputs.Item(X).ParmTypeID = 4 Then m.Inputs.Item(X).ParmTypeID = 1
            
            AddRowToGrid X, m.Inputs.Item(X)
            
            ' Make sure that Market1 floats to the top (DAJ: 5/31/2002)
            If .TextMatrix(X, IGCol(eIGCol_InputName)) = "Market1" And X <> 1 Then
                .RowPosition(X) = 1
            End If
        Next X
        
        If m.Inputs.Count >= 1 Then
            .Select 1, IGCol(eIGCol_Sort), .Rows - 1
            .AutoSize 0, .Cols - 1
        End If
        
        SetColumnWidths
        .Redraw = lRedraw
    End With
   
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.LoadInputsGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRowToGrid
'' Description: Add an input to the grid
'' Inputs:      Row to fill, Input to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddRowToGrid(ByVal pRow As Long, pInput As Object)
On Error GoTo ErrSection:
    
    Dim lRedraw As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .TextMatrix(pRow, IGCol(eIGCol_InputName)) = pInput.ParmName
        If pInput.ParmName = "Market1" Then
            .TextMatrix(pRow, IGCol(eIGCol_Sort)) = " "
        Else
            .TextMatrix(pRow, IGCol(eIGCol_Sort)) = pInput.ParmName
        End If
        .Cell(flexcpFontBold, pRow, IGCol(eIGCol_InputName)) = True
        .Cell(flexcpForeColor, pRow, IGCol(eIGCol_InputName)) = vbBlack
        .TextMatrix(pRow, IGCol(eIGCol_ParmTypeID)) = pInput.ParmTypeID
        .TextMatrix(pRow, IGCol(eIGCol_ParmID)) = pInput.ParmID
        .TextMatrix(pRow, IGCol(eIGCol_RuleID)) = m.Rule.RuleID '= pInput.RuleID
        .TextMatrix(pRow, IGCol(eIGCol_ParmDesc)) = pInput.ParmDesc
        .TextMatrix(pRow, IGCol(eIGCol_FromVal)) = ""
        .TextMatrix(pRow, IGCol(eIGCol_ToVal)) = ""
        .TextMatrix(pRow, IGCol(eIGCol_DefaultValue)) = ""
        .TextMatrix(pRow, IGCol(eIGCol_Req)) = pInput.Required
        .TextMatrix(pRow, IGCol(eIGCol_ListID)) = pInput.ListID
            
        'Set the value (or default if one doesn't exist).  The bars and
        'trades type structure is always "Market1" and "Trades"
        Select Case pInput.ParmTypeID
        
            Case kSN_RetBars
                .TextMatrix(pRow, IGCol(eIGCol_DefaultValue)) = pInput.ParmName
        
            Case kSN_RetTrades
                .TextMatrix(pRow, IGCol(eIGCol_DefaultValue)) = pInput.ParmName
                
            Case kSN_RetNumericConstant
                If pInput.Value = "" Then
                    .TextMatrix(pRow, IGCol(eIGCol_DefaultValue)) = FormatNum(Val(pInput.DefaultValue))
                Else
                    .TextMatrix(pRow, IGCol(eIGCol_DefaultValue)) = FormatNum(Val(pInput.Value))
                End If
                ColorCell pRow, IGCol(eIGCol_DefaultValue)
                
                '.TextMatrix(pRow, IGCol(eIGCol_FromVal) = FormatNum(Val(pInput.FromValue))
                '.TextMatrix(pRow, IGCol(eIGCol_ToVal) = FormatNum(Val(pInput.ToValue))
                'ColorCell pRow, IGCol(eIGCol_FromVal)
                'ColorCell pRow, IGCol(eIGCol_ToVal)
                
            Case Else
        End Select
        
        If pInput.ParmTypeID = kSN_RetBars Then .RowHidden(pRow) = True
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmRule.AddRowToGrid", eGDRaiseError_Raise
    Resume ErrExit:

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorCell
'' Description: Color the given cell appropriately
'' Inputs:      Row and Column of cell to color
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorCell(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    If ValOfText(vsInputs.TextMatrix(Row, Col)) < 0 Then
        vsInputs.Cell(flexcpForeColor, Row, Col) = vbRed
    Else
        vsInputs.Cell(flexcpForeColor, Row, Col) = vbBlack
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.ColorCell", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetColumnWidths
'' Description: Set the column widths on the inputs grid appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetColumnWidths()
On Error GoTo ErrSection:

    Dim lRedraw As Long
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        If .ColWidth(IGCol(eIGCol_InputName)) > 4500 Then .ColWidth(IGCol(eIGCol_InputName)) = 4500
        If .ColWidth(IGCol(eIGCol_DefaultValue)) < 1500 Then .ColWidth(IGCol(eIGCol_DefaultValue)) = 1500
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.SetColumnWidths", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewParms
'' Description: Compiles a list of new inputs
'' Inputs:      None
'' Returns:     List of new inputs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewParms() As String
On Error GoTo ErrSection:
    
    Dim X       As Long
    Dim Y       As Long
    Dim Fnd     As Boolean
    Dim Parms   As String
    
    Parms = ""
    With vsInputs
        For Y = 1 To .Rows - 1
            
            'Search grid containing new inputs and look for a matching
            'inputs from the saved array.
            Fnd = False
            For X = 1 To m.Inputs.Count
                If m.Inputs.Item(X).ParmName = .TextMatrix(Y, IGCol(eIGCol_InputName)) Then
                    Fnd = True
                    Exit For
                End If
            Next X
            
            'Save value found, restore the value from the saved array.
            If Not Fnd Then
                If .TextMatrix(Y, IGCol(eIGCol_ParmTypeID)) <> kSN_RetBars Then
                    Parms = Parms & _
                        .TextMatrix(Y, IGCol(eIGCol_InputName)) & Chr(13) & Chr(10)
                End If
            End If
        Next Y
    End With
    
    NewParms = Parms

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmRule.NewParms", eGDRaiseError_Raise
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the print preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lIndex As Long, lTemp As Long
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .FontUnderline = True
        .Text = vbLf & "Rule:"
        .FontUnderline = False
        .Text = "    " & Trim(m.strName) & vbLf
        .FontUnderline = False
        .Font.Bold = False
        .Font.Size = 12
        
        .Text = vbLf & Editor1.Text & vbLf
        .Text = "THEN "
        'If opBuy.value = True Then .Text = "BUY at " Else .Text = "SELL at "
        Select Case cboAction.ListIndex
            Case Action(eRA_LongEntry), Action(eRA_ShortExit)
                .Text = "BUY at "
            Case Action(eRA_LongExit), Action(eRA_ShortEntry)
                .Text = "SELL at "
        End Select
        
        Select Case cboOrderPlacement.Text
            Case "Market", "Market on close"
                .Text = cboOrderPlacement.Text & vbLf
            Case "Stop with Limit"
                .Text = Editor2.Text & " STOP, " & edLimit.Text & " LIMIT" & vbLf
            Case "Stop with Limit Close Only"
                .Text = Editor2.Text & " STOP, " & edLimit.Text & " LIMIT Close Only" & vbLf
            Case Else
                .Text = Editor2.Text & " " & cboOrderPlacement.Text & vbLf
        End Select
        
        If chkCanExitOnEntryBar.Enabled Then
            If chkCanExitOnEntryBar = vbChecked Then
                .Text = "Can Exit On Entry Bar = True" & vbLf
            Else
                .Text = "Can Exit On Entry Bar = False" & vbLf
            End If
        End If
        
        .Font.Size = 14
        .Font.Bold = True
        .FontUnderline = True
        .Text = vbLf & "Default Values for Rule Inputs:" & vbLf
        .FontUnderline = False
        .Font.Bold = False
        .Font.Size = 12
        ''.Text = vbLf
        
        lTemp = 0&
        For lIndex = 1 To vsInputs.Rows - 1
            If vsInputs.RowHidden(lIndex) = False Then
                .Text = "     " & vsInputs.Cell(flexcpText, lIndex, IGCol(eIGCol_InputName)) & " = "
                .Text = vsInputs.Cell(flexcpText, lIndex, IGCol(eIGCol_DefaultValue)) & vbLf
                lTemp = lTemp + 1
            End If
        Next lIndex
        
        If lTemp = 0 Then
            .Text = "( No Inputs for this Rule )" & vbLf
        End If
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.GenerateReport", eGDRaiseError_Raise
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Form called from, Rule ID to Load, System ID for that rule,
''              System Library ID for that rule, Next default numbers for names
'' Returns:     Rule ID edited
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strCalledFrom As String, Optional ByVal lRuleID As Long = 0, _
            Optional ByVal lSystemID As Long = 0, Optional ByVal lSysLibID As Long = 0, _
            Optional ByVal lLE As Long = 0, Optional ByVal lLX As Long = 0, _
            Optional ByVal lSE As Long = 0, Optional ByVal lSX As Long = 0, _
            Optional frmSysMgr As Form = Nothing, _
            Optional strTextToPaste As String) As Long
On Error GoTo ErrSection:

    m.strCalledFrom = strCalledFrom
    m.lSystemNumber = lSystemID
    m.lSystemLibID = lSysLibID
    m.lLongEntry = lLE
    m.lLongExit = lLX
    m.lShortEntry = lSE
    m.lShortExit = lSX
    Set m.frmSysMgr = frmSysMgr

    'If lRuleID = 0 Then
    '    Add
    'Else
    '    LoadRec lRuleID
    'End If
    Set m.Rule = New cRule
    m.Rule.RuleID = lRuleID
    m.Rule.Load
    Load
    
    Screen.MousePointer = vbDefault
    EnableToolbar False
        
    FormResize Me
    ShowForm Me, False, frmMain
    
    If Len(strTextToPaste) > 0 Then
        If lSystemID > 0 Then chkGlobal.Value = False
        Editor1.Text = strTextToPaste
        Editor1_Change
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.ShowMe", eGDRaiseError_Raise

End Function

Public Function ShowFromSysMgr(ByVal Rule As cRule, Optional ByVal lSystemID As Long = 0, _
            Optional ByVal lSysLibID As Long = 0, Optional ByVal lLE As Long = 0, _
            Optional ByVal lLX As Long = 0, Optional ByVal lSE As Long = 0, _
            Optional ByVal lSX As Long = 0, Optional frmSysMgr As Form = Nothing, _
            Optional ByVal strTextToPaste As String, _
            Optional ByVal lActionIdx As Long = -1) As Long
On Error GoTo ErrSection:

    m.strCalledFrom = "frmSystemManager"
    m.lSystemNumber = lSystemID
    m.lSystemLibID = lSysLibID
    m.lLongEntry = lLE
    m.lLongExit = lLX
    m.lShortEntry = lSE
    m.lShortExit = lSX
    Set m.frmSysMgr = frmSysMgr

    Set m.Rule = Rule
    Load
    
    Screen.MousePointer = vbDefault
    EnableToolbar False
    FormResize Me
    If Len(strTextToPaste) > 0 Then
        Editor1.Text = strTextToPaste
        Editor1_Change
    End If
    If lActionIdx >= 0 Then cboAction.ListIndex = lActionIdx
    
    ShowForm Me, False, frmMain
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.ShowMe", eGDRaiseError_Raise

End Function

Private Sub EnableToolbar(ByVal bEnabled As Boolean)
On Error GoTo ErrSection:

    With tbToolbar
        .Tools("ID_Save").Enabled = bEnabled
        'original code - save awhile then remove 09-14-2004
        '.Tools("ID_SaveAs").Enabled = (Trim(m.strName <> ""))
        '.Tools("ID_Rename").Enabled = (Trim(m.strName <> ""))
        '.Tools("ID_SaveFavorite").Enabled = Not bEnabled
        'end original code save
        .Tools("ID_SaveAs").Visible = (Len(Trim(m.strName)) > 0)
        .Tools("ID_Rename").Visible = (Len(Trim(m.strName)) > 0)
        If Len(Trim(m.strName)) = 0 Then
            .Tools("ID_SaveFavorite").Visible = False
        Else
            .Tools("ID_SaveFavorite").Visible = (m.lSystemNumber <> 0)
        End If
        .Tools("ID_Diagram").Enabled = Not .Tools("ID_Verify").Enabled
        .Tools("ID_CondBuilder").Visible = (Len(Trim(m.strName)) = 0)

        
        'disable toolbox button if strategy wizard is in use
        If m.frmSysMgr Is Nothing Then
            .Tools("ID_ToolBox").Enabled = True
        Else
            If TypeOf m.frmSysMgr Is frmSystemManager Then
                .Tools("ID_ToolBox").Enabled = Not m.frmSysMgr.StrategyWizard
            End If
        End If
    End With
    
    ' Hide the advanced tab for a local rule...
    vsIndexTab1.TabVisible(Tabs(eTab_Advanced)) = m.bShared

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRule.EnableToolbar", eGDRaiseError_Raise
    
End Sub

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
        
    Dim bSkipAutoIf As Boolean
    Dim strResponse As String
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        If WindowState = vbMinimized Then WindowState = vbNormal
    
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

Private Function Load() As Boolean
On Error GoTo ErrSection:

    Dim strSystem As String

    m.lCurrentRuleID = m.Rule.RuleID
    m.bNewRule = (m.Rule.RuleID = 0)
    
    With m.Rule
        If Not g.Security.CanEdit(.SecurityLevel, .Password) Then GoTo ErrExit:
        
        ClearForm
        m.strName = .Name
        m.bNewRule = (.RuleID = 0)
        
        'Zero for System number means global
        If .SystemNumber <> 0 Or m.strCalledFrom = "frmSystemManager" Then
            chkGlobal.Value = vbUnchecked
            m.bShared = False
        Else
            chkGlobal.Value = vbChecked
            m.bShared = True
        End If
        chkGlobal.Visible = False
        If m.lSystemNumber <> 0 Then
            strSystem = SystemNameForID(m.lSystemNumber)
            If strSystem <> "" Then
                vsSystemName.Caption = "Local Rule for " & Replace(strSystem, "&", "&&")
            Else
                vsSystemName.Caption = "Local Rule for New Strategy"
            End If
        Else
            vsSystemName.Caption = "Building Block Rule"
        End If
            
        If .RuleID <> 0 Then
            Editor1.TextRTF = .GetRTF(.Cond)
            edLimit.TextRTF = .GetRTF(.Price2RTF)
            
            If .BuySell Then
                If .RuleType = 0 Then
'                    If IsDefaultRuleName(.Name) Then m.lLongEntry = DefaultRuleNumber(.Name)
                    cboAction.ListIndex = Action(eRA_LongEntry)
                Else
'                    If IsDefaultRuleName(.Name) Then m.lShortExit = DefaultRuleNumber(.Name)
                    cboAction.ListIndex = Action(eRA_ShortExit)
                End If
            Else
                If .RuleType = 0 Then
'                    If IsDefaultRuleName(.Name) Then m.lShortEntry = DefaultRuleNumber(.Name)
                    cboAction.ListIndex = Action(eRA_ShortEntry)
                Else
'                    If IsDefaultRuleName(.Name) Then m.lLongExit = DefaultRuleNumber(.Name)
                    cboAction.ListIndex = Action(eRA_LongExit)
                End If
            End If
            
            If .ExitOnEntryBar = True Then chkCanExitOnEntryBar.Value = vbChecked Else chkCanExitOnEntryBar.Value = vbUnchecked
            If Len(.OrderPlacement) = 0 Then
                .OrderPlacement = "Market"
                InfBox "The Order Placement for|" & .Name & "|is invalid.  It will be set to a Market order.", "!", , "Warning"
            End If
            cboOrderPlacement.Text = .OrderPlacement
            Editor2.TextRTF = .GetRTF(.Price)
            
            'Temporary, remove after converting all rules from rtf stored format into
            'tokened text format 7-17-2000 MT
            If Left(.Cond, 1) = "{" Then
                Editor1.TextRTF = .Cond
                edLimit.TextRTF = .Price2RTF
                Editor2.TextRTF = .Price
            End If
            
            ' Take care of pyramiding stuff 5/18/2001 DAJ
            If .RuleType = 0 Then
                txtContractsToEnter.Text = Format(.NumberContracts, "#,##0")
            Else
                If .ExitBasedOnEachTrade Then
                    optTrade.Value = True
                Else
                    optPosition.Value = True
                    If .AsPercentOfPosition Then
                        optPercent.Value = True
                        txtPercent.Text = CStr(.NumberContracts)
                    Else
                        optNumContracts.Value = True
                        txtContractsToExit.Text = Format(.NumberContracts, "#,##0")
                    End If
                End If
            End If
        
            ' Select the appropriate item in the rule category combo box...
            If .CategoryID > 0 Then
                cboCategory.Text = RuleCategoryFromID(.CategoryID)
            ElseIf .RuleType = 0 Then
                cboCategory.Text = "Other Entries"
            Else
                cboCategory.Text = "Other Exits"
            End If
        End If
    End With
    
    SetEditorCaption Me, "Rule", m.strName
    
    'Initiliaze inputs grid
    Set m.Inputs = m.Rule.Inputs
    LoadInputsGrid
    
    'Build advanced view
    With m.Rule
        edAdvanced.TextRTF = .GetRTF(EZToAdvanced(.BuySell, .OrderPlacement, _
            .Cond, .Price, .Price2RTF))
        chkAdvanced_Click
    End With
    
    tbToolbar.Tools("ID_Verify").Enabled = m.Rule.Reverify
    tbToolbar.Tools("ID_Diagram").Enabled = Not m.Rule.Reverify
    EnableToolbar m.Rule.Reverify
    
    Load = True
    
ErrExit:
    m.lReturnValue = LockWindowUpdate(0)
    Exit Function
    
ErrSection:
    m.lReturnValue = LockWindowUpdate(0)
    RaiseError "frmRule.Load", eGDRaiseError_Raise

End Function

Private Function VerifyUsage(ByVal hRefsArray As Long) As Boolean
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
    
    If astrFunctions.Size > 0 Then
        VerifyUsage = False
        InfBox "The following function(s) cannot be used in|strategy testing:||" & _
                astrFunctions.JoinFields("|") & "|", , , "Validation Error"
    Else
        VerifyUsage = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.VerifyUsage", eGDRaiseError_Raise
    
End Function

Private Function VerifyFunctionUsage(ByVal lID As Long, Optional astrFunctions As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim rsFunction As Recordset
    Dim lPos As Long

    VerifyFunctionUsage = True

    Set rsFunction = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
            "WHERE [FunctionID]=" & Str(lID) & ";", dbOpenDynaset)
    If Not (rsFunction.BOF And rsFunction.EOF) Then
        If GetBit(rsFunction!Usage, 2) = False Then 'And (rsFunction!Usage <> 0) Then
            If Not astrFunctions Is Nothing Then
                If astrFunctions.BinarySearch(rsFunction!FunctionName, lPos) = False Then
                    astrFunctions.Add rsFunction!FunctionName, lPos
                End If
            End If
            VerifyFunctionUsage = False
        End If
    End If

ErrExit:
    Set rsFunction = Nothing
    Exit Function
    
ErrSection:
    Set rsFunction = Nothing
    RaiseError "frmRule.VerifyFunctionUsage", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CloseOnlyOrders
'' Description: Determine if the rule references Next Bar Low, High, or Close
'' Inputs:      None
'' Returns:     True if rule must be a Close Only, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CloseOnlyOrders() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim alFunctions As New cGdArray     ' List of Function ID's in this Rule's "tree"
    Dim lIndex As Long                  ' Index into a for loop
    Dim strName As String               ' Name of the function
    
    bReturn = False
    
    ' Get the list of functions used by this rule either directly or indirectly...
    alFunctions.Create eGDARRAY_Longs
    Set alFunctions = RuleRefs(m.Rule.CondRefs)
    
    ' Walk through all of the functions looking for an occurance of Next Bar High,
    ' Next Bar Low, or Next Bar Close...
    For lIndex = 0 To alFunctions.Size - 1
        strName = FunctionNameFromID(alFunctions(lIndex))
        Select Case UCase(strName)
            Case "NEXT BAR LOW", "NEXT BAR HIGH", "NEXT BAR CLOSE"
                bReturn = True
                Exit For
            
        End Select
    Next lIndex
    
    CloseOnlyOrders = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.CloseOnlyOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCategories
'' Description: Load the category combo box with the available categories
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCategories()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    ' Clear out the combo box...
    cboCategory.Clear
    
    ' Load the combo from the items in the Rule Categories table in the database...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleCategories];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!CategoryID > 0 Then
            cboCategory.AddItem rs!CategoryName
            cboCategory.ItemData(cboCategory.NewIndex) = rs!CategoryID
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmRule.LoadCategories", eGDRaiseError_Raise
    
End Sub

Public Property Let CondBuilderExpr(ByVal strExpr As String)
On Error Resume Next
    
   Editor1.Text = strExpr
   tbToolbar.Tools("ID_CondBuilder").Visible = False

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionOfIf
'' Description: Position of the IF in the editor text
'' Inputs:      None
'' Returns:     Position of IF if there, else zero
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PositionOfIf() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lPos As Long                    ' Position of IF in the string
    Dim strText As String               ' Text to look through
    Dim strCharBefore As String         ' Character before
    Dim strCharAfter As String          ' Character after
    
    lReturn = 0&
    strText = UCase(" " & Editor1.Text & " ")
    
    lPos = InStr(strText, "IF")
    Do While lPos > 0
        strCharBefore = Mid(strText, lPos - 1, 1)
        strCharAfter = Mid(strText, lPos + 2, 1)
        
        If (strCharBefore = " ") Or (strCharBefore = vbCr) Or (strCharBefore = vbLf) Then
            If (strCharAfter = " ") Or (strCharAfter = vbCr) Or (strCharAfter = vbLf) Then
                lReturn = lPos
                Exit Do
            End If
        End If
    
        lPos = InStr(lPos + 1, strText, "IF")
    Loop
    
    PositionOfIf = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.PositionOfIf"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyHighlightBarReport
'' Description: Perform a verifcation for the HighlightBar reporter
'' Inputs:      Error
'' Returns:     True if success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VerifyHighlightBarReport(strError As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim bPosition As Boolean            ' Perform pyramiding by position or trade?
    Dim bPercent As Boolean             ' Perform pyramiding by percent of position?
    Dim strNum1 As String               ' Contracts to enter
    Dim strNum2 As String               ' Contracts to exit
    Dim strNum3 As String               ' Percent of position to exit
    Dim Expr As cExpression             ' Expression object for doing the verification
    Dim Condition As String
    Dim EntryPrice As String
    Dim StopLimitPrice As String
    Dim CondLate As Boolean
    Dim ActionLate As Boolean
    Dim OrderTypeInd As Byte
    Dim Buy As Byte
    
    bReturn = False
    strError = ""
    
    If (Trim(Editor1.Text) = "") Or (Trim(UCase(Editor1.Text)) = "IF") Then
        strError = "Condition not specified"
        bReturn = False
    Else
        Editor1.Text = FixPeriodInMarkets(Editor1.Text)
    
        ' Save current input values entered by user...
        SaveGridValues
        
        ' Save the pyramiding values
        bPosition = optPosition.Value
        bPercent = optPercent.Value
        strNum1 = txtContractsToEnter.Text
        strNum2 = txtContractsToExit.Text
        strNum3 = txtPercent.Text
        
        ' Make sure that a price condition is specified for non-market orders and
        ' a with limit price condition is specified for with limit orders...
        If (Editor2.Visible = True) And (Len(Editor2.Text) = 0) Then
            strError = "Order price not specified"
            bReturn = False
        ElseIf (Editor2.Visible = True) And (IsBooleanExpression(Editor2.Text) = True) Then
            strError = "Order price cannot be a boolean expression"
            bReturn = False
        ElseIf (edLimit.Visible = True) And (Len(edLimit.Text) = 0) Then
            strError = "With Limit price not specified"
            bReturn = False
        ElseIf (edLimit.Visible = True) And (IsBooleanExpression(edLimit.Text) = True) Then
            strError = "With Limit price cannot be a boolean expression"
            bReturn = False
        Else
            'Shut things off, get ready for verifying rule
            Screen.MousePointer = vbHourglass
            m.lReturnValue = LockWindowUpdate(Me.hWnd)
            edAdvanced.TurnOffEditing
            
            'Verify...
            Set Expr = New cExpression
            Expr.PortfolioNavigator = False
            Expr.Functions = g.Functions
            
            'Advanced editor
            If chkAdvanced.Value = vbChecked Then
                Expr.ValidateRule edAdvanced.Text
            Else
                'EZ editor
                Buy = cboAction.ListIndex = Action(eRA_LongEntry) Or cboAction.ListIndex = Action(eRA_ShortExit)
                Expr.ValidateRule BuildRule(Editor1.Text, Buy, Editor2.Text, cboOrderPlacement.ListIndex, edLimit.Text)
            End If
            
            '.GetFIDs=Returns the function ID's used in the entire rule...
            m.Rule.CondRefs = Expr.GetFIDs
            
            If InputsValidForHighlightBarReport(Expr.Inputs) = False Then
                strError = "Inputs are invalid"
                bReturn = False
            ElseIf Expr.LateCondition = True Then
                strError = "Condition cannot reference late calculating functions"
                bReturn = False
            ElseIf CloseOnlyOrders = True Then
                strError = "Condition cannot reference Next Bar Low,|Next Bar High, or Next Bar Close"
                bReturn = False
            Else
                'Save Late calculating flags...
                m.Rule.LateCondition = Expr.LateCondition
                m.Rule.LateAction = Expr.LateAction
                
                'Parse verified Rule back into text boxes...
                AdvancedToEZ True, Expr.EditText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
                m.Rule.Cond = Condition
                m.Rule.Price = EntryPrice
                m.Rule.Price2RTF = StopLimitPrice
                m.Rule.OrderPlacement = cboOrderPlacement.Text 'OrderTypeInd
                
                'Parse Codedtext used by engine and save appropriate pieces...
                AdvancedToEZ False, Expr.CodedText, Buy, OrderTypeInd, Condition, EntryPrice, StopLimitPrice
                m.Rule.CondCoded = Condition
                If EntryPrice = "" Then
                    m.Rule.PriceCoded = "N/A"
                Else
                    m.Rule.PriceCoded = EntryPrice
                End If
                If StopLimitPrice = "" Then
                    m.Rule.Price2Coded = "N/A"
                Else
                    m.Rule.Price2Coded = StopLimitPrice
                End If
                
                'Load the inputs from condition and Actions into grid...Restore
                'any changes input values that already exist in the grid
                Set m.Inputs = Expr.Inputs
                LoadInputsGrid
                RestoreGridValues
                
                ' Make sure that the rule does not refrence any function that is not meant
                ' for strategy testing...
                If VerifyUsage(m.Rule.CondRefs) = False Then
                    strError = "Condition references functions that cannot be used in strategy testing"
                    bReturn = False
                Else
                    ' Restore the pyramiding values
                    optPosition.Value = bPosition
                    optPercent.Value = bPercent
                    txtContractsToEnter.Text = strNum1
                    txtContractsToExit.Text = strNum2
                    txtPercent.Text = strNum3
                    
                    m.Rule.Reverify = False
                    tbToolbar.Tools("ID_Verify").Enabled = False
                    tbToolbar.Tools("ID_Diagram").Enabled = True
                    
                    Screen.MousePointer = vbDefault
                    m.lReturnValue = LockWindowUpdate(0)
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    VerifyHighlightBarReport = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.VerifyHighlightBarReport"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InputsValidForHighlightBarReport
'' Description: Determine if the given inputs are valid for a highlight bar report
'' Inputs:      Inputs
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InputsValidForHighlightBarReport(ByVal Inputs As cInputs) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strValidInputs As String        ' Valid inputs

    strValidInputs = ",MARKET1,DAILY,WEEKLY,MONTHLY,"

    bReturn = True
    If Inputs Is Nothing Then
        bReturn = False
    ElseIf Inputs.Count = 0 Then
        bReturn = False
    ElseIf UCase(Inputs.Item(1).ParmName) <> "MARKET1" Then
        bReturn = False
    Else
        For lIndex = 2 To Inputs.Count
            If InStr(strValidInputs, "," & UCase(Inputs.Item(lIndex).ParmName) & ",") = 0 Then
                bReturn = False
                Exit For
            End If
        Next lIndex
    End If
    
    InputsValidForHighlightBarReport = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRule.InputsValidForHighlightBarReport"
    
End Function

