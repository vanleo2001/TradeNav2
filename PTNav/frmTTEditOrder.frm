VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTTEditOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buy/Sell"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraTime 
      Height          =   555
      Left            =   5700
      TabIndex        =   41
      Top             =   2460
      Width           =   4035
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
      Caption         =   "frmTTEditOrder.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":0028
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":0048
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optOnOpen 
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   240
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
         Caption         =   "frmTTEditOrder.frx":0064
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":0092
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":00B2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optOnClose 
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   240
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
         Caption         =   "frmTTEditOrder.frx":00CE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":00FE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":011E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAnyTime 
         Height          =   255
         Left            =   180
         TabIndex        =   42
         Top             =   240
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
         Caption         =   "frmTTEditOrder.frx":013A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmTTEditOrder.frx":016A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":018A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboLots 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2700
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
      Tip             =   "frmTTEditOrder.frx":01A6
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":01C6
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkAutoJournal 
      Height          =   255
      Left            =   9060
      TabIndex        =   103
      Top             =   1980
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
      Caption         =   "frmTTEditOrder.frx":01E2
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmTTEditOrder.frx":021C
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":023C
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboExchanges 
      Height          =   315
      Left            =   1080
      TabIndex        =   102
      Top             =   2340
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
      Tip             =   "frmTTEditOrder.frx":0258
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":0278
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin VB.Timer tmrWait 
      Left            =   8880
      Top             =   2640
   End
   Begin HexUniControls.ctlUniFrameWL fraAdvanced 
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   3180
      Width           =   10155
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
      Caption         =   "frmTTEditOrder.frx":0294
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":02B4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":02D4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkAdvanced 
         Height          =   255
         Left            =   8520
         TabIndex        =   53
         Top             =   0
         Width           =   1695
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
         Caption         =   "frmTTEditOrder.frx":02F0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":0334
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":0354
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkExitAll 
         Height          =   255
         Left            =   6780
         TabIndex        =   52
         Top             =   0
         Width           =   1695
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
         Caption         =   "frmTTEditOrder.frx":0370
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":03BA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":03DA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkAutoExit 
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   0
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
         Caption         =   "frmTTEditOrder.frx":03F6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":0436
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":0456
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSession 
      Height          =   555
      Left            =   4020
      TabIndex        =   38
      Top             =   2520
      Width           =   4035
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
      Caption         =   "frmTTEditOrder.frx":0472
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":04A0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":04C0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optElectronic 
         Height          =   195
         Left            =   780
         TabIndex        =   40
         Top             =   240
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
         Caption         =   "frmTTEditOrder.frx":04DC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":0510
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":0530
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPit 
         Height          =   220
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "frmTTEditOrder.frx":054C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmTTEditOrder.frx":0572
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":0592
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab tabSpecialOrders 
      Height          =   2295
      Left            =   180
      TabIndex        =   54
      Top             =   3600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
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
      Caption         =   "Order Cancels Order|Triggered By Order|Triggered by Condition|Trailing Stop|Contingency"
      Align           =   0
      Appearance      =   1
      CurrTab         =   2
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin HexUniControls.ctlUniFrameWL fraContingency 
         Height          =   1920
         Left            =   10080
         TabIndex        =   8
         Top             =   45
         Width           =   9045
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
         Caption         =   "frmTTEditOrder.frx":05AE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditOrder.frx":05CE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":05EE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraContingencyStop 
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   1080
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
            Caption         =   "frmTTEditOrder.frx":060A
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTEditOrder.frx":062A
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":064A
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraStopTif 
               Height          =   255
               Left            =   2760
               TabIndex        =   12
               Top             =   0
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
               Caption         =   "frmTTEditOrder.frx":0666
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmTTEditOrder.frx":0686
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":06A6
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optStopDay 
                  Height          =   255
                  Left            =   60
                  TabIndex        =   14
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
                  Caption         =   "frmTTEditOrder.frx":06C2
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmTTEditOrder.frx":06E8
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTEditOrder.frx":0708
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optStopGTC 
                  Height          =   255
                  Left            =   780
                  TabIndex        =   25
                  Top             =   0
                  Width           =   735
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
                  Caption         =   "frmTTEditOrder.frx":0724
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmTTEditOrder.frx":074A
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTEditOrder.frx":076A
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniCheckXP chkStopLoss 
               Height          =   195
               Left            =   60
               TabIndex        =   27
               Top             =   0
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
               Caption         =   "frmTTEditOrder.frx":0786
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":07D6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":07F6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optStopDollar 
               Height          =   255
               Left            =   480
               TabIndex        =   28
               Top             =   240
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
               Caption         =   "frmTTEditOrder.frx":0812
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":084C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":086C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optStopPoints 
               Height          =   255
               Left            =   480
               TabIndex        =   29
               Top             =   480
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
               Caption         =   "frmTTEditOrder.frx":0888
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":08B4
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":08D4
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin NavSuite.gdPriceEditor peStopLoss 
               Height          =   375
               Left            =   1920
               TabIndex        =   30
               Top             =   300
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   661
            End
            Begin HexUniControls.ctlUniLabelXP lblStopOther 
               Height          =   195
               Left            =   3360
               Top             =   390
               Width           =   1875
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
               Caption         =   "frmTTEditOrder.frx":08F0
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTEditOrder.frx":091C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":093C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraContingencyProfit 
            Height          =   735
            Left            =   120
            TabIndex        =   31
            Top             =   300
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
            Caption         =   "frmTTEditOrder.frx":0958
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTEditOrder.frx":0978
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":0998
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraProfitTif 
               Height          =   255
               Left            =   2760
               TabIndex        =   32
               Top             =   0
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
               Caption         =   "frmTTEditOrder.frx":09B4
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmTTEditOrder.frx":09D4
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":09F4
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optProfitDay 
                  Height          =   255
                  Left            =   60
                  TabIndex        =   56
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
                  Caption         =   "frmTTEditOrder.frx":0A10
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmTTEditOrder.frx":0A36
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTEditOrder.frx":0A56
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optProfitGTC 
                  Height          =   255
                  Left            =   780
                  TabIndex        =   85
                  Top             =   0
                  Width           =   735
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
                  Caption         =   "frmTTEditOrder.frx":0A72
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmTTEditOrder.frx":0A98
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmTTEditOrder.frx":0AB8
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniRadioXP optProfitPoints 
               Height          =   255
               Left            =   480
               TabIndex        =   86
               Top             =   480
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
               Caption         =   "frmTTEditOrder.frx":0AD4
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":0B00
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":0B20
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optProfitDollar 
               Height          =   255
               Left            =   480
               TabIndex        =   94
               Top             =   240
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
               Caption         =   "frmTTEditOrder.frx":0B3C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":0B76
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":0B96
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkProfitTarget 
               Height          =   195
               Left            =   60
               TabIndex        =   99
               Top             =   0
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
               Caption         =   "frmTTEditOrder.frx":0BB2
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":0C0A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":0C2A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin NavSuite.gdPriceEditor peProfitTarget 
               Height          =   375
               Left            =   1920
               TabIndex        =   101
               Top             =   300
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
            End
            Begin HexUniControls.ctlUniLabelXP lblProfitOther 
               Height          =   195
               Left            =   3360
               Top             =   390
               Width           =   1875
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
               Caption         =   "frmTTEditOrder.frx":0C46
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTEditOrder.frx":0C72
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":0C92
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblContingency 
            Height          =   195
            Left            =   180
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
            Caption         =   "frmTTEditOrder.frx":0CAE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditOrder.frx":0D16
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":0D36
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraTrail 
         Height          =   1920
         Left            =   9780
         TabIndex        =   81
         Top             =   45
         Width           =   9045
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
         Caption         =   "frmTTEditOrder.frx":0D52
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditOrder.frx":0D72
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":0D92
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtPointTrail 
            Height          =   285
            Left            =   4320
            TabIndex        =   76
            Top             =   960
            Width           =   1020
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditOrder.frx":0DAE
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
            Tip             =   "frmTTEditOrder.frx":0DDE
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":0DFE
         End
         Begin HexUniControls.ctlUniRadioXP optTrailPoints 
            Height          =   255
            Left            =   1080
            TabIndex        =   75
            Top             =   960
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
            Caption         =   "frmTTEditOrder.frx":0E1A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmTTEditOrder.frx":0E88
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":0EA8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optTrailDollar 
            Height          =   255
            Left            =   1080
            TabIndex        =   73
            Top             =   600
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
            Caption         =   "frmTTEditOrder.frx":0EC4
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":0F2C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":0F4C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTrail 
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   270
            Width           =   7935
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
            Caption         =   "frmTTEditOrder.frx":0F68
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":1056
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1076
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdScrollBar sbPointTrail 
            Height          =   360
            Left            =   5340
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   915
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDollarTrail 
            Height          =   285
            Left            =   4320
            TabIndex        =   74
            Top             =   600
            Width           =   1215
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditOrder.frx":1092
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
            Tip             =   "frmTTEditOrder.frx":10B2
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":10D2
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraCondition 
         Height          =   1920
         Left            =   45
         TabIndex        =   80
         Top             =   45
         Width           =   9045
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
         Caption         =   "frmTTEditOrder.frx":10EE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditOrder.frx":110E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":112E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRichTextBoxXP rtfCondition 
            Height          =   555
            Left            =   480
            TabIndex        =   98
            Top             =   1260
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   979
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditOrder.frx":114A
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
            Tip             =   "frmTTEditOrder.frx":116A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":118A
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
         Begin HexUniControls.ctlUniButtonImageXP cmdEditCustom 
            Height          =   315
            Left            =   1920
            TabIndex        =   97
            Top             =   900
            Width           =   1095
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
            Caption         =   "frmTTEditOrder.frx":11A6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":11D4
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":11F4
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkCustomCondition 
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   930
            Width           =   1695
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
            Caption         =   "frmTTEditOrder.frx":1210
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":1250
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1270
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraPriceCheck 
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   480
            Width           =   5895
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
            Caption         =   "frmTTEditOrder.frx":128C
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTEditOrder.frx":12AC
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":12CC
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkPrice 
               Height          =   255
               Left            =   0
               TabIndex        =   95
               Top             =   53
               Width           =   255
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
               Caption         =   "frmTTEditOrder.frx":12E8
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":1314
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1334
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdCondSymLookup 
               Height          =   255
               Left            =   3270
               TabIndex        =   91
               Top             =   53
               Width           =   240
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
               Caption         =   "frmTTEditOrder.frx":1350
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":1390
               Style           =   1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":13B0
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniComboImageXP cboConditionField 
               Height          =   315
               Left            =   300
               TabIndex        =   90
               Top             =   23
               Width           =   1455
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
               Tip             =   "frmTTEditOrder.frx":13CC
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":13EC
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniComboImageXP cboConditionOperator 
               Height          =   315
               Left            =   3660
               TabIndex        =   89
               Top             =   23
               Width           =   735
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
               Tip             =   "frmTTEditOrder.frx":1408
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1428
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtConditionPrice 
               Height          =   285
               Left            =   4500
               TabIndex        =   88
               Top             =   38
               Width           =   1020
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTEditOrder.frx":1444
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
               Tip             =   "frmTTEditOrder.frx":1474
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1494
            End
            Begin gdOCX.gdScrollBar gdCondPrice 
               Height          =   360
               Left            =   5520
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   0
               Width           =   210
               _ExtentX        =   370
               _ExtentY        =   635
            End
            Begin HexUniControls.ctlUniTextBoxXP txtConditionSymbol 
               Height          =   315
               Left            =   2100
               TabIndex        =   92
               Top             =   23
               Width           =   1440
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTTEditOrder.frx":14B0
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
               Tip             =   "frmTTEditOrder.frx":14E4
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1504
            End
            Begin HexUniControls.ctlUniLabelXP lblOf 
               Height          =   255
               Left            =   1860
               Top             =   60
               Width           =   255
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
               Caption         =   "frmTTEditOrder.frx":1520
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTEditOrder.frx":1544
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1564
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraConditionTime 
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   120
            Width           =   8895
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
            Caption         =   "frmTTEditOrder.frx":1580
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTEditOrder.frx":15A0
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":15C0
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectDate gdEndTime 
               Height          =   300
               Left            =   5400
               TabIndex        =   71
               Top             =   0
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   529
               ShowDayOfWeek   =   0   'False
               ShowPM          =   1
               ShowTime        =   1
               Value           =   2
            End
            Begin HexUniControls.ctlUniCheckXP chkEndTime 
               Height          =   255
               Left            =   4500
               TabIndex        =   70
               Top             =   30
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
               Caption         =   "frmTTEditOrder.frx":15DC
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":160C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":162C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectDate gdBeginTime 
               Height          =   300
               Left            =   900
               TabIndex        =   69
               Top             =   0
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   529
               ShowDayOfWeek   =   0   'False
               ShowPM          =   1
               ShowTime        =   1
               Value           =   2
            End
            Begin HexUniControls.ctlUniCheckXP chkBeginTime 
               Height          =   255
               Left            =   0
               TabIndex        =   68
               Top             =   30
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
               Caption         =   "frmTTEditOrder.frx":1648
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":1678
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1698
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblEndLocalTime 
               Height          =   195
               Left            =   8100
               Top             =   60
               Width           =   735
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
               Caption         =   "frmTTEditOrder.frx":16B4
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTEditOrder.frx":16EA
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":170A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblBeginLocalTime 
               Height          =   195
               Left            =   3600
               Top             =   60
               Width           =   735
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
               Caption         =   "frmTTEditOrder.frx":1726
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTTEditOrder.frx":175C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":177C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblDefaultPeriod 
            Height          =   255
            Left            =   3180
            Top             =   960
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
            Caption         =   "frmTTEditOrder.frx":1798
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditOrder.frx":17D6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":17F6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraTrigger 
         Height          =   1920
         Left            =   -9690
         TabIndex        =   79
         Top             =   45
         Width           =   9045
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
         Caption         =   "frmTTEditOrder.frx":1812
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditOrder.frx":1832
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":1852
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkMoveWithTrigger 
            Height          =   255
            Left            =   420
            TabIndex        =   61
            Top             =   1020
            Width           =   4695
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
            Caption         =   "frmTTEditOrder.frx":186E
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":18E0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1900
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtRelativePrice 
            Height          =   285
            Left            =   5160
            TabIndex        =   63
            Top             =   1500
            Width           =   1275
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditOrder.frx":191C
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
            Tip             =   "frmTTEditOrder.frx":193C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":195C
         End
         Begin HexUniControls.ctlUniCheckXP chkRelativePrice 
            Height          =   255
            Left            =   420
            TabIndex        =   62
            Top             =   1515
            Width           =   4695
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
            Caption         =   "frmTTEditOrder.frx":1978
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":1A1C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1A3C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTriggerConfirm 
            Height          =   255
            Left            =   5160
            TabIndex        =   64
            Top             =   120
            Width           =   3135
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
            Caption         =   "frmTTEditOrder.frx":1A58
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":1AC6
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1AE6
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraTriggerOptions 
            Height          =   855
            Left            =   5160
            TabIndex        =   65
            Top             =   480
            Width           =   3135
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
            Caption         =   "frmTTEditOrder.frx":1B02
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTTEditOrder.frx":1B46
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1B66
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optTriggerFull 
               Height          =   225
               Left            =   240
               TabIndex        =   67
               Top             =   510
               Width           =   2535
               _ExtentX        =   4471
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
               Caption         =   "frmTTEditOrder.frx":1B82
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTTEditOrder.frx":1BE0
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1C00
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optTriggerPartial 
               Height          =   255
               Left            =   240
               TabIndex        =   66
               Top             =   240
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
               Caption         =   "frmTTEditOrder.frx":1C1C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmTTEditOrder.frx":1C8A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTTEditOrder.frx":1CAA
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniComboImageXP cboTriggerOrders 
            Height          =   315
            Left            =   420
            TabIndex        =   60
            Top             =   600
            Width           =   4575
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
            Tip             =   "frmTTEditOrder.frx":1CC6
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1CE6
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkTriggerOrder 
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   270
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
            Caption         =   "frmTTEditOrder.frx":1D02
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":1D82
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1DA2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraCancel 
         Height          =   1920
         Left            =   -9990
         TabIndex        =   78
         Top             =   45
         Width           =   9045
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
         Caption         =   "frmTTEditOrder.frx":1DBE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditOrder.frx":1DDE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":1DFE
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkBrokerOCO 
            Height          =   195
            Left            =   420
            TabIndex        =   58
            Top             =   1020
            Width           =   3315
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
            Caption         =   "frmTTEditOrder.frx":1E1A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":1E8A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1EAA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboCancelOrders 
            Height          =   315
            Left            =   3660
            TabIndex        =   57
            Top             =   660
            Width           =   4095
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
            Tip             =   "frmTTEditOrder.frx":1EC6
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":1EE6
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkCancelOrder 
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   270
            Width           =   8235
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
            Caption         =   "frmTTEditOrder.frx":1F02
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTTEditOrder.frx":2008
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":2028
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblCancelNote 
            Height          =   255
            Left            =   420
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
            Caption         =   "frmTTEditOrder.frx":2044
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTTEditOrder.frx":20BE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":20DE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1875
      Left            =   9120
      TabIndex        =   45
      Top             =   60
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
      Caption         =   "frmTTEditOrder.frx":20FA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":211A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":213A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdJournal 
         Height          =   435
         Left            =   0
         TabIndex        =   48
         Top             =   960
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
         Caption         =   "frmTTEditOrder.frx":2156
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":2186
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":21A6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   49
         Top             =   1440
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
         Caption         =   "frmTTEditOrder.frx":21C2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":21F0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2210
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   46
         Top             =   0
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
         Caption         =   "frmTTEditOrder.frx":222C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":2266
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2286
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPark 
         Height          =   435
         Left            =   0
         TabIndex        =   47
         Top             =   480
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
         Caption         =   "frmTTEditOrder.frx":22A2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":22D8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":22F8
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPosition 
      Height          =   1335
      Left            =   240
      TabIndex        =   24
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
      Caption         =   "frmTTEditOrder.frx":2314
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":2334
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":2354
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   1380
         TabIndex        =   26
         Top             =   180
         Width           =   1815
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
         Tip             =   "frmTTEditOrder.frx":2370
         Sorted          =   -1  'True
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2390
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label10 
         Height          =   255
         Left            =   300
         Top             =   240
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
         Caption         =   "frmTTEditOrder.frx":23AC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":23DC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":23FC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAction 
         Height          =   255
         Left            =   1380
         Top             =   780
         Width           =   1875
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
         Caption         =   "frmTTEditOrder.frx":2418
         BackColor       =   -2147483633
         ForeColor       =   255
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":2458
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2478
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   195
         Left            =   180
         Top             =   780
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
         Caption         =   "frmTTEditOrder.frx":2494
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":24CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":24EE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblCurPos 
         Height          =   255
         Left            =   1380
         Top             =   540
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
         Caption         =   "frmTTEditOrder.frx":250A
         BackColor       =   -2147483633
         ForeColor       =   4210752
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":2540
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2560
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   60
         Top             =   540
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
         Caption         =   "frmTTEditOrder.frx":257C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":25BE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":25DE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNewPos 
         Height          =   255
         Left            =   1380
         Top             =   1020
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
         Caption         =   "frmTTEditOrder.frx":25FA
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":2632
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2652
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   255
         Left            =   120
         Top             =   1020
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
         Caption         =   "frmTTEditOrder.frx":266E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":26A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":26C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraDuration 
      Height          =   615
      Left            =   4020
      TabIndex        =   33
      Top             =   1800
      Width           =   4035
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
      Caption         =   "frmTTEditOrder.frx":26E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":2714
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":2734
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optDay 
         Height          =   220
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
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
         Caption         =   "frmTTEditOrder.frx":2750
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmTTEditOrder.frx":2776
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2796
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdExpiration 
         Height          =   315
         Left            =   1800
         TabIndex        =   37
         Top             =   210
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Value           =   38143
      End
      Begin HexUniControls.ctlUniRadioXP optExpiration 
         Height          =   220
         Left            =   1530
         TabIndex        =   36
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "frmTTEditOrder.frx":27B2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":27DC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":283C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGTC 
         Height          =   220
         Left            =   780
         TabIndex        =   35
         Top             =   270
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "frmTTEditOrder.frx":2858
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":287E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":28BE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraOrder 
      Height          =   1635
      Left            =   4080
      TabIndex        =   13
      Top             =   60
      Width           =   3435
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
      Caption         =   "frmTTEditOrder.frx":28DA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":2922
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":2942
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optMIT 
         Height          =   220
         Left            =   0
         TabIndex        =   104
         Top             =   1110
         Width           =   1665
         _ExtentX        =   2937
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
         Caption         =   "frmTTEditOrder.frx":295E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":299E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":29BE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VB.PictureBox picOrder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   2940
         ScaleHeight     =   1380
         ScaleWidth      =   480
         TabIndex        =   23
         Top             =   0
         Width           =   480
         Begin VB.Line lineHL 
            BorderColor     =   &H00808080&
            BorderWidth     =   3
            X1              =   210
            X2              =   210
            Y1              =   300
            Y2              =   1020
         End
         Begin VB.Line lineO 
            BorderColor     =   &H00808080&
            BorderWidth     =   3
            X1              =   210
            X2              =   60
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Line lineC 
            BorderColor     =   &H00808080&
            BorderWidth     =   3
            X1              =   360
            X2              =   210
            Y1              =   660
            Y2              =   660
         End
      End
      Begin HexUniControls.ctlUniTextBoxXP txtBelow 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   795
         Width           =   1020
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditOrder.frx":29DA
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
         Tip             =   "frmTTEditOrder.frx":2A0A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2A2A
      End
      Begin gdOCX.gdScrollBar sbAbove 
         Height          =   360
         Left            =   2700
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   225
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniTextBoxXP txtAbove 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   270
         Width           =   1020
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditOrder.frx":2A46
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
         Tip             =   "frmTTEditOrder.frx":2A76
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2A96
      End
      Begin HexUniControls.ctlUniRadioXP optBelow 
         Height          =   220
         Left            =   0
         TabIndex        =   19
         Top             =   840
         Width           =   1665
         _ExtentX        =   2937
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
         Caption         =   "frmTTEditOrder.frx":2AB2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":2AF6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2B16
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAbove 
         Height          =   220
         Left            =   0
         TabIndex        =   15
         Top             =   300
         Width           =   1665
         _ExtentX        =   2937
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
         Caption         =   "frmTTEditOrder.frx":2B32
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":2B74
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2B94
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdScrollBar sbBelow 
         Height          =   360
         Left            =   2700
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   750
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniRadioXP optMarket 
         Height          =   220
         Left            =   0
         TabIndex        =   18
         Top             =   570
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "frmTTEditOrder.frx":2BB0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmTTEditOrder.frx":2C04
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2C24
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optStopLimit 
         Height          =   220
         Left            =   0
         TabIndex        =   22
         Top             =   1380
         Width           =   2895
         _ExtentX        =   5106
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
         Caption         =   "frmTTEditOrder.frx":2C40
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":2CA2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2CC2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   255
         Left            =   0
         Top             =   30
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
         Caption         =   "frmTTEditOrder.frx":2CDE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":2D12
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2D32
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin gdOCX.gdScrollBar sbPrice 
      Height          =   360
      Left            =   7860
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPrice 
      Height          =   285
      Left            =   7980
      TabIndex        =   82
      Top             =   2040
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTTEditOrder.frx":2D4E
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
      Tip             =   "frmTTEditOrder.frx":2D7E
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":2D9E
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
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
      Caption         =   "frmTTEditOrder.frx":2DBA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTEditOrder.frx":2DDA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":2DFA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraQuantity 
         Height          =   675
         Left            =   840
         TabIndex        =   3
         Top             =   60
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
         Caption         =   "frmTTEditOrder.frx":2E16
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTTEditOrder.frx":2E36
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":2E56
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtQty 
            Height          =   315
            Left            =   0
            TabIndex        =   5
            Top             =   345
            Width           =   780
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditOrder.frx":2E72
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
            Tip             =   "frmTTEditOrder.frx":2E9C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":2EBC
         End
         Begin vsOcx6LibCtl.vsElastic vsCalc 
            Height          =   315
            Left            =   0
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Use money management calculator to determine quantity"
            Top             =   0
            Width           =   990
            _ExtentX        =   1746
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
            Appearance      =   1
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   600
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   192
            ForeColorDisabled=   -2147483631
            Caption         =   "Quantity"
            Align           =   0
            Appearance      =   1
            AutoSizeChildren=   0
            BorderWidth     =   4
            ChildSpacing    =   5
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
            PicturePos      =   7
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            _GridInfo       =   ""
            Begin VB.Image imgCalc 
               Height          =   240
               Left            =   705
               Picture         =   "frmTTEditOrder.frx":2ED8
               Top             =   45
               Width           =   240
            End
         End
         Begin gdOCX.gdScrollBar sbQty 
            Height          =   360
            Left            =   780
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   315
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin HexUniControls.ctlUniTextBoxXP txtExitPos 
            Height          =   315
            Left            =   0
            TabIndex        =   100
            Top             =   300
            Width           =   780
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTTEditOrder.frx":3462
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
            Tip             =   "frmTTEditOrder.frx":3492
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTTEditOrder.frx":34B2
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   255
         Left            =   3330
         TabIndex        =   11
         Top             =   390
         Width           =   240
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
         Caption         =   "frmTTEditOrder.frx":34CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":3500
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":3520
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTTEditOrder.frx":353C
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
         Tip             =   "frmTTEditOrder.frx":3570
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":3590
      End
      Begin HexUniControls.ctlUniRadioXP optSell 
         Height          =   220
         Left            =   0
         TabIndex        =   2
         Top             =   540
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "frmTTEditOrder.frx":35AC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   255
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":35D6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":35F6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optBuy 
         Height          =   220
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "frmTTEditOrder.frx":3612
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Pressed         =   0   'False
         Tip             =   "frmTTEditOrder.frx":363A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":365A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   195
         Left            =   1800
         Top             =   420
         Width           =   345
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
         Caption         =   "frmTTEditOrder.frx":3676
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":369A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":36BA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   2340
         Top             =   90
         Width           =   795
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
         Caption         =   "frmTTEditOrder.frx":36D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":3702
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":3722
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAt 
         Height          =   195
         Left            =   3600
         Top             =   420
         Width           =   255
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
         Caption         =   "frmTTEditOrder.frx":373E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":3762
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":3782
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   195
         Left            =   780
         Top             =   60
         Width           =   795
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
         Caption         =   "frmTTEditOrder.frx":379E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTTEditOrder.frx":37D0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTTEditOrder.frx":37F0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblLot 
      Height          =   195
      Left            =   240
      Top             =   2760
      Width           =   795
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
      Caption         =   "frmTTEditOrder.frx":380C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTTEditOrder.frx":3838
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":3858
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblExchange 
      Height          =   195
      Left            =   240
      Top             =   2400
      Width           =   795
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
      Caption         =   "frmTTEditOrder.frx":3874
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTTEditOrder.frx":38A6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTEditOrder.frx":38C6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmTTEditOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTEditOrder.frm
'' Description: Allow the user to edit/view an order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/21/2009   DAJ         Only give the wrong side of the market warning if
''                          the trade settings say to give the warning
'' 05/06/2009   DAJ         Allow for refreshing the order
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 11/20/2009   DAJ         Fixed bugs related to one account, no accounts
'' 12/14/2009   DAJ         Make conditional price regional settings friendly
'' 01/19/2010   DAJ         Fix for conditional price with Bonds (#5570)
'' 03/11/2010   DAJ         Added numbars required stuff for conditional orders (#5580)
'' 03/11/2010   DAJ         Use global Order Strategies collection, colors
'' 03/11/2010   DAJ         Made sure the Enter key works for the default button
'' 04/21/2010   DAJ         Took out flag file check for allowing Broker OCO's
'' 05/06/2010   DAJ         Check order price against triggering price if applicable
'' 06/03/2010   DAJ         Fixed persisting of Trigger Confirm/Trigger Partial flags (#5760)
'' 09/29/2010   DAJ         Changed reference to trigger confirm flag
'' 09/30/2010   DAJ         Sort the OCO and OTO combo boxes by date descending
'' 10/28/2010   DAJ         Changed default prices, Swap Stop/Limit as trigger changes
'' 12/01/2010   DAJ         Only load snapshot orders into the OCO/OTO combos
'' 05/26/2011   DAJ         On new order, use global trigger confirm flag
'' 06/14/2011   DAJ         When loading combos, load snapshot and parked orders
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 12/09/2011   DAJ         Persist the last trailing stop dollar amount the user typed in
'' 01/31/2012   DAJ         Disable all but price for automated orders
'' 03/06/2012   DAJ         When verifying, check to see if order can be moved to new price
'' 03/16/2012   DAJ         Don't allow user to change accounts on a non-parked existing order
'' 04/11/2012   DAJ         Don't allow user to check the custom condition without supplying a condition
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/05/2012   DAJ         Load lots combo regardless of current broker if lot number passed in
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 08/03/2012   DAJ         Remove Alaron
'' 09/11/2012   DAJ         Owner name in lot combo, Lot combo keyed by Lot ID
'' 09/14/2012   DAJ         Update the position based on the selected lot
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Contingency Orders
'' 12/18/2012   DAJ         Moved order status verification to the bottom of VerifyOrder
'' 01/18/2013   DAJ         Don't allow automated trading for spreads
'' 01/18/2013   DAJ         Broker held OCO for Interactive Brokers
'' 01/25/2013   DAJ         Check for above/below to be Null instead of Zero
'' 03/04/2013   DAJ         Removed 'Label2' because it was not used, but creating issues ( #6796 )
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 04/15/2013   DAJ         Only allow auto exits, TSOG, and automated trading if enabled for streaming
'' 06/27/2013   DAJ         Don't allow contingency orders for market orders
'' 08/01/2013   DAJ         Change to whether or not order is considered conditional
'' 08/28/2013   DAJ         If user clicks Submit, but not connected to broker, ask if they want to Park
'' 09/30/2013   DAJ         Removed Xpress, Fixed connection check upon submit for game mode
'' 01/17/2014   DAJ         Fix for OCO/OTO issue not having complete lists in combo ( #6949 )
'' 01/17/2014   DAJ         Handle negative ID's for OCO/OTO when parked ( #6949 )
'' 01/31/2014   DAJ         Have user specify TIF on contingecy orders
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
'' 08/27/2014   DAJ         Don't allow user to edit TIF values on existing non-parked order
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 11/14/2014   DAJ         Added support for MIT, On-Close, and On-Open orders
'' 12/04/2014   DAJ         Remove enabled symbol check for trading
'' 01/26/2015   DAJ         Added seconds to the Trigger and Expire time controls
'' 01/27/2015   DAJ         Fix for Sell Stop with Limit order
'' 10/08/2015   DAJ         Don't swap stop and limit if triggering order is a market
'' 10/26/2015   DAJ         Fix for CurrentPositionAfterTrigger -- keep walking up through all trigger orders
'' 11/06/2015   DAJ         New move with trigger flag on triggered orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kOnIcon$ = "kProcessingOn" ' "kCheckOn"
Private Const kOffIcon$ = "kProcessingOff" ' "kCheckOff"

Public Enum eGDTTEditOrderMode
    eGDTTEditOrderMode_Normal = 0
    eGDTTEditOrderMode_FromAlert
    eGDTTEditOrderMode_GameNewOrder
    eGDTTEditOrderMode_GameEditOrder
    eGDTTEditOrderMode_FromOrderGroup
End Enum

Private Enum eGDTabs
    eGDTab_Cancel = 0
    eGDTab_Trigger
    eGDTab_Condition
    eGDTab_Trail
    eGDTab_Contingency
End Enum

Private Type mPrivate
    Order As cPtOrder                   ' Order object for this order
    nOrderStatus As eTT_OrderStatus     ' Status of the order when form was loaded
    Mode As eGDTTEditOrderMode          ' Mode for this instance
    nReturn As eGDEditOrderReturn       ' Return value from the form

    Bars As cGdBars                     ' Bar properties for the chosen symbol
    Price As cPriceEditor               ' Editor for order price
    Above As cPriceEditor               ' Editor for "above" price
    Below As cPriceEditor               ' Editor for "below" price
    Qty As cPriceEditor                 ' Editor for quantity
    strPos As String                    ' Position of order (EL, ES, XL, XS)
    nCurPosition As Long                ' Current position of symbol in account
    bLoading As Boolean                 ' Is the form currently loading?
    strNumBarsInfo As String            ' Number of bars required information
    lGamePosition As Long               ' Game position
    
    CondBars As cGdBars                 ' Bar properties for conditional symbol
    CondPrice As cPriceEditor           ' Editor for the conditional price
    TrailAmount As cPriceEditor         ' Editor for the trailing point amount

    ListLoading As cListLoading         ' Lists for the TradeSense OCX
    bChange As Boolean                  ' Change or don't change the other control
    bInProgress As Boolean              ' Are we already in here?
    bSettingAutoExit As Boolean         ' Are we setting an auto exit?
    bNewOrder As Boolean                ' Is this a new order?
    bProfitDollar As Boolean            ' Is the Profit Target control set as dollars?
    bStopDollar As Boolean              ' Is the Stop Loss control set as dollars?
End Type
Private m As mPrivate

Public Property Get Symbol() As String
    Symbol = Trim(txtSymbol.Text)
End Property

Public Property Get Broker() As eTT_AccountType
    If cboAccounts.ListIndex >= 0 Then
        Broker = g.Broker.AccountTypeForName(cboAccounts.Text)
    Else
        Broker = -1
    End If
End Property

Private Function Tabs(ByVal nTab As eGDTabs) As Long
    Tabs = nTab
End Function

Private Function ConditionType(ByVal nCond As eGDConditionTypes) As Long
    ConditionType = nCond
End Function

Private Property Get AccountID() As Long
    If cboAccounts.ListIndex >= 0 Then
        AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    Else
        AccountID = 0&
    End If
End Property

Private Property Get SymbolOrSymbolID() As Variant
    
    Dim lSymbolID As Long               ' Symbol ID
    
    If Len(Trim(txtSymbol.Text)) > 0 Then
        lSymbolID = GetSymbolID(Trim(txtSymbol.Text))
        If lSymbolID = 0& Then
            SymbolOrSymbolID = Trim(txtSymbol.Text)
        Else
            SymbolOrSymbolID = lSymbolID
        End If
    Else
        SymbolOrSymbolID = ""
    End If
    
End Property

Public Property Get OrderID() As Long
    If Not m.Order Is Nothing Then
        OrderID = m.Order.OrderID
    Else
        OrderID = 0
    End If
End Property

Private Property Get FeedYardLotID() As String
    
    Dim strReturn As String             ' Return value for the property

    strReturn = ""
    If cboLots.ListIndex >= 0 Then
        If cboLots.ItemData(cboLots.ListIndex) >= 0 Then
            strReturn = Str(cboLots.ItemData(cboLots.ListIndex))
        End If
    End If
    
    FeedYardLotID = strReturn
    
End Property

Private Property Get ProfitDollar() As Double
    
    Dim dReturn As Double               ' Return value for the property

    If m.bProfitDollar = True Then
        dReturn = peProfitTarget.Price
    Else
        dReturn = peProfitTarget.Price * m.Bars.TickValuePerMove
    End If
    
    ProfitDollar = dReturn
    
End Property

Private Property Get ProfitPoints() As Double
    
    Dim dReturn As Double               ' Return value for the property

    If m.bProfitDollar = False Then
        dReturn = peProfitTarget.Price
    ElseIf m.Bars.TickValuePerMove = 0 Then
        dReturn = 0
    Else
        dReturn = peProfitTarget.Price / m.Bars.TickValuePerMove
    End If
    
    ProfitPoints = dReturn
    
End Property

Private Property Get StopDollar() As Double
    
    Dim dReturn As Double               ' Return value for the property

    If m.bStopDollar = True Then
        dReturn = peStopLoss.Price
    Else
        dReturn = peStopLoss.Price * m.Bars.TickValuePerMove
    End If
    
    StopDollar = dReturn
    
End Property

Private Property Get StopPoints() As Double
    
    Dim dReturn As Double               ' Return value for the property

    If m.bStopDollar = False Then
        dReturn = peStopLoss.Price
    ElseIf m.Bars.TickValuePerMove = 0 Then
        dReturn = 0
    Else
        dReturn = peStopLoss.Price / m.Bars.TickValuePerMove
    End If
    
    StopPoints = dReturn
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Order, Buy/Sell, Mode, Game Position, Market, Hide Park Button?,
''              Save?, Feed Yard Lot ID
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Order As cPtOrder, Optional ByVal nBuy As Byte = 255, _
    Optional ByVal Mode As eGDTTEditOrderMode = eGDTTEditOrderMode_Normal, _
    Optional ByVal nGamePosition As Long = 0, _
    Optional ByVal strMarket As String, _
    Optional ByVal bHideParkButton As Boolean = False, _
    Optional ByVal bSave As Boolean = True, _
    Optional ByVal strFeedYardLotID As String = "") As eGDEditOrderReturn
On Error GoTo ErrSection:

    Dim i&, strType$, strText$, bLowStop As Boolean
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim dMarketPrice As Double          ' Market price for the order

    If m.bInProgress = False Then
        m.bInProgress = True
    
        m.Mode = Mode
    
        optBuy.ForeColor = g.ChartGlobals.nLongColor
        optSell.ForeColor = g.ChartGlobals.nShortColor

        If IsGameMode Then
            'disable accounts combo if in game mode
            cboAccounts.Clear
            cboAccounts.BackColor = vbRed
            cboAccounts.Font.Bold = True
            cboAccounts.AddItem "Game Mode"
            cboAccounts.ListIndex = 0
            cboAccounts.Enabled = False
        Else
            Set m.Order = Order
            m.nOrderStatus = Order.Status
            
            ' Load accounts combo from the database...
            PopulateAccountsCbo cboAccounts, Order.AccountID, True
            
            ' If only 1 account, then disable...
            If cboAccounts.ListCount = 0 Then
                InfBox "There are currently no accounts to trade", "!", , "Order Error"
                m.bInProgress = False
                Exit Function
            End If
            With cboAccounts
                ' Only enable the accounts combo box if there is more than one account in the
                ' combo and it is either a new order or a parked order...
                .Enabled = (.ListCount > 1) And ((Order.OrderID <= 0) Or (Order.Status = eTT_OrderStatus_Parked))
                If .ListIndex < 0 Then .ListIndex = 0
            End With
        End If
                
        ' defaults
        txtSymbol = Trim(UCase(Order.Symbol))
        Set m.Bars = New cGdBars
        If Len(txtSymbol) > 0 Then
            'SetBarProperties m.Bars, txtSymbol
            DM_GetBars m.Bars, txtSymbol.Text, , LastDailyDownload
            g.RealTime.SpliceBars m.Bars
        End If
        If (Order.AccountID > 0) Or (IsGameMode = True) Then
            SetAccountCombo Order.AccountID
        ElseIf cboAccounts.ListCount = 1 Then
            cboAccounts.ListIndex = 0&
        Else
            cboAccounts.ListIndex = -1&
        End If
        
        If (m.Mode = eGDTTEditOrderMode_FromAlert) Or (m.Mode = eGDTTEditOrderMode_Normal) Or (m.Mode = eGDTTEditOrderMode_FromOrderGroup) Then
            LoadCancelOrdersCombo
            LoadTriggerOrdersCombo
        End If
        
        m.lGamePosition = nGamePosition
        
        m.bNewOrder = (Order.OrderID <= 0)
        If Order.OrderID <> 0 Then
            If Order.Buy Then
                optBuy = True
            Else
                optSell = True
            End If
            Select Case Order.OrderType
                Case eTT_OrderType_Market, eTT_OrderType_MarketOnClose
                    optMarket.Value = True
                    
                    optAnyTime.Value = (Order.OrderType = eTT_OrderType_Market)
                    optOnClose.Value = (Order.OrderType = eTT_OrderType_MarketOnClose)
                    optOnOpen.Value = (Order.OrderType = eTT_OrderType_MarketOnOpen)
            
                Case eTT_OrderType_StopWithLimit, eTT_OrderType_StopWithLimitCloseOnly
                    optStopLimit = True
                    If optBuy Then
                        bLowStop = True
                    End If
                    
                    optAnyTime.Value = (Order.OrderType = eTT_OrderType_StopWithLimit)
                    optOnClose.Value = (Order.OrderType = eTT_OrderType_StopWithLimitCloseOnly)
                    optOnOpen.Value = False
                
                Case eTT_OrderType_Stop, eTT_OrderType_StopCloseOnly
                    If optBuy Then
                        optAbove = True
                    Else
                        optBelow = True
                        bLowStop = True
                    End If
                    
                    optAnyTime.Value = (Order.OrderType = eTT_OrderType_Stop)
                    optOnClose.Value = (Order.OrderType = eTT_OrderType_StopCloseOnly)
                    optOnOpen.Value = False
                
                Case eTT_OrderType_Limit, eTT_OrderType_LimitCloseOnly
                    If optBuy Then
                        optBelow = True
                    Else
                        optAbove = True
                        bLowStop = True
                    End If
                    
                    optAnyTime.Value = (Order.OrderType = eTT_OrderType_Limit)
                    optOnClose.Value = (Order.OrderType = eTT_OrderType_LimitCloseOnly)
                    optOnOpen.Value = False
                
                Case eTT_OrderType_MIT
                    optMIT = True
                    If Not optBuy Then
                        bLowStop = True
                    End If
                    
                    optAnyTime.Value = True
                    optOnClose.Value = False
                    optOnOpen.Value = False
                    
            End Select
            
            If Order.Expiration = 0& Then
                optGTC.Value = True
                gdExpiration.Value = ExpirationDateForSymbol
            ElseIf Order.Expiration = -1& Then
                optDay.Value = True
                gdExpiration.Value = ExpirationDateForSymbol
            ElseIf Order.Expiration < 0& Then
                optDay.Value = True
                gdExpiration.Value = Order.Expiration * -1&
            Else
                optExpiration.Value = True
                gdExpiration.Value = Order.Expiration
            End If
            
            If UCase(Order.Session) = "E" Then
                optElectronic.Value = True
            Else
                optPit.Value = True
            End If
            
            If Order.CancelOrderID <> 0 Then
                chkCancelOrder.Value = vbChecked
                SetCancelOrderCombo Order.CancelOrderID
                tabSpecialOrders.TabPicture(Tabs(eGDTab_Cancel)) = Picture16(ToolbarIcon(kOnIcon))
            ElseIf Order.BrokerCancelOrderID <> 0 Then
                chkCancelOrder.Value = vbChecked
                SetCancelOrderCombo Abs(Order.BrokerCancelOrderID)
                tabSpecialOrders.TabPicture(Tabs(eGDTab_Cancel)) = Picture16(ToolbarIcon(kOnIcon))
            Else
                chkCancelOrder.Value = vbUnchecked
                tabSpecialOrders.TabPicture(Tabs(eGDTab_Cancel)) = Picture16(ToolbarIcon(kOffIcon))
            End If
        
            If Order.TriggerOrderID <> 0 Then
                chkTriggerOrder.Value = vbChecked
                SetTriggerOrderCombo Order.TriggerOrderID
                tabSpecialOrders.TabPicture(Tabs(eGDTab_Trigger)) = Picture16(ToolbarIcon(kOnIcon))
            Else
                chkTriggerOrder.Value = vbUnchecked
                tabSpecialOrders.TabPicture(Tabs(eGDTab_Trigger)) = Picture16(ToolbarIcon(kOffIcon))
            End If
            
            SetTriggerOptions Order.TriggerOptions
        Else
            If nBuy = 0 Then
                optSell = True
            ElseIf nBuy <> 255 Then
                optBuy = True
            Else
                optSell = False
                optBuy = False
            End If
            
            optMarket.Value = True
            optAnyTime.Value = True
            
            optDay.Value = True
            gdExpiration.Value = ExpirationDateForSymbol
            
            optPit.Value = True
            
            chkCancelOrder.Value = vbUnchecked
            tabSpecialOrders.TabPicture(Tabs(eGDTab_Cancel)) = Picture16(ToolbarIcon(kOffIcon))
            
            chkTriggerOrder.Value = vbUnchecked
            tabSpecialOrders.TabPicture(Tabs(eGDTab_Trigger)) = Picture16(ToolbarIcon(kOffIcon))
            
            optTriggerPartial.Value = True
            CheckBoxValue(chkTriggerConfirm) = g.Broker.ConfirmTriggered
            CheckBoxValue(chkMoveWithTrigger) = True
       End If
        
        If IsGameMode = False Then
            SetConditionOptions Order.ConditionOptions
            SetTrailControls
            SetContingencyControls Order.Contingency
        End If
            
        If m.Mode = eGDTTEditOrderMode_GameNewOrder And Len(strMarket) > 0 Then
            txtBelow = strMarket
            txtAbove = strMarket
        Else
            If bLowStop Then
                txtBelow = CStr(Order.StopPrice)
                If Order.OrderType = eTT_OrderType_MIT Then
                    txtAbove = CStr(Order.MitPrice)
                Else
                    txtAbove = CStr(Order.LimitPrice)
                End If
            Else
                txtAbove = CStr(Order.StopPrice)
                If Order.OrderType = eTT_OrderType_MIT Then
                    txtBelow = CStr(Order.MitPrice)
                Else
                    txtBelow = CStr(Order.LimitPrice)
                End If
            End If
        End If
               
        Set m.Price = New cPriceEditor
        Set m.Qty = New cPriceEditor
        txtQty = Str(Order.RemainingQuantity)
        'm.Qty.Init sbQty, txtQty, Nothing, Order.RemainingQuantity
        g.Broker.InitQuantityEditor m.Qty, sbQty, txtQty, AccountID, SymbolOrSymbolID, Order.RemainingQuantity
                    
        'SetPriceRange True, True
        
        Set m.Above = New cPriceEditor
        Set m.Below = New cPriceEditor
        Set m.CondPrice = New cPriceEditor
        Set m.TrailAmount = New cPriceEditor
        
        If Trim(txtAbove.Text) = Str(kNullData) Then
            If Trim(txtBelow.Text) = Str(kNullData) Then
                InitAbove m.Bars(eBARS_Close, m.Bars.Size - 1)
            Else
                InitAbove ValOfText(txtBelow)
            End If
        Else
            InitAbove ValOfText(txtAbove)
        End If
        If Trim(txtBelow.Text) = Str(kNullData) Then
            If Trim(txtAbove.Text) = Str(kNullData) Then
                InitBelow m.Bars(eBARS_Close, m.Bars.Size - 1)
            Else
                InitBelow ValOfText(txtAbove)
            End If
        Else
            InitBelow ValOfText(txtBelow)
        End If
        
        If IsGameMode = True Then
            m.nCurPosition = nGamePosition
        Else
            If Trim(txtConditionPrice.Text) = "0" Then
                InitConditionalPrice m.CondBars(eBARS_Close, m.CondBars.Size - 1)
            Else
                InitConditionalPrice ValOfText(txtConditionPrice.Text)
            End If
            If Trim(txtPointTrail.Text) = "0" Then
                m.TrailAmount.Init sbPointTrail, txtPointTrail, m.Bars
            Else
                m.TrailAmount.Init sbPointTrail, txtPointTrail, m.Bars, ValOfText(txtPointTrail.Text)
            End If
            'm.nCurPosition = CurrentPosition(cboAccounts.ItemData(cboAccounts.ListIndex), txtSymbol.Text)
            m.nCurPosition = CurrentPositionAfterTrigger
        End If
        
        If Order.ExitPos > 0& Then
            chkExitAll.Value = vbChecked
        Else
            chkExitAll.Value = vbUnchecked
        End If
            
        FixControls (Order.RemainingQuantity = 0)
        
        If IsOpenOrder(Order.Status) = False Then
            cmdOK.Enabled = False
            cmdPark.Enabled = False
        ElseIf HasBeenSent(Order.Status) Then
            optBuy.Enabled = False
            optSell.Enabled = False
            txtSymbol.Enabled = False
            cmdLookup.Enabled = False
            
            Enable chkTriggerOrder, NotSent(m.nOrderStatus)
            Enable cboTriggerOrders, NotSent(m.nOrderStatus)
            Enable chkBeginTime, NotSent(m.nOrderStatus)
            Enable chkPrice, NotSent(m.nOrderStatus)
            Enable chkCustomCondition, NotSent(m.nOrderStatus)
            Enable chkTrail, NotSent(m.nOrderStatus)
            
            If chkCancelOrder = vbChecked Or chkTriggerOrder = vbChecked Or chkBeginTime = vbChecked Or chkEndTime = vbChecked Or chkPrice = vbChecked Or chkCustomCondition = vbChecked Or chkTrail = vbChecked Then
                chkAdvanced.Value = vbChecked
                chkAdvanced.Enabled = False
            End If
            
'            If (g.Broker.AccountTypeForID(Order.AccountID) <> eTT_AccountType_LindWaldock) And (g.Broker.AccountTypeForID(Order.AccountID) <> eTT_AccountType_ManExpress) Then
                optBelow.Enabled = False
                optAbove.Enabled = False
                optMarket.Enabled = False
                optStopLimit.Enabled = False
                optMIT.Enabled = False
'            End If
        End If
        
        Select Case m.Mode
            Case eGDTTEditOrderMode_Normal
                cmdOK.Caption = "Submit &Order"
                cmdOK.Visible = True
                cmdPark.Caption = "&Park Order"
                cmdPark.Visible = Not bHideParkButton
                If Order.IsAutomated Then
                    optBuy.Enabled = False
                    optSell.Enabled = False
                    txtSymbol.Visible = True
                    txtSymbol.Enabled = False
                    cmdLookup.Visible = True
                    cmdLookup.Enabled = False
                    cboAccounts.Enabled = False
                    fraDuration.Visible = True
                    CheckBoxValue(chkAdvanced) = False
                    chkAdvanced.Visible = True
                    chkAdvanced.Enabled = False
                Else
                    txtSymbol.Visible = True
                    txtSymbol.Enabled = True
                    cmdLookup.Visible = True
                    cmdLookup.Enabled = True
                    fraDuration.Visible = True
                    chkAdvanced.Visible = True
                End If
                
            Case eGDTTEditOrderMode_FromAlert
                cmdOK.Caption = "Save &Order"
                cmdOK.Visible = True
                cmdPark.Caption = "&Park Order"
                cmdPark.Visible = False
                cmdCancel.Top = cmdPark.Top
                txtSymbol.Visible = True
                txtSymbol.Enabled = True            '01-29-2007: Pete wants ability to change symbols
                cmdLookup.Visible = True
                cmdLookup.Enabled = True
                fraDuration.Visible = True
                chkAdvanced.Value = vbUnchecked
                chkAdvanced.Visible = False
                
            Case eGDTTEditOrderMode_GameNewOrder
                cmdOK.Caption = "Place &Order"
                cmdOK.Visible = True
                cmdPark.Caption = "&Park Order"
                cmdPark.Visible = False
                cmdCancel.Top = cmdPark.Top
                txtSymbol.Visible = True
                txtSymbol.Enabled = False
                cmdLookup.Visible = False
                cmdLookup.Enabled = False
                optGTC.Value = True
                fraDuration.Visible = False
                chkAdvanced.Visible = False
                
            Case eGDTTEditOrderMode_GameEditOrder
                cmdOK.Caption = "Save &Order"
                cmdOK.Visible = True
                cmdPark.Caption = "&Delete Order"
                cmdPark.Visible = Not bHideParkButton
                txtSymbol.Visible = True
                txtSymbol.Enabled = False
                cmdLookup.Visible = False
                cmdLookup.Enabled = False
                optGTC.Value = True
                fraDuration.Visible = False
                chkAdvanced.Visible = False
                
            Case eGDTTEditOrderMode_FromOrderGroup
                cmdOK.Caption = "Save &Order"
                cmdOK.Visible = True
                cmdPark.Caption = "&Park Order"
                cmdPark.Visible = False
                cmdCancel.Top = cmdPark.Top
                txtSymbol.Visible = True
                txtSymbol.Enabled = True
                cmdLookup.Visible = True
                cmdLookup.Enabled = True
                fraDuration.Visible = True
                chkAdvanced.Value = vbUnchecked
                chkAdvanced.Visible = False
        
        End Select
        
        If IsGameMode = True Then
            cmdJournal.Visible = False
        Else
            cmdJournal.Visible = (Order.OrderID > 0)
            EnableAdvancedControls
        End If
        
        EnableQuantityControls
        EnableDurationControls
        EnableOrderTimeControls
        
        ShowSessionControls
        ShowExchangeControls
        'If (ShowLotControls = True) Or (Len(strFeedYardLotID) > 0) Then
            LoadLotsCombo strFeedYardLotID
        'End If
        
        GetContractInformation
        
        m.bLoading = True
        ShowForm Me, eForm_ActModal, frmMain
        
        If m.nReturn <> eGDEditOrderReturn_Cancel Then
            With Order
                If IsGameMode = True Then
                    If m.strPos = "ES" Or m.strPos = "EL" And m.nReturn = eGDEditOrderReturn_Submit Then
                        .Enter = True
                    End If
                Else
                    .AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
                    If Len(.GenesisOrderID) = 0 Then .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(.AccountID)) 'cboAccounts.Text)
                End If
                .Buy = optBuy.Value
                .SymbolOrSymbolID = txtSymbol.Text
                .Quantity = m.Qty.Price + .FillQuantity
                Select Case True
                    Case optMarket
                        If optAnyTime.Value = True Then
                            .OrderType = eTT_OrderType_Market
                        ElseIf optOnClose.Value = True Then
                            .OrderType = eTT_OrderType_MarketOnClose
                        ElseIf optOnOpen.Value = True Then
                            .OrderType = eTT_OrderType_MarketOnOpen
                        End If
                        
                    Case optStopLimit
                        If optAnyTime.Value = True Then
                            .OrderType = eTT_OrderType_StopWithLimit
                        ElseIf optOnClose.Value = True Then
                            .OrderType = eTT_OrderType_StopWithLimitCloseOnly
                        End If
                        
                        If optBuy Then
                            .StopPrice = m.Below.Price
                            .LimitPrice = m.Above.Price
                        Else
                            .StopPrice = m.Above.Price
                            .LimitPrice = m.Below.Price
                        End If

                    Case optAbove
                        If optBuy Then
                            If optAnyTime.Value = True Then
                                .OrderType = eTT_OrderType_Stop
                            ElseIf optOnClose.Value = True Then
                                .OrderType = eTT_OrderType_StopCloseOnly
                            End If
                            
                            .StopPrice = m.Above.Price
                            .LimitPrice = 0
                        Else
                            If optAnyTime.Value = True Then
                                .OrderType = eTT_OrderType_Limit
                            ElseIf optOnClose.Value = True Then
                                .OrderType = eTT_OrderType_LimitCloseOnly
                            End If
                            
                            .LimitPrice = m.Above.Price
                            .StopPrice = 0
                        End If
                    
                    Case optBelow
                        If optBuy Then
                            If optAnyTime.Value = True Then
                                .OrderType = eTT_OrderType_Limit
                            ElseIf optOnClose.Value = True Then
                                .OrderType = eTT_OrderType_LimitCloseOnly
                            End If
                            
                            .LimitPrice = m.Below.Price
                            .StopPrice = 0
                        Else
                            If optAnyTime.Value = True Then
                                .OrderType = eTT_OrderType_Stop
                            ElseIf optOnClose.Value = True Then
                                .OrderType = eTT_OrderType_StopCloseOnly
                            End If
                            
                            .StopPrice = m.Below.Price
                            .LimitPrice = 0
                        End If
                    
                    Case optMIT
                        If optBuy Then
                            .OrderType = eTT_OrderType_MIT
                            .MitPrice = m.Below.Price
                            .StopPrice = 0
                        Else
                            .OrderType = eTT_OrderType_MIT
                            .MitPrice = m.Above.Price
                            .StopPrice = 0
                        End If
                End Select
                
                Select Case True
                    Case optGTC
                        .Expiration = 0&
                    Case optDay
                        .Expiration = ExpirationDateForSymbol * -1&
                    Case optExpiration
                        .Expiration = gdExpiration.Value
                End Select
                
                .Session = ""
                
                If cboExchanges.ListIndex >= 0 Then
                    .Exchange = cboExchanges.Text
                Else
                    .Exchange = ""
                End If
                
                ' Let the SendOrder routine set the order date and do a RefreshOrder so that the
                ' Order Submitted alert works.  04/05/2007 DAJ
                ''.OrderDate = Now

                If chkCancelOrder.Value = vbChecked Then
                    If chkBrokerOCO.Value = vbChecked Then
                        .CancelOrderID = 0
                        If Abs(.BrokerCancelOrderID) <> cboCancelOrders.ItemData(cboCancelOrders.ListIndex) Then
                            .BrokerCancelOrderID = cboCancelOrders.ItemData(cboCancelOrders.ListIndex) * -1
                        End If
                    Else
                        .CancelOrderID = cboCancelOrders.ItemData(cboCancelOrders.ListIndex)
                        .BrokerCancelOrderID = 0
                    End If
                Else
                    .CancelOrderID = 0&
                    .BrokerCancelOrderID = 0&
                End If
                
                If chkTriggerOrder.Value = vbChecked Then
                    .TriggerOrderID = cboTriggerOrders.ItemData(cboTriggerOrders.ListIndex)
                    If optTriggerPartial.Value = True Then
                        .TriggerOptions = "1"
                    Else
                        .TriggerOptions = "0"
                    End If
                    If chkTriggerConfirm.Value = vbChecked Then
                        .TriggerOptions = .TriggerOptions & ",1"
                    Else
                        .TriggerOptions = .TriggerOptions & ",0"
                    End If
                    If chkRelativePrice.Value = vbChecked Then
                        .TriggerOptions = .TriggerOptions & "," & Trim(txtRelativePrice.Text)
                    Else
                        .TriggerOptions = .TriggerOptions & ","
                    End If
                    .TriggerOptions = .TriggerOptions & "," & Str(CheckBoxValue(chkMoveWithTrigger))
                Else
                    .TriggerOrderID = 0&
                    .TriggerOptions = ""
                End If
                
                .ConditionOptions = GetConditionOptions
                
                If chkTrail.Value = vbChecked Then
                    If optTrailDollar Then
                        .TrailOptions = "0,0" ' & Str(ConvertTimeZone(g.RealTime.FeedTime, "NY", m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))
                        .TrailAmount = ValOfText(txtDollarTrail.Text)
                    Else
                        .TrailOptions = "1,0" ' & Str(ConvertTimeZone(g.RealTime.FeedTime, "NY", m.Bars.Prop(eBARS_ExchangeTimeZoneInf)))
                        .TrailAmount = m.TrailAmount.Price
                    End If
                    
''                    .StopPrice = .TrailingStopValue(True)
                    dMarketPrice = .MarketPrice(False)
                    .StopPrice = .TrailingStopValue(dMarketPrice, dMarketPrice)
                Else
                    .TrailAmount = 0#
                End If
                
                If chkExitAll.Value = vbChecked Then
                    .ExitPos = 100&
                Else
                    .ExitPos = 0&
                End If
                
                .Contingency = GetContingencyControls
                
                If IsGameMode = True Then
                    'do nothing
                Else
                    If bSave Then
                        If (m.Mode <> eGDTTEditOrderMode_FromAlert) And (m.Mode <> eGDTTEditOrderMode_FromOrderGroup) Then
                            .Save
                        End If
                    End If
                    
                    SetIniFileProperty "LastAccount", Order.AccountID, "TTSummary", g.strIniFile
                    
                    If m.nCurPosition = 0 Then SetDefaultEntryForSymbol .SymbolID, .Symbol, .Quantity
                End If
            
                If (m.nReturn = eGDEditOrderReturn_Submit) And (m.bNewOrder = True) And (cboLots.ListIndex > 0) Then
                    g.CattleBridge.SetUpNewOrder .GenesisOrderID, Str(cboLots.ItemData(cboLots.ListIndex))
                End If
            End With
        End If
        
        ShowMe = m.nReturn
        m.bInProgress = False
    End If
    
ErrExit:
    If m.bInProgress = False Then Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTEditOrder.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExchanges
'' Description: Load the exchange combo and set the default exchange
'' Inputs:      Exchange List, Default Exchange
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetExchanges(ByVal strExchangeList As String, ByVal strDefaultExchange As String)
On Error GoTo ErrSection:

    Dim astrExchanges As cGdArray       ' Array of exchanges
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Split out the exchange information into an array...
    Set astrExchanges = New cGdArray
    astrExchanges.SplitFields strExchangeList, ","
    
    With cboExchanges
        .Clear
        
        ' Load up combo box with the comma delimited list of exchanges in strExchangeList...
        For lIndex = 0 To astrExchanges.Size - 1
            If Len(astrExchanges(lIndex)) > 0 Then
                cboExchanges.AddItem astrExchanges(lIndex)
            End If
        Next lIndex
        
        ' Set the combo box to the string passed in strDefaultExchange if it exists...
        If Len(strDefaultExchange) > 0 Then
            For lIndex = 0 To cboExchanges.ListCount - 1
                If cboExchanges.List(lIndex) = strDefaultExchange Then
                    cboExchanges.ListIndex = lIndex
                    Exit For
                End If
            Next lIndex
        End If
    End With
    
    InfBox ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetExchanges"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshOrder
'' Description: Refresh the local order with the order passed in
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    If m.Order.OrderID = Order.OrderID Then
        Set m.Order = Order
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.RefreshOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: Reset controls when the account changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        GetContractInformation
    
        'm.nCurPosition = CurrentPosition(cboAccounts.ItemData(cboAccounts.ListIndex), txtSymbol.Text)
        m.nCurPosition = CurrentPositionAfterTrigger
        FixControls True
        
        LoadCancelOrdersCombo
        LoadTriggerOrdersCombo
        
        FixControls
        
        EnableDurationControls
        EnableOrderTimeControls
        
        ShowSessionControls
        ShowExchangeControls
        ShowLotControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cboAccounts.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboLots_Click
'' Description: Handle the user changing the feed yard lot
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboLots_Click()
On Error GoTo ErrSection:

    If Visible Then
        m.nCurPosition = CurrentPositionAfterTrigger
        FixControls True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cboLots_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboTriggerOrders_Click
'' Description: Handle the user changing the triggering order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboTriggerOrders_Click()
On Error GoTo ErrSection:

    If Visible Then
        SwapStopLimit
    
        m.nCurPosition = CurrentPositionAfterTrigger
        FixControls True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cboTriggerOrders_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAdvanced_Click
'' Description: Show/Hide the special order tabs as applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAdvanced_Click()
On Error GoTo ErrSection:

    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkAdvanced.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAutoExit_Click
'' Description: Allow user to turn on/turn off an auto exit
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAutoExit_Click()
On Error GoTo ErrSection:

    Dim strAutoExit As String           ' Auto Exit selected
    
    If Visible And (Not m.bSettingAutoExit) Then
        If chkAutoExit.Value = vbChecked Then
            If CanActivateAutomatedItem(AccountID, SymbolOrSymbolID, "Auto Exit", "Order Form") Then
                strAutoExit = frmOrderStrategies.ShowMe(AccountID, SymbolOrSymbolID)
                If Len(strAutoExit) > 0 Then
                    g.OrderStrategies.ActivateExit AccountID, SymbolOrSymbolID, strAutoExit
                    chkAutoExit.Caption = "A&uto Exit: " & FileBase(strAutoExit)
                Else
                    chkAutoExit.Value = vbUnchecked
                    chkAutoExit.Caption = "A&uto Exit: " & "None"
                End If
            Else
                chkAutoExit.Value = vbUnchecked
                chkAutoExit.Caption = "A&uto Exit: " & "None"
            End If
        Else
            g.OrderStrategies.DeactivateExit AccountID, SymbolOrSymbolID, True, "Turned off from Order Form"
            chkAutoExit.Caption = "A&uto Exit: " & "None"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkAutoExit_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAutoJournal_Click
'' Description: Allow user to turn on/turn off the automatic journal popup
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAutoJournal_Click()
On Error GoTo ErrSection:

    If Visible Then
        g.Broker.AutoJournalPopUp = (chkAutoJournal.Value = vbChecked)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkAutoJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkBeginTime_Click
'' Description: Allow user to turn on/turn off a trigger time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkBeginTime_Click()
On Error GoTo ErrSection:

    If chkBeginTime = vbChecked Or chkEndTime = vbChecked Or chkPrice = vbChecked Or chkCustomCondition = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOffIcon))
    End If

    If Me.Visible Then
        EnableAdvancedControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkBeginTime.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkBrokerOCO_Click
'' Description: Save the value if the user changes it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkBrokerOCO_Click()
On Error GoTo ErrSection:

    If CheckBoxValue(chkBrokerOCO) <> g.Broker.HoldOcoAtBroker(cboAccounts.Text) Then
        g.Broker.HoldOcoAtBroker(cboAccounts.Text) = CheckBoxValue(chkBrokerOCO)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkBrokerOCO_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCancelOrder_Click
'' Description: Set the tab picture as the user clicks on the check box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCancelOrder_Click()
On Error GoTo ErrSection:

    If Visible Then
        If chkCancelOrder.Value = vbChecked Then
            tabSpecialOrders.TabPicture(Tabs(eGDTab_Cancel)) = Picture16(ToolbarIcon(kOnIcon))
        Else
            tabSpecialOrders.TabPicture(Tabs(eGDTab_Cancel)) = Picture16(ToolbarIcon(kOffIcon))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkCancelOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCustomCondition_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCustomCondition_Click()
On Error GoTo ErrSection:

    If chkBeginTime = vbChecked Or chkEndTime = vbChecked Or chkPrice = vbChecked Or chkCustomCondition = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOffIcon))
    End If

    If Me.Visible Then
        If (chkCustomCondition.Value = vbChecked) And (g.RealTime.IsServerActive(True) = False) Then
            InfBox "You cannot set up a conditional order unless real-time streaming is active", "!", , "Order Error"
            chkCustomCondition.Value = vbUnchecked
            GoTo ErrExit
        End If
        
        EnableAdvancedControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkCustomCondition_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkEndTime_Click
'' Description: Allow user to turn on/turn off an expiration time
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkEndTime_Click()
On Error GoTo ErrSection:

    If chkBeginTime = vbChecked Or chkEndTime = vbChecked Or chkPrice = vbChecked Or chkCustomCondition = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOffIcon))
    End If

    If Me.Visible Then
        EnableAdvancedControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkEndTime.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkExitAll_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkExitAll_Click()
On Error GoTo ErrSection:

    If Visible Then FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkExitAll_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkPrice_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkPrice_Click()
On Error GoTo ErrSection:

    If chkBeginTime = vbChecked Or chkEndTime = vbChecked Or chkPrice = vbChecked Or chkCustomCondition = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOffIcon))
    End If

    If Me.Visible Then
        If (chkPrice.Value = vbChecked) And (g.RealTime.IsServerActive(True) = False) Then
            InfBox "You cannot set up a conditional order unless real-time streaming is active", "!", , "Order Error"
            chkPrice.Value = vbUnchecked
            GoTo ErrExit
        End If
        
        EnableAdvancedControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkPrice_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTrail_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTrail_Click()
On Error GoTo ErrSection:

    If chkTrail.Value = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Trail)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Trail)) = Picture16(ToolbarIcon(kOffIcon))
    End If
    
    If Me.Visible Then
        If (chkTrail.Value = vbChecked) And (g.RealTime.IsServerActive(True) = False) Then
            InfBox "You cannot set up a trailing stop order unless real-time streaming is active", "!", , "Order Error"
            chkTrail.Value = vbUnchecked
            GoTo ErrExit
        End If
        
        EnableAdvancedControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkTrail.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTriggerConfirm_Click
'' Description: Change the value of the global trigger confirmation
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTriggerConfirm_Click()
On Error GoTo ErrSection:

    If Visible Then
        g.Broker.ConfirmTriggered = CheckBoxValue(chkTriggerConfirm)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkTriggerConfirm_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkTriggerOrder_Click
'' Description: Set the picture as the user turns the check on and off
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkTriggerOrder_Click()
On Error GoTo ErrSection:
    
    Dim dTriggerPrice As Double         ' Trigger price
    
    If chkTriggerOrder.Value = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Trigger)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Trigger)) = Picture16(ToolbarIcon(kOffIcon))
    End If
    
    If Visible Then
        SwapStopLimit
    End If
    
    m.nCurPosition = CurrentPositionAfterTrigger
    FixControls
    EnableAdvancedControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.chkTriggerOrder_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel out of the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.nReturn = eGDEditOrderReturn_Cancel
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCondSymLookup_Click
'' Description: Allow the user to lookup a condition symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCondSymLookup_Click()
On Error GoTo ErrSection:

    LookupConditionSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdCondSymLookup.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditCustom_Click
'' Description: Allow the user to edit the custom condition
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditCustom_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the custom form
    Dim strPeriod As String             ' Default period for Market1
    Dim bOverride As Boolean            ' Override the number of bars calculated?
    Dim lNumBarsCalc As Long            ' Number of bars calculated
    Dim lNumBarsOver As Long            ' Number of bars override
    
    strPeriod = Parse(lblDefaultPeriod.Caption, ":", 2)
    If Len(m.strNumBarsInfo) = 0 Then
        bOverride = False
        lNumBarsCalc = 0&
        lNumBarsOver = 0&
    Else
        bOverride = CBool(Parse(m.strNumBarsInfo, "|", 1))
        lNumBarsCalc = CLng(Val(Parse(m.strNumBarsInfo, "|", 2)))
        lNumBarsOver = CLng(Val(Parse(m.strNumBarsInfo, "|", 3)))
    End If
    
    strReturn = frmCustomCondition.ShowMe(rtfCondition.Text, txtSymbol.Text, strPeriod, bOverride, lNumBarsCalc, lNumBarsOver)
    If Len(strReturn) > 0 Then
        SetConditionRTF strReturn
        lblDefaultPeriod.Caption = "Default Period: " & strPeriod
        m.strNumBarsInfo = Str(bOverride) & "|" & Str(lNumBarsCalc) & "|" & Str(lNumBarsOver)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdEditCustom_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdJournal_Click
'' Description: Allow the user to add/edit/view journal entries for this order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdJournal_Click()
On Error GoTo ErrSection:

    g.TnJournal.ShowOrderJournal m.Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdJournal.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to lookup a symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdLookup.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the user to save and submit the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    MoveFocus cmdOK
    
    ' 03/11/2010 DAJ: Need a DoEvents here so that the LostFocus event on the
    ' cPriceEditor text boxes can trigger propertly (#5659)...
    DoEvents
    
    If VerifyOrder(False) = True Then
        If (IsGameMode = True) Or (g.Broker.ConnectionStatusForAccount(AccountID) = eGDConnectionStatus_Connected) Then
            m.nReturn = eGDEditOrderReturn_Submit
        ElseIf InfBox("You are not currently connected to this account.  Would you like to Park the order instead?", "?", "+Yes|-No", "Error") = "Y" Then
            m.nReturn = eGDEditOrderReturn_Park
        End If
    
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPark_Click
'' Description: Allow the user to save and park the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPark_Click()
On Error GoTo ErrSection:

    If VerifyOrder(True) = True Then
        m.nReturn = eGDEditOrderReturn_Park
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.cmdPark.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Make sure that the quantity has the focus when first started
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bLoading Then
        MoveFocus txtQty
        m.bLoading = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTTEditOrder.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyPress
'' Description: Handle some shortcut key presses for the form
'' Inputs:      ASCII version of key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

#If 0 Then
    Select Case UCase(Chr(KeyAscii))
        Case "B"
            KeyAscii = 0
            If optBuy.Enabled Then optBuy = True
        Case "S"
            KeyAscii = 0
            If optSell.Enabled Then optSell = True
        Case "M"
            KeyAscii = 0
            If optMarket.Enabled Then optMarket = True
            MoveFocus txtPrice
        Case "U"
            KeyAscii = 0
            If optAbove.Enabled Then optAbove = True
            MoveFocus txtPrice
        Case "D"
            KeyAscii = 0
            If optBelow.Enabled Then optBelow = True
            MoveFocus txtPrice
        Case "P"
            KeyAscii = 0
            MoveFocus txtPrice
        Case "Q"
            KeyAscii = 0
            MoveFocus txtQty
    End Select
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement information from the ini file
    
    Icon = Picture16(ToolbarIcon("kDollarSign"))
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTTEditOrder", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
        'If Width < 9525 Then Width = 9525
        If Width < 10515 Then Width = 10515
    End If
    tabSpecialOrders.CurrTab = GetIniFileProperty("CurrTab", Tabs(eGDTab_Cancel), "TTEditOrder", g.strIniFile)
    chkAdvanced.Value = GetIniFileProperty("Advanced", vbUnchecked, "TTEditOrder", g.strIniFile)
    
    imgCalc.ToolTipText = vsCalc.ToolTipText
    
    With tabSpecialOrders
        .align = asBottom
        .Style = tsSlanted
        .BoldCurrent = True
    End With
    
    With cboConditionField
        .AddItem "Session Open"
        .AddItem "Session High"
        .AddItem "Session Low"
        .AddItem "Last Price"
        
        .ListIndex = 3
    End With
    
    With cboConditionOperator
        .AddItem "="
        .AddItem "<>"
        .AddItem ">="
        .AddItem "<="
        .AddItem ">"
        .AddItem "<"
        
        .ListIndex = 0
    End With
    
    If m.ListLoading Is Nothing Then
        'Load internally generated TradeSense lists (Symbols, etc.)
        Set m.ListLoading = New cListLoading
        m.ListLoading.Load
    End If
    
    ''tabSpecialOrders.TabVisible(Tabs(eGDTab_Condition)) = FileExist(AddSlash(App.Path) & "ShowCond.FLG")
    ''tabSpecialOrders.TabVisible(Tabs(eGDTab_Trail)) = FileExist(AddSlash(App.Path) & "ShowCond.FLG")
    
    chkRelativePrice.Visible = False
    txtRelativePrice.Visible = False
    
    chkExitAll.Visible = False
    
    tmrWait.Interval = 1000
    tmrWait.Enabled = False

    If g.Broker.AutoJournalPopUp Then
        chkAutoJournal.Value = vbChecked
    Else
        chkAutoJournal.Value = vbUnchecked
    End If
    chkAutoJournal.ToolTipText = "Automatically show journal dialog when an order gets manually submitted, amended, or cancelled."
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceRange
'' Description: Set the price ranges accordingly
'' Inputs:      Reset Price?, Reset Quantity?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceRange(Optional ByVal bResetPrice As Boolean = True, Optional ByVal bResetQty As Boolean = False)

    Dim i&, nLastBar&, dPrice#, dHigh#, dLow#, dMinMove#, strText$, nQty&
    Dim Chart As cChart
   
Exit Sub
    
    i = cboAccounts.ItemData(cboAccounts.ListIndex)
    Set Chart = Forms(i).Chart
    With Chart
        m.nCurPosition = .Position
        DisplayPosition m.nCurPosition, lblCurPos
        If optBuy + optSell = 0 Then
            If m.nCurPosition > 0 Then
                optSell = True
            Else
                optBuy = True
            End If
        End If
        #If 0 Then
        If m.nCurPosition > 0 Then
            ' Long
            'optAction(0).Caption = "EXIT &Long  (Sell 1)"
            'optAction(1).Caption = "ENTER &Short  (Sell 2)"
            optBuy.Enabled = False
            optSell.Enabled = True
            optSell = True
        ElseIf m.nCurPosition < 0 Then
            ' Short
            'optAction(0).Caption = "ENTER &Long  (Buy 2)"
            'optAction(1).Caption = "EXIT &Short  (Buy 1)"
            optBuy.Enabled = True
            optSell.Enabled = False
            optBuy = True
        Else
            ' None
            'optAction(0).Caption = "Enter &LONG  (Buy 1)"
            'optAction(1).Caption = "Enter &SHORT  (Sell 1)"
            optBuy.Enabled = True
            optSell.Enabled = True
        End If
        #End If
        'DrawTriangle optAction(0)
        'DrawTriangle optAction(1)
        
        If bResetQty Then
            nQty = Abs(m.nCurPosition)
            If nQty = 0 Then nQty = 1
            m.Qty = nQty
        End If
        nQty = ValOfText(txtQty) 'm.Qty
                
        If nQty = 0 Then
            lblNewPos.Caption = ""
            strText = ""
        ElseIf optSell Then
            optAbove.Caption = "LIMIT" ' - if price gets up to"
            optBelow.Caption = "STOP" ' - if price gets down to"
            DisplayPosition m.nCurPosition - nQty, lblNewPos
            If nQty > Abs(m.nCurPosition) Then
                m.strPos = "ES"
            Else
                m.strPos = "XL"
            End If
            If m.nCurPosition < 0 Then
                m.strPos = "ES"
                strText = "ADD to Short"
            ElseIf m.nCurPosition = 0 Then
                m.strPos = "ES"
                strText = "ENTER Short"
            ElseIf nQty = Abs(m.nCurPosition) Then
                m.strPos = "XL"
                strText = "EXIT from Long"
            ElseIf nQty < Abs(m.nCurPosition) Then
                m.strPos = "XL"
                strText = "REDUCE Long"
            Else
                m.strPos = "ES"
                strText = "REVERSE to Short"
            End If
        Else
            optAbove.Caption = "STOP" ' - if price gets up to"
            optBelow.Caption = "LIMIT" ' - if price gets down to"
            DisplayPosition m.nCurPosition + nQty, lblNewPos
            If m.nCurPosition > 0 Then
                m.strPos = "EL"
                strText = "ADD to Long"
            ElseIf m.nCurPosition = 0 Then
                m.strPos = "EL"
                strText = "ENTER Long"
            ElseIf nQty = Abs(m.nCurPosition) Then
                m.strPos = "XS"
                strText = "EXIT from Short"
            ElseIf nQty < Abs(m.nCurPosition) Then
                m.strPos = "XS"
                strText = "REDUCE Short"
            Else
                m.strPos = "EL"
                strText = "REVERSE to Long"
            End If
        End If
        lblAction.Caption = strText
        'If InStr(m.strPos, "S") Then
        If optSell Then
            lblAction.ForeColor = g.ChartGlobals.nShortColor
        Else
            lblAction.ForeColor = g.ChartGlobals.nLongColor
        End If
        
        dMinMove = .Bars.Prop(eBARS_TickMove) * .Bars.Prop(eBARS_MinMoveInTicks)
        nLastBar = .LastGoodDataBar(False)
        dHigh = .Bars(eBARS_High, nLastBar)
        dLow = .Bars(eBARS_Low, nLastBar)
        dPrice = .Bars(eBARS_Close, nLastBar)
        If optAbove Then
            dLow = dPrice + dMinMove
            dPrice = dHigh + dMinMove
            dHigh = 999999
            ''txtPrice.Top = optAbove.Top - 45
        ElseIf optBelow Then
            dHigh = dPrice - dMinMove
            dPrice = dLow - dMinMove
            dLow = dMinMove
            ''txtPrice.Top = optBelow.Top - 45
        Else
            ''txtPrice.Top = optMarket.Top - 45
        End If
    End With
    With m.Price
        If .Price >= dLow And .Price <= dHigh And Not bResetPrice Then
            dPrice = .Price '(keep same price if in range)
        End If
        ''.Init sbPrice, txtPrice, Chart.Bars, dPrice, dLow, dHigh
    End With
    
    DrawOrder
    
    Set Chart = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel the unload and let the ShowMe handle it
'' Inputs:      Whether to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.nReturn = eGDEditOrderReturn_Cancel
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lWidthDiff As Long              ' Difference between the scale width and width
    Dim lHeightDiff As Long             ' Difference between the scale height and height
    Dim lTop As Long                    ' Top of the tab

    lWidthDiff = Width - ScaleWidth
    lHeightDiff = Height - ScaleHeight
    
    If IsGameMode = True Then
        tabSpecialOrders.Visible = False
        Move Left, Top, Width, fraPosition.Height * 2 + 60      '4128
    Else
        fraTime.Move fraSession.Left, fraSession.Top
        fraTime.Visible = True
        
        If fraSession.Visible Or fraTime.Visible Or lblExchange.Visible Or lblLot.Visible Then
            fraAdvanced.Move 120, fraSession.Top + fraSession.Height + 60
        Else
            fraAdvanced.Move 120, fraDuration.Top + fraDuration.Height + 60
        End If
        
        If lblLot.Visible Then
            If lblExchange.Visible Then
                lblLot.Top = lblExchange.Top + 360
                cboLots.Top = cboExchanges.Top + 360
            Else
                lblLot.Top = lblExchange.Top
                cboLots.Top = cboExchanges.Top
            End If
        End If
        
        lTop = fraAdvanced.Top + fraAdvanced.Height + 60
        
        If chkAdvanced.Value = vbChecked Then
            tabSpecialOrders.Visible = True
            Move Left, Top, Width, lTop + tabSpecialOrders.Height + lHeightDiff + 60
        Else
            tabSpecialOrders.Visible = False
            Move Left, Top, Width, lTop + lHeightDiff
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Cleanup after ourselves when the form gets unloaded
'' Inputs:      Whether to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrWait.Enabled = False

    Set m.Order = Nothing
    Set m.Above = Nothing
    Set m.Below = Nothing
    Set m.Price = Nothing
    Set m.Qty = Nothing

    SetIniFileProperty "frmTTEditOrder", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "CurrTab", tabSpecialOrders.CurrTab, "TTEditOrder", g.strIniFile
    SetIniFileProperty "Advanced", chkAdvanced.Value, "TTEditOrder", g.strIniFile
    SetIniFileProperty "DollarTrail", ValOfText(txtDollarTrail.Text), "TTEditOrder", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    imgCalc_Click
'' Description: Allow the user to do some money management calculations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub imgCalc_Click()
On Error GoTo ErrSection:

    MMcalc

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.imgCalc_Click", eGDRaiseError_Show

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAbove_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAbove_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
        MoveFocus txtAbove
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optAbove.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optBelow_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optBelow_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
        MoveFocus txtBelow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optBelow.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optBuy_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optBuy_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls True
        MoveFocus txtQty
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optBuy.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optElectronic_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optElectronic_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optElectronic_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optExpiration_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optExpiration_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optExpiration.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optGTC_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optGTC_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optGTC.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optMarket_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optMarket_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
        MoveFocus txtQty
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optMarket.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optMIT_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optMIT_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
        If txtAbove.Visible Then
            MoveFocus txtAbove
        Else
            MoveFocus txtBelow
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optMIT_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optPit_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optPit_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optPit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optProfitDollar_Click
'' Description: Handle the user changing over to a Profit Target dollar amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optProfitDollar_Click()
On Error GoTo ErrSection:

    Dim dDollar As Double               ' Current value of the profit target in dollars
    
    If Visible Then
        If m.bProfitDollar = False Then
            dDollar = ProfitDollar
            InitProfitTarget True, dDollar
            m.bProfitDollar = True
            
            SetProfitOtherLabel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optProfitDollar_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optProfitPoints_Click
'' Description: Handle the user changing over to a Profit Target point amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optProfitPoints_Click()
On Error GoTo ErrSection:

    Dim dPoints As Double               ' Current value of the profit target in points
    
    If Visible Then
        If m.bProfitDollar = True Then
            dPoints = ProfitPoints
            InitProfitTarget False, dPoints
            m.bProfitDollar = False
            
            SetProfitOtherLabel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optProfitPoints_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSell_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSell_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls True
        MoveFocus txtQty
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optSell.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optStopDollar_Click
'' Description: Handle the user changing over to a Stop Loss dollar amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optStopDollar_Click()
On Error GoTo ErrSection:

    Dim dDollar As Double               ' Current value of the profit target in dollars
    
    If Visible Then
        If m.bStopDollar = False Then
            dDollar = StopDollar
            InitStopLoss True, dDollar
            m.bStopDollar = True
            
            SetStopOtherLabel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optStopDollar_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optStopLimit_Click
'' Description: Fix other controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optStopLimit_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        FixControls
        MoveFocus txtAbove
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optStopLimit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optStopPoints_Click
'' Description: Handle the user changing over to a Stop Loss point amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optStopPoints_Click()
On Error GoTo ErrSection:

    Dim dPoints As Double               ' Current value of the profit target in points
    
    If Visible Then
        If m.bStopDollar = True Then
            dPoints = StopPoints
            InitStopLoss False, dPoints
            m.bStopDollar = False
            
            SetStopOtherLabel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optStopPoints_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DrawOrder
'' Description: Draw the order with a bar
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DrawOrder()
On Error GoTo ErrSection:

    Dim nMid&, nAbove&, nBelow&

    With picOrder
        ' init things for drawing
        .Cls
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .DrawWidth = 1
        .DrawStyle = vbSolid
        'If InStr(m.strPos, "S") Then
        If optSell Then
            .ForeColor = g.ChartGlobals.nShortColor
        Else
            .ForeColor = g.ChartGlobals.nLongColor
        End If
        nMid = .ScaleHeight / 2 - 1
        nAbove = nMid - (optMarket.Top - optAbove.Top) / Screen.TwipsPerPixelY
        nBelow = nMid + (optMarket.Top - optAbove.Top) / Screen.TwipsPerPixelY
        
        If optStopLimit Then
            If optBuy Then
                DrawOHLC nBelow + 8, nMid + 8, nBelow + 22, nBelow + 12
            Else
                DrawOHLC nAbove - 8, nMid - 8, nAbove - 22, nAbove - 12
            End If
            DrawTriangle nAbove
            DrawTriangle nBelow
            ' draw connecting line
            .DrawWidth = 1
            .DrawStyle = vbDot
            .CurrentX = lineHL.X1 + 6 '11
            .CurrentY = nBelow - 4
            picOrder.Line -(.CurrentX, nAbove + 4)
        ElseIf optAbove Then
            DrawOHLC nBelow, nAbove + 4, nBelow + 8, nMid
            DrawTriangle nAbove
        ElseIf optBelow Then
            DrawOHLC nAbove, nAbove - 8, nBelow - 4, nMid
            DrawTriangle nBelow
        ElseIf optMIT Then
            If optBuy Then
                DrawOHLC nAbove, nAbove - 8, nBelow - 4, nMid
                DrawTriangle nBelow
            Else
                DrawOHLC nBelow, nAbove + 4, nBelow + 8, nMid
                DrawTriangle nAbove
            End If
        ElseIf optBuy Then
            DrawOHLC nBelow, nAbove + 4, nBelow + 8, nMid
            DrawTriangle nMid, -1
        Else
            DrawOHLC nAbove, nAbove - 8, nBelow - 4, nMid
            'DrawOHLC nBelow, nAbove + 8, nBelow + 8, nMid
            DrawTriangle nMid, -1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.DrawOrder", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DrawOHLC
'' Description: Draw the Open/High/Low/Close bar
'' Inputs:      Values of Open, High, Low, and Close
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DrawOHLC(ByVal nOpen&, ByVal nHigh&, ByVal nLow&, ByVal nClose&)
On Error GoTo ErrSection:

    lineHL.Y1 = nHigh
    lineHL.Y2 = nLow
    lineO.Y1 = nOpen
    lineO.Y2 = nOpen
    lineC.Y1 = nClose
    lineC.Y2 = nClose

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.DrawOHLC", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DrawTriangle
'' Description: Draw the triangle for the order price
''              iDir: 1 = right of bar and points left, -1 = left of bar and points right
'' Inputs:      Start Y and Direction
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DrawTriangle(ByVal nStartY&, Optional ByVal iDir% = 1)
On Error GoTo ErrSection:

    Dim i&

    With picOrder
        ' init things to start the triangle
        .DrawWidth = 1
        .DrawStyle = vbSolid
        .CurrentY = nStartY
        .CurrentX = lineHL.X1 + 2 * iDir
        
        ' draw the triangle
        For i = 8 To 2 Step -2
            picOrder.Line -(.CurrentX + i * iDir, .CurrentY - i)
            picOrder.Line -(.CurrentX, .CurrentY + 2 * i)
            picOrder.Line -(.CurrentX - i * iDir, .CurrentY - i)
            If i = 6 Then
                ' if an exit, only do 2 loops to leave hollow triangle
                If InStr(m.strPos, "X") Then Exit For
            End If
            If i = 2 Then
                picOrder.Line -(.CurrentX + 2 * iDir, .CurrentY)
            Else
                picOrder.Line -(.CurrentX + iDir, .CurrentY)
            End If
        Next
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.DrawTriangle", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayPosition
'' Description: Display the position
'' Inputs:      Position, Position label
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayPosition(ByVal nPosition&, lblPosition As ctlUniLabelXP)
On Error GoTo ErrSection:

    With lblPosition
        If nPosition > 0 Then
            .Caption = "Long " & Str(nPosition)
            .ForeColor = g.ChartGlobals.nLongColor
        ElseIf nPosition < 0 Then
            .Caption = "Short " & Str(Abs(nPosition))
            .ForeColor = g.ChartGlobals.nShortColor
        Else
            .Caption = "None"
            .ForeColor = &H404040    ' &H80000012
        End If
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.DisplayPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTrailDollar_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTrailDollar_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableAdvancedControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optTrailDollar.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTrailPoints_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTrailPoints_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableAdvancedControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.optTrailPoints.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    peProfitTarget_Changed
'' Description: As the profit target changes, change the "other" label
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub peProfitTarget_Changed()
On Error GoTo ErrSection:

    SetProfitOtherLabel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.peProfitTarget_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    peStopLoss_Changed
'' Description: As the stop loss changes, change the "other" label
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub peStopLoss_Changed()
On Error GoTo ErrSection:

    SetStopOtherLabel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.peStopLoss_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDollarTrail_LostFocus
'' Description: As the dollar trail changes, change the point trail
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDollarTrail_LostFocus()
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the text box
    
    If (Not m.TrailAmount Is Nothing) Or (m.bChange = False) Then
        dValue = ValOfText(txtDollarTrail.Text)
    
        If m.Bars.Prop(eBARS_TickValue) <> 0 Then
            m.TrailAmount.Price = m.Bars.Prop(eBARS_TickMove) * (dValue / m.Bars.Prop(eBARS_TickValue))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtDollarTrail.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPointTrail_Change
'' Description: As the user changes the point trail, change the dollar trail
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPointTrail_Change()
On Error GoTo ErrSection:

    Dim dDollar As Double               ' Dollar amount for the points
    
    If Not m.TrailAmount Is Nothing Then
        If m.Bars.Prop(eBARS_TickMove) <> 0 Then
            dDollar = m.Bars.Prop(eBARS_TickValue) * (m.TrailAmount.Price / m.Bars.Prop(eBARS_TickMove))
        End If
        m.bChange = False
        txtDollarTrail.Text = Format(dDollar, "$#,##0.00")
        m.bChange = True
    
        If (Not m.Above Is Nothing) And (Not m.Bars Is Nothing) And chkTrail = vbChecked Then
            m.Above.Price = m.Bars(eBARS_Close, m.Bars.Size - 1) + m.TrailAmount.Price
        End If
        If (Not m.Below Is Nothing) And (Not m.Bars Is Nothing) And chkTrail = vbChecked Then
            m.Below.Price = m.Bars(eBARS_Close, m.Bars.Size - 1) - m.TrailAmount.Price
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtPointTrail.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtAbove_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtAbove_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtAbove

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtAbove.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtBelow_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtBelow_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtBelow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtBelow.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtConditionSymbol_Click
'' Description: Allow the user to lookup a condition symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtConditionSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    LookupConditionSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtConditionSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtConditionSymbol_KeyPress
'' Description: If the user types in the text box, bring up the symbol selector
'' Inputs:      Ascii Value of Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtConditionSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupConditionSymbol KeyAscii
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtConditionSymbol.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtQty_Change
'' Description: Fix the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtQty_Change()
On Error GoTo ErrSection:

    If Me.Visible Then
        'SetPriceRange False
        FixControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtQty.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtQty_GotFocus
'' Description: Select the text of the text box when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtQty_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtQty.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: If the user clicks on the text box, bring up the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: Select the text of the text box when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtQty.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: If the user types in the text box, bring up the symbol selector
'' Inputs:      Ascii Value of Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupSymbol KeyAscii
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtSymbol.KeyPress", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_LostFocus
'' Description: Do some setup when the symbol changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_LostFocus()
On Error GoTo ErrSection:
    
    If txtSymbol <> m.Bars.Prop(eBARS_Symbol) Then
        txtSymbol = ConvertToTradeSymbol(txtSymbol, Date)
        DM_GetBars m.Bars, txtSymbol.Text, , LastDailyDownload
        g.RealTime.SpliceBars m.Bars
        InitAbove m.Bars(eBARS_Close, m.Bars.Size - 1)
        InitBelow m.Bars(eBARS_Close, m.Bars.Size - 1)
        'If cboAccounts.ListIndex >= 0 Then
        '    m.nCurPosition = g.Broker.CurrentPosition(cboAccounts.ItemData(cboAccounts.ListIndex), txtSymbol.Text, m.Order.AutoTradeItemID)
        'Else
        '    m.nCurPosition = 0&
        'End If
        m.nCurPosition = CurrentPositionAfterTrigger
        If Me.Visible Then
            FixControls True
        End If
        
        EnableDurationControls
        EnableOrderTimeControls
        
        ShowSessionControls
        ShowExchangeControls
        ShowLotControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.txtSymbol.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixControls
'' Description: Setup the controls based on all the values
'' Inputs:      Reset Quantity?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixControls(Optional ByVal bResetQty As Boolean = False)
On Error GoTo ErrSection:

    Dim bAbove As Boolean, bBelow As Boolean, bLowStop As Boolean
    Dim nQty&, eAcctType As eTT_AccountType
    Dim strAbove$, strBelow$, strText$
    Static bInProgress As Boolean
    Dim strAutoExitName As String       ' Auto exit name to select
    Dim Account As cPtAccount           ' Account object
    
    If bInProgress Then Exit Sub '(so not recursive)
    bInProgress = True
    
    If optBuy Then
        cmdOK.BackColor = kBuyColor
    ElseIf optSell Then
        cmdOK.BackColor = kSellColor
    Else
        cmdOK.BackColor = cmdPark.BackColor
    End If
    
    ' display position info
    If chkTriggerOrder.Value = vbChecked Then
        Label6.Caption = "After Trigger:"
    Else
        Label6.Caption = "Current Position:"
    End If
    DisplayPosition m.nCurPosition, lblCurPos
    If optBuy + optSell = 0 Then
        If m.nCurPosition > 0 Then
            optSell = True
            bResetQty = True
        ElseIf m.nCurPosition < 0 Then
            optBuy = True
            bResetQty = True
        End If
    End If
    
    If chkExitAll.Value = vbChecked Then
        nQty = m.nCurPosition
    Else
        If bResetQty Then SetDefaultQuantity
        nQty = ValOfText(txtQty) 'm.Qty
    End If
            
    If nQty = 0 Or optBuy + optSell = 0 Then
        lblNewPos.Caption = ""
        strText = ""
    ElseIf optSell Then
        DisplayPosition m.nCurPosition - nQty, lblNewPos
        If nQty > Abs(m.nCurPosition) Then
            m.strPos = "ES"
        Else
            m.strPos = "XL"
        End If
        If m.nCurPosition < 0 Then
            m.strPos = "ES"
            strText = "ADD to Short"
        ElseIf m.nCurPosition = 0 Then
            m.strPos = "ES"
            strText = "ENTER Short"
        ElseIf nQty = Abs(m.nCurPosition) Then
            m.strPos = "XL"
            strText = "EXIT from Long"
        ElseIf nQty < Abs(m.nCurPosition) Then
            m.strPos = "XL"
            strText = "REDUCE Long"
        Else
            m.strPos = "ES"
            strText = "REVERSE to Short"
        End If
    Else
        DisplayPosition m.nCurPosition + nQty, lblNewPos
        If m.nCurPosition > 0 Then
            m.strPos = "EL"
            strText = "ADD to Long"
        ElseIf m.nCurPosition = 0 Then
            m.strPos = "EL"
            strText = "ENTER Long"
        ElseIf nQty = Abs(m.nCurPosition) Then
            m.strPos = "XS"
            strText = "EXIT from Short"
        ElseIf nQty < Abs(m.nCurPosition) Then
            m.strPos = "XS"
            strText = "REDUCE Short"
        Else
            m.strPos = "EL"
            strText = "REVERSE to Long"
        End If
    End If
    
    EnableQuantityControls
    
    If chkExitAll.Value = vbChecked Then
        txtQty.Visible = False
        txtExitPos.Top = txtQty.Top
        txtExitPos.Visible = True
    Else
        txtQty.Visible = True
        txtExitPos.Visible = False
    End If
    
    lblAction.Caption = strText
    'If InStr(m.strPos, "S") Then
    If optSell Then
        lblAction.ForeColor = g.ChartGlobals.nShortColor
    Else
        lblAction.ForeColor = g.ChartGlobals.nLongColor
    End If
    
    If cboAccounts.ListIndex >= 0 Then
        eAcctType = g.Broker.AccountTypeForID(cboAccounts.ItemData(cboAccounts.ListIndex))
    Else
        eAcctType = eTT_AccountType_SimStream
    End If
    
    ' Enable/show controls and appropriate captions
    If optBuy Or optSell Then
        cmdOK.Enabled = True
        cmdPark.Enabled = True
        optStopLimit.Enabled = (chkTrail.Value = vbUnchecked)
        optAbove.Enabled = (chkTrail.Value = vbUnchecked)
        optBelow.Enabled = (chkTrail.Value = vbUnchecked)
        optMIT.Enabled = (chkTrail.Value = vbUnchecked)
    Else
        cmdOK.Enabled = False
        cmdPark.Enabled = False
        optStopLimit.Enabled = False
        optAbove.Enabled = False
        optBelow.Enabled = False
        optMIT.Enabled = False
        optMarket.Value = True
    End If
    
    If chkTrail.Value = vbChecked Then
        If optBuy.Value = True Then
            optAbove.Value = True
        Else
            optBelow.Value = True
        End If
        m.Above.Price = m.Bars(eBARS_Close, m.Bars.Size - 1) + m.TrailAmount.Price
        m.Below.Price = m.Bars(eBARS_Close, m.Bars.Size - 1) - m.TrailAmount.Price
    End If
    
    If optSell Then bLowStop = True

    If optStopLimit Then
        strAbove = ": at or below"
        strBelow = ": at or above"
        bAbove = True
        bBelow = True
        bLowStop = Not bLowStop
    Else
        strAbove = ": at or above"
        strBelow = ": at or below"
        
        If optMIT Then
            If optBuy Then
                bAbove = False
                bBelow = True
            Else
                bAbove = True
                bBelow = False
            End If
        Else
            bAbove = optAbove
            bBelow = optBelow
        End If
    End If
    
    If bLowStop Then
        strAbove = "LIMIT" & strAbove
        strBelow = "STOP" & strBelow
        
        optMIT.Top = 300
        optAbove.Top = 570
        optMarket.Top = 840
        optBelow.Top = 1110
        
        txtBelow.Top = 1065
        sbBelow.Top = 1027
        If optMIT.Value = True Then
            txtAbove.Top = 270
            sbAbove.Top = 225
        Else
            txtAbove.Top = 525
            sbAbove.Top = 487
        End If
    Else
        strAbove = "STOP" & strAbove
        strBelow = "LIMIT" & strBelow
        
        optAbove.Top = 300
        optMarket.Top = 570
        optBelow.Top = 840
        optMIT.Top = 1110
        
        txtAbove.Top = 255
        sbAbove.Top = 217
        If optMIT.Value = True Then
            txtBelow.Top = 1065
            sbBelow.Top = 1027
        Else
            txtBelow.Top = 795
            sbBelow.Top = 757
        End If
    End If
    
    optStopLimit.Top = 1380
    
    optAbove.Caption = strAbove
    optBelow.Caption = strBelow
    
    If optBuy Then
        optMIT.Caption = "MIT: at or below"
    Else
        optMIT.Caption = "MIT: at or above"
    End If
    
    txtAbove.Visible = bAbove
    sbAbove.Visible = bAbove
    txtBelow.Visible = bBelow
    sbBelow.Visible = bBelow

    DrawOrder
    
    If optBuy Then
        chkTrail.Caption = "Submit a Stop order that will stay a fixed amount above the lowest low since the order was submitted."
    ElseIf optSell Then
        chkTrail.Caption = "Submit a Stop order that will stay a fixed amount below the highest high since the order was submitted."
    End If
    
    ' Set the auto exit
    m.bSettingAutoExit = True
    strAutoExitName = g.OrderStrategies.ExitForAccountAndSymbol(AccountID, SymbolOrSymbolID)
    If Len(strAutoExitName) > 0 Then
        chkAutoExit.Caption = "A&uto Exit: " & strAutoExitName
        chkAutoExit.Value = vbChecked
    Else
        chkAutoExit.Caption = "A&uto Exit: " & "None"
        chkAutoExit.Value = vbUnchecked
    End If
    m.bSettingAutoExit = False
    
    ' DAJ 09/09/2009: Only allow the user to hold an OCO at the broker for PFG...
    If cboAccounts.ListIndex >= 0 Then
        Set Account = g.Broker.Account(cboAccounts.ItemData(cboAccounts.ListIndex))
    Else
        Set Account = Nothing
    End If
    If Account Is Nothing Then
        chkBrokerOCO.Value = vbUnchecked
        chkBrokerOCO.Visible = False
    ElseIf (g.Broker.IsIbBroker(Account.AccountType) = True) And (m.bNewOrder = False) Then
        ' For Interactive Brokers, you can only setup an OCO for a brand new
        ' order, so if it isn't a new order, disable the check box and leave it
        ' as unchecked...
        CheckBoxValue(chkBrokerOCO) = False
        chkBrokerOCO.Visible = g.Broker.BrokerAllowsOCO(Account.AccountType)
        chkBrokerOCO.Enabled = False
    Else
        CheckBoxValue(chkBrokerOCO) = g.Broker.HoldOcoAtBroker(cboAccounts.Text)
        chkBrokerOCO.Visible = g.Broker.BrokerAllowsOCO(Account.AccountType)
        chkBrokerOCO.Enabled = chkBrokerOCO.Visible
    End If
    
    'If eAcctType = eTT_AccountType_Photon Then
    '    optExpiration.Value = True
    '    gdExpiration.Value = Date
    '    optGTC.Enabled = False
    '    optExpiration.Enabled = False
    '    gdExpiration.Enabled = False
    'Else
    '    optGTC.Enabled = True
    '    optExpiration.Enabled = True
    '    gdExpiration.Value = ExpirationDateForSymbol
    '    Enable gdExpiration, optExpiration
    'End If
    EnableDurationControls
    EnableOrderTimeControls
    
    ShowSessionControls
    ShowExchangeControls
    ShowLotControls
    
    ' 06/27/2013 DAJ: Since contingency orders get submitted with the original order, they use the
    ' order price to calculate their value.  Since market orders don't have an order price, the
    ' contingency orders are being submitted with null prices.  So, we decided not to allow them
    ' if the order is a market order...
    If optMarket.Value = True Then
        CheckBoxValue(chkProfitTarget) = False
        Disable chkProfitTarget
        CheckBoxValue(chkStopLoss) = False
        Disable chkStopLoss
    Else
        Enable chkProfitTarget
        Enable chkStopLoss
    End If

    bInProgress = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.FixControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Lookup a symbol for the user to trade
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LookupSymbol(Optional ByVal KeyAscii As Long = 0&)
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    Dim strSymbol As String
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, "Symbol to Buy/Sell", , , True)
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol to Buy/Sell", False, False, True)
    End If
    If astrSymbol.Size > 0 Then
        strSymbol = ConvertToTradeSymbol(astrSymbol(0), Date)
        'If (InStr(Trim(astrSymbol(0)), " ") > 0) And (InStr(Trim(astrSymbol(0)), "-") > 0) And (Not FileExist(AddSlash(App.Path) & "TradeFO.FLG")) Then
        '    InfBox "Future Options are not currently allowed|to be traded", "!", , "Error"
        'ElseIf strSymbol <> UCase(Trim(txtSymbol.Text)) Then
        If strSymbol <> UCase(Trim(txtSymbol.Text)) Then
            txtSymbol.Text = strSymbol
            DM_GetBars m.Bars, txtSymbol.Text, , LastDailyDownload
            g.RealTime.SpliceBars m.Bars
            InitAbove m.Bars(eBARS_Close, m.Bars.Size - 1)
            InitBelow m.Bars(eBARS_Close, m.Bars.Size - 1)
            
            'If cboAccounts.ListIndex >= 0 Then
            '    m.nCurPosition = g.Broker.CurrentPosition(cboAccounts.ItemData(cboAccounts.ListIndex), txtSymbol.Text, m.Order.AutoTradeItemID)
            'Else
            '    m.nCurPosition = 0&
            'End If
            m.nCurPosition = CurrentPositionAfterTrigger
            
            If (optDay.Value = True) Or (optGTC.Value = True) Then
                gdExpiration.Value = ExpirationDateForSymbol
            End If
            
            GetContractInformation
            
            If Me.Visible Then
                FixControls True
            End If
            
            InitProfitTarget optProfitDollar.Value, peProfitTarget.Price
            SetProfitOtherLabel
            
            InitStopLoss optStopDollar.Value, peStopLoss.Price
            SetStopOtherLabel
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.LookupSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupConditionSymbol
'' Description: Lookup a symbol for the user to trade
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LookupConditionSymbol(Optional ByVal KeyAscii As Long = 0&)
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(txtConditionSymbol.Text, False, True, "Condition Symbol")
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Condition Symbol", False, False)
    End If
    If astrSymbol.Size > 0 Then
        txtConditionSymbol.Text = astrSymbol(0)
        DM_GetBars m.CondBars, txtConditionSymbol.Text, , LastDailyDownload
        g.RealTime.SpliceBars m.CondBars
        InitConditionalPrice m.CondBars(eBARS_Close, m.CondBars.Size - 1)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.LookupConditionSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsCalc_Click
'' Description: Allow the user to do some money management calculations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsCalc_Click()
On Error GoTo ErrSection:

    MMcalc

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.vsCalc_Click", eGDRaiseError_Show
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MMcalc
'' Description: Allow the user to do some money management calculations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MMcalc()
On Error GoTo ErrSection:

    Dim nContracts As Long

    If cboAccounts.ListIndex >= 0 Then
        nContracts = frmMmCalc.ShowMe(txtSymbol, cboAccounts.ItemData(cboAccounts.ListIndex))
        If nContracts > 0 Then
            If optSell Then
                nContracts = -nContracts
            End If
            ' if reversing position, add current position in order to exit and enter new position
            If m.nCurPosition * nContracts < 0 Then
                nContracts = Abs(nContracts) + Abs(m.nCurPosition)
            End If
            txtQty = Abs(nContracts)
            FixControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.MMcalc", eGDRaiseError_Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountCombo
'' Description: Set the account combo to the given account ID
'' Inputs:      Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAccountCombo(ByVal lAccountID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With cboAccounts
        For lIndex = 0 To .ListCount - 1
            If .ItemData(lIndex) = lAccountID Then
                .ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetAccountCombo", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpirationDateForSymbol
'' Description: Given the symbol and the time of day, figure out the "proper"
''              default expiration date
'' Inputs:      None
'' Returns:     Default Expiration Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExpirationDateForSymbol() As Double
On Error GoTo ErrSection:

    Dim strExchangeTimeInf As String    ' String of exchange time zone information
    Dim dNowInExchangeTime As Double    ' The current time in exchange time zone
    Dim dSessionEnd As Double           ' Session End time in exchange time zone
    Dim lReturn As Long                 ' Return value
    
    If Len(Trim(txtSymbol.Text)) > 0 Then
        strExchangeTimeInf = m.Bars.Prop(eBARS_ExchangeTimeZoneInf)
        dNowInExchangeTime = ConvertTimeZone(Now, , strExchangeTimeInf)
        dSessionEnd = Int(dNowInExchangeTime) + (m.Bars.Prop(eBARS_DefaultEndTime) / 1440#)
        
        If dNowInExchangeTime > dSessionEnd Then
            lReturn = Int(dNowInExchangeTime) + 1
            Do While Not IsWeekday(lReturn)
                lReturn = lReturn + 1
            Loop
        Else
            lReturn = Int(dNowInExchangeTime)
        End If
    Else
        lReturn = Date
    End If

    ExpirationDateForSymbol = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.ExpirationDateForSymbol", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetDefaultQuantity
'' Description: Set the default quantity based on symbol, position, and account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetDefaultQuantity()
On Error GoTo ErrSection:

    Dim lQuantity As Long               ' Quantity to set the control to

    lQuantity = Abs(m.nCurPosition)
    If lQuantity = 0 Then
        lQuantity = GetDefaultEntryForSymbol(m.Bars.Prop(eBARS_SymbolID), m.Bars.Prop(eBARS_Symbol))
        If lQuantity = 0 Then
            lQuantity = g.Broker.DefaultOrderQuantity(AccountID, SymbolOrSymbolID)
        End If
    End If
    
    g.Broker.InitQuantityEditor m.Qty, sbQty, txtQty, AccountID, SymbolOrSymbolID, lQuantity

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetDefaultQuantity", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateSecTypeForAccount
'' Description: Is this a valid security type for the account?
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateSecTypeForAccount()
On Error GoTo ErrSection:

    Dim Account As New cPtAccount       ' Temporary account object
    
    If cboAccounts.ListIndex >= 0 Then
        If Account.Load(cboAccounts.ItemData(cboAccounts.ListIndex)) Then
            Select Case m.Bars.SecurityType
                Case "F"
                    If GetBit(Account.SecTypeMask, 1) = False Then
                        MoveFocus cboAccounts
                        Err.Raise vbObjectError + 1000, , "You cannot trade Futures in the " & cboAccounts.Text & " account"
                    End If
                
                Case "S"
                    If GetBit(Account.SecTypeMask, 2) = False Then
                        MoveFocus cboAccounts
                        Err.Raise vbObjectError + 1000, , "You cannot trade Stocks in the " & cboAccounts.Text & " account"
                    End If
                
                Case "I"
                    If IsForex(m.Bars.Prop(eBARS_Symbol)) Then
                        If GetBit(Account.SecTypeMask, 3) = False Then
                            MoveFocus cboAccounts
                            Err.Raise vbObjectError + 1000, , "You cannot trade Forex Symbols in the " & cboAccounts.Text & " account"
                        End If
                    Else
                        MoveFocus cboAccounts
                        Err.Raise vbObjectError + 1000, , "You cannot trade Indexes in the " & cboAccounts.Text & " account"
                    End If
                
                Case "FO"
                    If GetBit(Account.SecTypeMask, 4) = False Then
                        MoveFocus cboAccounts
                        Err.Raise vbObjectError + 1000, , "You cannot trade Future Options in the " & cboAccounts.Text & " account"
                    End If
                
                Case "SO"
                    If GetBit(Account.SecTypeMask, 5) = False Then
                        MoveFocus cboAccounts
                        Err.Raise vbObjectError + 1000, , "You cannot trade Stock Options in the " & cboAccounts.Text & " account"
                    End If
            
            End Select
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.ValidateSecTypeForAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCancelOrdersCombo
'' Description: Load up the cancel orders combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCancelOrdersCombo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID
    Dim Order As New cPtOrder           ' Temporary order object
    
    cboCancelOrders.Clear
    
    If (cboAccounts.ListCount > 0) And (cboAccounts.ListIndex >= 0) Then
        lAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)

        If (m.Order.CancelOrderID = 0) And (m.Order.BrokerCancelOrderID = 0) Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [AccountID]=" & Str(lAccountID) & " AND ([IsSnapshot]<>0 OR [Status]=" & Str(eTT_OrderStatus_Parked) & " OR [Status]=" & Str(eTT_OrderStatus_TriggerPending) & ") AND ([CancelOrderID]=0 AND [BrokerCancelOrderID]=0) ORDER BY [OrderDate] DESC;", dbOpenDynaset)
        ElseIf m.Order.CancelOrderID <> 0 Then
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [AccountID]=" & Str(lAccountID) & " AND ([IsSnapshot]<>0 OR [Status]=" & Str(eTT_OrderStatus_Parked) & " OR [Status]=" & Str(eTT_OrderStatus_TriggerPending) & ") AND (([CancelOrderID]=0 AND [BrokerCancelOrderID]=0) OR [OrderID]=" & Str(m.Order.CancelOrderID) & ") ORDER BY [OrderDate] DESC;", dbOpenDynaset)
        Else
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [AccountID]=" & Str(lAccountID) & " AND ([IsSnapshot]<>0 OR [Status]=" & Str(eTT_OrderStatus_Parked) & " OR [Status]=" & Str(eTT_OrderStatus_TriggerPending) & ") AND (([CancelOrderID]=0 AND [BrokerCancelOrderID]=0) OR [OrderID]=" & Str(Abs(m.Order.BrokerCancelOrderID)) & ") ORDER BY [OrderDate] DESC;", dbOpenDynaset)
        End If
        
        Do While Not rs.EOF
            If IsOpenOrder(rs!Status) And (rs!OrderID <> m.Order.OrderID) Then
                If Order.Load(rs!OrderID) Then
                    ' Cannot allow an IB order that is already in a broker-held OCO situation...
                    If (g.Broker.IsIbBroker(Order.Broker) = False) Or (Order.BrokerCancelOrderID = 0) Then
                        cboCancelOrders.AddItem Order.OrderText & " - " & Order.BrokerID
                        cboCancelOrders.ItemData(cboCancelOrders.NewIndex) = Order.OrderID
                    End If
                End If
            End If
            
            rs.MoveNext
        Loop
    End If
        
    If cboCancelOrders.ListCount = 0 Then
        cboCancelOrders.AddItem "<There are no Open Orders in this Account>"
        chkCancelOrder.Value = vbUnchecked
        chkCancelOrder.Enabled = False
        cboCancelOrders.ListIndex = 0
        cboCancelOrders.Enabled = False
        lblCancelNote.Enabled = False
    Else
        chkCancelOrder.Enabled = True
        cboCancelOrders.ListIndex = 0
        cboCancelOrders.Enabled = True
        lblCancelNote.Enabled = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.LoadCancelOrdersCombo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCancelOrderCombo
'' Description: Set the cancel order combo to the given order ID
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCancelOrderCombo(ByVal lOrderID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With cboCancelOrders
        For lIndex = 0 To .ListCount - 1
            If .ItemData(lIndex) = Abs(lOrderID) Then
                .ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetCancelOrderCombo", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadTriggerOrdersCombo
'' Description: Load up the trigger orders combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTriggerOrdersCombo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lAccountID As Long              ' Account ID
    Dim Order As New cPtOrder           ' Temporary order object
    Dim strOrders As String             ' String of orders that have been added
    
    cboTriggerOrders.Clear
    
    If cboAccounts.ListIndex >= 0 Then
        lAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] " & _
                    "WHERE [AccountID]=" & Str(lAccountID) & " AND ([IsSnapshot]<>0 OR [Status]=" & Str(eTT_OrderStatus_Parked) & " OR [Status]=" & Str(eTT_OrderStatus_TriggerPending) & ") ORDER BY [OrderDate] DESC;", dbOpenDynaset)
        Do While Not rs.EOF
            If IsOpenOrder(rs!Status) And (rs!OrderID <> m.Order.OrderID) Then
                If Order.Load(rs!OrderID) Then
                    cboTriggerOrders.AddItem Order.OrderText & " - " & Order.BrokerID
                    cboTriggerOrders.ItemData(cboTriggerOrders.NewIndex) = Order.OrderID
                    strOrders = "," & Str(Order.OrderID) & ","
                End If
            End If
            
            rs.MoveNext
        Loop
    End If
    
    ' Make sure that the triggering order is in the combo box as well...
    If (m.Order.TriggerOrderID <> 0) And (InStr(strOrders, "," & Str(m.Order.TriggerOrderID) & ",") = 0) Then
        Set Order = New cPtOrder
        If Order.Load(m.Order.TriggerOrderID) Then
            cboTriggerOrders.AddItem Order.OrderText & " - " & Order.BrokerID
            cboTriggerOrders.ItemData(cboTriggerOrders.NewIndex) = Order.OrderID
            strOrders = "," & Str(Order.OrderID) & ","
        End If
    End If
    
    If cboTriggerOrders.ListCount = 0 Then
        cboTriggerOrders.AddItem "<There are no Open Orders in this Account>"
        chkTriggerOrder.Value = vbUnchecked
        chkTriggerOrder.Enabled = False
        cboTriggerOrders.ListIndex = 0
        cboTriggerOrders.Enabled = False
        optTriggerPartial.Enabled = False
        optTriggerFull.Enabled = False
        chkTriggerConfirm.Enabled = False
    Else
        chkTriggerOrder.Enabled = True
        cboTriggerOrders.ListIndex = 0
        cboTriggerOrders.Enabled = True
        optTriggerPartial.Enabled = True
        optTriggerFull.Enabled = True
        chkTriggerConfirm.Enabled = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.LoadTriggerOrdersCombo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTriggerOrderCombo
'' Description: Set the trigger order combo to the given order ID
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTriggerOrderCombo(ByVal lOrderID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With cboTriggerOrders
        For lIndex = 0 To .ListCount - 1
            If .ItemData(lIndex) = Abs(lOrderID) Then
                .ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetTriggerOrderCombo", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTriggerOptions
'' Description: Set the trigger options
'' Inputs:      Trigger Options string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTriggerOptions(ByVal strOptions As String)
On Error GoTo ErrSection:

    Dim lConfirm As Long                ' Does the user want to confirm order?
    Dim lPartial As Long                ' Trigger on partial fill?
    
    lPartial = CLng(Val(Parse(strOptions, ",", 1)))
    lConfirm = CLng(Val(Parse(strOptions, ",", 2)))
    
    CheckBoxValue(chkTriggerConfirm) = g.Broker.ConfirmTriggered
    
    If lPartial = 0 Then
        optTriggerPartial.Value = False
        optTriggerFull.Value = True
    Else
        optTriggerPartial.Value = True
        optTriggerFull.Value = False
    End If
    
    If Val(Parse(strOptions, ",", 3)) = 0 Then
        chkRelativePrice.Value = vbUnchecked
        txtRelativePrice.Text = ""
    Else
        chkRelativePrice.Value = vbChecked
        txtRelativePrice.Text = Parse(strOptions, ",", 3)
    End If
    
    CheckBoxValue(chkMoveWithTrigger) = CBool(Val(Parse(strOptions, ",", 4)))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetTriggerOptions", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConditionOptions
'' Description: Set the condition options
'' Inputs:      Condition Options string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetConditionOptions(ByVal strOptions As String)
On Error GoTo ErrSection:

    Dim astrOptions As New cGdArray     ' Array of options
    Dim astrPrice As New cGdArray       ' Array of price options
    
    astrOptions.SplitFields strOptions, vbTab
    astrPrice.SplitFields astrOptions(5), ";"
    
    If astrOptions(0) = "1" Then chkBeginTime.Value = vbChecked Else chkBeginTime.Value = vbUnchecked
    If Len(astrOptions(1)) > 0 Then gdBeginTime.Value = Val(astrOptions(1)) Else gdBeginTime.Value = Date
    If astrOptions(2) = "1" Then chkEndTime.Value = vbChecked Else chkEndTime.Value = vbUnchecked
    If Len(astrOptions(3)) > 0 Then gdEndTime.Value = Val(astrOptions(3)) Else gdEndTime.Value = Date
    If astrOptions(4) = "1" Then chkPrice.Value = vbChecked Else chkPrice.Value = vbUnchecked
    If Len(astrPrice(0)) > 0 Then cboConditionField.Text = astrPrice(0) Else cboConditionField.Text = "Last Price"
    If Len(astrPrice(1)) > 0 Then txtConditionSymbol.Text = astrPrice(1) Else txtConditionSymbol.Text = txtSymbol.Text
    If Len(astrPrice(2)) > 0 Then cboConditionOperator.Text = astrPrice(2) Else cboConditionOperator.Text = "="
    If Len(astrPrice(3)) > 0 Then txtConditionPrice.Text = Format(Val(astrPrice(3))) Else txtConditionPrice.Text = "0"
    If astrOptions(6) = "1" Then chkCustomCondition.Value = vbChecked Else chkCustomCondition.Value = vbUnchecked
    If Len(astrOptions(7)) > 0 Then SetConditionRTF (astrOptions(7)) Else rtfCondition.Text = ""
    If Len(astrOptions(8)) > 0 Then lblDefaultPeriod.Caption = "Default Period: " & astrOptions(8) Else lblDefaultPeriod.Caption = "Default Period: Daily"
    If Len(astrOptions(9)) > 0 Then m.strNumBarsInfo = astrOptions(9) Else m.strNumBarsInfo = ""
    
    Set m.CondBars = New cGdBars
    DM_GetBars m.CondBars, txtConditionSymbol.Text
    
    If chkBeginTime = vbChecked Or chkEndTime = vbChecked Or chkPrice = vbChecked Or chkCustomCondition = vbChecked Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOnIcon))
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Condition)) = Picture16(ToolbarIcon(kOffIcon))
    End If

ErrExit:
    Set astrOptions = Nothing
    Set astrPrice = Nothing
    Exit Sub
    
ErrSection:
    Set astrOptions = Nothing
    Set astrPrice = Nothing
    RaiseError "frmTTEditOrder.SetConditionOptions", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetConditionOptions
'' Description: Get the condition options from the user interface
'' Inputs:      None
'' Returns:     Condition Options
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetConditionOptions() As String
On Error GoTo ErrSection:
    
    Dim astrOptions As New cGdArray     ' Array of options
    Dim astrPrice As New cGdArray       ' Array of price options
    
    astrOptions.Create eGDARRAY_Strings
    astrPrice.Create eGDARRAY_Strings
    
    If chkBeginTime.Value = vbChecked Then
        astrOptions(0) = "1"
        astrOptions(1) = Str(gdBeginTime.Value)
    Else
        astrOptions(0) = "0"
        astrOptions(1) = ""
    End If
    If chkEndTime.Value = vbChecked Then
        astrOptions(2) = "1"
        astrOptions(3) = Str(gdEndTime.Value)
    Else
        astrOptions(2) = "0"
        astrOptions(3) = ""
    End If
    If chkPrice.Value = vbChecked Then
        astrOptions(4) = "1"
        astrPrice(0) = cboConditionField.Text
        astrPrice(1) = txtConditionSymbol.Text
        astrPrice(2) = cboConditionOperator.Text
        astrPrice(3) = Str(m.CondPrice.Price)
        astrOptions(5) = astrPrice.JoinFields(";")
    Else
        astrOptions(4) = "0"
        astrOptions(5) = ""
    End If
    If chkCustomCondition.Value = vbChecked Then
        astrOptions(6) = "1"
        astrOptions(7) = rtfCondition.Text
        astrOptions(8) = Parse(lblDefaultPeriod.Caption, ":", 2)
        astrOptions(9) = m.strNumBarsInfo
    Else
        astrOptions(6) = "0"
        astrOptions(7) = ""
        astrOptions(8) = ""
        astrOptions(9) = ""
    End If
    
    GetConditionOptions = astrOptions.JoinFields(vbTab)
    
ErrExit:
    Set astrOptions = Nothing
    Set astrPrice = Nothing
    Exit Function
    
ErrSection:
    Set astrOptions = Nothing
    Set astrPrice = Nothing
    RaiseError "frmTTEditOrder.GetConditionOptions", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetTrailControls
'' Description: Set the trailing stop controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTrailControls()
On Error GoTo ErrSection:

    Dim dTrailAmount As Double          ' Trailing stop dollar amount

    With m.Order
        If .TrailAmount = 0 Then
            chkTrail.Value = vbUnchecked
            tabSpecialOrders.TabPicture(Tabs(eGDTab_Trail)) = Picture16(ToolbarIcon(kOffIcon))
            
            dTrailAmount = GetIniFileProperty("DollarTrail", 1500#, "TTEditOrder", g.strIniFile)
            txtDollarTrail.Text = Format(dTrailAmount, "$#,##0.00")
            
            If m.Bars.Prop(eBARS_TickValue) <> 0 Then
                txtPointTrail.Text = (dTrailAmount / m.Bars.Prop(eBARS_TickValue)) * m.Bars.Prop(eBARS_TickMove)
            End If
        Else
            chkTrail.Value = vbChecked
            tabSpecialOrders.TabPicture(Tabs(eGDTab_Trail)) = Picture16(ToolbarIcon(kOnIcon))
            If CLng(Val(Parse(.TrailOptions, ",", 1))) = 0 Then
                optTrailDollar.Value = True
                optTrailPoints.Value = False
                
                txtDollarTrail.Text = Format(.TrailAmount, "$#,##0.00")
                If m.Bars.Prop(eBARS_TickValue) <> 0 Then
                    txtPointTrail.Text = (.TrailAmount / m.Bars.Prop(eBARS_TickValue)) * m.Bars.Prop(eBARS_TickMove)
                End If
            Else
                optTrailDollar.Value = False
                optTrailPoints.Value = True
                
                txtPointTrail.Text = Str(.TrailAmount)
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetTrailControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetContingencyControls
'' Description: Set the contingency controls
'' Inputs:      Contingency Options
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetContingencyControls(ByVal ContingencyOptions As cContingencyOrders)
On Error GoTo ErrSection:
    
    Dim dDefault As Double              ' Default value
    
    If (ContingencyOptions.UseProfitTarget = False) And (ContingencyOptions.UseStopLoss = False) Then
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Contingency)) = Picture16(ToolbarIcon(kOffIcon))
        
        CheckBoxValue(chkProfitTarget) = False
        m.bProfitDollar = True
        optProfitDollar.Value = True
        optProfitPoints.Value = False
        
        dDefault = GetIniFileProperty("LastProfitDollar", 1500#, "TTEditOrder", g.strIniFile)
        InitProfitTarget True, dDefault
    
        CheckBoxValue(chkStopLoss) = False
        m.bStopDollar = True
        optStopDollar.Value = True
        optStopPoints.Value = False
        
        dDefault = GetIniFileProperty("LastStopLoss", 1500#, "TTEditOrder", g.strIniFile)
        InitStopLoss True, dDefault
    Else
        tabSpecialOrders.TabPicture(Tabs(eGDTab_Contingency)) = Picture16(ToolbarIcon(kOnIcon))
        
        With ContingencyOptions
            CheckBoxValue(chkProfitTarget) = .UseProfitTarget
            optProfitDollar.Value = .ProfitInDollars
            optProfitPoints.Value = Not .ProfitInDollars
            m.bProfitDollar = optProfitDollar.Value
            
            If optProfitDollar.Value = True Then
                InitProfitTarget True, .ProfitDollarAmount
            Else
                InitProfitTarget False, .ProfitPointsAmount
            End If
            
            optProfitDay.Value = (.ProfitTif = eTT_TimeInForce_Day)
            optProfitGTC.Value = (.ProfitTif = eTT_TimeInForce_GTC)
        
            CheckBoxValue(chkStopLoss) = .UseStopLoss
            optStopDollar.Value = .StopInDollars
            optStopPoints.Value = Not .StopInDollars
            m.bStopDollar = optStopDollar.Value
            
            If optStopDollar.Value = True Then
                InitStopLoss True, .StopDollarAmount
            Else
                InitStopLoss False, .StopPointsAmount
            End If
            
            optStopDay.Value = (.StopTif = eTT_TimeInForce_Day)
            optStopGTC.Value = (.StopTif = eTT_TimeInForce_GTC)
        End With
    End If
    
    SetProfitOtherLabel
    SetStopOtherLabel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetContingencyControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContingencyControls
'' Description: Get the contingency information from the contingency controls
'' Inputs:      None
'' Returns:     Contingency Options
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetContingencyControls() As cContingencyOrders
On Error GoTo ErrSection:

    Dim ContingencyOptions As cContingencyOrders    ' Contingency options
    
    If m.Order Is Nothing Then
        Set ContingencyOptions = New cContingencyOrders
    Else
        Set ContingencyOptions = m.Order.Contingency.MakeCopy
    End If
    If (m.bNewOrder = False) Or (CheckBoxValue(chkProfitTarget) Or CheckBoxValue(chkStopLoss)) Then
        With ContingencyOptions
            .UseProfitTarget = CheckBoxValue(chkProfitTarget)
            .ProfitInDollars = optProfitDollar.Value
            .ProfitDollarAmount = ProfitDollar
            .ProfitPointsAmount = ProfitPoints
            If optProfitDay.Value = True Then
                .ProfitTif = eTT_TimeInForce_Day
            ElseIf optProfitGTC.Value = True Then
                .ProfitTif = eTT_TimeInForce_GTC
            End If
            SetIniFileProperty "LastProfitDollar", .ProfitDollarAmount, "TTEditOrder", g.strIniFile
            
            
            .UseStopLoss = CheckBoxValue(chkStopLoss)
            .StopInDollars = optStopDollar.Value
            .StopDollarAmount = StopDollar
            .StopPointsAmount = StopPoints
            If optStopDay.Value = True Then
                .StopTif = eTT_TimeInForce_Day
            ElseIf optStopGTC.Value = True Then
                .StopTif = eTT_TimeInForce_GTC
            End If
            SetIniFileProperty "LastStopLoss", .StopDollarAmount, "TTEditOrder", g.strIniFile
        End With
    End If
    
    Set GetContingencyControls = ContingencyOptions

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.GetContingencyControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableAdvancedControls
'' Description: Enable/Disable the advanced controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableAdvancedControls()
On Error GoTo ErrSection:

    Enable gdBeginTime, (chkBeginTime.Value = vbChecked)
    Enable lblBeginLocalTime, (chkBeginTime.Value = vbChecked)
    Enable gdEndTime, (chkEndTime.Value = vbChecked)
    Enable lblEndLocalTime, (chkEndTime.Value = vbChecked)
    
    Enable cboConditionField, (chkPrice = vbChecked)
    Enable lblOf, (chkPrice = vbChecked)
    Enable txtConditionSymbol, (chkPrice = vbChecked)
    Enable cmdCondSymLookup, (chkPrice = vbChecked)
    Enable cboConditionOperator, (chkPrice = vbChecked)
    Enable txtConditionPrice, (chkPrice = vbChecked)
    Enable gdCondPrice, (chkPrice = vbChecked)
    
    Enable cmdEditCustom, (chkCustomCondition = vbChecked)
    Enable lblDefaultPeriod, (chkCustomCondition = vbChecked)
    Enable rtfCondition, (chkCustomCondition = vbChecked)
    
    Enable optTrailDollar, (chkTrail.Value = vbChecked)
    Enable txtDollarTrail, optTrailDollar.Value And (chkTrail.Value = vbChecked)
    Enable optTrailPoints, (chkTrail.Value = vbChecked)
    Enable txtPointTrail, optTrailPoints.Value And (chkTrail.Value = vbChecked)
    Enable sbPointTrail, optTrailPoints.Value And (chkTrail.Value = vbChecked)
    
    If chkTrail.Value = vbChecked Then
        If optBuy.Value = True Then
            optAbove.Value = True
        Else
            optBelow.Value = True
        End If
        If (Not m.Above Is Nothing) And (Not m.Bars Is Nothing) Then
            m.Above.Price = m.Bars(eBARS_Close, m.Bars.Size - 1) + m.TrailAmount.Price
        End If
        If (Not m.Below Is Nothing) And (Not m.Bars Is Nothing) Then
            m.Below.Price = m.Bars(eBARS_Close, m.Bars.Size - 1) - m.TrailAmount.Price
        End If
        
        CheckBoxValue(chkMoveWithTrigger) = True
    End If
    
    chkMoveWithTrigger.Enabled = CheckBoxValue(chkTriggerOrder) And (Not CheckBoxValue(chkTrail))
    
    If m.Order.IsAutomated Then
        Enable optAbove, optAbove.Value
        Enable optMarket, True
        Enable optBelow, optBelow.Value
        Enable optMIT, optMIT.Value
        Enable optStopLimit, optStopLimit.Value
    Else
        Enable optAbove, (chkTrail.Value = vbUnchecked)
        Enable optMarket, (chkTrail.Value = vbUnchecked)
        Enable optBelow, (chkTrail.Value = vbUnchecked)
        Enable optStopLimit, (chkTrail.Value = vbUnchecked)
        Enable optMIT, (chkTrail.Value = vbUnchecked)
        Enable txtAbove, (chkTrail.Value = vbUnchecked)
        Enable sbAbove, (chkTrail.Value = vbUnchecked)
        Enable txtBelow, (chkTrail.Value = vbUnchecked)
        Enable sbBelow, (chkTrail.Value = vbUnchecked)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.EnableAdvancedControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConditionRTF
'' Description: Set the condition rtf box
'' Inputs:      Expression
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetConditionRTF(ByVal strExpression As String)
On Error GoTo ErrSection:

    Dim Expr As New cExpression         ' Expression object
    Dim Func As New cFunction           ' Temporary function object
    
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule strExpression
        rtfCondition.TextRTF = Func.GetRTF(.EditText)
    End With
    
ErrExit:
    Set Expr = Nothing
    Set Func = Nothing
    Exit Sub
    
ErrSection:
    Set Expr = Nothing
    Set Func = Nothing
    RaiseError "frmTTEditOrder.SetConditionRTF"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyOrder
'' Description: Verify the order before allowing it to be submitted or parked
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyOrder(ByVal bPark As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strReturn As String             ' Return from an InfBox
    Dim Order As New cPtOrder           ' Temporary order object
    Dim lAccountID As Long              ' Account ID selected
    Dim nOrderType As eTT_OrderType     ' Order type selected
    Dim nTimeInForce As eTT_TimeInForce ' Time in force selected
    Dim dMarketPrice As Double          ' Market price for validating order prices
    Dim dOrderPrice As Double           ' New order price
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerExchange As String     ' Broker exchange

    If Len(Trim(txtSymbol.Text)) = 0 Then
        InfBox "You must supply a symbol for this to be a valid order", "!", , "Order Error"
        MoveFocus txtSymbol
        Exit Function
    End If
    
    If cboAccounts.ListIndex = -1& Then
        MoveFocus cboAccounts
        InfBox "Please supply an account for this order", "!", , "Order Error"
        Exit Function
    End If
    
    If IsGameMode = True Then
        'aardvark 3009 fix (order object is not needed and not set for game mode)
        If m.Qty <= 0 Then
            InfBox "A quantity must be specified.", "e", , "Error"
            MoveFocus txtQty
        Else
            VerifyOrder = True
        End If
        Exit Function
    End If
    
    lAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    
    If g.Broker.IsTradeableSymbol(lAccountID, SymbolOrSymbolID) = False Then
        g.Broker.ShowUnknownSymbolError GetSymbol(SymbolOrSymbolID), g.Broker.AccountTypeForID(AccountID), "EditOrder.VerifyOrder", True
        Exit Function
    End If
    
    ' 12/04/2014 DAJ: With the changes in exchange fees coming from the exchange, there is more of a
    ' possibility that traders could have data turned off, but trading turned on for symbols.  Because of
    ' this, we need to let them trade symbols that they don't get data for...
    'If g.Broker.IsEnabledSymbol(lAccountID, SymbolOrSymbolID, strBrokerBase, strBrokerExchange) = False Then
    '    g.Broker.ShowNotEnabledForSymbolError GetSymbol(SymbolOrSymbolID), g.Broker.AccountTypeForID(AccountID), strBrokerBase, strBrokerExchange, "EditOrder.VerifyOrder", True
    '    Exit Function
    'End If

    ' Check to see if the order type is allowed on the given symbol in the given account...
    nOrderType = OrderTypeSelected
    If g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, nOrderType) = False Then
        InfBox OrderType(nOrderType) & " orders not allowed for " & txtSymbol.Text & " on the " & g.Broker.BrokerName(g.Broker.AccountTypeForID(lAccountID)) & " servers", "!", , "Order Type Error"
        Exit Function
    End If
    
    ' Check to see if the time in force is allowed on the given symbol in the given account...
    nTimeInForce = TimeInForceSelected
    If g.Broker.TimeInForceAllowed(lAccountID, txtSymbol.Text, nTimeInForce) = False Then
        InfBox TimeInForce(nTimeInForce) & " orders not allowed for " & txtSymbol.Text & " on the " & g.Broker.BrokerName(g.Broker.AccountTypeForID(lAccountID)) & " servers", "!", , "Time in Force Error"
        Exit Function
    End If

    If chkExitAll.Value = vbUnchecked Then
        If m.Qty <= 0 Then
            InfBox "A quantity must be specified.", "e", , "Error"
            MoveFocus txtQty
            Exit Function
        End If
    End If
    
    ' 05/06/2010 DAJ: If this order is set to trigger from another order, use that order's
    ' order price instead of the current market price to validate the prices on this order
    ' since the market will be closer to the triggering order when it fills and triggers this
    ' order...
    dMarketPrice = TriggerPrice
    If dMarketPrice = 0# Then
        dMarketPrice = g.RealTime.LastKnownPrice(txtSymbol.Text)
    End If
    
    If (optStopLimit.Value = True) Then
        If m.Above.Price < m.Below.Price Then
            If optBuy.Value = True Then
                InfBox "The Limit value must be greater|than the Stop value", "!", , "Order Error"
            Else
                InfBox "The Stop value must be greater|than the Limit value", "!", , "Order Error"
            End If
            Exit Function
        End If
        If optBuy.Value = True Then
            dOrderPrice = m.Below.Price
        Else
            dOrderPrice = m.Above.Price
        End If
    ElseIf (optAbove.Value = True) Then
        If m.Above.Price <= dMarketPrice Then
            strReturn = "Y"
            If optBuy.Value = True Then
                If (g.Broker.WarnStopWrongSide = True) And (bPark = False) Then
                    strReturn = InfBox("The Stop Price for a Buy order should be greater than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                End If
            Else
                If (g.Broker.WarnStopWrongSide = True) And (bPark = False) Then
                    strReturn = InfBox("The Limit Price for a Sell order should be greater than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                End If
            End If
            If strReturn = "N" Then Exit Function
        End If
        dOrderPrice = m.Above.Price
    ElseIf (optBelow.Value = True) Then
        If m.Below.Price >= dMarketPrice Then
            strReturn = "Y"
            If optBuy.Value = True Then
                If (g.Broker.WarnStopWrongSide = True) And (bPark = False) Then
                    strReturn = InfBox("The Limit Price for a Buy order should be less than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                End If
            Else
                If (g.Broker.WarnStopWrongSide = True) And (bPark = False) Then
                    strReturn = InfBox("The Stop Price for a Sell order should be less than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                End If
            End If
            If strReturn = "N" Then Exit Function
        End If
        dOrderPrice = m.Below.Price
    ElseIf (optMIT.Value = True) Then
        strReturn = "Y"
        If optBuy.Value = True Then
            dOrderPrice = m.Below.Price
            If m.Below.Price >= dMarketPrice Then
                If (g.Broker.WarnStopWrongSide = True) And (bPark = False) Then
                    strReturn = InfBox("The Market if Touched Price for a Buy order should be less than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                End If
            End If
        Else
            dOrderPrice = m.Above.Price
            If m.Above.Price <= dMarketPrice Then
                If (g.Broker.WarnStopWrongSide = True) And (bPark = False) Then
                    strReturn = InfBox("The Market if Touched Price for a Sell order should be greater than the current market price.  Submitting the order in this way could lead to an immediate fill.||Do you want to continue anyway?|", "?", "+Yes|-No", "Order Confirmation")
                End If
            End If
        End If
    End If
    
    ' Check to see if the order can be moved to the new price...
    If (Not m.Order Is Nothing) And (bPark = False) Then
        If m.Order.OrderID > 0 Then
            If (CheckBoxValue(chkTriggerOrder) = False) And (CheckBoxValue(chkBeginTime) = False) And (CheckBoxValue(chkPrice) = False) And (CheckBoxValue(chkCustomCondition) = False) Then
                If CanMoveOrder(m.Order, , , dOrderPrice) = False Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    If (chkCancelOrder.Value = vbChecked) And (chkTriggerOrder.Value = vbChecked) Then
        If cboTriggerOrders.ItemData(cboTriggerOrders.ListIndex) = cboCancelOrders.ItemData(cboCancelOrders.ListIndex) Then
            InfBox "The same order cannot be used as both a trigger order and a cancel order", "!", , "Order Error"
            Exit Function
        End If
    End If
    
    If (chkCancelOrder.Value = vbChecked) And (m.Order.OrderID > 0) Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(cboCancelOrders.ItemData(cboCancelOrders.ListIndex)) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            If rs!TriggerOrderID = m.Order.OrderID Then
                InfBox "You cannot cancel an order that will be triggered by this order", "!", , "Order Error"
                Exit Function
            End If
        End If
    End If
    
    If (chkTriggerOrder.Value = vbChecked) And (optMarket.Value = True) Then
        If Order.Load(cboTriggerOrders.ItemData(cboTriggerOrders.ListIndex)) Then
            If (Order.Symbol = txtSymbol.Text) And (Order.IsConditional(False) = False) Then
                If InfBox("Triggering a market order from an order with the same symbol may cause undesirable results.||Are you sure you want to do this?|", "?", "+Yes|No", "Order Confirmation") = "N" Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    If (chkCustomCondition.Value = vbChecked) And (Len(Trim(rtfCondition.Text)) = 0) Then
        tabSpecialOrders.CurrTab = Tabs(eGDTab_Condition)
        MoveFocus chkCustomCondition
        InfBox "You must either supply a custom condition or turn off the 'Custom Condition' option", "!", , "Order Error"
        Exit Function
    End If
    
    ' DAJ 09/20/2012: Since we never did fully implement this, I am going to comment it out for now...
    ''ValidateSecTypeForAccount
    
    ' DAJ 12/18/2012: Moved this block down to the bottom of the function so that it is the
    ' last thing we verify.  Customer ran into a situation where the order filled while he had
    ' a question up about being on the wrong side of the market and we didn't verify the order
    ' status after he answered 'Yes'...
    If m.Order.Status <> m.nOrderStatus Then
        Select Case m.Order.Status
            'Case eTT_OrderStatus_Partial
            Case eTT_OrderStatus_Filled
                InfBox "This order has filled while you were editing it and can no longer be edited", "i", , "Order Status Change"
                m.nReturn = eGDEditOrderReturn_Cancel
                Hide
                Exit Function
            Case eTT_OrderStatus_Cancelled
                InfBox "This order was cancelled while you were editing it and can no longer be edited", "i", , "Order Status Change"
                m.nReturn = eGDEditOrderReturn_Cancel
                Hide
                Exit Function
            Case eTT_OrderStatus_Rejected
                InfBox "This order was rejected while you were editing it and can no longer be edited", "i", , "Order Status Change"
                m.nReturn = eGDEditOrderReturn_Cancel
                Hide
                Exit Function
            Case eTT_OrderStatus_BalCancelled
                InfBox "This balance of this order was cancelled while you were editing it and can no longer be edited", "i", , "Order Status Change"
                m.nReturn = eGDEditOrderReturn_Cancel
                Hide
                Exit Function
            Case eTT_OrderStatus_Expired
                InfBox "This order expired while you were editing it and can no longer be edited", "i", , "Order Status Change"
                m.nReturn = eGDEditOrderReturn_Cancel
                Hide
                Exit Function
        End Select
    End If
    
    VerifyOrder = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.VerifyOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPosition
'' Description: Determine the current position for the symbol/account/lot
'' Inputs:      None
'' Returns:     Current Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CurrentPosition() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strSymbol As String             ' Symbol
    Dim strAccountNumber As String      ' Account number for the selected account
    Dim strFeedYardLotID As String      ' Feed yard lot ID
    
    lReturn = 0&
    If IsGameMode = True Then
        lReturn = m.lGamePosition
    ElseIf (cboAccounts.ListIndex >= 0) And (Len(txtSymbol.Text) > 0) Then
        strFeedYardLotID = FeedYardLotID
        strSymbol = Trim(txtSymbol.Text)
        strAccountNumber = g.Broker.GetAccountNumber(AccountID)
        
        If Len(strFeedYardLotID) > 0 Then
            lReturn = g.CattleBridge.Position(strFeedYardLotID, strAccountNumber, Broker, strSymbol)
        Else
            lReturn = g.Broker.CurrentPosition(AccountID, strSymbol, m.Order.AutoTradeItemID)
        End If
    End If
    
    CurrentPosition = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.CurrentPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentPositionAfterTrigger
'' Description: Determine the current position taking into account the
''              triggering order if applicable
'' Inputs:      None
'' Returns:     Current Position after Trigger
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CurrentPositionAfterTrigger() As Long
On Error GoTo ErrSection:

    Dim lPosition As Long               ' Current Position
    Dim Trigger As New cPtOrder         ' Triggering order
    Dim lOrderID As Long                ' ID of the order to load
    
    If cboAccounts.ListIndex >= 0 Then
        'lPosition = g.Broker.CurrentPosition(cboAccounts.ItemData(cboAccounts.ListIndex), txtSymbol.Text, m.Order.AutoTradeItemID)
        lPosition = CurrentPosition
        If chkTriggerOrder.Value = vbChecked Then
            lOrderID = cboTriggerOrders.ItemData(cboTriggerOrders.ListIndex)
            Do While lOrderID > 0
                If Trigger.Load(lOrderID) Then
                    If Trigger.Symbol = txtSymbol.Text Then
                        If Trigger.Buy Then
                            lPosition = lPosition + Trigger.Quantity - Trigger.FillQuantity
                        Else
                            lPosition = lPosition - Trigger.Quantity + Trigger.FillQuantity
                        End If
                    End If
                    
                    lOrderID = Trigger.TriggerOrderID
                Else
                    lOrderID = 0&
                End If
            Loop
        End If
    End If
    
    CurrentPositionAfterTrigger = lPosition

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.CurrentPositionAfterTrigger"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableDurationControls
'' Description: Enable/Disable the duration controls as per the account selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableDurationControls()
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID for the chosen account
    
    lAccountID = AccountID
    If (OrderIsAutomated = True) Or (OrderIsWorking = True) Then
        optDay.Enabled = False
        optGTC.Enabled = False
        optExpiration.Enabled = False
    ElseIf lAccountID >= 0 Then
        optDay.Enabled = g.Broker.TimeInForceAllowed(lAccountID, txtSymbol.Text, eTT_TimeInForce_Day)
        optGTC.Enabled = g.Broker.TimeInForceAllowed(lAccountID, txtSymbol.Text, eTT_TimeInForce_GTC)
        optExpiration.Enabled = g.Broker.TimeInForceAllowed(lAccountID, txtSymbol.Text, eTT_TimeInForce_GTD)
    Else
        optDay.Enabled = True
        optGTC.Enabled = True
        optExpiration.Enabled = True
    End If

    Enable gdExpiration, optExpiration

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.EnableDurationControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableOrderTimeControls
'' Description: Enable/Disable the order time controls as per the account selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableOrderTimeControls()
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID for the chosen account
    
    lAccountID = AccountID
    If (OrderIsAutomated = True) Or (OrderIsWorking = True) Or (lAccountID = 0) Then
        Enable optAnyTime, False
        Enable optOnClose, False
        Enable optOnClose, False
    Else
        Enable optAnyTime, True
        
        If optMarket.Value = True Then
            Enable optOnClose, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_MarketOnClose)
            Enable optOnOpen, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_MarketOnOpen)
        
        ElseIf optAbove.Value = True Then
            If optBuy.Value = True Then
                Enable optOnClose, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_StopCloseOnly)
            ElseIf optBuy.Value = False Then
                Enable optOnClose, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_LimitCloseOnly)
            Else
                Enable optOnClose, False
            End If
            
            Enable optOnOpen, False
        
        ElseIf optBelow.Value = True Then
            If optBuy.Value = True Then
                Enable optOnClose, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_LimitCloseOnly)
            ElseIf optBuy.Value = False Then
                Enable optOnClose, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_StopCloseOnly)
            Else
                Enable optOnClose, False
            End If
        
            Enable optOnOpen, False
        
        ElseIf optMIT.Value = True Then
            Enable optOnClose, False
            Enable optOnOpen, False
        
        ElseIf optStopLimit.Value = True Then
            Enable optOnClose, g.Broker.OrderTypeAllowed(lAccountID, txtSymbol.Text, eTT_OrderType_StopWithLimitCloseOnly)
            Enable optOnOpen, False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.EnableOrderTimeControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableQuantityControls
'' Description: Enable/Disable the quantity controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableQuantityControls()
On Error GoTo ErrSection:

    Dim bEnable As Boolean              ' Enable the controls?

    If Not m.Order Is Nothing Then      'aardvark 6620 (gamemode does not have/use this object
        bEnable = (chkExitAll.Value = vbUnchecked) And (m.Order.IsAutomated = False)
    
        chkExitAll.Enabled = (UCase(Left(m.strPos, 1)) = "X") And (m.Order.IsAutomated = False)
        'If chkExitAll.Enabled = False Then chkExitAll.Value = vbUnchecked
    End If
    
    If IsGameMode = True Then
        bEnable = True          'aardvark issue 6664
    End If
    
    Enable txtQty, bEnable
    Enable vsCalc, bEnable
    Enable imgCalc, bEnable
    Enable sbQty, bEnable
    Enable fraQuantity, bEnable

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.EnableQuantityControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowSessionControls
'' Description: Show/Hide the session controls as per the account/symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowSessionControls()
On Error GoTo ErrSection:

    Dim nAcctType As eTT_AccountType    ' Account type for the selected account
    Dim strSymbol As String             ' Symbol from text box
    Dim strBaseSymbol As String         ' Base symbol
    Dim bShowSession As Boolean         ' Show session frame?
    Dim bPit As Boolean                 ' Current value of optPit
    
    If cboAccounts.ListIndex >= 0 Then
        nAcctType = g.Broker.AccountTypeForID(cboAccounts.ItemData(cboAccounts.ListIndex))
    Else
        nAcctType = eTT_AccountType_SimStream
    End If
    
    bShowSession = False
    fraSession.Visible = bShowSession
    Form_Resize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.ShowSessionControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeSelected
'' Description: Order type that is selected in the user interface
'' Inputs:      None
'' Returns:     Order Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderTypeSelected() As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function

    Select Case True
        Case optMarket
            nReturn = eTT_OrderType_Market
        Case optStopLimit
            nReturn = eTT_OrderType_StopWithLimit
        Case optAbove
            If optBuy Then
                nReturn = eTT_OrderType_Stop
            Else
                nReturn = eTT_OrderType_Limit
            End If
        Case optBelow
            If optBuy Then
                nReturn = eTT_OrderType_Limit
            Else
                nReturn = eTT_OrderType_Stop
            End If
        Case optMIT
            nReturn = eTT_OrderType_MIT
    End Select
    
    OrderTypeSelected = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.OrderTypeSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceSelected
'' Description: Time in Force that is selected in the user interface
'' Inputs:      None
'' Returns:     Time in Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TimeInForceSelected() As eTT_TimeInForce
On Error GoTo ErrSection:

    Dim nReturn As eTT_TimeInForce      ' Time in force selected
    
    Select Case True
        Case optDay
            nReturn = eTT_TimeInForce_Day
        Case optGTC
            nReturn = eTT_TimeInForce_GTC
        Case optExpiration
            nReturn = eTT_TimeInForce_GTD
    End Select
    
    TimeInForceSelected = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.TimeInForceSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInformation
'' Description: Get the contract information (if applicable) for given symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetContractInformation()
On Error GoTo ErrSection:

    If (cboAccounts.ListIndex >= 0) And (Len(Trim(txtSymbol.Text)) > 0) Then
        g.Broker.GetContractInfo g.Broker.AccountTypeForName(cboAccounts.Text), Trim(txtSymbol.Text), True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.GetContractInformation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowExchangeControls
'' Description: Show or hide the exchange controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowExchangeControls()
On Error GoTo ErrSection:

    Dim bVisible As Boolean             ' Do we want the exchange controls visible?
    
    bVisible = g.Broker.ShowExchangeControls(cboAccounts.Text, txtSymbol.Text)
    
    lblExchange.Visible = bVisible
    cboExchanges.Visible = bVisible
    
    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.ShowExchangeControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLotControls
'' Description: Show or hide the lot controls as appropriate
'' Inputs:      None
'' Returns:     True if visible, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ShowLotControls() As Boolean
On Error GoTo ErrSection:

    Dim bVisible As Boolean             ' Do we want the exchange controls visible?
    
    bVisible = (g.CattleBridge.CattleFormLoaded = True) And (m.bNewOrder = True)
    
    lblLot.Visible = bVisible
    cboLots.Visible = bVisible
    
    Form_Resize
    
    ShowLotControls = bVisible

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.ShowLotControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TriggerPrice
'' Description: Order price for the triggering order (if applicable)
'' Inputs:      None
'' Returns:     Trigger Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TriggerPrice() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    Dim Order As New cPtOrder           ' Triggering order
    
    dReturn = 0#
    If chkTriggerOrder.Value = vbChecked Then
        If Order.Load(cboTriggerOrders.ItemData(cboTriggerOrders.ListIndex)) Then
            dReturn = Order.OrderPrice(True)
        End If
    End If
    
    TriggerPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.TriggerPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SwapStopLimit
'' Description: Swap the Stop and Limit as appropriate as user changes trigger
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SwapStopLimit()
On Error GoTo ErrSection:

    Dim dTriggerPrice As Double         ' Trigger price

    If m.Order.OrderID <= 0 Then
        If (chkTriggerOrder.Value = vbChecked) Then
            dTriggerPrice = TriggerPrice
        Else
            dTriggerPrice = g.RealTime.LastKnownPrice(txtSymbol.Text)
        End If
        
        ' DAJ 10/08/2015: If the triggering order is a market order or the trigger price couldn't be determined
        ' for some reason, don't try to swap the stop and limit...
        If (dTriggerPrice <> kNullData) And (dTriggerPrice <> 0) Then
            If (optAbove.Value = True) And (m.Above.Price < dTriggerPrice) Then
                optBelow.Value = True
            ElseIf (optBelow.Value = True) And (m.Below.Price > dTriggerPrice) Then
                optAbove.Value = True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SwapStopLimit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLotsCombo
'' Description: Load the lots combo
'' Inputs:      Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLotsCombo(ByVal strFeedYardLotID As String)
On Error GoTo ErrSection:

    If g.CattleBridge Is Nothing Then
        cboLots.Clear
        cboLots.AddItem "None"
        cboLots.ItemData(cboLots.NewIndex) = -1&
        cboLots.ListIndex = 0
    Else
        g.CattleBridge.LoadLotsCombo cboLots, strFeedYardLotID, "None"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.LoadLotsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetProfitOtherLabel
'' Description: Set the profit other label
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetProfitOtherLabel()
On Error GoTo ErrSection:

    If optProfitDollar.Value = True Then
        lblProfitOther.Caption = " = " & m.Bars.PriceDisplay(ProfitPoints) & " points"
    Else
        lblProfitOther.Caption = " = " & Format(ProfitDollar, "$#,##0.00") & " dollars"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetProfitOtherLabel"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetStopOtherLabel
'' Description: Set the stop other label
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetStopOtherLabel()
On Error GoTo ErrSection:

    If optStopDollar.Value = True Then
        lblStopOther.Caption = " = " & m.Bars.PriceDisplay(StopPoints) & " points"
    Else
        lblStopOther.Caption = " = " & Format(StopDollar, "$#,##0.00") & " dollars"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.SetStopOtherLabel"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitProfitTarget
'' Description: Initialize the profit target price editor
'' Inputs:      For Dollars?, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitProfitTarget(ByVal bForDollars As Boolean, ByVal dValue As Double)
On Error GoTo ErrSection:

    If bForDollars = True Then
        peProfitTarget.Init Nothing, dValue, m.Bars.TickValue, , , m.Bars.TickValue
    Else
        peProfitTarget.Init m.Bars, dValue, m.Bars.MinMove
        peProfitTarget.AsTradingUnits = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.InitProfitTarget"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitStopLoss
'' Description: Initialize the stop loss price editor
'' Inputs:      For Dollars?, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitStopLoss(ByVal bForDollars As Boolean, ByVal dValue As Double)
On Error GoTo ErrSection:

    If bForDollars = True Then
        peStopLoss.Init Nothing, dValue, m.Bars.TickValue, , , m.Bars.TickValue
    Else
        peStopLoss.Init m.Bars, dValue, m.Bars.MinMove
        peStopLoss.AsTradingUnits = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.InitStopLoss"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitAbove
'' Description: Initialize the above price editor
'' Inputs:      Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitAbove(ByVal dValue As Double)
On Error GoTo ErrSection:

    Dim dMin As Double                  ' Minimum value for the price editor
    Dim bShowIfZero As Boolean          ' Show the price in the editor if it is zero?

    If IsSpreadSymbol(m.Bars.Prop(eBARS_Symbol)) Then
        dMin = -999999#
        bShowIfZero = True
    Else
        dMin = m.Bars.MinMove
        bShowIfZero = False
    End If
    
    m.Above.Init sbAbove, txtAbove, m.Bars, dValue, dMin, , , bShowIfZero

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.InitAbove"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitBelow
'' Description: Initialize the above price editor
'' Inputs:      Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitBelow(ByVal dValue As Double)
On Error GoTo ErrSection:

    Dim dMin As Double                  ' Minimum value for the price editor
    Dim bShowIfZero As Boolean          ' Show the price in the editor if it is zero?

    If IsSpreadSymbol(m.Bars.Prop(eBARS_Symbol)) Then
        dMin = -999999#
        bShowIfZero = True
    Else
        dMin = m.Bars.MinMove
        bShowIfZero = False
    End If
    
    m.Below.Init sbBelow, txtBelow, m.Bars, dValue, dMin, , , bShowIfZero

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.InitBelow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitConditionalPrice
'' Description: Initialize the conditional price editor
'' Inputs:      Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitConditionalPrice(ByVal dValue As Double)
On Error GoTo ErrSection:

    Dim dMin As Double                  ' Minimum value for the price editor
    Dim bShowIfZero As Boolean          ' Show the price in the editor if it is zero?

    If IsSpreadSymbol(m.Bars.Prop(eBARS_Symbol)) Then
        dMin = -999999#
        bShowIfZero = True
    Else
        dMin = m.Bars.MinMove
        bShowIfZero = False
    End If
    
    m.CondPrice.Init gdCondPrice, txtConditionPrice, m.CondBars, dValue, dMin, , , bShowIfZero

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTEditOrder.InitConditionalPrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetQuantityControl
'' Description: Set the quantity control based on the current symbol and account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetQuantityControl() As Long
On Error GoTo ErrSection:

    Dim lMinQuantity As Long            ' Minimum quantity for the account/symbol selected
    Dim lCurrentPosition As Long        ' Current position
    Dim bExit As Boolean                ' Order will be an exit
    
    lMinQuantity = g.Broker.MinimumOrderQuantity(AccountID, SymbolOrSymbolID)
    lCurrentPosition = CurrentPositionAfterTrigger
    bExit = (((optBuy.Value = True) And (lCurrentPosition < 0)) Or ((optSell.Value = True) And (lCurrentPosition > 0)))
    
    If (bExit = True) And (lCurrentPosition < lMinQuantity) Then
        lMinQuantity = lCurrentPosition
    End If
    
    m.Qty.Init sbQty, txtQty, Nothing, , lMinQuantity
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.SetQuantityControl"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGameMode
'' Description: Determine if the mode is currently a game mode
'' Inputs:      None
'' Returns:     True if Game Mode, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsGameMode() As Boolean
On Error GoTo ErrSection:

    IsGameMode = ((m.Mode = eGDTTEditOrderMode_GameNewOrder) Or (m.Mode = eGDTTEditOrderMode_GameEditOrder))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.IsGameMode"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIsAutomated
'' Description: Determine if the order for this form is automated
'' Inputs:      None
'' Returns:     True if Automated, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderIsAutomated() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not m.Order Is Nothing Then
        bReturn = m.Order.IsAutomated
    End If
    
    OrderIsAutomated = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.OrderIsAutomated"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIsWorking
'' Description: Determine if the order for this form is working
'' Inputs:      None
'' Returns:     True if Working, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderIsWorking() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Not m.Order Is Nothing Then
        bReturn = ((m.Order.OrderID > 0) And (m.Order.Status <> eTT_OrderStatus_Parked))
    End If
    
    OrderIsWorking = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTEditOrder.OrderIsWorking"
    
End Function


