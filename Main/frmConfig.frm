VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Settings"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   2235
      TabIndex        =   64
      Top             =   4185
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
      Caption         =   "frmConfig.frx":000C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmConfig.frx":0038
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmConfig.frx":0058
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   0
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
         Caption         =   "frmConfig.frx":0074
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConfig.frx":009E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConfig.frx":00BE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   0
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
         Caption         =   "frmConfig.frx":00DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConfig.frx":0108
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConfig.frx":0128
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab vstCfg 
      Height          =   4050
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   7144
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
      Caption         =   "&Account|Auto &Update|&Real Time|&Tick Data|&Misc."
      Align           =   0
      Appearance      =   1
      CurrTab         =   4
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
      Picture(0)      =   "frmConfig.frx":0144
      Picture(1)      =   "frmConfig.frx":029E
      Picture(2)      =   "frmConfig.frx":03F8
      Picture(3)      =   "frmConfig.frx":0552
      Picture(4)      =   "frmConfig.frx":0AEC
      Begin HexUniControls.ctlUniFrameWL fraMisc 
         Height          =   3630
         Left            =   45
         TabIndex        =   63
         Top             =   375
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   6403
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmConfig.frx":1086
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmConfig.frx":10A6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmConfig.frx":10C6
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdTheme 
            Height          =   405
            Left            =   2760
            TabIndex        =   2
            Top             =   3120
            Width           =   1575
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
            Caption         =   "frmConfig.frx":10E2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":1120
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1140
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkDivAdjust 
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   1875
            Width           =   5355
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
            Caption         =   "frmConfig.frx":115C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":120A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":122A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkChartsAutoSize 
            Height          =   255
            Left            =   600
            TabIndex        =   6
            Top             =   2250
            Width           =   5955
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
            Caption         =   "frmConfig.frx":1246
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":12F6
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1316
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboArchive 
            Height          =   315
            Left            =   4080
            TabIndex        =   8
            Top             =   1500
            Width           =   1515
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
            Tip             =   "frmConfig.frx":1332
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1352
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdAppBk 
            Height          =   405
            Left            =   4680
            TabIndex        =   12
            Top             =   3120
            Width           =   1695
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
            Caption         =   "frmConfig.frx":136E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":13AA
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":13CA
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkLocalTimeZone 
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   240
            Width           =   5955
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
            Caption         =   "frmConfig.frx":13E6
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmConfig.frx":14A0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":14C0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdConnection 
            Height          =   405
            Left            =   360
            TabIndex        =   58
            Top             =   3120
            Width           =   1935
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
            Caption         =   "frmConfig.frx":14DC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":1522
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1542
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtSymbolExpire 
            Height          =   300
            Left            =   3900
            TabIndex        =   48
            Top             =   630
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmConfig.frx":155E
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
            Tip             =   "frmConfig.frx":1580
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":15A0
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDaysToSave 
            Height          =   300
            Left            =   3900
            TabIndex        =   52
            Top             =   1080
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmConfig.frx":15BC
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
            Tip             =   "frmConfig.frx":15E0
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1600
         End
         Begin HexUniControls.ctlUniCheckXP chkSymbolExpire 
            Height          =   255
            Left            =   600
            TabIndex        =   47
            Top             =   660
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
            Caption         =   "frmConfig.frx":161C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmConfig.frx":1688
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":16A8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkScans 
            Height          =   255
            Left            =   5400
            TabIndex        =   50
            Top             =   720
            Visible         =   0   'False
            Width           =   5955
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
            Caption         =   "frmConfig.frx":16C4
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":1774
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1794
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor clrAltRow 
            Height          =   315
            Left            =   6060
            TabIndex        =   55
            Top             =   1260
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Color           =   13168895
            CustomColor     =   13168895
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDeveloper 
            Height          =   345
            Left            =   2520
            TabIndex        =   57
            Top             =   3143
            Visible         =   0   'False
            Width           =   3210
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmConfig.frx":17B0
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
            Tip             =   "frmConfig.frx":17D0
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":17F0
         End
         Begin HexUniControls.ctlUniCheckXP chkExtendBlankBars 
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   2820
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
            Caption         =   "frmConfig.frx":180C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":18A2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":18C2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSnapToDots 
            Height          =   255
            Left            =   1020
            TabIndex        =   27
            Top             =   2520
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
            Caption         =   "frmConfig.frx":18DE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":1980
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":19A0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label6 
            Height          =   255
            Left            =   600
            Top             =   1560
            Width           =   3735
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
            Caption         =   "frmConfig.frx":19BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":1A3C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1A5C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDeleteFiles 
            Height          =   255
            Left            =   600
            Top             =   1140
            Width           =   3255
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
            Caption         =   "frmConfig.frx":1A78
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":1AF0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1B10
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDays 
            Height          =   255
            Left            =   4500
            Top             =   1140
            Width           =   435
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
            Caption         =   "frmConfig.frx":1B2C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":1B56
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1B76
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   4500
            Top             =   660
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
            Caption         =   "frmConfig.frx":1B92
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":1BC4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1BE4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   255
            Left            =   5880
            Top             =   1020
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
            Caption         =   "frmConfig.frx":1C00
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":1C68
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1C88
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDeveloper 
            Height          =   240
            Left            =   600
            Top             =   3255
            Visible         =   0   'False
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
            Caption         =   "frmConfig.frx":1CA4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":1CF2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1D12
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin vsOcx6LibCtl.vsElastic vseMisc 
         Height          =   3630
         Left            =   -7590
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   375
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   6403
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
         Align           =   0
         Appearance      =   0
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
         Begin HexUniControls.ctlUniFrameWL fraPrecisionTick 
            Height          =   1695
            Left            =   360
            TabIndex        =   31
            Top             =   1440
            Width           =   6195
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
            Caption         =   "frmConfig.frx":1D2E
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":1DBE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":1DDE
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtAutoTickFill 
               Height          =   315
               Left            =   3540
               TabIndex        =   32
               Top             =   1110
               Width           =   435
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":1DFA
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
               Tip             =   "frmConfig.frx":1E1E
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":1E3E
            End
            Begin HexUniControls.ctlUniCheckXP chkTickDataFill 
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   270
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
               Caption         =   "frmConfig.frx":1E5A
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmConfig.frx":1F1A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":1F3A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optAutoTickFill 
               Height          =   220
               Left            =   480
               TabIndex        =   36
               Top             =   1140
               Width           =   3135
               _ExtentX        =   5530
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
               Caption         =   "frmConfig.frx":1F56
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmConfig.frx":1FC0
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":1FE0
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optMinTickFill 
               Height          =   220
               Left            =   480
               TabIndex        =   39
               Top             =   840
               Width           =   5235
               _ExtentX        =   9234
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
               Caption         =   "frmConfig.frx":1FFC
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":20A8
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":20C8
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optFullTickFill 
               Height          =   220
               Left            =   480
               TabIndex        =   41
               Top             =   570
               Width           =   4995
               _ExtentX        =   8811
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
               Caption         =   "frmConfig.frx":20E4
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":2186
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":21A6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label22 
               Height          =   255
               Left            =   4020
               Top             =   1140
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
               Caption         =   "frmConfig.frx":21C2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":2210
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2230
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label23 
               Height          =   255
               Left            =   900
               Top             =   1380
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
               Caption         =   "frmConfig.frx":224C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":22AA
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":22CA
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraScrubLevel 
            Height          =   1035
            Left            =   1080
            TabIndex        =   40
            Top             =   180
            Width           =   4815
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
            Caption         =   "frmConfig.frx":22E6
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":2322
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":2342
            RightToLeft     =   0   'False
            Begin MSComctlLib.Slider slScrubLevel 
               Height          =   450
               Left            =   2220
               TabIndex        =   43
               Top             =   240
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   794
               _Version        =   393216
               LargeChange     =   1
            End
            Begin HexUniControls.ctlUniLabelXP Label16 
               Height          =   195
               Left            =   600
               Top             =   600
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
               Caption         =   "frmConfig.frx":235E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":239C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":23BC
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label15 
               Height          =   195
               Left            =   3960
               Top             =   720
               Width           =   315
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
               Caption         =   "frmConfig.frx":23D8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":23FE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":241E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label14 
               Height          =   195
               Left            =   3000
               Top             =   720
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
               Caption         =   "frmConfig.frx":243A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":2466
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2486
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label13 
               Height          =   195
               Left            =   2280
               Top             =   720
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
               Caption         =   "frmConfig.frx":24A2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":24C8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":24E8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTickScrubbing 
               Height          =   195
               Left            =   420
               Top             =   360
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
               Caption         =   "frmConfig.frx":2504
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":254E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":256E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
      Begin vsOcx6LibCtl.vsElastic vseRealTime 
         Height          =   3630
         Left            =   -7890
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   375
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   6403
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
         Align           =   0
         Appearance      =   0
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
         Begin HexUniControls.ctlUniFrameWL fraRealTime 
            Height          =   2655
            Left            =   180
            TabIndex        =   65
            Top             =   240
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
            Caption         =   "frmConfig.frx":258A
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":2604
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":2624
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraRecalc 
               Height          =   1455
               Left            =   1620
               TabIndex        =   42
               Top             =   1140
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
               Caption         =   "frmConfig.frx":2640
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmConfig.frx":2698
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":26B8
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtRecalcSeconds 
                  Height          =   285
                  Left            =   1500
                  TabIndex        =   44
                  Top             =   570
                  Width           =   435
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmConfig.frx":26D4
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
                  Tip             =   "frmConfig.frx":26F8
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmConfig.frx":2718
               End
               Begin HexUniControls.ctlUniRadioXP optRecalcSeconds 
                  Height          =   220
                  Left            =   300
                  TabIndex        =   45
                  Top             =   600
                  Width           =   1395
                  _ExtentX        =   2461
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
                  Caption         =   "frmConfig.frx":2734
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmConfig.frx":2768
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmConfig.frx":2788
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optRecalcBar 
                  Height          =   220
                  Left            =   300
                  TabIndex        =   46
                  Top             =   300
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
                  Caption         =   "frmConfig.frx":27A4
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmConfig.frx":27DC
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmConfig.frx":27FC
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optRecalcTick 
                  Height          =   220
                  Left            =   300
                  TabIndex        =   49
                  Top             =   900
                  Width           =   2775
                  _ExtentX        =   4895
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
                  Caption         =   "frmConfig.frx":2818
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   -1  'True
                  Tip             =   "frmConfig.frx":287C
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmConfig.frx":289C
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label19 
                  Height          =   195
                  Left            =   660
                  Top             =   1110
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
                  Caption         =   "frmConfig.frx":28B8
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   0
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmConfig.frx":291C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmConfig.frx":293C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP Label18 
                  Height          =   195
                  Left            =   1980
                  Top             =   600
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
                  Caption         =   "frmConfig.frx":2958
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmConfig.frx":2986
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmConfig.frx":29A6
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniComboImageXP cboFeeds 
               Height          =   315
               Left            =   3600
               TabIndex        =   38
               Top             =   570
               Width           =   1695
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
               Tip             =   "frmConfig.frx":29C2
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":29E2
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkActivate 
               Height          =   255
               Left            =   1080
               TabIndex        =   37
               Top             =   600
               Width           =   2535
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
               Caption         =   "frmConfig.frx":29FE
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":2A58
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2A78
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblRTDescription 
               Height          =   495
               Left            =   600
               Top             =   0
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
               Caption         =   "frmConfig.frx":2A94
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":2B94
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2BB4
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniLabelXP lblNoRealTime 
            Height          =   735
            Left            =   60
            Top             =   2580
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
            Caption         =   "frmConfig.frx":2BD0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmConfig.frx":2CF4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":2D14
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin vsOcx6LibCtl.vsElastic vseAutoDownload 
         Height          =   3630
         Left            =   -8190
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   375
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   6403
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
         Align           =   0
         Appearance      =   0
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
         Begin HexUniControls.ctlUniFrameWL fraDisplayTime 
            Height          =   375
            Left            =   1020
            TabIndex        =   51
            Top             =   120
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
            Caption         =   "frmConfig.frx":2D30
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":2D50
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":2D70
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optNewYorkTime 
               Height          =   255
               Left            =   3300
               TabIndex        =   53
               Top             =   120
               Width           =   1755
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
               Caption         =   "frmConfig.frx":2D8C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmConfig.frx":2DC6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2DE6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optLocalTime 
               Height          =   255
               Left            =   1980
               TabIndex        =   54
               Top             =   120
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
               Caption         =   "frmConfig.frx":2E02
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":2E36
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2E56
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label4 
               Height          =   255
               Left            =   180
               Top             =   120
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
               Caption         =   "frmConfig.frx":2E72
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":2EBE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2EDE
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraDailyDownload 
            Height          =   1335
            Left            =   675
            TabIndex        =   20
            Top             =   600
            Width           =   5535
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
            Caption         =   "frmConfig.frx":2EFA
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":2F32
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":2F52
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectDate gdStartAuto 
               Height          =   300
               Left            =   3240
               TabIndex        =   22
               Top             =   270
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ShowDayOfWeek   =   0   'False
               ShowDate        =   0
               ShowTime        =   2
               MinDate         =   0
               MaxDate         =   0.999988425925926
               Value           =   0
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTryHours 
               Height          =   285
               Left            =   2940
               TabIndex        =   26
               Top             =   660
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":2F6E
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
               Tip             =   "frmConfig.frx":2F8E
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":2FAE
            End
            Begin HexUniControls.ctlUniTextBoxXP txtTryInterval 
               Height          =   285
               Left            =   1500
               TabIndex        =   24
               Top             =   660
               Width           =   495
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":2FCA
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
               Tip             =   "frmConfig.frx":2FEA
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":300A
            End
            Begin HexUniControls.ctlUniCheckXP chkAutoDownload 
               Height          =   255
               Left            =   360
               TabIndex        =   21
               Top             =   300
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
               Caption         =   "frmConfig.frx":3026
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":308A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":30AA
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblDURealTimeMsg 
               Height          =   255
               Left            =   180
               Top             =   1020
               Visible         =   0   'False
               Width           =   5175
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
               Caption         =   "frmConfig.frx":30C6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":3162
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3182
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label12 
               Height          =   255
               Left            =   3540
               Top             =   690
               Visible         =   0   'False
               Width           =   1575
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
               Caption         =   "frmConfig.frx":319E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":31EC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":320C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label11 
               Height          =   255
               Left            =   2100
               Top             =   690
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
               Caption         =   "frmConfig.frx":3228
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":327A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":329A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTryInterval 
               Height          =   255
               Left            =   720
               Top             =   690
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
               Caption         =   "frmConfig.frx":32B6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":32E8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3308
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraQuoteBoard 
            Height          =   1335
            Left            =   675
            TabIndex        =   28
            Top             =   2040
            Width           =   5535
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
            Caption         =   "frmConfig.frx":3324
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":336A
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":338A
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtQuoteInterval 
               Height          =   285
               Left            =   3600
               TabIndex        =   30
               Top             =   300
               Width           =   495
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":33A6
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
               Tip             =   "frmConfig.frx":33C6
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":33E6
            End
            Begin HexUniControls.ctlUniCheckXP chkAutoQuotes 
               Height          =   255
               Left            =   360
               TabIndex        =   29
               Top             =   300
               Width           =   3195
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
               Caption         =   "frmConfig.frx":3402
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":3470
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3490
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectDate gdStartTime 
               Height          =   300
               Left            =   1140
               TabIndex        =   33
               Top             =   660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ShowDayOfWeek   =   0   'False
               ShowCalendar    =   0   'False
               ShowDate        =   0
               ShowTime        =   2
               MinDate         =   0
               MaxDate         =   0.999988425925926
               Value           =   0
            End
            Begin gdOCX.gdSelectDate gdEndTime 
               Height          =   300
               Left            =   2700
               TabIndex        =   35
               Top             =   660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ShowDayOfWeek   =   0   'False
               ShowCalendar    =   0   'False
               ShowDate        =   0
               ShowTime        =   2
               MinDate         =   0
               MaxDate         =   0.999988425925926
               Value           =   0
            End
            Begin HexUniControls.ctlUniLabelXP lblQbRealTimeMsg 
               Height          =   195
               Left            =   180
               Top             =   1050
               Visible         =   0   'False
               Width           =   5175
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
               Caption         =   "frmConfig.frx":34AC
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":355E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":357E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label2 
               Height          =   255
               Left            =   660
               Top             =   720
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
               Caption         =   "frmConfig.frx":359A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":35C2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":35E2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label3 
               Height          =   255
               Left            =   2460
               Top             =   720
               Width           =   135
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
               Caption         =   "frmConfig.frx":35FE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":3622
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3642
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label9 
               Height          =   255
               Left            =   4200
               Top             =   330
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
               Caption         =   "frmConfig.frx":365E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":368C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":36AC
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
      Begin vsOcx6LibCtl.vsElastic vseAccount 
         Height          =   3630
         Left            =   -8490
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   375
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   6403
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
         Align           =   0
         Appearance      =   0
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
         Begin HexUniControls.ctlUniFrameWL fraAccount 
            Height          =   1425
            Left            =   180
            TabIndex        =   56
            Top             =   120
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
            Caption         =   "frmConfig.frx":36C8
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":3700
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":3720
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtCustomerID 
               Height          =   285
               Left            =   1200
               TabIndex        =   3
               Top             =   300
               Width           =   855
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":373C
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
               Tip             =   "frmConfig.frx":375C
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":377C
            End
            Begin HexUniControls.ctlUniTextBoxXP txtDataServID 
               Height          =   285
               Left            =   3300
               TabIndex        =   5
               Top             =   300
               Width           =   495
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":3798
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
               Tip             =   "frmConfig.frx":37BE
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":37DE
            End
            Begin HexUniControls.ctlUniTextBoxXP txtMachineID 
               Height          =   285
               Left            =   1200
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   1020
               Width           =   2595
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   -1  'True
               Text            =   "frmConfig.frx":37FA
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
               Tip             =   "frmConfig.frx":381A
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":383A
            End
            Begin HexUniControls.ctlUniTextBoxXP txtPassword 
               Height          =   285
               Left            =   1200
               TabIndex        =   7
               Top             =   660
               Width           =   1755
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmConfig.frx":3856
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
               PasswordChar    =   "*"
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmConfig.frx":3876
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3896
            End
            Begin HexUniControls.ctlUniLabelXP lblPassword 
               Height          =   255
               Left            =   180
               Top             =   690
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
               Caption         =   "frmConfig.frx":38B2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":38E6
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3906
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblMachineID 
               Height          =   255
               Left            =   180
               Top             =   1035
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
               Caption         =   "frmConfig.frx":3922
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":3958
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3978
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblDataServiceID 
               Height          =   255
               Left            =   2280
               Top             =   330
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
               Caption         =   "frmConfig.frx":3994
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":39D0
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":39F0
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblCustomerID 
               Height          =   255
               Left            =   180
               Top             =   330
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
               Caption         =   "frmConfig.frx":3A0C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":3A46
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3A66
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraNextGen 
            Height          =   1245
            Left            =   4380
            TabIndex        =   66
            Top             =   300
            Visible         =   0   'False
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
            Caption         =   "frmConfig.frx":3A82
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":3AC8
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":3AE8
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optClassic 
               Height          =   220
               Left            =   1260
               TabIndex        =   11
               Top             =   300
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "frmConfig.frx":3B04
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":3B32
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3B52
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optNextGen 
               Height          =   220
               Left            =   180
               TabIndex        =   10
               Top             =   300
               Width           =   1155
               _ExtentX        =   2037
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
               Caption         =   "frmConfig.frx":3B6E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":3B9C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3BBC
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label10 
               Height          =   615
               Left            =   120
               Top             =   540
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
               Caption         =   "frmConfig.frx":3BD8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":3C8E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3CAE
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdTransfer 
            Height          =   975
            Left            =   4560
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   975
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
            Caption         =   "frmConfig.frx":3CCA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmConfig.frx":3D16
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":3DA8
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraAuth 
            Height          =   1995
            Left            =   180
            TabIndex        =   18
            Top             =   1590
            Visible         =   0   'False
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   3519
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmConfig.frx":3DC4
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmConfig.frx":3E00
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmConfig.frx":3E20
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optConnectionInfo 
               Height          =   240
               Left            =   4200
               TabIndex        =   15
               Top             =   190
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmConfig.frx":3E3C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":3E84
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3EA4
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optPurchased 
               Height          =   240
               Left            =   2520
               TabIndex        =   14
               Top             =   190
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmConfig.frx":3EC0
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":3F02
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3F22
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optEnablements 
               Height          =   240
               Left            =   720
               TabIndex        =   13
               Top             =   195
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmConfig.frx":3F3E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmConfig.frx":3F80
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3FA0
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtCodes 
               Height          =   1410
               Left            =   60
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   525
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   2487
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Enabled         =   -1  'True
               Locked          =   -1  'True
               Text            =   "frmConfig.frx":3FBC
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
               ScrollBars      =   2
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmConfig.frx":3FDC
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":3FFC
            End
            Begin HexUniControls.ctlUniLabelXP lblDisplayCodes 
               Height          =   240
               Left            =   60
               Top             =   190
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmConfig.frx":4018
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   0
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmConfig.frx":4048
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmConfig.frx":4068
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmConfig.frm
'' Description: Form to allow the user to change certain program settings
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/03/2012   DAJ         Remove Photon
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eConfigTabs
    eNoChange = -1
    eAccountTab = 0
    eSnapQuoteTab = 1
    eRealTimeTab = 2
    eTickDataTab = 3
    eMiscTab = 4
End Enum

Private Type mPrivate
    bOK As Boolean
    eCurrTab As eConfigTabs
    bAutoDownloadDirty As Boolean
    bAutoRefreshDirty As Boolean
    bFormLoadDone As Boolean
End Type
Private m As mPrivate

Private Function Tabs(ByVal lTab As eConfigTabs) As Long
    Tabs = lTab
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkActivate_Click
'' Description: If the user clicks on the Activate Real Time check box, enable/
''              disable the appropriate controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkActivate_Click()
On Error GoTo ErrSection:

    EnableControls
    If chkActivate Then
        vstCfg.TabPicture(eRealTimeTab) = Picture16(ToolbarIcon("kGreenLight"))
    Else
        vstCfg.TabPicture(eRealTimeTab) = Picture16(ToolbarIcon("kRedLight"))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkActivate.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAutoDownload_Click
'' Description: If the user clicks on the Auto Download check box, enable/
''              disable the appropriate controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAutoDownload_Click()
On Error GoTo ErrSection:

    'Bring up filter download form if user has never seen it
    If chkAutoDownload.Value = 1 And m.bFormLoadDone Then
        If Len(GetIniFileProperty("DownloadExclude", "", "General", g.strIniFile)) < 1 Then
            If Len(GetIniFileProperty("DownloadInclude", "", "General", g.strIniFile)) < 1 Then
                If ExtremeCharts = 1 Or HasModule("TRANS") Then
                    frmFilterDownload.ShowMe True
                Else
                    frmFilterDownload.ShowMe
                End If
            End If
        End If
    End If
    
    m.bAutoDownloadDirty = True
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkAutoDownload.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAutoQuotes_Click
'' Description: If the user clicks on the Auto Quotes check box, enable/disable
''              the appropriate controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAutoQuotes_Click()
On Error GoTo ErrSection:

    m.bAutoRefreshDirty = True
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkAutoQuotes.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkChartsAutoSize_Click()
On Error GoTo ErrSection:
    
    If Not Me.Visible Then Exit Sub
    
    If chkChartsAutoSize.Value = vbChecked Then
        If chkSnapToDots.Value = vbChecked Then
            chkSnapToDots.Value = vbUnchecked
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkChartsAutoSize_Click", eGDRaiseError_Show

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkScans_Click
'' Description: If the user clicks on the Scans check box, enable/disable the
''              appropriate controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkScans_Click()
On Error GoTo ErrSection:
    
    Dim strMsg$
    
    If Me.Visible Then
        If chkScans.Value <> 0 Then
            strMsg = "NOTE: when filters are enabled, some additional time will be taken after each daily update to recalculate the criteria for every symbol."
            If InfBox(strMsg, "i", "+OK|-Cancel", "Enabling Filters") = "C" Then
                chkScans.Value = 0
            End If
        End If
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkScans.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkSnapToDots_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    
    If chkSnapToDots.Value = vbChecked Then
        If chkChartsAutoSize.Value = vbChecked Then
            chkChartsAutoSize.Value = vbUnchecked
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkSnapToDots_Click", eGDRaiseError_Show

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkSymbolExpire_Click
'' Description: If the user clicks on the Symbol Expire check box, enable/disable
''              the appropriate controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkSymbolExpire_Click()
On Error GoTo ErrSection:
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkSymbolExpire.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub chkTickDataFill_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.chkTickDataFill_Click"
    Resume ErrExit
End Sub

Private Sub cmdAppBk_Click()
On Error GoTo ErrSection:

    frmAppBkCfg.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdAppBk_Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide
    'Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdConnection_Click()
On Error GoTo ErrSection:

    frmHTTP.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdConnection.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSave_Click
'' Description: If the user clicks on the Save button, save the information to
''              the registry and unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    Dim i&, d#
    Dim strKey As String                ' Key into the registry
    Dim strPassword As String           ' Password to store in registry
    Dim bRedoSymbolAccess As Boolean    ' Do we need to reset access flags?
    Dim strText As String
    Dim bUpdateCharts As Boolean        ' Do we need to update charts?

    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdSave

    ' TLB 11/9/2011: "1" is no longer a valid CustomerID (since some were inadvertantly putting their DataServiceID here)
    d = ValOfText(txtCustomerID)
    If d < 2 Or d > 999999 Then
        SelectAll txtCustomerID
        InfBox "Invalid Customer ID", "!", , "Error"
        'txtCustomerID.Text = ""
        MoveFocus txtCustomerID
        Exit Sub
    End If
    d = ValOfText(txtDataServID)
    If d < 1 Or d > 999 Then
        SelectAll txtDataServID
        InfBox "Invalid Data Service", "!", , "Error"
        MoveFocus txtDataServID
        Exit Sub
    End If

    ' Confirm password if has changed
    strPassword = Trim(UCase(txtPassword))
    If strPassword <> txtPassword.Tag Then
        strText = InfBox("Please enter password again to confirm ...", _
                        "?", , "Confirm Password", , , , , , "p")
        If UCase(Trim(strText)) <> strPassword Then
            ' if mismatch, restore old password
            txtPassword = txtPassword.Tag
            strText = "Second password did not match."
            If Len(txtPassword) > 0 Then
                strText = strText & "|The previous password has been restored."
            End If
            Beep
            InfBox strText, "e", , "Password Error"
            MoveFocus txtPassword
            Exit Sub
        End If
        txtPassword = strPassword
    End If

    ' Save all settings
    If lblCustomerID.Caption = "&User Name:" Then
        SetUserInfo Trim(txtCustomerID.Text), Trim(txtPassword.Text)
    Else
        'RI_SetDataServiceID CLng(ValOfText(txtDataServID.Text))
        RI_SetDataServiceID CLng(ValOfText(txtCustomerID.Text) * 1000 + ValOfText(txtDataServID.Text))
        RI_SetUserPassword txtPassword.Text
    End If

    ' Check if symbol access has changed
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    If ValOfText(txtSymbolExpire) <> GetRegistryValue(rkLocalMachine, strKey, "SymbolExpire", 31#) Then bRedoSymbolAccess = True
    If chkSymbolExpire <> GetRegistryValue(rkLocalMachine, strKey, "chkSymbolExpire", chkSymbolExpire) Then bRedoSymbolAccess = True

    If chkLocalTimeZone <> GetRegistryValue(rkLocalMachine, strKey, "DisplayLocalTimeZone", chkLocalTimeZone) Then
        bUpdateCharts = True
        g.bShowInLocalTimeZone = chkLocalTimeZone
        If FormIsLoaded("frmTTSummary") Then frmTTSummary.RefreshForm
    End If

    ' Save General settings
    SetRegistryValue rkLocalMachine, strKey, "DeleteZipFiles", ValOfText(txtDaysToSave)
    SetRegistryValue rkLocalMachine, strKey, "SymbolExpire", ValOfText(txtSymbolExpire)
    SetRegistryValue rkLocalMachine, strKey, "chkSymbolExpire", chkSymbolExpire
    SetRegistryValue rkLocalMachine, strKey, "DisplayLocalTimeZone", chkLocalTimeZone
    'SetRegistryValue rkLocalMachine, strKey, "DaysInScan", ValOfText(txtScanDays.Text)
    ScansEnabled = chkScans.Value
    'SetRegistryValue rkLocalMachine, strKey, "AltGridRowColor", clrAltRow.Color
    'ALT_GRID_ROW_COLOR = clrAltRow.Color
    
    If optRecalcSeconds.Value Then
        i = Round(ValOfText(txtRecalcSeconds))
        If i < 1 Then i = 1
        If i > 999 Then i = 999
        g.nRecalcIndRT = i
    ElseIf optRecalcBar.Value Then
        g.nRecalcIndRT = -1
    Else
        g.nRecalcIndRT = 0
    End If
    SetIniFileProperty "RecalcIndRT", g.nRecalcIndRT, "General", g.strIniFile

If 1 Then
    SetRegistryValue rkLocalMachine, strKey, "ScrubLevel", slScrubLevel.Value
    If slScrubLevel.Value <> g.iScrubLevel Then bUpdateCharts = True
    g.iScrubLevel = slScrubLevel.Value
ElseIf slScrubLevel.Value <> 0 Then
'temporary code
    slScrubLevel = 0
    InfBox "Auto-scrubbing will be enabled in the next version.", "i", , "Under Construction"
End If

    ' Save Auto Downloading information to registry
    SetRegistryValue rkLocalMachine, strKey, "LocalTime", optLocalTime.Value
    If chkAutoDownload.Enabled Then
        SetRegistryValue rkLocalMachine, strKey, "AutoUpdate", chkAutoDownload.Value
    End If
    SetRegistryValue rkLocalMachine, strKey, "AutoStart", gdStartAuto.Value
    SetRegistryValue rkLocalMachine, strKey, "TryInterval", CLng(ValOfText(txtTryInterval.Text))
    SetRegistryValue rkLocalMachine, strKey, "TryHours", CLng(ValOfText(txtTryHours.Text))
    If m.bAutoDownloadDirty Then CalcNextTryTime

    ' Save Snap Quote information to registry
    If chkAutoQuotes.Enabled Then
        SetRegistryValue rkLocalMachine, strKey, "AutoQuotes", chkAutoQuotes.Value
    End If
    SetRegistryValue rkLocalMachine, strKey, "QuoteInterval", ValOfText(txtQuoteInterval.Text)
    SetRegistryValue rkLocalMachine, strKey, "QuoteStart", gdStartTime.Value
    SetRegistryValue rkLocalMachine, strKey, "QuoteEnd", gdEndTime.Value
    If m.bAutoRefreshDirty Then CalcNextQuoteRefresh
    
    ' Save Real Time information to the registry
    SetRegistryValue rkLocalMachine, strKey, "ActivateRT", chkActivate.Value
    SetRegistryValue rkLocalMachine, strKey, "ActiveFeedRT", cboFeeds.Text

    ' Save PrecisionTick information to the registry
    If chkTickDataFill.Value <> GetRegistryValue(rkLocalMachine, strKey, "TickDataFill", vbChecked) Then
        SetRegistryValue rkLocalMachine, strKey, "TickDataFill", chkTickDataFill.Value
        bUpdateCharts = True
    End If
    If optFullTickFill Then
        i = 1
    ElseIf optMinTickFill Then
        i = 2
    Else
        i = 0
    End If
    If i <> GetRegistryValue(rkLocalMachine, strKey, "TickFillMode", 0) Then
        SetRegistryValue rkLocalMachine, strKey, "TickFillMode", i
        bUpdateCharts = True
    End If
    i = Val(txtAutoTickFill)
    If i < 1 Then i = 30 '(default)
    If i <> GetRegistryValue(rkLocalMachine, strKey, "TickFillModeMinutes", 30) Then
        SetRegistryValue rkLocalMachine, strKey, "TickFillModeMinutes", i
        bUpdateCharts = True
    End If
    SetRegistryValue rkLocalMachine, strKey, "Developer", txtDeveloper.Text
    
    i = cboArchive.ListIndex
    If i >= 0 And i < cboArchive.ListCount Then
        SetIniFileProperty "ArchivePrompt", cboArchive.ItemData(i), "General", g.strIniFile
    End If
    
    i = (chkDivAdjust.Value <> 0)
    If i <> g.bDivAdjust Then
        g.bDivAdjust = i
        SetIniFileProperty "DividendAdjust", g.bDivAdjust, "General", g.strIniFile
        bUpdateCharts = True
    End If
    
    ' NextGen (salmon) -- if frame not visible then don't set so will just keep default
    If fraNextGen.Visible Then
        If g.RealTime.UseNextGen <> optNextGen.Value Then
            g.RealTime.UseNextGen = optNextGen.Value
            If g.RealTime.Active Then
                If optNextGen.Value Then
                    strText = "The NextGen setting will be enabled the next time streaming is restarted."
                Else
                    strText = "The Classic setting will be enabled the next time streaming is restarted."
                End If
                InfBox strText, "!", , "Data refresh setting"
            End If
        End If
    End If
    
    ' Redo symbol access flags if changed
    If bRedoSymbolAccess Then
        Screen.MousePointer = vbHourglass
        Me.Hide
        DoEvents
        
        InfBox "i=t ; w=NOWAIT ; Rebuilding symbol list ..."
        'InfBox "i=t ; w=NOWAIT ; Reloading symbols ..."
        g.SymbolPool.Load False
        frmSymbolGrid.RefreshGrid
        
        '???? frmQuotes.TotalRefresh True
    
        ' Now force visible charts to update
        UpdateVisibleCharts
        
        Screen.MousePointer = vbDefault
        InfBox ""
    
    ElseIf bUpdateCharts Then
        Screen.MousePointer = vbHourglass
        Me.Hide
        DoEvents
        
        g.RealTime.RefreshAllFormData True
        'UpdateVisibleCharts
        'frmQuotes.TotalRefresh False
        
        Screen.MousePointer = vbDefault
    Else
        ' reload combo box
        frmSymbolGrid.LoadCombo
        Me.Hide
    End If
    
    'flags for Elliot Wave special request features:
    '   chart mode autosize, save blank bars always, snap charts to background dots
    g.ChartGlobals.bChartModeAutoSize = (-1) * chkChartsAutoSize.Value
    g.ChartGlobals.bExtForecastBars = (-1) * chkExtendBlankBars.Value
    g.ChartGlobals.bSnapToDots = (-1) * chkSnapToDots.Value

    If g.ChartGlobals.bSnapToDots Then
        g.ChartGlobals.bChartModeAutoSize = False   'precautionary, theoretically should already be false
        '03-01-2013: sometimes chartnavigator.ini gets renamed or removed, this results in bitmap info reverting to defaults
        '   if defaults is not actually what was last generated then you get an issue with desktop flashing & all charts
        '   get shoved into upper-left of application window non-stop (Richard ran into this problem)
        LoadAppBkImage True
    End If

    m.bOK = True
        
    'Unload Me
    
    g.RealTime.Init , "frmConfig"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdTheme_Click()
On Error GoTo ErrSection:

    frmTheme.ShowMe
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdTheme_Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdTransfer_Click
'' Description: If the user clicks on the transfer button, attempt to start
''              a Data Service Transfer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTransfer_Click()
On Error GoTo ErrSection:

    ' Hide the form so that the status window is active
    Me.Hide

    ' Save any information that might have changed
    If lblCustomerID.Caption = "&User Name:" Then
        SetUserInfo Trim(txtCustomerID.Text), Trim(txtPassword.Text)
    Else
        RI_SetDataServiceID CLng(ValOfText(txtCustomerID.Text) * 1000 + ValOfText(txtDataServID.Text))
        RI_SetUserPassword txtPassword.Text
    End If

    ' Try to transfer the data service
    frmStatus.Status = eStatus_Initialized
    frmStatus.AddDetail "Transferring data service"
    TransferDataService
    frmStatus.Status = eStatus_Completed
    frmStatus.AddDetail "Finished"
    
    ' Show this form modally again
    ShowForm Me, True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.cmdTransfer.Click", eGDRaiseError_Show
    Resume ErrExit
        
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
    RaiseError "frmConfig.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, load up the controls with values from
''              the registry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the registry
    Dim nDeleteZip As Long              ' Number of days of zip files to keep
    Dim strTime As String               ' Time from the registry
    Dim strDirectory As String          ' Real time directory
    Dim lIndex As Long
    Dim nID&, i&, nDays&
    Dim aDirs As New cGdArray
    Static bAlreadyDone As Boolean
              
    ' Center the form
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
'''    frmStatus.ShowMe
        
    vstCfg.FirstTab = 0
    If Not bAlreadyDone Then
        bAlreadyDone = True
        vstCfg.CurrTab = 0
    End If
    
    'set flag telling auto download click event not to show filter download form
    m.bFormLoadDone = False
    
    ' Initialize the auto-start downloading date control
    With gdStartAuto
        .ShowDate = NoDate
        .ShowDayOfWeek = False
        .ShowTime = HourMinute
    End With
       
    ' Initialize the feeds combo box
    aDirs.GetMatchingFiles AddSlash(App.Path) & "..\RealTime\*", False, True
    For lIndex = 0 To aDirs.Size - 1
        strDirectory = aDirs(lIndex)
        If Right(strDirectory, 1) = "\" Then
            cboFeeds.AddItem Left(strDirectory, Len(strDirectory) - 1)
        End If
    Next
    If cboFeeds.ListCount > 0 And (HasModule("RTG") Or HasModule("RTE")) Then
        fraRealTime.Visible = True
        lblNoRealTime.Visible = False
        cboFeeds.ListIndex = 0
        If g.RealTime.Active Then
            vstCfg.TabPicture(eRealTimeTab) = Picture16(ToolbarIcon("kGreenLight"))
        Else
            vstCfg.TabPicture(eRealTimeTab) = Picture16(ToolbarIcon("kRedLight"))
        End If
        lblQbRealTimeMsg.Visible = True
        lblDURealTimeMsg.Visible = True
    Else
        fraRealTime.Visible = False
        lblNoRealTime.Visible = True
        CenterTheControl lblNoRealTime, fraRealTime
        vstCfg.TabPicture(eRealTimeTab) = Picture16(ToolbarIcon("kRedLight"))
        lblQbRealTimeMsg.Visible = False
        lblDURealTimeMsg.Visible = False
        fraQuoteBoard.Height = fraQuoteBoard.Height - 180
    End If
    vstCfg.TabPicture(eSnapQuoteTab) = Picture16(ToolbarIcon("ID_Download"))
    
    ' Salmon or not
    If HasModule("RTG") And (g.lLCD > 0) Then
        If g.RealTime.UseNextGen Then
            optNextGen.Value = True
        Else
            optClassic.Value = True
        End If
        fraNextGen.Visible = True
    Else
        fraNextGen.Visible = False
    End If
    
    ' Authorizations
    If g.lLCD > 0 Then
        fraAuth.Visible = True
    Else
        fraAuth.Visible = False
    End If
    
    ' Load Account settings
    lblCustomerID.Caption = "C&ustomer ID:"
    lblDataServiceID.Visible = True
    txtDataServID.Visible = True
    txtMachineID.Locked = True
    txtMachineID.BackColor = vbButtonFace
    txtPassword.PasswordChar = "*"
    txtMachineID.Text = UCase(RI_GetMachineID)
    nID = RI_GetDataServiceID
    If nID > 0 Then
        txtCustomerID.Text = Format(nID \ 1000, "0")
        txtDataServID.Text = Format(nID Mod 1000, "000")
        txtPassword.Text = RI_GetUserPassword
        'cmdTransfer.Visible = True
    Else
        txtDataServID.Text = "001"
        txtCustomerID.Text = ""
        txtPassword.Text = ""
        cmdTransfer.Visible = False
    End If
    
    ' store original password (so can confirm if changed)
    txtPassword.Tag = UCase(Trim(txtPassword))
    
    ' Get General settings from registry
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    chkLocalTimeZone = GetRegistryValue(rkLocalMachine, strKey, "DisplayLocalTimeZone", vbChecked)
    txtDaysToSave = CStr(GetRegistryValue(rkLocalMachine, strKey, "DeleteZipFiles", 31#))
    txtSymbolExpire = CStr(GetRegistryValue(rkLocalMachine, strKey, "SymbolExpire", 31#))
    chkSymbolExpire = GetRegistryValue(rkLocalMachine, strKey, "chkSymbolExpire", vbChecked)
    'txtScanDays.Text = CStr(GetRegistryValue(rkLocalMachine, strKey, "DaysInScan", 365#))
    chkScans.Value = Abs(ScansEnabled)
    'clrAltRow.CustomColor = &HC8F0FF
    'clrAltRow.Color = GetRegistryValue(rkLocalMachine, strKey, "AltGridRowColor", &HC8F0FF)
    slScrubLevel = GetRegistryValue(rkLocalMachine, strKey, "ScrubLevel", 5&)
''slScrubLevel.Value = 0
    
    ' Get Auto Downloading information from the registry
    optLocalTime = GetRegistryValue(rkLocalMachine, strKey, "LocalTime", False)
    If g.RealTime.Active Then
        ' auto-downloading is ON if realtime is active
        chkAutoDownload.Enabled = False
        chkAutoDownload.Value = 1
    Else
        chkAutoDownload.Enabled = True
        chkAutoDownload = GetRegistryValue(rkLocalMachine, strKey, "AutoUpdate", vbChecked)
    End If
    gdStartAuto = GetRegistryValue(rkLocalMachine, strKey, "AutoStart", TimeSerial(18, 30, 0))
    txtTryInterval = CStr(GetRegistryValue(rkLocalMachine, strKey, "TryInterval", 5))
    txtTryHours = CStr(GetRegistryValue(rkLocalMachine, strKey, "TryHours", 2))
    
    ' Get Snap Quote information from registry
    'If g.RealTime.Active Then
    If 0 Then
        ' auto QB-refresh is OFF if realtime is active
        chkAutoQuotes.Enabled = False
        chkAutoQuotes.Value = 0
    Else
        chkAutoQuotes.Enabled = True
        chkAutoQuotes = GetRegistryValue(rkLocalMachine, strKey, "AutoQuotes", vbUnchecked)
    End If
    txtQuoteInterval = CStr(GetRegistryValue(rkLocalMachine, strKey, "QuoteInterval", 10#))
    gdStartTime = GetRegistryValue(rkLocalMachine, strKey, "QuoteStart", Date + TimeSerial(8, 20, 0))
    gdEndTime = GetRegistryValue(rkLocalMachine, strKey, "QuoteEnd", Date + TimeSerial(16, 30, 0))
    
    ' Get Real Time information from the registry
    chkActivate = GetRegistryValue(rkLocalMachine, strKey, "ActivateRT", vbUnchecked)
    strDirectory = GetRegistryValue(rkLocalMachine, strKey, "ActiveFeedRT", cboFeeds.Text)
    For lIndex = 0 To cboFeeds.ListCount - 1
        If UCase(cboFeeds.List(lIndex)) = UCase(strDirectory) Then
            cboFeeds.ListIndex = lIndex
            Exit For
        End If
    Next lIndex
    
    ' Get System Navigator information from the registry
    chkTickDataFill = GetRegistryValue(rkLocalMachine, strKey, "TickDataFill", vbChecked)
    Select Case GetRegistryValue(rkLocalMachine, strKey, "TickFillMode", 0)
    Case 1
        optFullTickFill = True
    Case 2
        optMinTickFill = True
    Case Else
        optAutoTickFill = True
    End Select
    txtAutoTickFill = GetRegistryValue(rkLocalMachine, strKey, "TickFillModeMinutes", 30)
    txtDeveloper = GetRegistryValue(rkLocalMachine, strKey, "Developer", "")
    
    ' Enable the appropriate controls
    EnableControls
       
    m.bAutoDownloadDirty = False
    m.bAutoRefreshDirty = False
    
    ' How often to recalc indicators when realtime
    If g.nRecalcIndRT > 0 And g.nRecalcIndRT < 1000 Then
        optRecalcSeconds.Value = True
        txtRecalcSeconds.Text = Str(g.nRecalcIndRT)
    ElseIf g.nRecalcIndRT < 0 Then
        optRecalcBar.Value = True
        txtRecalcSeconds.Text = "10"
    Else
        optRecalcTick.Value = True
        txtRecalcSeconds.Text = "10"
    End If
    
    ' If the quote board is not showing, enable the quote board button
    '''cmdQuoteBoard.Enabled = Not FormIsLoaded("frmQuotes")
    
    ' Archive prompt
    With cboArchive
        .AddItem "Every time"
        .ItemData(.ListCount - 1) = 0
        .AddItem "Once a day"
        .ItemData(.ListCount - 1) = 1
        .AddItem "Once a week"
        .ItemData(.ListCount - 1) = 7
        .AddItem "Once a month"
        .ItemData(.ListCount - 1) = 31
        .AddItem "Never"
        .ItemData(.ListCount - 1) = -1
        nDays = GetIniFileProperty("ArchivePrompt", 0, "General", g.strIniFile)
        .ListIndex = 0
        For i = 0 To .ListCount - 1
            If .ItemData(i) = nDays Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
    
    'flags for Elliot Wave special request features
    '   keep charts as fixed sized pixels, save blank bars always, snap charts to dots
    chkChartsAutoSize.Value = Abs(g.ChartGlobals.bChartModeAutoSize)
    chkExtendBlankBars.Value = Abs(g.ChartGlobals.bExtForecastBars)
    chkSnapToDots.Value = Abs(g.ChartGlobals.bSnapToDots)
    
    chkDivAdjust.Value = Abs(g.bDivAdjust)
    chkDivAdjust.Visible = AllowDivAndMF
    
    'set flag telling auto download click event it's okay to show filter download form
    m.bFormLoadDone = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.Form.Load", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Me.Hide
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Resize()
On Error Resume Next

End Sub

Private Sub gdEndTime_Changed()
On Error GoTo ErrSection:

    m.bAutoRefreshDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.gdEndTime.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub gdStartAuto_Changed()
On Error GoTo ErrSection:
    
    m.bAutoDownloadDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.gdStartAuto.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub gdStartTime_Changed()
On Error GoTo ErrSection:

    m.bAutoRefreshDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.gdStartTime.Changed", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub lblDisplayCodes_Click()
On Error GoTo ErrSection:

    optEnablements.Value = False
    optPurchased.Value = False
    optConnectionInfo.Value = False
    txtCodes.Text = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.lblDisplayCodes_Click"
    Resume ErrExit
End Sub

Private Sub optAutoTickFill_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optAutoTickFill_Click"
    Resume ErrExit
End Sub

Private Sub optClassic_Click()
On Error GoTo ErrSection:

    ResetNextGen
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optClassic_Click"
End Sub

Private Sub optConnectionInfo_Click()
On Error GoTo ErrSection:

    txtCodes.Text = RI_GetMessage & vbCrLf & vbCrLf & RI_GetComment

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optConnectionInfo_Click"
    Resume ErrExit
End Sub

Private Sub optEnablements_Click()
On Error GoTo ErrSection:
    
    Dim strText$
    
    SetColumnWidthForControl txtCodes, 55
    
    strText = Trim(g.strAuthorizationString)
    strText = Replace(strText, ",,", ",")
    Do While Left(strText, 1) = ","
        strText = Trim(Mid(strText, 2))
    Loop
    Do While Right(strText, 1) = ","
        strText = Trim(Left(strText, Len(strText) - 1))
    Loop
    strText = Trim(Replace(strText, ",", " " & vbTab))
    txtCodes.Text = strText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optEnablements_Click"
    Resume ErrExit
End Sub

Private Sub optFullTickFill_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optFullTickFill_Click"
    Resume ErrExit
End Sub

Private Sub optLocalTime_Click()
On Error GoTo ErrSection:

    optLocalTime.Font.Bold = True
    optNewYorkTime.Font.Bold = False

    If Not Me.Visible Then Exit Sub
    
    gdStartTime = ConvertTimeZone(gdStartTime, "NY", "")
    gdEndTime = ConvertTimeZone(gdEndTime, "NY", "")
    gdStartAuto = ConvertTimeZone(gdStartAuto, "NY", "")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optLocalTime_Click"
    Resume ErrExit
        
End Sub

Private Sub optMinTickFill_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optMinTickFill_Click"
    Resume ErrExit
End Sub

Private Sub optNewYorkTime_Click()
On Error GoTo ErrSection:

    optLocalTime.Font.Bold = False
    optNewYorkTime.Font.Bold = True

    If Not Me.Visible Then Exit Sub
    
    gdStartTime = ConvertTimeZone(gdStartTime)
    gdEndTime = ConvertTimeZone(gdEndTime)
    gdStartAuto = ConvertTimeZone(gdStartAuto)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optNewYorkTime.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub optNextGen_Click()
On Error GoTo ErrSection:

    ResetNextGen
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optNextGen_Click"
End Sub

Private Sub optPurchased_Click()
On Error GoTo ErrSection:

    SetColumnWidthForControl txtCodes, 34
    txtCodes.Text = Trim(Replace(DM_GetPurchased(True), ",", " " & vbTab))
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optPurchased_Click"
End Sub

Private Sub optRecalcBar_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optRecalcBar_Click"
End Sub

Private Sub optRecalcSeconds_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optRecalcSeconds_Click"
End Sub

Private Sub optRecalcTick_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.optRecalcTick_Click"
End Sub

Private Sub slScrubLevel_Change()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.slScrubLevel.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDataServID_LostFocus()
On Error GoTo ErrSection:

    Dim d#
    
    d = ValOfText(txtDataServID)
    If d < 1 Or d > 999 Then
        d = 1
    Else
        d = Int(d)
    End If
    txtDataServID = Format(d, "000")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.txtDataServID.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.txtPassword.GotFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtQuoteInterval_LostFocus
'' Description: If the user leaves the Quote Interval box with a value less
''              than 10, change it back to 10
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtQuoteInterval_LostFocus()
On Error GoTo ErrSection:

    If ValOfText(txtQuoteInterval.Text) < 10# Then
        txtQuoteInterval.Text = "10"
    End If
    m.bAutoRefreshDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.txtQuoteInterval.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtSymbolExpire_LostFocus()
On Error GoTo ErrSection:

    Dim iDays&
    
    iDays = ValOfText(txtSymbolExpire)
    If iDays <= 0 Then
        txtSymbolExpire = CStr(GetRegistryValue(rkLocalMachine, _
            "Software\Genesis Financial Data Services\Navigator Suite\General", _
            "SymbolExpire", 31#))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.txtSymbolExpire.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtTryHours_LostFocus()
On Error GoTo ErrSection:

    If ValOfText(txtTryHours.Text) < 1# Then
        txtTryHours.Text = "1"
    End If
    m.bAutoDownloadDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.txtTryHours.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTryInterval_LostFocus
'' Description: Don't let the user enter in an interval less than 5 minutes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTryInterval_LostFocus()
On Error GoTo ErrSection:

    If ValOfText(txtTryInterval.Text) < 5# Then
        txtTryInterval.Text = "5"
    End If
    m.bAutoDownloadDirty = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.txtTryInterval.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vstCfg_Switch
'' Description: When the user switches tabs, save what tab they are now on
'' Inputs:      Tab coming from, Tab going to, Whether or not to cancel change
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vstCfg_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    m.eCurrTab = NewTab
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.vstCfg.Switch", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls based on the value of other controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim i&

    ' Auto Downloading stuff
    Enable gdStartAuto, chkAutoDownload
    Enable txtTryInterval, chkAutoDownload
    Enable txtTryHours, chkAutoDownload
    
    ' Auto Quote Board Refresh stuff
    Enable txtQuoteInterval, chkAutoQuotes
    Enable gdStartTime, chkAutoQuotes
    Enable gdEndTime, chkAutoQuotes
    
    ' Real Time stuff
    'Enable cboFeeds, chkActivate
    Enable txtRecalcSeconds, optRecalcSeconds
    
    ' Precision Tick
    Enable optFullTickFill, chkTickDataFill
    Enable optMinTickFill, chkTickDataFill
    Enable optAutoTickFill, chkTickDataFill
    Enable txtAutoTickFill, (chkTickDataFill <> 0 And optAutoTickFill <> 0)
    
    ' Misc stuff
    Enable txtSymbolExpire, chkSymbolExpire
    'Enable txtScanDays, chkScans
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.EnableControls", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Shows the form with the given tab (defaults to last tab shown)
'' Inputs:      Tab to show
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional eConfigTab As eConfigTabs = eNoChange) As Boolean
On Error GoTo ErrSection:

    Dim bHasTick As Boolean

    bHasTick = HasModule("FT") Or HasModule("ST") Or HasModule("IT")
    vstCfg.TabVisible(Tabs(eTickDataTab)) = bHasTick
    If ExtremeCharts = 1 And Not bHasTick Then
        vstCfg.TabVisible(Tabs(eRealTimeTab)) = False
    Else
        vstCfg.TabVisible(Tabs(eRealTimeTab)) = True
    End If
    
    ' Only show the scrub level frame if the user has tick data (DAJ: 03/24/2003)...
    'fraScrubLevel.Visible = (HasModule("FT") Or HasModule("IT") Or HasModule("ST")) And FileExist("C:\Common\Ask32.EXE")
        
    vstCfg.FirstTab = 0
    If eConfigTab = eNoChange Then
        vstCfg.CurrTab = m.eCurrTab
    Else
        vstCfg.CurrTab = eConfigTab
    End If
    
    EnableControls

    cmdTheme.Enabled = IsAtLeastVista

    ShowForm Me, True
        
    ShowMe = m.bOK
    Unload Me

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConfig.ShowMe", eGDRaiseError_Raise
        
End Function

Private Sub ResetNextGen()
On Error GoTo ErrSection:

    Dim bShowMsg As Boolean
        
    If Me.Visible And g.RealTime.Active Then
        If g.RealTime.UseNextGen Then
            If Not optNextGen.Value Then
                optNextGen.Value = True
                bShowMsg = True
            End If
        ElseIf Not optClassic.Value Then
            optClassic.Value = True
            bShowMsg = True
        End If
        If bShowMsg Then
            InfBox "Streaming must first be turned off before this setting can be changed.", "e", , "Error"
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmConfig.ResetNextGen"
End Sub

