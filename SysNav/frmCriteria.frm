VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCriteria 
   Caption         =   "Criteria Editor"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   Icon            =   "frmCriteria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8835
   Begin HexUniControls.ctlUniFrameWL fraAdvanced 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   5745
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmCriteria.frx":038A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCriteria.frx":03CC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCriteria.frx":03EC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraNumDays 
         Height          =   615
         Left            =   120
         TabIndex        =   0
         Top             =   720
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
         Caption         =   "frmCriteria.frx":0408
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmCriteria.frx":0428
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0448
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtNumDays 
            Height          =   315
            Left            =   4200
            TabIndex        =   7
            Top             =   60
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmCriteria.frx":0464
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
            Tip             =   "frmCriteria.frx":0486
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":04A6
         End
         Begin HexUniControls.ctlUniTextBoxXP txtOverride 
            Height          =   315
            Left            =   4260
            TabIndex        =   3
            Top             =   180
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmCriteria.frx":04C2
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
            Tip             =   "frmCriteria.frx":04E4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":0504
         End
         Begin HexUniControls.ctlUniRadioXP optOverride 
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            Top             =   300
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
            Caption         =   "frmCriteria.frx":0520
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmCriteria.frx":055E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":057E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optAutoDetect 
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   300
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
            Caption         =   "frmCriteria.frx":059A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmCriteria.frx":05EC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":060C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblNumBars1 
            Height          =   195
            Left            =   120
            Top             =   60
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
            Caption         =   "frmCriteria.frx":0628
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmCriteria.frx":06BC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":06DC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraBasedOn 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
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
         Caption         =   "frmCriteria.frx":06F8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmCriteria.frx":0718
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0738
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optDaily 
            Height          =   255
            Left            =   1080
            TabIndex        =   4
            Top             =   120
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
            Caption         =   "frmCriteria.frx":0754
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmCriteria.frx":078A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":07AA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optWeekly 
            Height          =   255
            Left            =   2160
            TabIndex        =   5
            Top             =   120
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
            Caption         =   "frmCriteria.frx":07C6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmCriteria.frx":083A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":085A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   120
            Top             =   120
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
            Caption         =   "frmCriteria.frx":0876
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmCriteria.frx":08A8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":08C8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSymbolGroups 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   4875
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmCriteria.frx":08E4
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmCriteria.frx":0904
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0924
         RightToLeft     =   0   'False
         Begin MSComctlLib.ImageCombo cboSymbolGroups 
            Height          =   330
            Left            =   2160
            TabIndex        =   9
            Top             =   0
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Text            =   "ImageCombo1"
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   120
            Top             =   60
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
            Caption         =   "frmCriteria.frx":0940
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmCriteria.frx":099C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmCriteria.frx":09BC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraDisplay 
      Height          =   1935
      Left            =   6000
      TabIndex        =   10
      Top             =   1800
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
      Caption         =   "frmCriteria.frx":09D8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCriteria.frx":0A12
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCriteria.frx":0A32
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtDecimalPlaces 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   750
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmCriteria.frx":0A4E
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
         Tip             =   "frmCriteria.frx":0A6E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0A8E
      End
      Begin HexUniControls.ctlUniRadioXP optCustomRound 
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   780
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
         Caption         =   "frmCriteria.frx":0AAA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCriteria.frx":0ADA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0AFA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTradingUnits 
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1200
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
         Caption         =   "frmCriteria.frx":0B16
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCriteria.frx":0B50
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0B70
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAutoRound 
         Height          =   255
         Left            =   180
         TabIndex        =   11
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
         Caption         =   "frmCriteria.frx":0B8C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmCriteria.frx":0BC0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0BE0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDecimals 
         Height          =   255
         Left            =   1680
         Top             =   780
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
         Caption         =   "frmCriteria.frx":0BFC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCriteria.frx":0C2C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCriteria.frx":0C4C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   7320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   9
      DisplayContextMenu=   0   'False
      Tools           =   "frmCriteria.frx":0C68
      ToolBars        =   "frmCriteria.frx":0F62
   End
   Begin NavTradeSenseV3.Editor Editor1 
      Height          =   1275
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   2249
   End
   Begin HexUniControls.ctlUniCheckXP chkAdvanced 
      Height          =   220
      Left            =   5940
      TabIndex        =   16
      Top             =   120
      Width           =   1755
      _ExtentX        =   3096
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
      Caption         =   "frmCriteria.frx":1173
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmCriteria.frx":11B7
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmCriteria.frx":11D7
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblEditor 
      Height          =   225
      Left            =   120
      Top             =   120
      Width           =   4995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmCriteria.frx":11F3
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCriteria.frx":1297
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCriteria.frx":12B7
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCriteria.frm
'' Description: Form for the management of Criteria
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/12/2010   DAJ         Utilize new AutoDetect object for auto detection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Const DONTLOADACTIONS = ""

Private Type mPrivate
    strCodedText As String
    strFormattedText As String
    strDescription As String
    
    bIsBoolean As Boolean
    bSkipAutoIf As Boolean
    ListLoading As cListLoading
    Function As cFunction
    
    Criteria As cCriteria
    strName As String
    bDirty As Boolean
    
    bOK As Boolean
    bModal As Boolean
End Type
Private m As mPrivate

Public Property Get ID() As String
    ID = m.Criteria.ID
End Property

Private Sub Save(ByVal strButton As String)
On Error GoTo ErrSection:
    
    Dim lNumDays As Long                ' Number of days necessary
    Dim bIsDirty As Boolean             ' Did the user change something?
    Dim strNewName As String            ' Return from the AskBox
    Dim strText As String               ' Prompt to send to the AskBox
    Dim bSaveAs As Boolean              ' Are we doing a Save As?
    Dim bDontHide As Boolean            ' Don't Hide the form
    Dim strReturn As String             ' Return from an InfBox
    
    ' Make sure that the user has entered an expression
    If Len(Editor1.Text) = 0 Then
        MoveFocus Editor1
        Err.Raise vbObjectError + 1000, , "Please enter an expression"
    End If
    
    ' Make sure that a Silver user cannot create more Criteria by using
    ' the Save As button...
    If strButton = "ID_SaveAs" Then
        If gdNumMatchingFiles(AddSlash(App.Path) & "Custom\Cus0*.SCN") >= 1 Then
            If Not HasGold(True, "Creating more custom Criteria") Then
                Exit Sub
            End If
        End If
    End If
    
    ' Only reverify if necessary (if a rule has changed)
    If tbToolbar.Tools("ID_Verify").Enabled Then
        If Not Verify Then Exit Sub 'bCancel = True
    End If
    
    ' Handle Rename/Save As
    m.strName = Trim(m.strName)
    If Len(m.strName) = 0 Then
        Select Case m.Criteria.UsageType
            Case eCriteria_FilterCriteria
                strText = "Save the current Criteria as..."
            Case eCriteria_QuoteBoardField
                strText = "Save the current Quote Board Field as..."
        End Select
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    ElseIf strButton = "ID_SaveAs" Then
        Select Case m.Criteria.UsageType
            Case eCriteria_FilterCriteria
                strText = "Save a copy of the current Criteria as..."
            Case eCriteria_QuoteBoardField
                strText = "Save a copy of the current Quote Board Field as..."
        End Select
        strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
        If Trim(UCase(strNewName)) <> UCase(m.strName) Then
            bSaveAs = True
        End If
    ElseIf strButton = "ID_Rename" Then
        Select Case m.Criteria.UsageType
            Case eCriteria_FilterCriteria
                strText = "Rename the current Criteria as..."
            Case eCriteria_QuoteBoardField
                strText = "Rename the current Quote Board Field as..."
        End Select
        strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & m.strName & " ; " & strText)
    Else
        strNewName = m.strName
    End If
    
    
    If Len(strNewName) = 0 Then
        Exit Sub 'Err.Raise vbObjectError + 1000, , "You must enter in a name for the filter"
    End If
    m.strName = Trim(strNewName)
    With m.Criteria
        If .UsageType = eCriteria_FilterCriteria Then
            SetEditorCaption Me, "Criteria", m.strName
        Else
            SetEditorCaption Me, "Quote Board Field", m.strName
        End If
    End With
    
    ' Calculate the number of days if not done already
    lNumDays = AutoDetect(m.strCodedText)
    txtNumDays.Text = Str(lNumDays)
    
    If (ValOfText(txtOverride.Text) < lNumDays) And optOverride.Value Then
        InfBox "Trade Navigator has determined that your criteria needs at least " & _
            Trim(CStr(lNumDays)) & " bars to run properly.  " & _
            "The value has been set accordingly.", _
            "i", , "Criteria"
        optAutoDetect = True
        txtOverride.Text = lNumDays
        bDontHide = True
    End If
    
    If (lNumDays = -1&) And (optAutoDetect Or (ValOfText(txtOverride.Text) <= 0)) Then
        InfBox "Trade Navigator could not automatically determine how many bars are needed to calculate " & _
                " the criteria.  Please specify an override for the number of necessary bars.", _
                "!", , "Criteria Error"
        optOverride = True
        MoveFocus txtOverride
        Exit Sub
    End If
    
    If bSaveAs Then
        Set m.Criteria = m.Criteria.MakeCopy
    Else
        'get the category ID & name from symbol pool
        Dim idx As Long
        
        idx = g.SymbolPool.Criterias.Index(m.Criteria.ID)
        If idx > 0 Then
            If Not g.SymbolPool.Criterias.Item(idx) Is Nothing Then
                m.Criteria.CategoryID = g.SymbolPool.Criterias.Item(idx).CategoryID
                m.Criteria.CategoryName = g.SymbolPool.Criterias.Item(idx).CategoryName
            End If
        End If
    End If
    With m.Criteria
        ' see if is now "dirty" (needs to be recalculated)
        bIsDirty = False
        If bSaveAs = True Then
            .ID = ""
            bIsDirty = True
        End If
        If .IsWeekly <> optWeekly.Value Then bIsDirty = True
        If .NumDaysCalc <> CLng(Trim(txtNumDays.Text)) Then bIsDirty = True
        If UCase(Trim(m.strCodedText)) <> UCase(Trim(.CodedText)) Then bIsDirty = True
        If cboSymbolGroups.SelectedItem.Key <> .GroupID Then bIsDirty = True
        If .NumDaysOverride <> CLng(ValOfText(Trim(txtOverride.Text))) Then bIsDirty = True
        
        .Name = Trim(m.strName)
        .Desc = Trim(m.strDescription)
        .NumDaysCalc = CLng(ValOfText(Trim(txtNumDays.Text)))
        If optOverride Then
            .NumDaysOverride = CLng(ValOfText(Trim(txtOverride.Text)))
        Else
            .NumDaysOverride = -1&
        End If
        .EnglishText = Trim(Editor1.Text)
        .CodedText = Trim(m.strCodedText)
        .FormattedText = Trim(m.strFormattedText)
        .IsBoolean = m.bIsBoolean
        .IsWeekly = optWeekly.Value
        
        ' Be careful not to mark an already dirty criteria to clean...
        If bIsDirty Then .IsDirty = True
        
        .GroupID = cboSymbolGroups.SelectedItem.Key
        
        Select Case True
            Case optAutoRound
                .PriceDisplay = eCriteria_AutoRound
            Case optCustomRound
                .PriceDisplay = eCriteria_RoundToDecimal
            Case optTradingUnits
                .PriceDisplay = eCriteria_TradingUnits
        End Select
        
        .DecimalPlaces = Val(txtDecimalPlaces.Text)
        
        .Save bIsDirty
    End With
    
    m.bOK = True
    EnableToolbar False
    
    ' If the criteria was dirty, then ask the user if they wish to recalculate the
    ' dirty criteria (DAJ: 03/26/2003) ... but only if ScansEnabled (TLB: 7/26/2011)
    If bIsDirty And m.Criteria.IsActive And ScansEnabled Then
        strReturn = InfBox("Would you like to recalculate this criteria now?", "?", "+Yes|-No", "Criteria")
        If strReturn = "Y" Then
            If g.SymbolPool.RecalcDirtyCriteria Then
                frmStatus.AddDetail "Finished"
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Save", eGDRaiseError_Raise

End Sub

Private Sub cboSymbolGroups_Click()
On Error GoTo ErrSection:
    
    EnableToolbar True
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCriteria.cboSymbolGroups.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkAdvanced_Click()
On Error GoTo ErrSection:

    fraAdvanced.Visible = (chkAdvanced.Value = vbChecked)
    Form_Resize
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.chkAdvanced.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Editor1_Change()
On Error GoTo ErrSection:
    
    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = (Len(Trim(Editor1.Text)) > 0)
    If Len(Editor1.Text) = 0 Then SendKeys " "

    ' Don't allow the user to use an assignment in the expression for now...
    If InStr(Editor1.Text, ":=") <> 0 Then
        InfBox "You cannot have an assignment operator in this expression.", "!", , "Expression Error"
        Editor1.Text = Replace(Editor1.Text, ":=", "")
        If Len(Editor1.Text) > 0 Then
            Editor1.SelStart = Len(Editor1.Text)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Editor1.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Editor1_EditFunction(FunctionID As Long, FunctionName As String, Found As Boolean)
On Error GoTo ErrSection:
    
    Err.Raise vbObjectError + 1000, , "Sub functions cannot be edited or added here"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Editor1.EditFunction", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Editor1_GotFocus()
On Error GoTo ErrSection:
    
    Set g.ActiveEditor = Editor1
        
    If m.ListLoading Is Nothing Then
        'Load internally generated TradeSense lists (Symbols, etc.)
        Set m.ListLoading = New cListLoading
        m.ListLoading.Load
    End If
    
    With Editor1
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .ShowNewFunction = Not m.bModal
        .Usage = 8             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With
    
    If Len(Trim(Editor1.Text)) = 0 And Not m.bSkipAutoIf Then
        Editor1.Text = ""
        SendKeys " "
    End If
    
    m.bSkipAutoIf = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Editor1.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Editor1_LostFocus()
On Error GoTo ErrSection:
    
    Set g.ActiveEditor = Nothing
    Editor1.RemoveTradeSense
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Editor1.LostFocus", eGDRaiseError_Show
    Resume ErrExit

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
    RaiseError "frmCriteria.Editor1.NewFunction", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:

    If GetActiveWindow = Me.hWnd Then MoveFocus Editor1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
    
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
    RaiseError "frmCriteria.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form from the ini file

    chkAdvanced.Value = GetIniFileProperty("Advanced", vbUnchecked, "Criteria", g.strIniFile)
    chkAdvanced_Click
    strPlacement = GetIniFileProperty("frmCriteria", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_Criteria"), , True)
    
    With tbToolbar
        .Tools("ID_Verify").Picture = Picture16(ToolbarIcon("kVerify"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Description").Picture = Picture16(ToolbarIcon("ID_News"))
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_CondBuilder").Picture = Picture16(ToolbarIcon("ID_ConditionBuilder"))
    End With
    
    Set m.Function = New cFunction
    With m.Function
        .FunctionID = 0
        .Load
    End With

    If m.ListLoading Is Nothing Then
        'Load internally generated TradeSense lists (Symbols, etc.)
        Set m.ListLoading = New cListLoading
        m.ListLoading.Load
    End If
    
    With Editor1
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .Usage = 8             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With
    
    cboSymbolGroups.ImageList = frmMain.img16
    cboSymbolGroups.Locked = True
    LoadCombo
    
    ' For now, hide the Advanced check box and turn it on...
    chkAdvanced.Value = vbChecked
    chkAdvanced.Visible = False
    
    txtNumDays.Locked = True
    txtNumDays.Enabled = False
    txtOverride.Move txtNumDays.Left, txtNumDays.Top
   
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCriteria.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Dim w&, h&
    
    If WindowState = vbMinimized Then
        If TypeOf ActiveControl Is Editor Then
            Set g.ActiveEditor = Nothing
            ActiveControl.RemoveTradeSense
        End If
    End If

    w = (fraAdvanced.Left * 3) + fraAdvanced.Width + fraDisplay.Width
    h = fraAdvanced.Height + 1200
    If LimitFormSize(Me, w, h) Then Exit Sub
    
    If chkAdvanced.Value = vbChecked Then
        With fraAdvanced
            .Move .Left, ScaleHeight - .Height - lblEditor.Top
        End With
        With fraDisplay
            .Move .Left, ScaleHeight - .Height - lblEditor.Top
        End With
        h = fraAdvanced.Top - Editor1.Top - lblEditor.Top
    Else
        h = ScaleHeight - Editor1.Top - lblEditor.Top
    End If
    
    With Editor1
        .Move .Left, .Top, ScaleWidth - (.Left * 2), h
    End With
    
    ''chkAdvanced.Left = txtDesc.Left + txtDesc.Width - chkAdvanced.Width
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Set m.Function = Nothing
    Set m.ListLoading = Nothing
    
    SetIniFileProperty "Advanced", chkAdvanced.Value, "Criteria", g.strIniFile
    SetIniFileProperty "frmCriteria", GetFormPlacement(Me), "Placement", g.strIniFile
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
       
    If UnloadMode <> vbFormCode Then
        If AskToSave Then
            Cancel = True
        ElseIf m.bModal Then
            Cancel = True
            Me.Hide
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCriteria.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optAutoDetect_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optAutoDetect_Click"
End Sub

Private Sub optAutoRound_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optAutoRound.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optCustomRound_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optCustomRound.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optDaily_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optDaily.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optOverride_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    MoveFocus txtOverride

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optOverride_Click"
End Sub

Private Sub optTradingUnits_Click()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optTradingUnits.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optWeekly_Click()
On Error GoTo ErrSection:

    EnableToolbar True
    tbToolbar.Tools("ID_Verify").Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.optWeekly.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim iCancel As Integer
    Dim strID$, frm As Form

    ToggleFocus Me, txtNumDays

    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            Save Tool.ID
        
        Case "ID_Verify"
            Verify
            
        Case "ID_Print"
            PrintMe
            
        Case "ID_Description"
            m.strDescription = frmNotes.ShowMe(m.strDescription, "Description")
            EnableToolbar True
        
        Case "ID_Toolbox"
            If Not AskToSave Then
                strID = m.Criteria.ID
                Unload Me
                frmToolbox.ShowMe eTab_Criteria, strID
            End If
        
        Case "ID_Close"
            If Not AskToSave Then
                If m.bModal Then
                    Me.Hide
                Else
                    Unload Me
                End If
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
    RaiseError "frmCriteria.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Function Verify() As Boolean
On Error GoTo ErrSection:
   
    Dim i&, strChk$
    Dim lNumDays As Long
    Dim strNotKnown As String
    Dim bExtraInputs As Boolean
    Dim strMsg          As String
    Dim wrkText         As String
    Dim Expr            As cExpression
    Dim Inputs          As cInputs
    Dim lErrNum As Long
    Dim strErrSource As String
    Dim strErrDesc As String
    
    FixPeriodInMarkets
 
    'Shut things off, get ready for verifying rule
    Screen.MousePointer = vbHourglass
    LockWindowUpdate Me.hWnd
    m.strCodedText = ""
    
    'Verify...
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule Editor1.Text
    End With
        
    'Convert to rich text
    With Editor1
        .TurnOffEditing
        wrkText = Expr.EditText
        .TextRTF = m.Function.GetRTF(wrkText)
        m.strFormattedText = Expr.EditText
        .ExprIsFormatted = True
        .SelStart = 999999
        .TurnOnEditing
    End With
        
    ' see if unwanted inputs exist
    bExtraInputs = False
    strNotKnown = ""
    If Not Expr.Inputs Is Nothing Then
        'ShowParmLine TradeSense
        'mFunction.TradeSenseUsage = TradeSense.Tag
        Set Inputs = Expr.Inputs
        For i = 1 To Expr.Inputs.Count
            strChk = UCase(Inputs.Item(i).ParmName)
            'If strChk <> "WEEKLY" And strChk <> "GC" And strChk <> "MARKET1" Then
            If ValidMarket(strChk) = False Then
                strNotKnown = strNotKnown & "|" & Inputs.Item(i).ParmName
                bExtraInputs = True
            End If
        Next
    End If
    If bExtraInputs Then
        EnableToolbar False
        tbToolbar.Tools("ID_Verify").Enabled = True
        Err.Raise vbObjectError + 1000, , "Unrecognized items in expression:|" & strNotKnown & "|"
    Else
        ' successful
        m.strCodedText = Expr.CodedText
        i = Expr.FunctionReturnType
        If i = 3 Or i = 6 Then
            m.bIsBoolean = True
        Else
            m.bIsBoolean = False
        End If
        
        LockWindowUpdate 0
        
        'If IsIDE Then
            FileFromString App.Path & "\Chk\Criteria.chk", m.strCodedText
        'End If
        
        If EngineVerify Then
            EnableToolbar True
            tbToolbar.Tools("ID_Verify").Enabled = False
            lNumDays = AutoDetect(m.strCodedText)
            txtNumDays.Text = Str(lNumDays)
            
            If lNumDays > 0 Then
                txtNumDays.Font.Bold = False
                txtNumDays.ForeColor = optDaily.ForeColor
                lblNumBars1.ForeColor = optDaily.ForeColor
                'lblNumBars2.ForeColor = optDaily.ForeColor
                'lblNumBars2.Caption = "(will be auto-detected when expression is verified)"
            Else
                txtNumDays.Font.Bold = True
                txtNumDays.ForeColor = RGB(0, 0, 128)
                lblNumBars1.ForeColor = RGB(0, 0, 128)
                'lblNumBars2.ForeColor = RGB(0, 0, 128)
                'lblNumBars2.Caption = "(please make sure the value you specify is adequate)"
                chkAdvanced = 1
                strMsg = "The expression is valid.|However, the number of past bars required to load cannot be auto-detected.  Please make sure the value you specify is adequate."
                InfBox strMsg, "i", , "Important"
                optOverride = True
                MoveFocus txtOverride
            End If
            Verify = True
        Else
            EnableToolbar False
            tbToolbar.Tools("ID_Verify").Enabled = True
        End If
    End If
    
    
ErrExit:
    Screen.MousePointer = vbDefault
    Set Expr = Nothing
    LockWindowUpdate 0
    Exit Function

ErrSection:
    Screen.MousePointer = vbDefault
    LockWindowUpdate 0
    
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        lErrNum = Err.Number
        strErrSource = Err.Source
        strErrDesc = Err.Description
        
        'Highlight error in advanced editor...
        If Expr.EditText <> "" Then
            With Editor1
                .TurnOffEditing
                wrkText = Expr.EditText
                .ExprIsFormatted = False
                .TextRTF = m.Function.GetRTF(wrkText)
                .ExprIsFormatted = True
                .TurnOnEditing
            End With
        End If
        
        Set Expr = Nothing
        Err.Raise lErrNum, strErrSource, strErrDesc
        'InfBox Err.Description, "e", , "Invalid Expression"
    Else
        Set Expr = Nothing
        RaiseError "frmCriteria.Verify", eGDRaiseError_Raise
    End If
    
End Function

Private Function EngineVerify() As Boolean
On Error GoTo ErrSection:

    Dim i&, rc&, strCodedText$
    Dim astrParms As New cGdArray
    Dim astrBarNames As New cGdArray
    Dim aScanExpr As New cGdArray
    Dim strError As String
    
    strCodedText = Trim(m.strCodedText)
    If Len(strCodedText) > 0 Then
        ' Init the expression evaluator with list of scan expressions
        aScanExpr.Add strCodedText
        
        If optDaily Then
            MarketsInExpressions aScanExpr, 0#, False, astrBarNames, Nothing, "Daily"
        Else
            MarketsInExpressions aScanExpr, 0#, False, astrBarNames, Nothing, "Weekly"
        End If
        
        'astrBarNames(0) = "Market1"
        'astrBarNames(1) = "Weekly"
        'astrBarNames(2) = "GC"
        astrParms(0) = "CriteriaVerify"
        If SetupExpressions(astrParms, astrBarNames, aScanExpr, strError) Then
            EngineVerify = True
        Else
            'InfBox "i=[] ; h=Verify ERROR ; An error occured with the 'engine' verification:|Error #" & CStr(rc)
            InfBox "An error occured with the 'engine' verification.|Message: " & strError, "[]", , "Verify ERROR"
        End If
        
        ' clear the expression evaluator when done with it
        SetupExpressions astrParms '(clear expressions)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCriteria.EngineVerify", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetect
'' Description: Attempt to automatically detect the number of bars necessary
'' Inputs:      Expression
'' Returns:     Number of Bars (-1 if not calculated)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AutoDetect(ByVal strExpression As String) As Long
On Error GoTo ErrSection:

    Dim AD As New cAutoDetect           ' Auto detect object
    Dim lReturn As Long                 ' Return value for the function
    
    If optDaily Then
        lReturn = AD.AutoDetect(strExpression)
    Else
        lReturn = AD.AutoDetect(strExpression, , "Weekly")
    End If
    
    AutoDetect = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCriteria.AutoDetect"
    
End Function

Private Sub txtDecimalPlaces_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.txtDecimalPlaces.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtNumDays_Change()
On Error GoTo ErrSection:

    EnableToolbar True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.txtNumDays.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub EnableToolbar(ByVal bEnableSave As Boolean)
On Error GoTo ErrSection:

    With tbToolbar
        .Tools("ID_Toolbox").Enabled = Not m.bModal
        .Tools("ID_Save").Enabled = bEnableSave
        If Not m.Criteria Is Nothing Then
            If m.Criteria.UsageType = eCriteria_FilterCriteria Then
                .Tools("ID_SaveAs").Enabled = (Trim(m.strName) <> "")
            Else
                .Tools("ID_SaveAs").Enabled = False
            End If
        End If
        .Tools("ID_Rename").Enabled = (Trim(m.strName) <> "")
    End With
    
    optAutoRound.Enabled = Not m.bIsBoolean
    optCustomRound.Enabled = Not m.bIsBoolean
    txtDecimalPlaces.Enabled = optCustomRound.Value And (Not m.bIsBoolean)
    lblDecimals.Enabled = Not m.bIsBoolean
    optTradingUnits.Enabled = Not m.bIsBoolean
    
    If optOverride Then
        txtNumDays.Visible = False
        txtOverride.Visible = True
    Else
        txtNumDays.Visible = True
        txtOverride.Visible = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.EnableToolbar", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the filters combo box with the symbol groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(Optional ByVal bShowFilters As Boolean = False)
On Error Resume Next

    Dim lIndex As Long                  ' Index for a for loop
    Dim strID As String                 ' Symbol pool ID for the field
    Dim strType As String               ' Type of thing (i.e. Filter, Criteria, etc)
    Dim strPicture As String            ' Picture to use in the combo box
    Dim strSelID As String              ' ID of the currently selected item
    Dim bSelExists As Boolean           ' Old selection still exists
    Dim iSortStart As Long              ' Where to start the sort
    Dim strItem As String               ' Item to add to the combo box
    Dim aItems As New cGdArray          ' Items to add to the combo box
    Dim obj As Object                   ' Symbol Pool Object
    Dim bScans As Boolean               ' Are we doing scans?
   
    bScans = ScansEnabled
        
    If cboSymbolGroups.ComboItems.Count > 0 Then
        strSelID = cboSymbolGroups.SelectedItem.Key
        cboSymbolGroups.ComboItems.Clear
    End If
    
    ' get list of items to put into combo list
    With g.SymbolPool
        For lIndex = 0 To .ArrayTable.NumFields - 1
            strID = .FieldID(lIndex)
            If Len(strID) = 0 Then
                strType = "" '???
            Else
                strType = Left(strID, 3)
                strPicture = ""
                Set obj = .PoolObject(strID)
                Select Case UCase(strType)
                    Case "GRP"
                        If strID <> "GRP:_FLAGS_.GRP" Then
                            strPicture = ToolbarIcon("ID_SymbolGroups")
                        End If
                    Case "FIL"
                        If bScans And bShowFilters Then
                            strPicture = ToolbarIcon("ID_Filters")
                        End If
                End Select
                If Len(strPicture) > 0 And obj.IsActive = True Then
                    If strID = strSelID Then
                        bSelExists = True
                    End If
                    
                    If iSortStart = 0 And lIndex >= g.SymbolPool.OtherFieldsStart Then
                        iSortStart = aItems.Size
                    End If
                    
                    aItems.Add .ArrayTable.FieldName(lIndex) & vbTab _
                            & strID & vbTab & strPicture
                End If
            End If
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        cboSymbolGroups.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next


    If bSelExists Then
        cboSymbolGroups.ComboItems(strSelID).Selected = True
    Else
        cboSymbolGroups.ComboItems(1).Selected = True
    End If

    cboSymbolGroups.Refresh

End Sub

Public Function ShowMe(ByVal strPath As String, ByVal strID As String, _
                    Optional ByVal bModal As Boolean = False, _
                    Optional ByVal eType As eCriteriaUsageType = eCriteria_FilterCriteria, _
                    Optional ByVal strTextToPaste As String = "") As String
On Error GoTo ErrSection:

    Dim i&
    Dim Expr As cExpression
    Dim bDirty As Boolean

    Set m.Criteria = New cCriteria
    
    m.bModal = bModal
    
    'TLB: The "-" acts as a flag to create a new criteria but defaulting to NOT active
    '(so won't ask to recalc after saving)
    If strID = "-" Then
        strID = ""
        m.Criteria.IsActive = False
    End If
    
    If Len(strID) > 0 Then
        If Not m.Criteria.FromFile(strPath, strID) Then
            Err.Raise vbObjectError + 1000, , strID & " could not be loaded"
        End If
    End If
    
    With m.Criteria
        bDirty = False
        If eType <> .UsageType Then
            .UsageType = eType
            .ID = ""
            bDirty = True
        End If
    
        m.strName = .Name
        m.strDescription = .Desc
        m.strFormattedText = .FormattedText
        If .EnglishText <> "" Then
            ' if override value < 0 then it hadn't been stored yet, so we want to determine if
            ' the current value had already been set manually or not (i.e. if it doesn't match
            ' the auto-detect value then we'll assume it had been manually overridden)
            If .FormattedText = "" Or .NumDaysOverride < 0 Then
                Editor1.Text = .EnglishText
                txtNumDays = ""
                'if error during verify, just resume so will leave them with
                'the bad stuff in Red on the screen (instead of unloading the form)
                On Error Resume Next
                Verify
                On Error GoTo ErrSection
                If .NumDaysOverride < 0 Then
                    ' if current NumDays > auto-detected days, then must have been manually set
                    ' (so assign current value to Override and reassign the auto-detect value)
                    If .NumDays > Val(txtNumDays) Then
                        .NumDaysOverride = .NumDays
                        .NumDaysCalc = Val(txtNumDays)
                    Else
                        .NumDaysOverride = 0
                    End If
                End If
            Else
                If m.Function Is Nothing Then Set m.Function = New cFunction
                Editor1.TextRTF = m.Function.GetRTF(.FormattedText)
            End If
        End If
        m.strCodedText = .CodedText
        m.bIsBoolean = .IsBoolean
        If .NumDaysCalc < 0 Then .NumDaysCalc = 0
        txtNumDays.Text = Str(.NumDaysCalc)
        If .NumDaysOverride <= 0 Then
            optAutoDetect = True
            txtOverride.Text = "0"
        Else
            optOverride = True
            txtOverride.Text = Str(.NumDaysOverride)
        End If
        If .IsWeekly Then optWeekly.Value = True Else optDaily.Value = True
        If .GroupID = "" Then
            ' default is All Symbols
            cboSymbolGroups.ComboItems("GRP:ALL SYMBOLS.GRP").Selected = True
        Else
            ' TLB 5/29/2005: must check to see if group still exists
            On Error Resume Next
            i = 99
            i = cboSymbolGroups.ComboItems(.GroupID).Selected
            On Error GoTo ErrSection:
            If i = 99 Then
                ' group no longer exists
                cboSymbolGroups.ComboItems.Add 1, .GroupID, "(not found)"
            End If
            cboSymbolGroups.ComboItems(.GroupID).Selected = True
        End If
        
        If .UsageType = eCriteria_FilterCriteria Then
            SetEditorCaption Me, "Criteria", .Name
        Else
            SetEditorCaption Me, "Quote Board Field", .Name
        End If
        
        Select Case .PriceDisplay
            Case eCriteria_AutoRound
                optAutoRound = True
                optCustomRound = False
                optTradingUnits = False
                
            Case eCriteria_RoundToDecimal
                optAutoRound = False
                optCustomRound = True
                optTradingUnits = False
            
            Case eCriteria_TradingUnits
                optAutoRound = False
                optCustomRound = False
                optTradingUnits = True
        
        End Select
        
        txtDecimalPlaces.Text = Str(.DecimalPlaces)
    End With
    
    tbToolbar.Tools("ID_Verify").Enabled = False
    EnableToolbar bDirty
    m.bOK = False
    If bModal = True Then
        ShowForm Me, eForm_ActModal, frmMain
    Else
        ShowForm Me, eForm_Nonmodal, frmMain
    End If
    If bModal Then
        If m.bOK Then ShowMe = m.Criteria.ID
        Unload Me
    End If

    If Len(strTextToPaste) > 0 Then
        Editor1.Text = strTextToPaste
        Editor1_Change
    End If

ErrExit:
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCriteria.ShowMe", eGDRaiseError_Raise

End Function

' Returns True if Cancelled
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    Dim bSkipAutoIf As Boolean
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        If WindowState = vbMinimized Then WindowState = vbNormal
    
        Set g.ActiveEditor = Nothing
        Editor1.RemoveTradeSense
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

Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "CNV Criteria", Me, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.PrintMe", eGDRaiseError_Raise
    
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

    With frmPrintPreview.vp
        .StartDoc
        
        ' Header and Footer
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Bold = True
        .Font.Size = 14
        .FontUnderline = True
        .Text = vbLf & "Filter:"
        .FontUnderline = False
        .Text = "    " & Trim(m.strName) & vbLf
        .Font.Size = 12
        .Font.Bold = False
        .Text = vbLf & "Description: " & Trim(m.strDescription) & vbLf
        If optDaily Then
            .Text = "Based On: Daily Bars" & vbLf
        Else
            .Text = "Based On: Weekly Bars" & vbLf
        End If
        .Text = "Bars Required: " & txtNumDays.Text & vbLf
        .Text = "Calculate For: " & cboSymbolGroups.Text & vbLf & vbLf
        
        .Font.Bold = True
        .Font.Size = 14
        .FontUnderline = True
        .Text = "Formula:" & vbLf
        .FontUnderline = False
        .Font.Size = 12
        .Font.Bold = False
        
        .Text = vbLf
        If frmPrintPreview.GoingToFile Then
            .Text = Editor1.Text
        Else
            .TextRTF = Editor1.TextRTF
        End If
        .Text = vbLf
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidMarket
'' Description: Determine whether the given market variable is valid for use
''              in criteria expressions
'' Inputs:      Market
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidMarket(ByVal strMarket As String) As Boolean
On Error GoTo ErrSection:
    
    Dim strSymbol As String             ' Symbol of the given market
    Dim strPeriod As String             ' Period of the given market
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    ValidMarket = False
    strMarket = UCase(Trim(strMarket))
    Select Case UCase(strMarket)
        Case "MARKET1", "WEEKLY", "GC"
            ValidMarket = True
        
        Case "DAILY", "MONTHLY"
            If optWeekly.Value = False Then
                ValidMarket = True
            End If
            
        Case Else
            If Left(strMarket, 1) = Chr(34) And Right(strMarket, 1) = Chr(34) Then
                strSymbol = Parse(Replace(strMarket, Chr(34), ""), ",", 1)
                strPeriod = Parse(Replace(strMarket, Chr(34), ""), ",", 2)
                ' TLB: daily/weekly can just be ignored since it's always supported
                If strPeriod = "WEEKLY" Or strPeriod = "DAILY" Then
                    strPeriod = ""
                End If

                If Len(strSymbol) > 0 And Len(strPeriod) > 0 Then
                    Bars.Prop(eBARS_PeriodicityStr) = strPeriod
                    If Bars.Prop(eBARS_Periodicity) < ePRD_Days Then
                        DM_GetBars Bars, strSymbol, strPeriod, LastDailyDownload - 5
                    Else
                        DM_GetBars Bars, strSymbol, strPeriod
                    End If
                    If Bars.Size > 0 Then ValidMarket = True
                ElseIf strSymbol = "-067" Then
                    ValidMarket = True ' TLB: we now allow for "-067" as a special case
                ElseIf Len(strSymbol) > 0 Then
                    If optDaily Then strPeriod = "Daily" Else strPeriod = "Weekly"
                    DM_GetBars Bars, strSymbol, strPeriod
                    If Bars.Size > 0 Then ValidMarket = True
                End If
            End If
        
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCriteria.ValidMarket", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPeriodInMarkets
'' Description: Fix the Period in "Of" expressions surrounded by quotes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixPeriodInMarkets()
On Error GoTo ErrSection:

    Dim astrTokens As New cGdArray      ' Array of space delimited tokens
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol of the market variable
    Dim strPeriod As String             ' Period of the market variable
    Dim Bars As New cGdBars             ' Temporary Bars structure
    
    astrTokens.SplitFields Editor1.Text, " "
    For lIndex = 0 To astrTokens.Size - 1
        If UCase(astrTokens(lIndex)) = "OF" Then
            If lIndex + 1 < astrTokens.Size Then
                If Left(astrTokens(lIndex + 1), 1) = Chr(34) And Right(astrTokens(lIndex + 1), 1) = Chr(34) Then
                    strSymbol = Parse(Replace(astrTokens(lIndex + 1), Chr(34), ""), ",", 1)
                    strPeriod = Parse(Replace(astrTokens(lIndex + 1), Chr(34), ""), ",", 2)
                    
                    If Len(strPeriod) > 0 Then
                        Bars.Prop(eBARS_PeriodicityStr) = strPeriod
                        strPeriod = Bars.Prop(eBARS_PeriodicityStr)
                        
                        astrTokens(lIndex + 1) = Chr(34) & strSymbol & "," & strPeriod & Chr(34)
                    Else
                        astrTokens(lIndex + 1) = Chr(34) & strSymbol & Chr(34)
                    End If
                End If
            End If
        End If
    Next lIndex
    Editor1.Text = astrTokens.JoinFields(" ")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.FixPeriodInMarkets", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOverride_Change
'' Description: When the control changes, set the dirty flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOverride_Change()
On Error GoTo ErrSection:

    EnableToolbar True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.txtOverride_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtOverride_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtOverride_GotFocus()
On Error GoTo ErrSection:
    
    SelectAll txtOverride

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCriteria.txtOverride_GotFocus"
    
End Sub

Public Property Let CondBuilderExpr(ByVal strExpr As String)
On Error Resume Next
    
   Editor1.Text = strExpr
   tbToolbar.Tools("ID_CondBuilder").Visible = False

End Property


