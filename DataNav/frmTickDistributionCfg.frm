VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTickDistributionCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraColorSelector 
      Height          =   615
      Left            =   2040
      TabIndex        =   16
      Top             =   120
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
      Caption         =   "frmTickDistributionCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTickDistributionCfg.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTickDistributionCfg.frx":0040
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectColor gdOutlineColor 
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP lblOutline 
         Height          =   255
         Left            =   120
         Top             =   180
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
         Caption         =   "frmTickDistributionCfg.frx":005C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTickDistributionCfg.frx":00AA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":00CA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdFont 
      Height          =   315
      Left            =   4455
      TabIndex        =   11
      Top             =   683
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
      Caption         =   "frmTickDistributionCfg.frx":00E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmTickDistributionCfg.frx":010E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmTickDistributionCfg.frx":012E
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   3375
      Left            =   225
      TabIndex        =   7
      Top             =   1163
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
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
      Caption         =   "Price|Bid/Ask|Order Bar|Quote Bar|Volume|Misc."
      Align           =   0
      Appearance      =   1
      CurrTab         =   5
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
      Begin HexUniControls.ctlUniFrameWL fraMisc 
         Height          =   3000
         Left            =   45
         TabIndex        =   41
         Top             =   330
         Width           =   5445
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTickDistributionCfg.frx":014A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistributionCfg.frx":0178
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":0198
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkRightToLeft 
            Height          =   255
            Left            =   600
            TabIndex        =   48
            Top             =   1800
            Visible         =   0   'False
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
            Caption         =   "frmTickDistributionCfg.frx":01B4
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":020C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":022C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkProfitLoss 
            Height          =   255
            Left            =   600
            TabIndex        =   46
            Top             =   1020
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
            Caption         =   "frmTickDistributionCfg.frx":0248
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0288
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":02A8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBlankRows 
            Height          =   315
            Left            =   3720
            TabIndex        =   42
            Top             =   420
            Width           =   675
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTickDistributionCfg.frx":02C4
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmTickDistributionCfg.frx":02E4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0304
         End
         Begin gdOCX.gdSelectColor gdTickLineColor 
            Height          =   315
            Left            =   2940
            TabIndex        =   44
            Top             =   1380
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Enabled         =   0   'False
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniCheckXP chkTickLine 
            Height          =   255
            Left            =   600
            TabIndex        =   47
            Top             =   1410
            Visible         =   0   'False
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
            Caption         =   "frmTickDistributionCfg.frx":0320
            Enabled         =   0   'False
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":035C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":037C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label20 
            Height          =   195
            Left            =   3120
            Top             =   1200
            Visible         =   0   'False
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
            Caption         =   "frmTickDistributionCfg.frx":0398
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":03D6
            Style           =   0
            Enabled         =   0   'False
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":03F6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label14 
            Height          =   255
            Left            =   840
            Top             =   450
            Width           =   2895
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":0412
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0478
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0498
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraVolume 
         Height          =   3000
         Left            =   -6090
         TabIndex        =   37
         Top             =   330
         Width           =   5445
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTickDistributionCfg.frx":04B4
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistributionCfg.frx":04E6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":0506
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL Frame1 
            Height          =   1155
            Left            =   120
            TabIndex        =   50
            Top             =   1500
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
            Caption         =   "frmTickDistributionCfg.frx":0522
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":0564
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0584
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtShowVolMin 
               Height          =   315
               Left            =   4020
               TabIndex        =   52
               Top             =   360
               Width           =   975
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTickDistributionCfg.frx":05A0
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
               Alignment       =   1
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmTickDistributionCfg.frx":05C0
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":05E0
            End
            Begin HexUniControls.ctlUniTextBoxXP txtShowVolMax 
               Height          =   315
               Left            =   4020
               TabIndex        =   51
               Top             =   735
               Width           =   975
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmTickDistributionCfg.frx":05FC
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
               Alignment       =   1
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frmTickDistributionCfg.frx":061C
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":063C
            End
            Begin HexUniControls.ctlUniLabelXP Label16 
               Height          =   255
               Left            =   120
               Top             =   420
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
               Caption         =   "frmTickDistributionCfg.frx":0658
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":06E2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":0702
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label17 
               Height          =   255
               Left            =   120
               Top             =   765
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
               Caption         =   "frmTickDistributionCfg.frx":071E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":07A2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":07C2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkVolumeText 
            Height          =   255
            Left            =   1140
            TabIndex        =   39
            Top             =   630
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
            Caption         =   "frmTickDistributionCfg.frx":07DE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":081E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":083E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkVolumeBar 
            Height          =   255
            Left            =   1140
            TabIndex        =   38
            Top             =   300
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
            Caption         =   "frmTickDistributionCfg.frx":085A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0898
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":08B8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdFloodColor 
            Height          =   315
            Left            =   2700
            TabIndex        =   40
            Top             =   930
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label19 
            Height          =   255
            Left            =   1140
            Top             =   960
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
            Caption         =   "frmTickDistributionCfg.frx":08D4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0914
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0934
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraQuoteBar 
         Height          =   3000
         Left            =   -6390
         TabIndex        =   34
         Top             =   330
         Width           =   5445
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTickDistributionCfg.frx":0950
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistributionCfg.frx":0986
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":09A6
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkQuoteBar 
            Height          =   255
            Left            =   960
            TabIndex        =   35
            Top             =   240
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
            Caption         =   "frmTickDistributionCfg.frx":09C2
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0A00
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0A20
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgQBarCols 
            Height          =   2040
            Left            =   960
            TabIndex        =   36
            Top             =   600
            Width           =   2685
            _cx             =   4736
            _cy             =   3598
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
      End
      Begin HexUniControls.ctlUniFrameWL fraOrderAcctBar 
         Height          =   3000
         Left            =   -6690
         TabIndex        =   26
         Top             =   330
         Width           =   5445
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTickDistributionCfg.frx":0A3C
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistributionCfg.frx":0A7A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":0A9A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkHighlightEquity 
            Height          =   220
            Left            =   2805
            TabIndex        =   18
            Top             =   2610
            Width           =   1845
            _ExtentX        =   3254
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
            Caption         =   "frmTickDistributionCfg.frx":0AB6
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0B00
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0B20
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdHighlightEquity 
            Height          =   315
            Left            =   4680
            TabIndex        =   24
            Top             =   2550
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            Color           =   14737632
            CustomColor     =   14737632
         End
         Begin HexUniControls.ctlUniCheckXP chkHighlightPos 
            Height          =   220
            Left            =   2805
            TabIndex        =   33
            Top             =   2280
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
            Caption         =   "frmTickDistributionCfg.frx":0B3C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0B80
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0BA0
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgOrderButtons 
            Height          =   1755
            Left            =   3120
            TabIndex        =   60
            Top             =   390
            Visible         =   0   'False
            Width           =   2355
            _cx             =   4154
            _cy             =   3096
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
         Begin HexUniControls.ctlUniCheckXP chkOpenEntries 
            Height          =   220
            Left            =   120
            TabIndex        =   43
            Top             =   1890
            Width           =   1815
            _ExtentX        =   3201
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
            Caption         =   "frmTickDistributionCfg.frx":0BBC
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0BFE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0C1E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdConfigOrder 
            Height          =   255
            Left            =   1815
            TabIndex        =   59
            Top             =   594
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
            Caption         =   "frmTickDistributionCfg.frx":0C3A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0C6C
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0C8C
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdConfigAcct 
            Height          =   255
            Left            =   1815
            TabIndex        =   58
            Top             =   240
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
            Caption         =   "frmTickDistributionCfg.frx":0CA8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0CDA
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0CFA
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkShowAvgEntry 
            Height          =   220
            Left            =   120
            TabIndex        =   56
            Top             =   1566
            Width           =   1815
            _ExtentX        =   3201
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
            Caption         =   "frmTickDistributionCfg.frx":0D16
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0D5A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0D7A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkOrdBarOnRight 
            Height          =   220
            Left            =   120
            TabIndex        =   55
            Top             =   918
            Width           =   2385
            _ExtentX        =   4207
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
            Caption         =   "frmTickDistributionCfg.frx":0D96
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0DE6
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0E06
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQty1 
            Height          =   315
            Left            =   420
            TabIndex        =   32
            Top             =   2400
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTickDistributionCfg.frx":0E22
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmTickDistributionCfg.frx":0E44
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0E64
         End
         Begin HexUniControls.ctlUniCheckXP chkOrderColumns 
            Height          =   220
            Left            =   120
            TabIndex        =   31
            Top             =   1242
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
            Caption         =   "frmTickDistributionCfg.frx":0E80
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":0ED4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0EF4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQty2 
            Height          =   315
            Left            =   1020
            TabIndex        =   30
            Top             =   2400
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTickDistributionCfg.frx":0F10
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmTickDistributionCfg.frx":0F32
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0F52
         End
         Begin HexUniControls.ctlUniTextBoxXP txtQty3 
            Height          =   315
            Left            =   1620
            TabIndex        =   29
            Top             =   2400
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTickDistributionCfg.frx":0F6E
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
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmTickDistributionCfg.frx":0F90
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":0FB0
         End
         Begin HexUniControls.ctlUniCheckXP chkAccountBar 
            Height          =   220
            Left            =   120
            TabIndex        =   28
            Top             =   270
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "frmTickDistributionCfg.frx":0FCC
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":100C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":102C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkOrderBar 
            Height          =   220
            Left            =   120
            TabIndex        =   27
            Top             =   594
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
            Caption         =   "frmTickDistributionCfg.frx":1048
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":1084
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":10A4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin gdOCX.gdSelectColor gdAvgEntryColor 
            Height          =   315
            Left            =   1995
            TabIndex        =   57
            Top             =   1560
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            Color           =   8454143
            CustomColor     =   8454143
         End
         Begin gdOCX.gdSelectColor gdHighlightPos 
            Height          =   315
            Left            =   4680
            TabIndex        =   45
            Top             =   2220
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            Color           =   14737632
            CustomColor     =   14737632
         End
         Begin VSFlex7LCtl.VSFlexGrid fgAcctCols 
            Height          =   1790
            Left            =   2850
            TabIndex        =   49
            Top             =   315
            Visible         =   0   'False
            Width           =   2355
            _cx             =   4154
            _cy             =   3157
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
         Begin HexUniControls.ctlUniLabelXP lblCfgGriid 
            Height          =   255
            Left            =   2850
            Top             =   120
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
            Caption         =   "frmTickDistributionCfg.frx":10C0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":1106
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1126
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label13 
            Height          =   255
            Left            =   480
            Top             =   2205
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
            Caption         =   "frmTickDistributionCfg.frx":1142
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":1190
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":11B0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraBidAsk 
         Height          =   3000
         Left            =   -6990
         TabIndex        =   25
         Top             =   330
         Width           =   5445
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTickDistributionCfg.frx":11CC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistributionCfg.frx":11FE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":121E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraText 
            Height          =   1140
            Left            =   1830
            TabIndex        =   53
            Top             =   240
            Width           =   1560
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":123A
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":1262
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1282
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectColor gdBidText 
               Height          =   315
               Left            =   690
               TabIndex        =   54
               Top             =   270
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdAskText 
               Height          =   315
               Left            =   690
               TabIndex        =   72
               Top             =   675
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Color           =   49152
               CustomColor     =   49152
            End
            Begin HexUniControls.ctlUniLabelXP Label4 
               Height          =   195
               Left            =   210
               Top             =   735
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
               Caption         =   "frmTickDistributionCfg.frx":129E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":12C4
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":12E4
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label5 
               Height          =   195
               Left            =   210
               Top             =   345
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
               Caption         =   "frmTickDistributionCfg.frx":1300
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1326
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1346
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraBackground 
            Height          =   1140
            Left            =   105
            TabIndex        =   78
            Top             =   240
            Width           =   1560
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":1362
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":1396
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":13B6
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectColor gdBidColor 
               Height          =   315
               Left            =   690
               TabIndex        =   79
               Top             =   270
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdAskColor 
               Height          =   315
               Left            =   690
               TabIndex        =   80
               Top             =   675
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Color           =   49152
               CustomColor     =   49152
            End
            Begin HexUniControls.ctlUniLabelXP Label7 
               Height          =   195
               Left            =   210
               Top             =   735
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
               Caption         =   "frmTickDistributionCfg.frx":13D2
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":13F8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1418
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label8 
               Height          =   195
               Left            =   210
               Top             =   345
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
               Caption         =   "frmTickDistributionCfg.frx":1434
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":145A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":147A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraMarketDepthColors 
            Height          =   2625
            Left            =   3495
            TabIndex        =   65
            Top             =   240
            Width           =   1860
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":1496
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":14CE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":14EE
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   1
               Left            =   1095
               TabIndex        =   66
               Top             =   693
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   2
               Left            =   1095
               TabIndex        =   67
               Top             =   1071
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   3
               Left            =   1095
               TabIndex        =   68
               Top             =   1449
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   5
               Left            =   1095
               TabIndex        =   69
               Top             =   2205
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   0
               Left            =   1095
               TabIndex        =   70
               Top             =   315
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdPriceLevelColor 
               Height          =   315
               Index           =   4
               Left            =   1095
               TabIndex        =   71
               Top             =   1827
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin HexUniControls.ctlUniLabelXP lblOther 
               Height          =   255
               Left            =   60
               Top             =   2235
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
               Caption         =   "frmTickDistributionCfg.frx":150A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1544
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1564
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblFifthBest 
               Height          =   255
               Left            =   60
               Top             =   1857
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
               Caption         =   "frmTickDistributionCfg.frx":1580
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":15B0
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":15D0
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblFourthBest 
               Height          =   255
               Left            =   60
               Top             =   1479
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
               Caption         =   "frmTickDistributionCfg.frx":15EC
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":161C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":163C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblThirdBest 
               Height          =   255
               Left            =   60
               Top             =   1101
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
               Caption         =   "frmTickDistributionCfg.frx":1658
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1688
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":16A8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblSecondBest 
               Height          =   255
               Left            =   60
               Top             =   723
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
               Caption         =   "frmTickDistributionCfg.frx":16C4
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":16F4
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1714
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblBestBidAsk 
               Height          =   255
               Left            =   60
               Top             =   345
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
               Caption         =   "frmTickDistributionCfg.frx":1730
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1768
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1788
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraOptHighlight 
            Height          =   1395
            Left            =   105
            TabIndex        =   61
            Top             =   1485
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
            Caption         =   "frmTickDistributionCfg.frx":17A4
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":17E6
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1806
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectColor gdLargestSizeColor 
               Height          =   315
               Left            =   2340
               TabIndex        =   73
               Top             =   690
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Color           =   49152
               CustomColor     =   49152
            End
            Begin HexUniControls.ctlUniRadioXP optHighlightByPrice 
               Height          =   220
               Left            =   180
               TabIndex        =   64
               Top             =   360
               Width           =   2715
               _ExtentX        =   4789
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
               Caption         =   "frmTickDistributionCfg.frx":1822
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":187A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":189A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optHighlightBySize 
               Height          =   220
               Left            =   180
               TabIndex        =   63
               Top             =   720
               Width           =   2160
               _ExtentX        =   3810
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
               Caption         =   "frmTickDistributionCfg.frx":18B6
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":190A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":192A
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optHighlightNone 
               Height          =   220
               Left            =   180
               TabIndex        =   62
               Top             =   1080
               Width           =   2100
               _ExtentX        =   3704
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
               Caption         =   "frmTickDistributionCfg.frx":1946
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":196E
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":198E
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraColors 
         Height          =   3000
         Left            =   -7290
         TabIndex        =   8
         Top             =   330
         Width           =   5445
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTickDistributionCfg.frx":19AA
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTickDistributionCfg.frx":19DC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":19FC
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraLadderVolStyle 
            Height          =   735
            Left            =   180
            TabIndex        =   74
            Top             =   2160
            Width           =   5085
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":1A18
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":1A54
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1A74
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optLadderVolStyle 
               Height          =   220
               Index           =   2
               Left            =   3720
               TabIndex        =   75
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "frmTickDistributionCfg.frx":1A90
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1AB8
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1AD8
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optLadderVolStyle 
               Height          =   220
               Index           =   1
               Left            =   1980
               TabIndex        =   76
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
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
               Caption         =   "frmTickDistributionCfg.frx":1AF4
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1B32
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1B52
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optLadderVolStyle 
               Height          =   220
               Index           =   0
               Left            =   120
               TabIndex        =   77
               Top             =   360
               Width           =   1695
               _ExtentX        =   2990
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
               Caption         =   "frmTickDistributionCfg.frx":1B6E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1BB2
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1BD2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
         End
         Begin HexUniControls.ctlUniRadioXP optBollinger 
            Height          =   220
            Left            =   180
            TabIndex        =   10
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "frmTickDistributionCfg.frx":1BEE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":1C3E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1C5E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optOneColor 
            Height          =   220
            Left            =   2850
            TabIndex        =   9
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "frmTickDistributionCfg.frx":1C7A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTickDistributionCfg.frx":1CC2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1CE2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraPriceColColors 
            Height          =   1425
            Left            =   180
            TabIndex        =   12
            Top             =   585
            Width           =   5085
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":1CFE
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":1D92
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":1DB2
            RightToLeft     =   0   'False
            Begin gdOCX.gdSelectColor gdBarColor 
               Height          =   315
               Left            =   120
               TabIndex        =   13
               Top             =   660
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   8388608
               CustomColor     =   8388608
            End
            Begin gdOCX.gdSelectColor gdUpColor 
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   300
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   49152
               CustomColor     =   49152
            End
            Begin gdOCX.gdSelectColor gdDownColor 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   1020
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniLabelXP Label1 
               Height          =   195
               Left            =   1140
               Top             =   720
               Width           =   2835
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmTickDistributionCfg.frx":1DCE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1E2A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1E4A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label2 
               Height          =   195
               Left            =   1140
               Top             =   360
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
               Caption         =   "frmTickDistributionCfg.frx":1E66
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1ED2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1EF2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label3 
               Height          =   195
               Left            =   1140
               Top             =   1080
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
               Caption         =   "frmTickDistributionCfg.frx":1F0E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":1F82
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":1FA2
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraPriceNeutral 
            Height          =   1425
            Left            =   180
            TabIndex        =   19
            Top             =   585
            Width           =   5085
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmTickDistributionCfg.frx":1FBE
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmTickDistributionCfg.frx":2028
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTickDistributionCfg.frx":2048
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optPriceTextFixed 
               Height          =   255
               Left            =   180
               TabIndex        =   21
               Top             =   990
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
               Caption         =   "frmTickDistributionCfg.frx":2064
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":20C6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":20E6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optPriceTextBicolor 
               Height          =   195
               Left            =   180
               TabIndex        =   20
               Top             =   660
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
               Caption         =   "frmTickDistributionCfg.frx":2102
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":2182
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":21A2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectColor gdPriceText 
               Height          =   315
               Left            =   3000
               TabIndex        =   22
               Top             =   960
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   0
               CustomColor     =   0
            End
            Begin gdOCX.gdSelectColor gdPriceBackground 
               Height          =   315
               Left            =   2460
               TabIndex        =   23
               Top             =   240
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               Color           =   49152
               CustomColor     =   49152
            End
            Begin HexUniControls.ctlUniLabelXP Label18 
               Height          =   195
               Left            =   960
               Top             =   300
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
               Caption         =   "frmTickDistributionCfg.frx":21BE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmTickDistributionCfg.frx":21FE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmTickDistributionCfg.frx":221E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSession 
      Height          =   975
      Left            =   495
      TabIndex        =   3
      Top             =   23
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
      Caption         =   "frmTickDistributionCfg.frx":223A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTickDistributionCfg.frx":2262
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTickDistributionCfg.frx":2282
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optCurrentSession 
         Height          =   220
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
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
         Caption         =   "frmTickDistributionCfg.frx":229E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTickDistributionCfg.frx":22E4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":2304
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdSessionDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   540
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Enabled         =   0   'False
         AllowWeekends   =   0   'False
         MaxDate         =   42611
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
      Begin HexUniControls.ctlUniRadioXP optDate 
         Height          =   220
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "frmTickDistributionCfg.frx":2320
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTickDistributionCfg.frx":2356
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":2376
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1860
      TabIndex        =   0
      Top             =   4583
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
      Caption         =   "frmTickDistributionCfg.frx":2392
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTickDistributionCfg.frx":23BE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTickDistributionCfg.frx":23DE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdClearAll 
         Height          =   375
         Left            =   2280
         TabIndex        =   81
         Top             =   60
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
         Caption         =   "frmTickDistributionCfg.frx":23FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTickDistributionCfg.frx":242C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":244C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   60
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
         Caption         =   "frmTickDistributionCfg.frx":2468
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTickDistributionCfg.frx":248C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":24AC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1170
         TabIndex        =   2
         Top             =   60
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
         Caption         =   "frmTickDistributionCfg.frx":24C8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTickDistributionCfg.frx":24F4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTickDistributionCfg.frx":2514
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmTickDistributionCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTickDistributionCfg.frm
'' Description: Form to allow the user to change price ladder settings
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kFormHt = 5580
Private Const kFormWd = 6075
Private Const kFraButtonSm = 2235
Private Const kFraButtonLg = 3315

Public Enum eLadderCfgTab
    eLadderTab_LastUsed = -1
    eLadderTab_Price
    eLadderTab_BidAskk
    eLadderTab_OrderBar
    eLadderTab_QbBar
    eLadderTab_Volume
    eLadderTab_Misc
End Enum

Private Type mPrivate
    frmTDGrid As frmTickDistribution
    bEditing As Boolean
    bNeutral As Boolean
    
    'these are to restore colors last set by user
    nBidColor As Long
    nAskColor As Long
    nBidTextColor As Long
    nAskTextColor As Long
    
    nBarColor As Long
    nUpColor As Long
    nDownColor As Long
    
    nPriceBackgroundColor As Long
    nPriceTextColor As Long
    
    nGridRow As Long            'for color selector
End Type

Private m As mPrivate

Private Sub chkAccountBar_Click()
On Error Resume Next:

    If Not Me.Visible Then Exit Sub
    
    If chkAccountBar.Value = vbUnchecked Then
        If chkOrderBar.Value = vbUnchecked Then
            ToggleConfigGrids False, True
        Else
            ToggleConfigGrids True
        End If
    Else
        ToggleConfigGrids False
    End If

End Sub

Private Sub chkOrderBar_Click()
On Error Resume Next:

    If Not Me.Visible Then Exit Sub
    
    If chkOrderBar.Value = vbUnchecked Then
        chkOrdBarOnRight.Enabled = False
        If chkAccountBar.Value = vbUnchecked Then
            ToggleConfigGrids False, True
        Else
            ToggleConfigGrids False
        End If
    Else
        chkOrdBarOnRight.Enabled = True
        ToggleConfigGrids True
    End If

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    
    If m.nGridRow > 0 Then m.frmTDGrid.OutlineCell m.nGridRow, 0, True
    Unload Me

End Sub

Private Sub cmdClearAll_Click()
On Error Resume Next

    m.frmTDGrid.OutlineCell -1, 0, True
    Unload Me

End Sub

Private Sub cmdConfigAcct_Click()
On Error Resume Next
    ToggleConfigGrids False
End Sub

Private Sub cmdConfigOrder_Click()
On Error Resume Next
    ToggleConfigGrids True
End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    'set font currently in use
    Me.Font.Name = m.frmTDGrid.GridFontName
    Me.Font.Size = m.frmTDGrid.GridFontSize
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.frmTDGrid.GridFontName = Me.Font.Name
        m.frmTDGrid.GridFontSize = Me.Font.Size
        m.frmTDGrid.RefreshGrid
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.cmdFont.Click", eGDRaiseError_Show
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim i&, strText$
    Dim nVolMin&, nVolMax&, nVolOnOff&
    
    Dim bDateChanged As Boolean
    Dim bColorChanged As Boolean
    Dim bVolChanged As Boolean
    Dim bOrdBarOnOffChanged As Boolean
    Dim bVolOnOffChanged As Boolean
    
    Dim ePrevObarMode As eGDOrderBarMode
    
    If VerifyQuantityPresets Then
        Me.Hide
    
        If m.nGridRow > 0 Then
            m.frmTDGrid.OutlineCell m.nGridRow, gdOutlineColor.Color, False
            GoTo ErrExit
        End If
        
        If optCurrentSession.Value <> m.frmTDGrid.IsCurrentSession Or _
            m.frmTDGrid.SessionDate <> gdSessionDate.Value Then
            bDateChanged = True
            Me.Hide
        End If
        m.frmTDGrid.SessionDate = gdSessionDate.Value
        m.frmTDGrid.IsCurrentSession = optCurrentSession.Value
        
        'if min/max volume changed then have to reload data table
        nVolMin = Int(ValOfText(txtShowVolMin.Text))
        nVolMax = Int(ValOfText(txtShowVolMax.Text))
        If nVolMin < 0 Then nVolMin = 0     'prevent negatives
        If nVolMax < 0 Then nVolMax = 0
        If nVolMin <> m.frmTDGrid.ShowVolMin Or nVolMax <> m.frmTDGrid.ShowVolMax Then
            m.frmTDGrid.ShowVolMin = nVolMin
            m.frmTDGrid.ShowVolMax = nVolMax
            bVolChanged = True
        End If
        
        bColorChanged = ColorChanged()
        'price col options & volume col flood colors
        m.frmTDGrid.FloodColor = gdFloodColor.Color
        If m.bNeutral Then
            m.frmTDGrid.BarColor = gdPriceBackground.Color
            m.frmTDGrid.UpColor = gdPriceBackground.Color
            m.frmTDGrid.DownColor = gdPriceBackground.Color
            If optPriceTextBicolor.Value = True Then
                m.frmTDGrid.FixedPriceColor = -1
            Else
                m.frmTDGrid.FixedPriceColor = gdPriceText.Color
            End If
        Else
            m.frmTDGrid.BarColor = gdBarColor.Color
            m.frmTDGrid.UpColor = gdUpColor.Color
            m.frmTDGrid.DownColor = gdDownColor.Color
            If gdPriceText.Color >= 0 Then
                m.frmTDGrid.FixedPriceColor = gdPriceText.Color
            End If
        End If
        
        If optLadderVolStyle(eLadderVol_BidAsk).Value = True Then
            m.frmTDGrid.LadderVolumeStyle = eLadderVol_BidAsk
        ElseIf optLadderVolStyle(eLadderVol_None).Value = True Then
            m.frmTDGrid.LadderVolumeStyle = eLadderVol_None
        Else
            m.frmTDGrid.LadderVolumeStyle = eLadderVol_LastTrade
        End If
        
        'bid/ask colors
        m.frmTDGrid.BidColor = gdBidColor.Color
        m.frmTDGrid.AskColor = gdAskColor.Color
        m.frmTDGrid.BidTextColor = gdBidText.Color
        m.frmTDGrid.AskTextColor = gdAskText.Color
        
        'market depth colors
        m.frmTDGrid.FirstColor = gdPriceLevelColor(0).Color
        m.frmTDGrid.SecondColor = gdPriceLevelColor(1).Color
        m.frmTDGrid.ThirdColor = gdPriceLevelColor(2).Color
        m.frmTDGrid.FourthColor = gdPriceLevelColor(3).Color
        m.frmTDGrid.FifthColor = gdPriceLevelColor(4).Color
        m.frmTDGrid.OtherColor = gdPriceLevelColor(5).Color
        m.frmTDGrid.LargestSizeColor = gdLargestSizeColor.Color
        
        m.frmTDGrid.FloodMktDepth = HighlightButtonsGet
                
        'show volume text or tick line options
        nVolOnOff = m.frmTDGrid.ShowVolumeBar
        m.frmTDGrid.ShowVolumeBar = chkVolumeBar.Value
        m.frmTDGrid.ShowVolumeText = chkVolumeText.Value
        m.frmTDGrid.ShowTickline = chkTickLine.Value
        m.frmTDGrid.TickLineColor = gdTickLineColor.Color
        m.frmTDGrid.TickLineRightToLeft = chkRightToLeft
        If nVolOnOff <> m.frmTDGrid.ShowVolumeBar Then bVolOnOffChanged = True
        
        'profit loss
        m.frmTDGrid.ShowProfitLoss = chkProfitLoss.Value
        
        'quote bar
        strText = ""
        m.frmTDGrid.ShowQuoteBar = chkQuoteBar.Value
        strText = ParseGridCtrl(fgQBarCols)
        m.frmTDGrid.QuoteBarHeader strText
            
        'account bar
        strText = ""
        m.frmTDGrid.ShowAccountBar = chkAccountBar.Value
        strText = ParseGridCtrl(fgAcctCols)
        m.frmTDGrid.AccountBarHeader strText
            
        'order bar
        ePrevObarMode = m.frmTDGrid.ShowOrderBar
        If chkOrderBar.Value = vbUnchecked Then
            If chkOrdBarOnRight.Value = vbChecked Then
                m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_LastShownOnRight
            Else
                m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_NotShown
            End If
        ElseIf chkOrdBarOnRight.Value = vbChecked Then
            m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_Right
        Else
            m.frmTDGrid.ShowOrderBar = eGDOrderBarMode_BottomNarrow
        End If
        
        If m.frmTDGrid.ShowOrderBar <> ePrevObarMode Then bOrdBarOnOffChanged = True
        
        strText = ""
        m.frmTDGrid.OrderColumns = chkOrderColumns.Value
        strText = ParseOrderButtonsGrid(fgOrderButtons)
        m.frmTDGrid.OrdBarCtrls = strText
        
        i = ValOfText(txtBlankRows.Text)
        If i > 0 And i <> m.frmTDGrid.BlankRows Then
            m.frmTDGrid.BlankRows = i
            bVolChanged = True         'set this to re-load table
        End If
    
        g.Broker.SetQuantityPresets m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, Int(ValOfText(txtQty1.Text)), Int(ValOfText(txtQty2.Text)), Int(ValOfText(txtQty3.Text))
        
        m.frmTDGrid.ShowOpenEntries = chkOpenEntries.Value
        m.frmTDGrid.ShowAvgEntry = chkShowAvgEntry.Value
        m.frmTDGrid.AvgEntryColor = gdAvgEntryColor.Color
        
        i = m.frmTDGrid.HilitePosColor
        If chkHighlightPos.Value = vbChecked Then
            m.frmTDGrid.HilitePosColor = gdHighlightPos.Color
        ElseIf i <> -2 Then
            m.frmTDGrid.HilitePosColor = -1 * gdHighlightPos.Color
        End If
                    
        i = m.frmTDGrid.HiliteEquityColor
        If chkHighlightEquity.Value = vbChecked Then
            m.frmTDGrid.HiliteEquityColor = gdHighlightEquity.Color
        ElseIf i <> -2 Then
            m.frmTDGrid.HiliteEquityColor = -1 * gdHighlightEquity.Color
        End If
        
        m.frmTDGrid.RefreshGrid bDateChanged, bColorChanged, m.frmTDGrid.SummaryBarHeight, bVolChanged, bOrdBarOnOffChanged, bVolOnOffChanged
    
        Unload Me
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.cmdOK.Click", eGDRaiseError_Show
    
End Sub

Private Sub fgOrderButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim r&
    
    If chkOrdBarOnRight.Value = vbChecked Then
        With fgOrderButtons
            If .Row >= .FixedRows And .Row < .Rows Then
                r = .Row
                .Select r, 0, r, .Cols - 1
                .Refresh
                .DragRow r
            End If
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistributionCfg.fgOrderButtons_MouseDown"

End Sub

Private Sub fgQBarCols_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    m.bEditing = True
End Sub

Private Sub fgQBarCols_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub fgQBarCols_Click()
    ToggleShow
End Sub

Private Sub Form_Load()
    
    'set form icon
    Me.Icon = Picture16(ToolbarIcon("ID_TickDistribution"), , True)
    
    g.Styler.StyleForm Me
    
    With fraPriceColColors
        fraPriceNeutral.Move .Left, .Top, .Width, .Height
    End With
    
    With fgAcctCols
        fgOrderButtons.Move .Left, .Top
    End With
    
End Sub

Private Sub ToggleShow()
On Error Resume Next:

    If m.bEditing Then
        m.bEditing = False
        Exit Sub
    End If

    With fgQBarCols
        If .Col = 0 Then
            If .Cell(flexcpChecked, .Row, .Col) = 1 Then
                .Cell(flexcpChecked, .Row, .Col) = 2
            Else
                .Cell(flexcpChecked, .Row, .Col) = 1
            End If
        End If
    End With

End Sub

Public Sub ShowMe(frmCaller As Form, Optional ByVal nRow& = -1, _
    Optional ByVal nLastUsedColor& = -1, Optional ByVal strPrice$ = "", _
    Optional ByVal eTab As eLadderCfgTab = eLadderTab_LastUsed, _
    Optional ByVal bShowAcctCfgGrid As Boolean = False)
On Error GoTo ErrSection:
   
    Dim i&
    Dim nBidAskOrdBtns As Long
    Dim lPreset1 As Long                ' First order quantity preset
    Dim lPreset2 As Long                ' Second order quantity preset
    Dim lPreset3 As Long                ' Third order quantity preset
    
    Set m.frmTDGrid = frmCaller
    m.nGridRow = nRow
    
    If nRow > 0 Then
        If nLastUsedColor = -1 Then nLastUsedColor = vbYellow
        cmdCancel.Caption = "Clear"
        
        fraButtons.Width = kFraButtonLg
        fraSession.Visible = False
        cmdFont.Visible = False
        vsIndexTab1.Visible = False
        
        Me.Caption = "Select Color"
        Me.BorderStyle = 4          'fixed toolwindow
        
        Me.Width = fraColorSelector.Width + 500
        Me.Height = fraColorSelector.Height + fraButtons.Height * 2
        
        lblOutline.Caption = "Outline " & strPrice & " with color:"
        gdOutlineColor.Color = nLastUsedColor
        
        fraColorSelector.Move Me.Width / 2 - fraColorSelector.Width / 2, 0
        fraButtons.Move Me.Width / 2 - fraButtons.Width / 2, fraColorSelector.Top + fraColorSelector.Height
        
        'RH commented out fraColorSelector.BorderStyle = 0
        fraColorSelector.Visible = True

        Me.Move m.frmTDGrid.Left + (m.frmTDGrid.Width - Me.Width) / 2, m.frmTDGrid.Top + (m.frmTDGrid.Height - Me.Height) / 2
        ShowForm Me
        GoTo ErrExit
        
    Else
        fraButtons.Width = kFraButtonSm
        fraSession.Visible = True
        cmdFont.Visible = True
        vsIndexTab1.Visible = True
        fraColorSelector.Visible = False
        
        fraButtons.Top = vsIndexTab1.Top + vsIndexTab1.Height + 50
        
        Me.Width = kFormWd
        Me.Height = kFormHt
        Me.BorderStyle = 3          'fixed dialog
    End If
    
    
    If eTab = eLadderTab_LastUsed Then
        'get last used tab
        eTab = GetIniFileProperty("SettingsTab", eLadderTab_Price, "Price Ladder", g.strIniFile)
    End If
    If eTab < eLadderTab_Price Or eTab > eLadderTab_Misc Then eTab = eLadderTab_Price
    
    'volume
    gdFloodColor.Color = m.frmTDGrid.FloodColor
    gdTickLineColor.Color = m.frmTDGrid.TickLineColor
    If m.frmTDGrid.ShowVolMin > 0 Then
        txtShowVolMin.Text = Str(m.frmTDGrid.ShowVolMin)
    Else
        txtShowVolMin.Text = ""
    End If
    If m.frmTDGrid.ShowVolMax > 0 Then
        txtShowVolMax.Text = Str(m.frmTDGrid.ShowVolMax)
    Else
        txtShowVolMax.Text = ""
    End If
    'price
    gdBarColor.Color = m.frmTDGrid.BarColor
    gdUpColor.Color = m.frmTDGrid.UpColor
    gdDownColor.Color = m.frmTDGrid.DownColor
    gdPriceBackground.Color = m.frmTDGrid.BarColor      'for one-color background
    m.nBarColor = m.frmTDGrid.BarColor
    m.nUpColor = m.frmTDGrid.UpColor
    m.nDownColor = m.frmTDGrid.DownColor
    m.nPriceBackgroundColor = m.frmTDGrid.BarColor
    m.nPriceTextColor = m.frmTDGrid.FixedPriceColor
    If m.frmTDGrid.FixedPriceColor >= 0 Then
        gdPriceText.Color = m.nPriceTextColor
        optPriceTextFixed.Value = True
    Else
        gdPriceText.Color = vbBlack
        optPriceTextBicolor.Value = True
    End If
    i = m.frmTDGrid.LadderVolumeStyle
    If i < eLadderVol_LastTrade Or i > eLadderVol_None Then i = eLadderVol_LastTrade
    optLadderVolStyle(i).Value = True
    
    'bid/ask
    gdBidColor.Color = m.frmTDGrid.BidColor
    gdAskColor.Color = m.frmTDGrid.AskColor
    gdBidText.Color = m.frmTDGrid.BidTextColor
    gdAskText.Color = m.frmTDGrid.AskTextColor
    m.nAskColor = m.frmTDGrid.AskColor
    m.nBidColor = m.frmTDGrid.BidColor
    m.nAskTextColor = m.frmTDGrid.AskTextColor
    m.nBidTextColor = m.frmTDGrid.BidTextColor
        
    'market depth colors
    gdPriceLevelColor(0).Color = m.frmTDGrid.FirstColor
    gdPriceLevelColor(1).Color = m.frmTDGrid.SecondColor
    gdPriceLevelColor(2).Color = m.frmTDGrid.ThirdColor
    gdPriceLevelColor(3).Color = m.frmTDGrid.FourthColor
    gdPriceLevelColor(4).Color = m.frmTDGrid.FifthColor
    gdPriceLevelColor(5).Color = m.frmTDGrid.OtherColor
    gdLargestSizeColor.Color = m.frmTDGrid.LargestSizeColor
    
    HighlightButtonsSet m.frmTDGrid.FloodMktDepth
        
    'volume text or tickline
    chkVolumeBar.Value = m.frmTDGrid.ShowVolumeBar
    chkVolumeText.Value = m.frmTDGrid.ShowVolumeText
    chkTickLine.Value = m.frmTDGrid.ShowTickline
    chkRightToLeft.Value = m.frmTDGrid.TickLineRightToLeft
    If m.frmTDGrid.SecType = "I" Then
        chkVolumeBar.Enabled = False
        chkVolumeText.Enabled = False
    Else
        chkVolumeBar.Enabled = True
        chkVolumeText.Enabled = True
    End If
    txtBlankRows.Text = Str(m.frmTDGrid.BlankRows)
    
    If m.frmTDGrid.IsCurrentSession Then
        optCurrentSession.Value = True
        optDate.Value = False
        gdSessionDate.Value = Date
    Else
        optCurrentSession.Value = False
        optDate.Value = True
        gdSessionDate.Value = m.frmTDGrid.SessionDate
    End If
    
    'show/hide profit loss
    chkProfitLoss.Value = m.frmTDGrid.ShowProfitLoss
    'show/hide quote bar
    chkQuoteBar.Value = m.frmTDGrid.ShowQuoteBar
    'order bar options
    chkHighlightPos.Visible = False
    chkHighlightEquity.Visible = False
    gdHighlightPos.Visible = False
    gdHighlightEquity.Visible = False
    
    i = m.frmTDGrid.HilitePosColor
    If i > 0 Then
        chkHighlightPos.Value = vbChecked
        gdHighlightPos.Color = m.frmTDGrid.HilitePosColor
    ElseIf i <> -2 Then
        gdHighlightPos.Color = Abs(m.frmTDGrid.HilitePosColor)
    End If
        
    i = m.frmTDGrid.HiliteEquityColor
    If i > 0 Then
        chkHighlightEquity.Value = vbChecked
        gdHighlightEquity.Color = m.frmTDGrid.HiliteEquityColor
    ElseIf i <> -2 Then
        gdHighlightEquity.Color = Abs(m.frmTDGrid.HiliteEquityColor)
    End If
    
    Select Case m.frmTDGrid.ShowOrderBar
        Case eGDOrderBarMode_LastShownOnRight
            chkOrderBar.Value = vbUnchecked
            chkOrdBarOnRight.Value = vbChecked
            chkOrdBarOnRight.Enabled = False
        Case eGDOrderBarMode_NotShown, eGDOrderBarMode_LastShownBottom
            chkOrderBar.Value = vbUnchecked
            chkOrdBarOnRight.Value = vbUnchecked
            chkOrdBarOnRight.Enabled = False
        Case eGDOrderBarMode_BottomWide, eGDOrderBarMode_BottomNarrow, eGDOrderBarMode_BottomContinuous
            chkOrderBar.Value = vbChecked
            chkOrdBarOnRight.Value = vbUnchecked
            chkOrdBarOnRight.Enabled = True
        Case Else
            chkOrderBar.Value = vbChecked
            chkOrdBarOnRight.Value = vbChecked
            chkOrdBarOnRight.Enabled = True
    End Select
    
    chkOrderColumns.Value = m.frmTDGrid.OrderColumns
    chkAccountBar.Value = m.frmTDGrid.ShowAccountBar

    chkOpenEntries.Value = m.frmTDGrid.ShowOpenEntries
    chkShowAvgEntry.Value = m.frmTDGrid.ShowAvgEntry
    gdAvgEntryColor.Color = m.frmTDGrid.AvgEntryColor
        
    g.Broker.GetQuantityPresets m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset1, lPreset2, lPreset3
    txtQty1.Text = Str(lPreset1)
    txtQty2.Text = Str(lPreset2)
    txtQty3.Text = Str(lPreset3)
        
    InitQBarGrid fgQBarCols, m.frmTDGrid.QBarColArray
    InitAccountGrid fgAcctCols, m.frmTDGrid.ABarColArray, m.frmTDGrid.SecType
    InitOrderButtonsGrid fgOrderButtons, m.frmTDGrid.OrdBarCtrls
    
    If IsNeutralScheme() Then
        fraPriceNeutral.Visible = True
        fraPriceColColors.Visible = False
        m.bNeutral = True
        optOneColor.Value = True
    Else
        fraPriceNeutral.Visible = False
        fraPriceColColors.Visible = True
        m.bNeutral = False
        optBollinger.Value = True
    End If
    
    vsIndexTab1.CurrTab = eTab
    
    If eTab = eLadderTab_OrderBar Then
        If bShowAcctCfgGrid Then
            ToggleConfigGrids False
        Else
            ToggleConfigGrids True
        End If
    End If
    
    Me.Move m.frmTDGrid.Left + (m.frmTDGrid.Width - Me.Width) / 2, m.frmTDGrid.Top + (m.frmTDGrid.Height - Me.Height) / 2
    ShowForm Me, True
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.ShowMe", eGDRaiseError_Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m.frmTDGrid = Nothing
    
    If vsIndexTab1.Visible Then
        SetIniFileProperty "SettingsTab", vsIndexTab1.CurrTab, "Price Ladder", g.strIniFile
    End If
End Sub

Private Sub gdAskColor_LostFocus()
    m.nAskColor = gdAskColor.Color
End Sub

Private Sub gdAskText_LostFocus()
    m.nAskTextColor = gdAskText.Color
End Sub

Private Sub gdBarColor_LostFocus()
    m.nBarColor = gdBarColor.Color
End Sub

Private Sub gdBidColor_LostFocus()
    m.nBidColor = gdBidColor.Color
End Sub

Private Sub gdBidText_LostFocus()
    m.nBidTextColor = gdBidText.Color
End Sub

Private Sub gdDownColor_LostFocus()
    m.nDownColor = gdDownColor.Color
End Sub

Private Sub gdOutlineColor_Changed()
    cmdOK_Click
End Sub

Private Sub gdPriceBackground_LostFocus()
    m.nPriceBackgroundColor = gdPriceBackground.Color
End Sub

Private Sub gdPriceText_LostFocus()
    m.nPriceTextColor = gdPriceText.Color
End Sub

Private Sub gdUpColor_LostFocus()
    m.nUpColor = gdUpColor.Color
End Sub

Private Sub optBollinger_Click()

    If Not fraPriceColColors.Visible Then
        If gdUpColor.Color = gdDownColor.Color And gdBarColor.Color Then
            gdUpColor.Color = RGB(0, 192, 0)
            gdDownColor.Color = RGB(255, 0, 0)
            gdBarColor.Color = RGB(0, 0, 128)
        End If
        fraPriceNeutral.Visible = False
        fraPriceColColors.Visible = True
    End If
    
    m.bNeutral = False

End Sub

Private Sub optCurrentSession_Click()
    gdSessionDate.Enabled = False
End Sub

Private Sub optDate_Click()
    gdSessionDate.Enabled = True
End Sub

Private Function ColorChanged() As Boolean
On Error GoTo ErrSection:

    Dim bChanged As Boolean
    
    If m.frmTDGrid.FloodColor <> gdFloodColor.Color Or _
      m.frmTDGrid.BarColor <> gdBarColor.Color Or _
      m.frmTDGrid.UpColor <> gdUpColor.Color Or _
      m.frmTDGrid.DownColor <> gdDownColor.Color Or _
      m.frmTDGrid.BidColor <> gdBidColor.Color Or _
      m.frmTDGrid.AskColor <> gdAskColor.Color Or _
      m.frmTDGrid.FirstColor <> gdPriceLevelColor(0).Color Or _
      m.frmTDGrid.SecondColor <> gdPriceLevelColor(1).Color Or _
      m.frmTDGrid.ThirdColor <> gdPriceLevelColor(2).Color Or _
      m.frmTDGrid.FourthColor <> gdPriceLevelColor(3).Color Or _
      m.frmTDGrid.FifthColor <> gdPriceLevelColor(4).Color Or _
      m.frmTDGrid.OtherColor <> gdPriceLevelColor(5).Color Or _
      m.frmTDGrid.BidTextColor <> gdBidText.Color Or _
      m.frmTDGrid.AskTextColor <> gdAskText.Color Or _
      m.frmTDGrid.FixedPriceColor <> gdPriceText.Color Or _
      m.frmTDGrid.TickLineRightToLeft <> chkRightToLeft.Value Then

            bChanged = True

    End If
    
    ColorChanged = bChanged
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistributionCfg.ColorChanged", eGDRaiseError_Raise

End Function

Private Function IsNeutralScheme() As Boolean
On Error GoTo ErrSection:

    Dim bIsNeutral As Boolean
        
    If m.nUpColor = m.nBarColor And m.nDownColor = m.nBarColor Then
       bIsNeutral = True
    End If
    
    IsNeutralScheme = bIsNeutral

    Exit Function

ErrSection:
    RaiseError "frmTickDistributionCfg.IsNeutralScheme"
    
End Function

Private Sub optHighlightByPrice_Click()
    HighlightButtonsSet eBidAskColorMode_ByPrice
End Sub

Private Sub optHighlightBySize_Click()
On Error Resume Next
    HighlightButtonsSet eBidAskColorMode_BySize
End Sub

Private Sub optHighlightNone_Click()
On Error Resume Next
    HighlightButtonsSet eBidAskColorMode_None
End Sub

Private Sub optOneColor_Click()
On Error Resume Next
    
    If Not fraPriceNeutral.Visible Then
        fraPriceNeutral.Visible = True
        fraPriceColColors.Visible = False
    End If
    
    m.bNeutral = True
    
End Sub

Private Sub HighlightButtonsSet(eMode As eBidAskColorMode)
On Error GoTo ErrSection:

    Dim bEnable As Boolean
    
    Select Case eMode
        Case eBidAskColorMode_ByPrice
            optHighlightNone.Value = False
            optHighlightBySize.Value = False
            optHighlightByPrice.Value = True
        Case eBidAskColorMode_BySize
            optHighlightNone.Value = False
            optHighlightBySize.Value = True
            optHighlightByPrice.Value = False
        Case Default
            optHighlightNone.Value = True
            optHighlightBySize.Value = False
            optHighlightByPrice.Value = False
    End Select
    
    gdLargestSizeColor.Enabled = optHighlightBySize.Value
    
    'market depth color controls
    bEnable = optHighlightByPrice.Value
    fraMarketDepthColors.Enabled = bEnable
    
    lblBestBidAsk.Enabled = bEnable
    lblSecondBest.Enabled = bEnable
    lblThirdBest.Enabled = bEnable
    lblFourthBest.Enabled = bEnable
    lblFifthBest.Enabled = bEnable
    lblOther.Enabled = bEnable
    
    gdPriceLevelColor(0).Enabled = bEnable
    gdPriceLevelColor(1).Enabled = bEnable
    gdPriceLevelColor(2).Enabled = bEnable
    gdPriceLevelColor(3).Enabled = bEnable
    gdPriceLevelColor(4).Enabled = bEnable

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistributionCfg.HighlightButtonsSet"
    
End Sub

Private Function HighlightButtonsGet() As eBidAskColorMode
    
    If optHighlightBySize.Value = True Then
        HighlightButtonsGet = eBidAskColorMode_BySize
    ElseIf optHighlightByPrice.Value = True Then
        HighlightButtonsGet = eBidAskColorMode_ByPrice
    Else
        HighlightButtonsGet = eBidAskColorMode_None
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTickDistributionCfg.HighlightButtonsGet"
    
End Function

Private Sub ToggleConfigGrids(ByVal bOrdbarConfig As Boolean, Optional ByVal bAllOff As Boolean = False)
On Error GoTo ErrSection:
    
    If bAllOff Then
        lblCfgGriid.Visible = False
        fgAcctCols.Visible = False
        
        fgOrderButtons.Visible = False
        
        chkHighlightPos.Visible = False
        chkHighlightEquity.Visible = False
        gdHighlightPos.Visible = False
        gdHighlightEquity.Visible = False
    Else
        fgAcctCols.Visible = Not bOrdbarConfig
        
        fgOrderButtons.Visible = bOrdbarConfig
        
        chkHighlightPos.Visible = bOrdbarConfig
        chkHighlightEquity.Visible = bOrdbarConfig
        gdHighlightPos.Visible = bOrdbarConfig
        gdHighlightEquity.Visible = bOrdbarConfig
        
        If bOrdbarConfig Then
            lblCfgGriid.Caption = "Order bar buttons:"
        Else
            lblCfgGriid.Caption = "Account bar columns:"
        End If
        lblCfgGriid.Visible = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTickDistributionCfg.ToggleConfigGrids"
    
End Sub

Private Function VerifyQuantityPresets() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lPreset As Long                 ' Preset value
    
    bReturn = True
    
    lPreset = Int(Val(txtQty1.Text))
    If g.Broker.ValidQuantity(m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset) = False Then
        MoveFocus txtQty1
        InfBox "The first quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
        bReturn = False
    End If
    
    If bReturn = True Then
        lPreset = Int(Val(txtQty2.Text))
        If g.Broker.ValidQuantity(m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset) = False Then
            MoveFocus txtQty2
            InfBox "The second quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
            bReturn = False
        End If
    End If
    
    If bReturn = True Then
        lPreset = Int(Val(txtQty3.Text))
        If g.Broker.ValidQuantity(m.frmTDGrid.TradeAccountID, m.frmTDGrid.SymbolOrSymbolID, lPreset) = False Then
            MoveFocus txtQty3
            InfBox "The third quantity preset is an invalid quantity", "!", , "Quantity Preset Error"
            bReturn = False
        End If
    End If
    
    VerifyQuantityPresets = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTickDistributionCfg.VerifyQuantityPresets"
    
End Function

Private Sub txtQty1_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.txtQty1_GotFocus"
    
End Sub

Private Sub txtQty2_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty2

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.txtQty2_GotFocus"
    
End Sub

Private Sub txtQty3_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtQty3

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTickDistributionCfg.txtQty3_GotFocus"
    
End Sub


