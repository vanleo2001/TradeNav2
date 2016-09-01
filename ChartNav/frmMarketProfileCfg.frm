VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmMarketProfileCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trade Profile Settings"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   5490
      Left            =   105
      TabIndex        =   2
      Top             =   90
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9684
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
      Caption         =   "Data|TPOs|Statistics|Steidlmayer"
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
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin HexUniControls.ctlUniFrameWL Frame2 
         Height          =   5115
         Left            =   7980
         TabIndex        =   49
         Top             =   330
         Width           =   6645
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMarketProfileCfg.frx":0000
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketProfileCfg.frx":002C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketProfileCfg.frx":004C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkSMPProfile 
            Height          =   255
            Left            =   720
            TabIndex        =   16
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
            Caption         =   "frmMarketProfileCfg.frx":0068
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketProfileCfg.frx":00AC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":00CC
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL Frame3 
            Height          =   1695
            Left            =   480
            TabIndex        =   51
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
            Caption         =   "frmMarketProfileCfg.frx":00E8
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketProfileCfg.frx":0108
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":0128
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtSMPMin 
               Height          =   285
               Left            =   3120
               TabIndex        =   53
               Top             =   360
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":0144
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
               Tip             =   "frmMarketProfileCfg.frx":0168
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0188
            End
            Begin HexUniControls.ctlUniTextBoxXP txtSMPBreak 
               Height          =   285
               Left            =   3120
               TabIndex        =   52
               Top             =   720
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":01A4
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
               Tip             =   "frmMarketProfileCfg.frx":01C6
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":01E6
            End
            Begin gdOCX.gdSelectColor gdSMPTPO 
               Height          =   315
               Left            =   3600
               TabIndex        =   56
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniLabelXP Label17 
               Height          =   315
               Left            =   360
               Top             =   1230
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
               Caption         =   "frmMarketProfileCfg.frx":0202
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":0280
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":02A0
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label18 
               Height          =   255
               Left            =   1440
               Top             =   390
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
               Caption         =   "frmMarketProfileCfg.frx":02BC
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":02FC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":031C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label19 
               Height          =   255
               Left            =   1440
               Top             =   750
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
               Caption         =   "frmMarketProfileCfg.frx":0338
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":037A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":039A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniCheckXP chkVolIterator 
            Height          =   255
            Left            =   720
            TabIndex        =   50
            Top             =   2760
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
            Caption         =   "frmMarketProfileCfg.frx":03B6
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketProfileCfg.frx":03F4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":0414
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL Frame4 
            Height          =   1815
            Left            =   480
            TabIndex        =   17
            Top             =   2760
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
            Caption         =   "frmMarketProfileCfg.frx":0430
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketProfileCfg.frx":0450
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":0470
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtIteratorHighVol 
               Height          =   285
               Left            =   4320
               TabIndex        =   18
               Top             =   1350
               Width           =   495
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":048C
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
               Tip             =   "frmMarketProfileCfg.frx":04B0
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":04D0
            End
            Begin HexUniControls.ctlUniTextBoxXP txtIteratorLookback 
               Height          =   285
               Left            =   1860
               TabIndex        =   19
               Top             =   1350
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":04EC
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
               Tip             =   "frmMarketProfileCfg.frx":0512
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0532
            End
            Begin HexUniControls.ctlUniLabelXP Label12 
               Height          =   615
               Left            =   480
               Top             =   360
               Width           =   4455
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":054E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":068A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":06AA
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label20 
               Height          =   255
               Left            =   2760
               Top             =   1365
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
               Caption         =   "frmMarketProfileCfg.frx":06C6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":070E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":072E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label16 
               Height          =   255
               Left            =   480
               Top             =   1365
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
               Caption         =   "frmMarketProfileCfg.frx":074A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":078C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":07AC
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraStats 
         Height          =   5115
         Left            =   7680
         TabIndex        =   8
         Top             =   330
         Width           =   6645
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMarketProfileCfg.frx":07C8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketProfileCfg.frx":07F8
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketProfileCfg.frx":0818
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectColor clrItemColor 
            Height          =   315
            Left            =   5880
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniFrameWL Frame1 
            Height          =   1455
            Left            =   360
            TabIndex        =   36
            Top             =   2760
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
            Caption         =   "frmMarketProfileCfg.frx":0834
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketProfileCfg.frx":0870
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":0890
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniCheckXP chkShadeVol 
               Height          =   375
               Left            =   720
               TabIndex        =   40
               Top             =   960
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
               Caption         =   "frmMarketProfileCfg.frx":08AC
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":08FC
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":091C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkVolumePercent 
               Height          =   330
               Left            =   720
               TabIndex        =   38
               Top             =   240
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
               Caption         =   "frmMarketProfileCfg.frx":0938
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":097E
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":099E
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkVolumeActual 
               Height          =   330
               Left            =   720
               TabIndex        =   37
               Top             =   600
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
               Caption         =   "frmMarketProfileCfg.frx":09BA
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":09FE
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0A1E
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectColor gdVolumeColor 
               Height          =   315
               Left            =   3360
               TabIndex        =   39
               Top             =   990
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               CustomColor     =   255
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid fgStats 
            Height          =   1755
            Left            =   960
            TabIndex        =   22
            Top             =   720
            Width           =   4815
            _cx             =   8493
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
      End
      Begin HexUniControls.ctlUniFrameWL fraCharSequence 
         Height          =   5115
         Left            =   7380
         TabIndex        =   7
         Top             =   330
         Width           =   6645
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMarketProfileCfg.frx":0A3A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketProfileCfg.frx":0A5A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketProfileCfg.frx":0A7A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraTPO 
            Height          =   4935
            Left            =   240
            TabIndex        =   23
            Top             =   120
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
            Caption         =   "frmMarketProfileCfg.frx":0A96
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketProfileCfg.frx":0AB6
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":0AD6
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtExtraRows 
               Height          =   285
               Left            =   2280
               TabIndex        =   20
               Top             =   4545
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":0AF2
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
               Tip             =   "frmMarketProfileCfg.frx":0B1C
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0B3C
            End
            Begin HexUniControls.ctlUniTextBoxXP txtMarginLeft 
               Height          =   285
               Left            =   2280
               TabIndex        =   21
               Top             =   4200
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":0B58
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
               Tip             =   "frmMarketProfileCfg.frx":0B82
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0BA2
            End
            Begin HexUniControls.ctlUniCheckXP chkGridLines 
               Height          =   255
               Left            =   3360
               TabIndex        =   27
               Top             =   4560
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
               Caption         =   "frmMarketProfileCfg.frx":0BBE
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":0BFC
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0C1C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniFrameWL Frame5 
               Height          =   1740
               Left            =   240
               TabIndex        =   28
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
               Caption         =   "frmMarketProfileCfg.frx":0C38
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmMarketProfileCfg.frx":0C58
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":0C78
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniRadioXP optSeqEight 
                  Height          =   375
                  Left            =   2880
                  TabIndex        =   29
                  Top             =   840
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
                  Caption         =   "frmMarketProfileCfg.frx":0C94
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0CDE
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":0CFE
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optSeqSix 
                  Height          =   285
                  Left            =   3810
                  TabIndex        =   31
                  Top             =   540
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
                  Caption         =   "frmMarketProfileCfg.frx":0D1A
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0D54
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":0D74
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniCheckXP chkShowCountTPO 
                  Height          =   255
                  Left            =   2880
                  TabIndex        =   33
                  Top             =   1380
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
                  Caption         =   "frmMarketProfileCfg.frx":0D90
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0DCC
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":0DEC
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniCheckXP chkBoxFirst 
                  Height          =   255
                  Left            =   255
                  TabIndex        =   35
                  Top             =   1380
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
                  Caption         =   "frmMarketProfileCfg.frx":0E08
                  Enabled         =   -1  'True
                  Align           =   0
                  CheckBackColor  =   -2147483643
                  CheckForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0E5C
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":0E7C
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optSeqSeven 
                  Height          =   375
                  Left            =   240
                  TabIndex        =   44
                  Top             =   840
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
                  Caption         =   "frmMarketProfileCfg.frx":0E98
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0EF4
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":0F14
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optSeqOne 
                  Height          =   285
                  Left            =   240
                  TabIndex        =   45
                  Top             =   240
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
                  Caption         =   "frmMarketProfileCfg.frx":0F30
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0F6C
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":0F8C
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP OptSeqTwo 
                  Height          =   285
                  Left            =   2025
                  TabIndex        =   46
                  Top             =   255
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
                  Caption         =   "frmMarketProfileCfg.frx":0FA8
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":0FE2
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":1002
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optSeqThree 
                  Height          =   285
                  Left            =   3810
                  TabIndex        =   47
                  Top             =   255
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
                  Caption         =   "frmMarketProfileCfg.frx":101E
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":1058
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":1078
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optSeqFour 
                  Height          =   285
                  Left            =   240
                  TabIndex        =   48
                  Top             =   540
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
                  Caption         =   "frmMarketProfileCfg.frx":1094
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":10CE
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":10EE
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniRadioXP optSeqFive 
                  Height          =   285
                  Left            =   2025
                  TabIndex        =   54
                  Top             =   540
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
                  Caption         =   "frmMarketProfileCfg.frx":110A
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":1144
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":1164
                  ShowFocus       =   -1  'True
                  RightToLeft     =   0   'False
               End
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdFont 
               Height          =   300
               Left            =   4380
               TabIndex        =   30
               Top             =   195
               Width           =   810
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
               Caption         =   "frmMarketProfileCfg.frx":1180
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":11A8
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":11C8
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniFrameWL fraGradientColor 
               Height          =   1695
               Left            =   240
               TabIndex        =   24
               Top             =   2400
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
               Caption         =   "frmMarketProfileCfg.frx":11E4
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmMarketProfileCfg.frx":1204
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":1224
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniComboImageXP cboColorScheme 
                  Height          =   315
                  Left            =   720
                  TabIndex        =   42
                  Top             =   480
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
                  Tip             =   "frmMarketProfileCfg.frx":1240
                  Sorted          =   0   'False
                  HScroll         =   0   'False
                  RoundedBorders  =   -1  'True
                  IconDim         =   16
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":1260
                  DropDownOnTextClick=   -1  'True
                  DropDownWidth   =   -1
                  RightToLeft     =   0   'False
               End
               Begin gdOCX.gdSelectColor gdColor1 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   25
                  Top             =   480
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin gdOCX.gdSelectColor gdOtherText 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   32
                  Top             =   840
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin gdOCX.gdSelectColor gdBackground 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   34
                  Top             =   1200
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin gdOCX.gdSelectColor gdColor2 
                  Height          =   315
                  Left            =   3705
                  TabIndex        =   26
                  Top             =   480
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   556
                  CustomColor     =   255
               End
               Begin HexUniControls.ctlUniLabelXP lblBackgroundColor 
                  Height          =   315
                  Left            =   720
                  Top             =   1230
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
                  Caption         =   "frmMarketProfileCfg.frx":127C
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   1
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":12BC
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":12DC
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblOtherColor 
                  Height          =   315
                  Left            =   720
                  Top             =   870
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
                  Caption         =   "frmMarketProfileCfg.frx":12F8
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   1
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":1338
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":1358
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblTPOColors 
                  Height          =   255
                  Left            =   720
                  Top             =   240
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
                  Caption         =   "frmMarketProfileCfg.frx":1374
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":13AA
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":13CA
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblColorLabel2 
                  Height          =   255
                  Left            =   3960
                  Top             =   210
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "frmMarketProfileCfg.frx":13E6
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":140C
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":142C
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblColorLabel1 
                  Height          =   255
                  Left            =   2760
                  Top             =   210
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "frmMarketProfileCfg.frx":1448
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   1
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmMarketProfileCfg.frx":1472
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmMarketProfileCfg.frx":1492
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin VB.Image imgGradientHorz 
                  Height          =   480
                  Left            =   3360
                  Picture         =   "frmMarketProfileCfg.frx":14AE
                  Top             =   210
                  Width           =   480
               End
            End
            Begin HexUniControls.ctlUniLabelXP Label3 
               Height          =   255
               Left            =   360
               Top             =   4560
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
               Caption         =   "frmMarketProfileCfg.frx":17B8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":1806
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":1826
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label4 
               Height          =   255
               Left            =   360
               Top             =   4200
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
               Caption         =   "frmMarketProfileCfg.frx":1842
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":188A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":18AA
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblCharSequence 
               Height          =   465
               Left            =   1020
               Top             =   120
               Width           =   3675
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":18C6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":1994
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":19B4
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin VB.Image imgGradientVert 
            Enabled         =   0   'False
            Height          =   600
            Left            =   6360
            Picture         =   "frmMarketProfileCfg.frx":19D0
            Top             =   240
            Visible         =   0   'False
            Width           =   225
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSession 
         Height          =   5115
         Left            =   45
         TabIndex        =   3
         Top             =   330
         Width           =   6645
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMarketProfileCfg.frx":213A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmMarketProfileCfg.frx":215A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMarketProfileCfg.frx":217A
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectDate gdSessionDate 
            Height          =   315
            Left            =   3120
            TabIndex        =   6
            Top             =   780
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            Enabled         =   0   'False
            AllowWeekends   =   0   'False
            MaxDate         =   42611
            MaxDateIsToday  =   -1  'True
            Value           =   37015
         End
         Begin HexUniControls.ctlUniFrameWL fraProfile 
            Height          =   3615
            Left            =   240
            TabIndex        =   9
            Top             =   1200
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
            Caption         =   "frmMarketProfileCfg.frx":2196
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmMarketProfileCfg.frx":21B6
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":21D6
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniComboBoxXP cboTicksPerRow 
               Height          =   315
               Left            =   2640
               TabIndex        =   55
               Top             =   1208
               Width           =   1080
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
               Tip             =   "frmMarketProfileCfg.frx":21F2
               Sorted          =   0   'False
               HScroll         =   0   'False
               Style           =   2
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonWidth     =   17
               Locked          =   0   'False
               SelBackColor    =   -2147483635
               SelForeColor    =   -2147483634
               TrapTab         =   0   'False
               ButtonStyle     =   -1
               SelectorStyle   =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2212
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               ManualStart     =   0   'False
               MaxLength       =   0
               RightToLeft     =   0   'False
               LeftMargin      =   0
               RightMargin     =   0
               SelectOnFocus   =   0   'False
            End
            Begin HexUniControls.ctlUniTextBoxXP txtDaysBack 
               Height          =   315
               Left            =   4395
               TabIndex        =   43
               Top             =   2625
               Width           =   870
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmMarketProfileCfg.frx":222E
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
               Tip             =   "frmMarketProfileCfg.frx":2264
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2284
            End
            Begin HexUniControls.ctlUniComboBoxXP cboIntervalDays 
               Height          =   315
               Left            =   4050
               TabIndex        =   10
               Top             =   630
               Width           =   1320
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
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
               Tip             =   "frmMarketProfileCfg.frx":22A0
               Sorted          =   0   'False
               HScroll         =   0   'False
               Style           =   2
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonWidth     =   17
               Locked          =   0   'False
               SelBackColor    =   -2147483635
               SelForeColor    =   -2147483634
               TrapTab         =   0   'False
               ButtonStyle     =   -1
               SelectorStyle   =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":22C0
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               ManualStart     =   0   'False
               MaxLength       =   0
               RightToLeft     =   0   'False
               LeftMargin      =   0
               RightMargin     =   0
               SelectOnFocus   =   0   'False
            End
            Begin HexUniControls.ctlUniComboBoxXP cboIntervalMinutes 
               Height          =   315
               Left            =   1815
               TabIndex        =   11
               Top             =   615
               Width           =   1320
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
               Tip             =   "frmMarketProfileCfg.frx":22DC
               Sorted          =   0   'False
               HScroll         =   0   'False
               Style           =   2
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonWidth     =   17
               Locked          =   0   'False
               SelBackColor    =   -2147483635
               SelForeColor    =   -2147483634
               TrapTab         =   0   'False
               ButtonStyle     =   -1
               SelectorStyle   =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":22FC
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               ManualStart     =   0   'False
               MaxLength       =   0
               RightToLeft     =   0   'False
               LeftMargin      =   0
               RightMargin     =   0
               SelectOnFocus   =   0   'False
            End
            Begin HexUniControls.ctlUniComboBoxXP cboDaysProfile 
               Height          =   315
               Left            =   4395
               TabIndex        =   15
               Top             =   3015
               Visible         =   0   'False
               Width           =   870
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
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
               Tip             =   "frmMarketProfileCfg.frx":2318
               Sorted          =   0   'False
               HScroll         =   0   'False
               Style           =   2
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonWidth     =   17
               Locked          =   0   'False
               SelBackColor    =   -2147483635
               SelForeColor    =   -2147483634
               TrapTab         =   0   'False
               ButtonStyle     =   -1
               SelectorStyle   =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2338
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               ManualStart     =   0   'False
               MaxLength       =   0
               RightToLeft     =   0   'False
               LeftMargin      =   0
               RightMargin     =   0
               SelectOnFocus   =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdTime 
               Height          =   315
               Left            =   4050
               TabIndex        =   14
               Top             =   1935
               Width           =   1320
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
               Caption         =   "frmMarketProfileCfg.frx":2354
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":238A
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":23AA
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optIntervalMinutes 
               Height          =   225
               Left            =   840
               TabIndex        =   13
               Top             =   645
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
               Caption         =   "frmMarketProfileCfg.frx":23C6
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":23F4
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2414
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optIntervalDays 
               Height          =   225
               Left            =   3315
               TabIndex        =   12
               Top             =   675
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
               Caption         =   "frmMarketProfileCfg.frx":2430
               Enabled         =   0   'False
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":2458
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2478
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label1 
               Height          =   210
               Left            =   1980
               Top             =   2670
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
               Caption         =   "frmMarketProfileCfg.frx":2494
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":24EA
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":250A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label5 
               Height          =   210
               Left            =   840
               Top             =   1260
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
               Caption         =   "frmMarketProfileCfg.frx":2526
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":256E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":258E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label11 
               Height          =   210
               Left            =   3960
               Top             =   1260
               Width           =   1680
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":25AA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":25E4
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2604
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label2 
               Height          =   210
               Left            =   840
               Top             =   360
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
               Caption         =   "frmMarketProfileCfg.frx":2620
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":2662
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2682
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label10 
               Height          =   210
               Left            =   840
               Top             =   2670
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
               Caption         =   "frmMarketProfileCfg.frx":269E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":26D8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":26F8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label6 
               Height          =   210
               Left            =   2070
               Top             =   1890
               Width           =   1080
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":2714
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":2748
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2768
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label8 
               Height          =   210
               Left            =   2190
               Top             =   3075
               Visible         =   0   'False
               Width           =   2040
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":2784
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":27D8
               Style           =   0
               Enabled         =   0   'False
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":27F8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label7 
               Height          =   210
               Left            =   1110
               Top             =   2145
               Width           =   2040
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":2814
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":2846
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2866
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTimeStart 
               Height          =   210
               Left            =   3315
               Top             =   1890
               Width           =   1470
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":2882
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":28AC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":28CC
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblTimeEnd 
               Height          =   210
               Left            =   3315
               Top             =   2145
               Width           =   1470
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmMarketProfileCfg.frx":28E8
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":2912
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":2932
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label9 
               Height          =   210
               Left            =   840
               Top             =   1890
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
               Caption         =   "frmMarketProfileCfg.frx":294E
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmMarketProfileCfg.frx":298A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmMarketProfileCfg.frx":29AA
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniRadioXP optDate 
            Height          =   220
            Left            =   1320
            TabIndex        =   5
            Top             =   840
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
            Caption         =   "frmMarketProfileCfg.frx":29C6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketProfileCfg.frx":29FC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":2A1C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optCurrentSession 
            Height          =   220
            Left            =   1320
            TabIndex        =   4
            Top             =   480
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
            Caption         =   "frmMarketProfileCfg.frx":2A38
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmMarketProfileCfg.frx":2A7E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmMarketProfileCfg.frx":2A9E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   375
      Left            =   3525
      TabIndex        =   1
      Top             =   5640
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
      Caption         =   "frmMarketProfileCfg.frx":2ABA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMarketProfileCfg.frx":2AE6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMarketProfileCfg.frx":2B06
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   375
      Left            =   2430
      TabIndex        =   0
      Top             =   5640
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
      Caption         =   "frmMarketProfileCfg.frx":2B22
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMarketProfileCfg.frx":2B46
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMarketProfileCfg.frx":2B66
      RightToLeft     =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMarketProfileCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumStatsCol
    eStatsCol_Show = 0
    eStatsCol_Item
    eStatsCol_Color
    eStatsCol_PenSize
    eStatsCol_VAPercent
End Enum

Private Enum enumStatsRow
    eStatsRow_POCVol = 1
    eStatsRow_POCTPO
    eStatsRow_VAVol
    eStatsRow_VATPO
    eStatsRow_StdDev
    eStatsRow_Mean
End Enum

Private Type mPrivate
    Bars As cGdBars
    frm As frmMarketProfile
    eSelSequence As MktProfile_Char_Sequence
    
    nMouseColDown As Long
    nMouseRowDown As Long
End Type
Private m As mPrivate

Public Sub ShowMe(frm As frmMarketProfile, Bars As cGdBars, ByVal bSMPEnabled As Boolean)

    Set m.Bars = Bars
    Set m.frm = frm
    
    If m.frm Is Nothing Then Exit Sub
    
    InitDataControls
    InitDisplayControls
    
    vsIndexTab1.TabVisible(3) = bSMPEnabled
    If frm.IsSMPProfile Then
        chkSMPProfile.Value = vbChecked
        vsIndexTab1.TabEnabled(2) = False
        EnableGradientControls False
    End If
    
    ' center form horizontally & somewhat close to the top of trade profile form
    Me.Move m.frm.Left + (m.frm.Width - Me.ScaleWidth) / 2, m.frm.Top + (m.frm.Height - Me.ScaleHeight) / 4
    
    If m.Bars Is Nothing Then
        cmdTime.Enabled = False
        cmdCancel.Enabled = False
        vsIndexTab1.CurrTab = 1             'first time used, make user choose a character sequence
    ElseIf bSMPEnabled Then
        vsIndexTab1.CurrTab = 3
    Else
        vsIndexTab1.CurrTab = 0
    End If
    
    ShowForm Me, eForm_Modal
    
End Sub

Private Sub InitDisplayControls()
    
    Dim eVolText As MktProfile_Vol_Text
        
    gdVolumeColor.Color = m.frm.ProfileProperty(eProfileProp_VolShadeColor)
    gdOtherText.Color = m.frm.ProfileProperty(eProfileProp_OtherTextColor)
    gdBackground.Color = m.frm.ProfileProperty(eProfileProp_BackgroundColor)
    
    gdSMPTPO.Color = m.frm.ProfileProperty(eProfileProp_SMPColor)
    
    If m.frm.ProfileProperty(eProfileProp_ColorScheme) = MktProf_Color_VolIterator Then
        chkVolIterator.Value = vbChecked
    Else
        chkVolIterator.Value = vbUnchecked
    End If
    
    chkGridLines.Value = m.frm.ProfileProperty(eProfileProp_GridLines)
    
    txtMarginLeft.Text = Str(m.frm.ProfileProperty(eProfileProp_MarginLeft))
    txtExtraRows.Text = Str(m.frm.ProfileProperty(eProfileProp_ExtraRows))
    chkBoxFirst.Value = m.frm.ProfileProperty(eProfileProp_BoxFirst)
    
    chkVolumePercent.Value = m.frm.ProfileProperty(eProfileProp_VolPercentShow)
    chkVolumeActual.Value = m.frm.ProfileProperty(eProfileProp_VolActualShow)
    chkShadeVol.Value = m.frm.ProfileProperty(eProfileProp_VolShade)
    chkShowCountTPO.Value = m.frm.ProfileProperty(eProfileProp_TPOCountShow)
    
    txtSMPMin.Text = Str(m.frm.ProfileProperty(eProfileProp_SMPMinTPO))
    txtSMPBreak.Text = Str(m.frm.ProfileProperty(eProfileProp_SMPBreakOut))
    
    txtIteratorLookback.Text = Str(m.frm.ProfileProperty(eProfileProp_IteratorLookback))
    txtIteratorHighVol.Text = Str(m.frm.ProfileProperty(eProfileProp_IteratorHighVol))
    
    InitGridStats
    InitCharControls
    InitFont

End Sub

Private Sub InitDataControls()

    If m.frm Is Nothing Then Exit Sub
    
    'profile interval
    cboIntervalMinutes.Clear
    cboIntervalMinutes.AddItem "120"
    cboIntervalMinutes.AddItem "60"
    cboIntervalMinutes.AddItem "30"
    cboIntervalMinutes.AddItem "15"
    cboIntervalMinutes.AddItem "10"
    cboIntervalMinutes.AddItem "5"
    
    cboIntervalDays.Clear
    cboIntervalDays.AddItem "1"
    cboIntervalDays.AddItem "2"
    cboIntervalDays.AddItem "3"
    cboIntervalDays.AddItem "4"
    cboIntervalDays.AddItem "5"
    
    cboDaysProfile.Clear
    cboDaysProfile.AddItem "1"
    cboDaysProfile.AddItem "2"
    cboDaysProfile.AddItem "3"
    cboDaysProfile.AddItem "4"
    cboDaysProfile.AddItem "5"
    cboDaysProfile.ListIndex = m.frm.DaysProfile - 1
    
    SetIntervalCtrls
        
    'profile days
    txtDaysBack.Text = Str(m.frm.DaysBack)
    
    If m.frm.SessionDate = 0 Then
        gdSessionDate.Value = Now
        optCurrentSession.Value = True
    Else
        gdSessionDate.Value = m.frm.SessionDate
        optDate.Value = True
    End If
            
    lblTimeStart.Caption = DateFormat(m.frm.TimeStart, NO_DATE, HH_MM)
    lblTimeEnd.Caption = DateFormat(m.frm.TimeEnd, NO_DATE, HH_MM)
    
    'yscale min move (ie ticks per row)
    cboTicksPerRow.Clear
    cboTicksPerRow.AddItem "Auto"
    cboTicksPerRow.AddItem "1"
    cboTicksPerRow.AddItem "2"
    cboTicksPerRow.AddItem "3"
    cboTicksPerRow.AddItem "4"
    cboTicksPerRow.AddItem "5"
    cboTicksPerRow.AddItem "6"
    cboTicksPerRow.AddItem "10"
    cboTicksPerRow.AddItem "15"
    cboTicksPerRow.AddItem "20"
    cboTicksPerRow.AddItem "25"
    cboTicksPerRow.AddItem "50"
    cboTicksPerRow.AddItem "75"
    cboTicksPerRow.AddItem "100"
    cboTicksPerRow.AddItem "200"
    
    If Str(m.frm.YScaleMinMove) <= 0 Then
        cboTicksPerRow.ListIndex = 0
    Else
        cboTicksPerRow.Text = Str(m.frm.YScaleMinMove)
    End If

End Sub

Private Sub InitCharControls()

    SetCharCtrls
    
    cboColorScheme.Clear
    cboColorScheme.AddItem "Gradient"
    cboColorScheme.AddItem "Rainbow"
    cboColorScheme.AddItem "Up/Down"
    cboColorScheme.AddItem "Buy/Sell Volume"
    
    If m.frm.ProfileProperty(eProfileProp_ColorScheme) = MktProf_Color_VolIterator Then
        cboColorScheme.ListIndex = 1
        chkVolIterator.Value = 1
    Else
        cboColorScheme.ListIndex = m.frm.ProfileProperty(eProfileProp_ColorScheme)
        chkVolIterator.Value = 0
    End If

End Sub

Private Sub InitFont()

    If m.frm Is Nothing Then Exit Sub
    
    Dim strName As String
    Dim nSize&, nBold&, nItalic&
    
    strName = m.frm.GridFontName
    nSize = m.frm.GridFontSize
    nBold = m.frm.GridFontBold
    nItalic = m.frm.GridFontItalic
    
    Me.Font.Name = strName
    Me.Font.Size = nSize
    Me.Font.Bold = (-1) * nBold
    Me.Font.Italic = (-1) * nItalic
    
End Sub

Private Sub cboColorScheme_Click()

    With cboColorScheme
        If .ListIndex = MktProf_Color_Gradient Then
            imgGradientHorz.Visible = True
        Else
            imgGradientHorz.Visible = False
        End If
        
        If .ListIndex = MktProf_Color_Rainbow Or .ListIndex = MktProf_Color_VolIterator Then
            lblColorLabel1.Visible = False
            lblColorLabel2.Visible = False
            gdColor1.Visible = False
            gdColor2.Visible = False
        Else
            lblColorLabel1.Visible = True
            lblColorLabel2.Visible = True
            gdColor1.Visible = True
            gdColor2.Visible = True
        End If
        
        Select Case .ListIndex
            
            Case MktProf_Color_Gradient:
                lblColorLabel1.Caption = "From"
                lblColorLabel2.Caption = "To"
                gdColor1.Color = m.frm.ProfileProperty(eProfileProp_GradientFrom)
                gdColor2.Color = m.frm.ProfileProperty(eProfileProp_GradientTo)
                
            Case MktProf_Color_OpenClose:
                lblColorLabel1.Caption = "Up"
                lblColorLabel2.Caption = "Down"
                gdColor1.Color = m.frm.ProfileProperty(eProfileProp_UpColor)
                gdColor2.Color = m.frm.ProfileProperty(eProfileProp_DownColor)
            
            Case MktProf_Color_BidAsk:
                lblColorLabel1.Caption = "Buy"
                lblColorLabel2.Caption = "Sell"
                gdColor1.Color = m.frm.ProfileProperty(eProfileProp_AskColor)
                gdColor2.Color = m.frm.ProfileProperty(eProfileProp_BidColor)
        
        End Select
    End With

End Sub

Private Sub chkSMPProfile_Click()
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    
    If chkSMPProfile.Value = vbChecked Then
        If vsIndexTab1.TabEnabled(2) Then vsIndexTab1.TabEnabled(2) = False
        EnableGradientControls False
    Else
        If Not vsIndexTab1.TabEnabled(2) Then
            vsIndexTab1.TabEnabled(2) = True
        End If
        EnableGradientControls True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfileCfg.chkSMPProfile_Click"

End Sub

'Private Sub cboIntervalMinutes_Click()
'On Error GoTo ErrSection:
'
'    If Not Me.Visible Then Exit Sub
'
'    If optSeqEight Then
'        If cboIntervalMinutes.Text <> "30" Then
'            If InfBox("Classic lettering requires 30min data interval. Change TPO character sequence?", "?", "Ok|Cancel", "Trade Profile") = "O" Then
'                optSeqOne.Value = True
'            Else
'                cboIntervalMinutes.ListIndex = 2
'            End If
'        End If
'    End If
'
'ErrExit:
'    Exit Sub
'
'ErrSection:
'    RaiseError "frmMarketProfileCfg.cboIntervalMinutes_Change"
'
'End Sub

Private Sub clrItemColor_Changed()

    Dim nColor As Long
    
    clrItemColor.Visible = False
    nColor = clrItemColor.Color
    If nColor = 0 Then nColor = -1  '0 is reserved color in flex grid control
    
    With fgStats
        If .Row >= .FixedRows And .Row < .Rows Then
            If .Col = eStatsCol_Color Then
                .Cell(flexcpBackColor, .Row, .Col) = nColor
            End If
        End If
    End With
    
End Sub

Private Sub clrItemColor_ColorClicked()
    clrItemColor.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click()
    CommonDialogFont Me.CommonDialog1, Me.Font
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrExit

    Dim strErr$, iDaysBack&, iInterval&, iDate&, iType&, iFirstDate&
    Dim strFontName$, nFontSize&, nBold&, nItalic&, nTicksPerRow&
    Dim NewBars As cGdBars
    Dim bSMPProfile As Boolean
    
'    If optSeqEight Then
'        If optIntervalDays Or Int(cboIntervalMinutes.Text) <> 30 Then
'            If InfBox("Data interval will be changed to 30 minutes for Classic Lettering.", "?", "Ok|Cancel", "TradeProfile") = "O" Then
'                If optIntervalDays Then optIntervalMinutes_Click
'                cboIntervalMinutes.ListIndex = 2
'                m.eSelSequence = MktProf_CHR_Classic
'            Else
'                SetIntervalCtrls
'                SetCharCtrls
'            End If
'        End If
'    End If
    
    If optIntervalMinutes.Value = True Then
        strErr = "Minutes interval must be between 5 and 1440."
        If IsAlpha(cboIntervalMinutes.Text) Then                '6538
            GoTo ErrExit
        Else
            iInterval = ValOfText(cboIntervalMinutes.Text)
            If iInterval < 5 Or iInterval > 1440 Then GoTo ErrExit
        End If
    Else
        iInterval = ValOfText(cboIntervalDays.Text)     'not yet implemented
    End If
    
    strErr = "Number of days back must be between 0 and 99"
    If IsAlpha(txtDaysBack.Text) Then GoTo ErrExit
    
    iDaysBack = ValOfText(txtDaysBack.Text)
    If iDaysBack < 0 Or iDaysBack > 99 Then GoTo ErrExit
    
    If optIntervalDays.Value = True Then iType = 1
    
    strErr = "Left margin must be a numeric value."
    If IsAlpha(txtMarginLeft) Then GoTo ErrExit
    
    strErr = "Extra rows must be a numeric value."
    If IsAlpha(txtExtraRows) Then GoTo ErrExit
    
    strErr = "Ticks per row must be an integer greater than zero."
    If cboTicksPerRow.ListIndex = 0 Then
        nTicksPerRow = -1
    Else
        nTicksPerRow = Int(ValOfText(cboTicksPerRow.Text))
        If nTicksPerRow <= 0 Then GoTo ErrExit
    End If
    
    strErr = ""
    If optDate.Value = True Then
        iDate = gdSessionDate.Value
        iFirstDate = g.SymbolPool.TickFirstDate(m.Bars.Prop(eBARS_SymbolID))
        If iFirstDate = kNullData Then
            strErr = "Unable to obtain Tick First Date for " & m.Bars.Prop(eBARS_Symbol) & " (" & Str(m.Bars.Prop(eBARS_SymbolID)) & ")"
        ElseIf iDate < iFirstDate Then
            strErr = "Invalid date: earliest available data is on " & DateFormat(iFirstDate)
        Else
            Set NewBars = New cGdBars
            NewBars.Prop(eBARS_Periodicity) = ePRD_EachTick
            
            GetAvailTickData NewBars, 0, m.Bars.Prop(eBARS_Symbol), m.Bars.Prop(eBARS_SymbolID), iDate, -1
            
            If NewBars.Size = 0 Then
                strErr = "Tick data not available for " & DateFormat(iDate)
            ElseIf NewBars(eBARS_Vol, NewBars.Size - 1) = kNullData Then
                strErr = "Volume data not available for " & DateFormat(iDate)
            End If
            
            Set NewBars = Nothing
        End If
    End If
    
    If vsIndexTab1.TabVisible(3) Then
        If Int(ValOfText(txtSMPBreak.Text)) >= Int(ValOfText(txtSMPMin.Text)) Then
            strErr = "Breakout periods must be less than minimum periods"
        End If
    End If
    
    If Len(strErr) > 0 Then GoTo ErrExit
        
    m.frm.ProfileInterval = iInterval
    If optDate.Value = True Then
        m.frm.SessionDate = iDate
    Else
        m.frm.SessionDate = 0
    End If
    m.frm.DaysBack = iDaysBack
    
    m.frm.ProfileProperty(eProfileProp_GridLines) = chkGridLines.Value
    m.frm.ProfileProperty(eProfileProp_BackgroundColor) = gdBackground.Color
    m.frm.ProfileProperty(eProfileProp_OtherTextColor) = gdOtherText.Color
    m.frm.ProfileProperty(eProfileProp_ExtraRows) = ValOfText(txtExtraRows.Text)
    m.frm.ProfileProperty(eProfileProp_MarginLeft) = ValOfText(txtMarginLeft.Text)
    
    m.frm.ProfileProperty(eProfileProp_IteratorLookback) = Int(ValOfText(txtIteratorLookback.Text))
    m.frm.ProfileProperty(eProfileProp_IteratorHighVol) = Int(ValOfText(txtIteratorHighVol.Text))
    
    m.frm.ProfileProperty(eProfileProp_SMPMinTPO) = Int(ValOfText(txtSMPMin.Text))
    m.frm.ProfileProperty(eProfileProp_SMPBreakOut) = Int(ValOfText(txtSMPBreak.Text))
    
    bSMPProfile = chkSMPProfile.Value
    If bSMPProfile Then
        If optSeqEight Then
            m.frm.CharacterSequence = MktProf_CHR_Classic
        Else
            m.frm.CharacterSequence = MktProf_CHR_SeqOne
        End If
        
        m.frm.ProfileProperty(eProfileProp_SMPColor) = gdSMPTPO.Color
        
        If chkVolIterator.Value = vbChecked Then
            m.frm.ProfileProperty(eProfileProp_ColorScheme) = MktProf_Color_VolIterator
        Else
            m.frm.ProfileProperty(eProfileProp_ColorScheme) = MktProf_Color_Rainbow
        End If
    Else
        If chkVolIterator.Value = vbChecked Then
            m.frm.ProfileProperty(eProfileProp_ColorScheme) = MktProf_Color_VolIterator
        Else
            m.frm.ProfileProperty(eProfileProp_ColorScheme) = cboColorScheme.ListIndex
            Select Case cboColorScheme.ListIndex
                Case MktProf_Color_Gradient
                    m.frm.ProfileProperty(eProfileProp_GradientFrom) = gdColor1.Color
                    m.frm.ProfileProperty(eProfileProp_GradientTo) = gdColor2.Color
                Case MktProf_Color_OpenClose
                    m.frm.ProfileProperty(eProfileProp_UpColor) = gdColor1.Color
                    m.frm.ProfileProperty(eProfileProp_DownColor) = gdColor2.Color
                Case MktProf_Color_BidAsk
                    m.frm.ProfileProperty(eProfileProp_AskColor) = gdColor1.Color
                    m.frm.ProfileProperty(eProfileProp_BidColor) = gdColor2.Color
            End Select
        End If
        
        m.frm.ProfileProperty(eProfileProp_VolActualShow) = chkVolumeActual.Value
        m.frm.ProfileProperty(eProfileProp_VolPercentShow) = chkVolumePercent.Value
        m.frm.ProfileProperty(eProfileProp_VolShade) = chkShadeVol.Value
        m.frm.ProfileProperty(eProfileProp_VolShadeColor) = gdVolumeColor.Color
        m.frm.ProfileProperty(eProfileProp_TPOCountShow) = chkShowCountTPO.Value
        
        m.frm.ProfileProperty(eProfileProp_BoxFirst) = chkBoxFirst.Value
        
        'character sequence
        If optSeqOne Then
            m.frm.CharacterSequence = MktProf_CHR_SeqOne
        ElseIf OptSeqTwo Then
            m.frm.CharacterSequence = MktProf_CHR_SeqTwo
        ElseIf optSeqThree Then
            m.frm.CharacterSequence = MktProf_CHR_SeqThree
        ElseIf optSeqFour Then
            m.frm.CharacterSequence = MktProf_CHR_SeqFour
        ElseIf optSeqFive Then
            m.frm.CharacterSequence = MktProf_CHR_SeqFive
        ElseIf optSeqSix Then
            m.frm.CharacterSequence = MktProf_CHR_SeqSix
        ElseIf optSeqEight Then
            m.frm.CharacterSequence = MktProf_CHR_Classic
        Else
            m.frm.CharacterSequence = MktProf_CHR_Blocks
        End If
        
        With fgStats
            'show
            m.frm.StatShow(eProfileStats_POCVol) = .Cell(flexcpChecked, eStatsRow_POCVol, eStatsCol_Show)
            m.frm.StatShow(eProfileStats_POCTPO) = .Cell(flexcpChecked, eStatsRow_POCTPO, eStatsCol_Show)
            m.frm.StatShow(eProfileStats_VAVolume) = .Cell(flexcpChecked, eStatsRow_VAVol, eStatsCol_Show)
            m.frm.StatShow(eProfileStats_VATPO) = .Cell(flexcpChecked, eStatsRow_VATPO, eStatsCol_Show)
            m.frm.StatShow(eProfileStats_StdDev) = .Cell(flexcpChecked, eStatsRow_StdDev, eStatsCol_Show)
            m.frm.StatShow(eProfileStats_Mean) = .Cell(flexcpChecked, eStatsRow_Mean, eStatsCol_Show)
            'color
            m.frm.StatColor(eProfileStats_POCVol) = .Cell(flexcpBackColor, eStatsRow_POCVol, eStatsCol_Color)
            m.frm.StatColor(eProfileStats_POCTPO) = .Cell(flexcpBackColor, eStatsRow_POCTPO, eStatsCol_Color)
            m.frm.StatColor(eProfileStats_VAVolume) = .Cell(flexcpBackColor, eStatsRow_VAVol, eStatsCol_Color)
            m.frm.StatColor(eProfileStats_VATPO) = .Cell(flexcpBackColor, eStatsRow_VATPO, eStatsCol_Color)
            m.frm.StatColor(eProfileStats_StdDev) = .Cell(flexcpBackColor, eStatsRow_StdDev, eStatsCol_Color)
            m.frm.StatColor(eProfileStats_Mean) = .Cell(flexcpBackColor, eStatsRow_Mean, eStatsCol_Color)
            'pen size
            m.frm.StatPenSize(eProfileStats_POCVol) = Int(Val(.TextMatrix(eStatsRow_POCVol, eStatsCol_PenSize)))
            m.frm.StatPenSize(eProfileStats_POCTPO) = Int(Val(.TextMatrix(eStatsRow_POCTPO, eStatsCol_PenSize)))
            m.frm.StatPenSize(eProfileStats_VAVolume) = Int(Val(.TextMatrix(eStatsRow_VAVol, eStatsCol_PenSize)))
            m.frm.StatPenSize(eProfileStats_VATPO) = Int(Val(.TextMatrix(eStatsRow_VATPO, eStatsCol_PenSize)))
            m.frm.StatPenSize(eProfileStats_StdDev) = Int(Val(.TextMatrix(eStatsRow_StdDev, eStatsCol_PenSize)))
            m.frm.StatPenSize(eProfileStats_Mean) = Int(Val(.TextMatrix(eStatsRow_Mean, eStatsCol_PenSize)))
            'percent
            m.frm.StatPercent(eProfileStats_VAVolume) = ValOfText(.TextMatrix(eStatsRow_VAVol, eStatsCol_VAPercent))
            m.frm.StatPercent(eProfileStats_VATPO) = ValOfText(.TextMatrix(eStatsRow_VATPO, eStatsCol_VAPercent))
        End With
    End If
        
    strFontName = Me.Font.Name
    nFontSize = Me.Font.Size
    nBold = Me.Font.Bold
    nItalic = Me.Font.Italic
    
    m.frm.YScaleMinMove = nTicksPerRow
    m.frm.UpdateSettings strFontName, nFontSize, nBold, nItalic, bSMPProfile
    
    DoEvents
    
    Unload Me

ErrExit:
    If Len(strErr) > 0 Then InfBox strErr, "E"
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfileCfg.cmdOK_Click"

End Sub

Private Sub cmdSMPFont_Click()
    CommonDialogFont Me.CommonDialog1, Me.Font
End Sub

Private Sub cmdTime_Click()

    Dim dTime As Double
    
    If Not m.Bars Is Nothing Then
        If frmStartStopTimes.ShowMe(m.Bars) = True Then
            dTime = m.Bars.Prop(eBARS_StartTime)
            lblTimeStart.Caption = DateFormat(dTime / 1440, NO_DATE, HH_MM)
            dTime = m.Bars.Prop(eBARS_EndTime)
            lblTimeEnd.Caption = DateFormat(dTime / 1440, NO_DATE, HH_MM)
        End If
    End If
    
    Me.SetFocus

End Sub

Private Sub InitGridStats()

    Dim nCheck&

    With fgStats
        .Redraw = flexRDNone
        SetupGrid fgStats, eGridMode_Grid
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .FixedRows = 1
        .Editable = flexEDKbdMouse
        .ScrollBars = flexScrollBarNone
        .Rows = 7
        .Cols = 5
        
        .Cell(flexcpFontName, 0, 0, .Rows - 1, .Cols - 1) = "Microsoft Sans Serif"
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        
        'column headers
        .TextMatrix(0, eStatsCol_Show) = "Show"
        .TextMatrix(0, eStatsCol_Item) = "Item"
        .TextMatrix(0, eStatsCol_Color) = "Color"
        .TextMatrix(0, eStatsCol_PenSize) = "Width"
        .TextMatrix(0, eStatsCol_VAPercent) = "VA %"
                
        'items
        .TextMatrix(eStatsRow_POCVol, eStatsCol_Item) = "POC (volume)"
        .TextMatrix(eStatsRow_POCTPO, eStatsCol_Item) = "POC (TPO)"
        .TextMatrix(eStatsRow_VAVol, eStatsCol_Item) = "Value area (volume)"
        .TextMatrix(eStatsRow_VATPO, eStatsCol_Item) = "Value area (TPO)"
        .TextMatrix(eStatsRow_StdDev, eStatsCol_Item) = "Standard deviation"
        .TextMatrix(eStatsRow_Mean, eStatsCol_Item) = "Mean price"
        
        'width (pen size)
        .TextMatrix(eStatsRow_POCVol, eStatsCol_PenSize) = Str(m.frm.StatPenSize(eProfileStats_POCVol))
        .TextMatrix(eStatsRow_POCTPO, eStatsCol_PenSize) = Str(m.frm.StatPenSize(eProfileStats_POCTPO))
        .TextMatrix(eStatsRow_VAVol, eStatsCol_PenSize) = Str(m.frm.StatPenSize(eProfileStats_VAVolume))
        .TextMatrix(eStatsRow_VATPO, eStatsCol_PenSize) = Str(m.frm.StatPenSize(eProfileStats_VATPO))
        .TextMatrix(eStatsRow_StdDev, eStatsCol_PenSize) = Str(m.frm.StatPenSize(eProfileStats_StdDev))
        .TextMatrix(eStatsRow_Mean, eStatsCol_PenSize) = Str(m.frm.StatPenSize(eProfileStats_Mean))
        
        'value area percent
        .TextMatrix(eStatsRow_POCVol, eStatsCol_VAPercent) = "n/a"
        .TextMatrix(eStatsRow_POCTPO, eStatsCol_VAPercent) = "n/a"
        .TextMatrix(eStatsRow_VAVol, eStatsCol_VAPercent) = m.frm.StatPercentStr(eProfileStats_VAVolume)
        .TextMatrix(eStatsRow_VATPO, eStatsCol_VAPercent) = m.frm.StatPercentStr(eProfileStats_VATPO)
        .TextMatrix(eStatsRow_StdDev, eStatsCol_VAPercent) = "n/a"
        .TextMatrix(eStatsRow_Mean, eStatsCol_VAPercent) = "n/a"
        
        'color
        .Cell(flexcpBackColor, eStatsRow_POCVol, eStatsCol_Color) = m.frm.StatColor(eProfileStats_POCVol)
        .Cell(flexcpBackColor, eStatsRow_POCTPO, eStatsCol_Color) = m.frm.StatColor(eProfileStats_POCTPO)
        .Cell(flexcpBackColor, eStatsRow_VAVol, eStatsCol_Color) = m.frm.StatColor(eProfileStats_VAVolume)
        .Cell(flexcpBackColor, eStatsRow_VATPO, eStatsCol_Color) = m.frm.StatColor(eProfileStats_VATPO)
        .Cell(flexcpBackColor, eStatsRow_StdDev, eStatsCol_Color) = m.frm.StatColor(eProfileStats_StdDev)
        .Cell(flexcpBackColor, eStatsRow_Mean, eStatsCol_Color) = m.frm.StatColor(eProfileStats_Mean)
        
        .Cell(flexcpPictureAlignment, 1, 0, .Rows - 1, 0) = flexAlignCenterCenter
        
        'the StatShow(x) property returns 0:no show, 1:show
        'translate to grid checkbox constant: 0=no checkbox, 1=checkbox checked, 2=checkbox unchecked
        nCheck = m.frm.StatShow(eProfileStats_POCVol)
        If nCheck = 0 Then nCheck = flexUnchecked
        .Cell(flexcpChecked, eStatsRow_POCVol, eStatsCol_Show) = nCheck
        
        nCheck = m.frm.StatShow(eProfileStats_POCTPO)
        If nCheck = 0 Then nCheck = flexUnchecked
        .Cell(flexcpChecked, eStatsRow_POCTPO, eStatsCol_Show) = nCheck
        
        nCheck = m.frm.StatShow(eProfileStats_VAVolume)
        If nCheck = 0 Then nCheck = flexUnchecked
        .Cell(flexcpChecked, eStatsRow_VAVol, eStatsCol_Show) = nCheck
        
        nCheck = m.frm.StatShow(eProfileStats_VATPO)
        If nCheck = 0 Then nCheck = flexUnchecked
        .Cell(flexcpChecked, eStatsRow_VATPO, eStatsCol_Show) = nCheck
        
        nCheck = m.frm.StatShow(eProfileStats_StdDev)
        If nCheck = 0 Then nCheck = flexUnchecked
        .Cell(flexcpChecked, eStatsRow_StdDev, eStatsCol_Show) = nCheck
        
        nCheck = m.frm.StatShow(eProfileStats_Mean)
        If nCheck = 0 Then nCheck = flexUnchecked
        .Cell(flexcpChecked, eStatsRow_Mean, eStatsCol_Show) = nCheck
        
        .ColWidth(eStatsCol_Show) = 600
        .ColWidth(eStatsCol_Item) = 2000
        .ColWidth(eStatsCol_Color) = 600
        .ColWidth(eStatsCol_PenSize) = 800
        .ColWidth(eStatsCol_VAPercent) = 800
        .Redraw = flexRDBuffered
        
    End With

End Sub

Private Sub fgStats_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next

    Select Case Col
        Case eStatsCol_Item
            Cancel = True
        Case eStatsCol_VAPercent
            If Row <> eStatsRow_VAVol And Row <> eStatsRow_VATPO Then Cancel = True
    End Select

End Sub

Private Sub fgStats_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim bReset As Boolean

    bReset = True
    With fgStats
        If .Row >= .FixedRows And .Row < .Rows And .Col >= .FixedCols And .Col < .Cols Then
            .Select .Row, .Col
            m.nMouseColDown = .Col
            m.nMouseRowDown = .Row
            If .TextMatrix(0, .Col) = "Color" Then
                clrItemColor.Width = .ColWidth(eStatsCol_Color)
                clrItemColor.Left = .Left + .ColWidth(eStatsCol_Show) + .ColWidth(eStatsCol_Item)
                clrItemColor.Top = .Top + .RowHeight(.Row) * .Row - 1
                clrItemColor.Color = .Cell(flexcpBackColor, .Row, .Col)
                clrItemColor.Visible = True
                bReset = False
            End If
        End If
    End With

    If bReset Then
        m.nMouseColDown = -1
        m.nMouseRowDown = -1
        clrItemColor.Visible = False
    End If

End Sub

Private Sub fgStats_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    With fgStats
        If .MouseCol <> eStatsCol_Color Then clrItemColor.Visible = False
    End With

End Sub

Private Sub fgStats_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If clrItemColor.Visible And m.nMouseColDown > 0 And m.nMouseRowDown > 0 And Not clrItemColor.DropDownVisible Then
        clrItemColor.UserControl_Click
    End If

End Sub

Private Sub Form_Load()
    Me.Icon = Picture16(ToolbarIcon("ID_MarketProfile"), , True)
    
    g.Styler.StyleForm Me
    
End Sub

Private Sub optCurrentSession_Click()
    gdSessionDate.Enabled = False
End Sub

Private Sub optDate_Click()
    gdSessionDate.Enabled = True
End Sub

Private Sub optIntervalDays_Click()
    cboIntervalDays.Enabled = True
    cboIntervalMinutes.Enabled = False
End Sub

Private Sub optIntervalMinutes_Click()
    cboIntervalMinutes.Enabled = True
    cboIntervalDays.Enabled = False
End Sub

Private Sub optSeqEight_Click()
On Error GoTo ErrSection

'    If optIntervalDays Or Int(cboIntervalMinutes.Text) <> 30 Then
'        If InfBox("Data interval will be changed to 30 minutes for Classic Lettering.", "?", "Ok|Cancel", "TradeProfile") = "O" Then
'            If optIntervalDays Then optIntervalMinutes_Click
'            cboIntervalMinutes.ListIndex = 2
'            m.eSelSequence = MktProf_CHR_Classic
'        Else
'            SetIntervalCtrls
'            SetCharCtrls
'        End If
'    Else
        m.eSelSequence = MktProf_CHR_Classic
'    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfileCfg.optSeqEight_Click"

End Sub

Private Sub optSeqFive_Click()
    m.eSelSequence = MktProf_CHR_SeqFive
End Sub

Private Sub optSeqFour_Click()
    m.eSelSequence = MktProf_CHR_SeqFour
End Sub

Private Sub optSeqOne_Click()
    m.eSelSequence = MktProf_CHR_SeqOne
End Sub

Private Sub optSeqSeven_Click()
    m.eSelSequence = MktProf_CHR_Blocks
End Sub

Private Sub optSeqSix_Click()
    m.eSelSequence = MktProf_CHR_SeqSix
End Sub

Private Sub optSeqThree_Click()
    m.eSelSequence = MktProf_CHR_SeqThree
End Sub

Private Sub OptSeqTwo_Click()
    m.eSelSequence = MktProf_CHR_SeqTwo
End Sub

Private Sub SetCharCtrls()

    Select Case m.frm.CharacterSequence
        Case MktProf_CHR_SeqTwo
            OptSeqTwo.Value = True
        Case MktProf_CHR_SeqThree
            optSeqThree.Value = True
        Case MktProf_CHR_SeqFour
            optSeqFour.Value = True
        Case MktProf_CHR_SeqFive
            optSeqFive.Value = True
        Case MktProf_CHR_SeqSix
            optSeqSix.Value = True
        Case MktProf_CHR_Blocks
            optSeqSeven.Value = True
        Case MktProf_CHR_Classic
            optSeqEight.Value = True
        Case Default
            optSeqOne.Value = True
    End Select

End Sub

Private Sub SetIntervalCtrls()

    If m.frm.IntervalType = 0 Then
        optIntervalMinutes.Value = True
        cboIntervalMinutes.Text = Str(m.frm.ProfileInterval)
        Select Case m.frm.ProfileInterval
            Case 120
                cboIntervalMinutes.ListIndex = 0
            Case 60
                cboIntervalMinutes.ListIndex = 1
            Case 30
                cboIntervalMinutes.ListIndex = 2
            Case 15
                cboIntervalMinutes.ListIndex = 3
            Case 10
                cboIntervalMinutes.ListIndex = 4
            Case Default
                cboIntervalMinutes.ListIndex = -1
        End Select
        cboIntervalDays.ListIndex = 0
    Else
        optIntervalDays.Value = True
        cboIntervalDays.Text = Str(m.frm.ProfileInterval)
        Select Case m.frm.ProfileInterval
            Case 1
                cboIntervalDays.ListIndex = 0
            Case 2
                cboIntervalDays.ListIndex = 1
            Case 3
                cboIntervalDays.ListIndex = 2
            Case 4
                cboIntervalDays.ListIndex = 3
            Case 5
                cboIntervalDays.ListIndex = 4
            Case Default
                cboIntervalDays.ListIndex = -1
        End Select
        cboIntervalMinutes.ListIndex = 0
    End If

End Sub

Private Sub EnableGradientControls(ByVal bOn As Boolean)
On Error GoTo ErrSection:

    'aardvark 6938
    lblTPOColors.Enabled = bOn
    lblOtherColor.Enabled = bOn
    lblBackgroundColor.Enabled = bOn
    cboColorScheme.Enabled = bOn
    
    lblColorLabel1.Enabled = bOn
    lblColorLabel2.Enabled = bOn
    gdColor1.Enabled = bOn
    gdColor2.Enabled = bOn
    
    gdOtherText.Enabled = bOn
    gdBackground.Enabled = bOn

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmMarketProfileCfg.EnableGradientControls"

End Sub

