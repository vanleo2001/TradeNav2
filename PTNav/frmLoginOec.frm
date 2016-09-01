VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmLoginOec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   915
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLoginOec.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginOec.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginOec.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdLogin 
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   480
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
         Caption         =   "frmLoginOec.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":0094
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":00B4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1080
         TabIndex        =   8
         Top             =   480
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
         Caption         =   "frmLoginOec.frx":00D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":00FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":011E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkShowIP 
         Height          =   435
         Left            =   2220
         TabIndex        =   9
         Top             =   480
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
         Caption         =   "frmLoginOec.frx":013A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":018A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":01AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAgree 
         Height          =   435
         Left            =   0
         Top             =   0
         Width           =   3495
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmLoginOec.frx":01C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginOec.frx":0286
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":02A6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraServerInfo 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLoginOec.frx":02C2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginOec.frx":0306
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginOec.frx":0326
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optDemo 
         Height          =   220
         Left            =   900
         TabIndex        =   12
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
         Caption         =   "frmLoginOec.frx":0342
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":036C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":038C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTest 
         Height          =   220
         Left            =   1740
         TabIndex        =   13
         Top             =   240
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
         Caption         =   "frmLoginOec.frx":03A8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":03D2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":03F2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLive 
         Height          =   220
         Left            =   180
         TabIndex        =   11
         Top             =   240
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
         Caption         =   "frmLoginOec.frx":040E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":0438
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":0458
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRestore 
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   3135
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
         Caption         =   "frmLoginOec.frx":0474
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":04DA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":04FA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtServerIP 
         Height          =   285
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   1635
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginOec.frx":0516
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
         Tip             =   "frmLoginOec.frx":0536
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":0556
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPort 
         Height          =   285
         Left            =   2700
         TabIndex        =   3
         Top             =   600
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginOec.frx":0572
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
         Tip             =   "frmLoginOec.frx":0592
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":05B2
      End
      Begin HexUniControls.ctlUniLabelXP lblServerIP 
         Height          =   195
         Left            =   180
         Top             =   645
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
         Caption         =   "frmLoginOec.frx":05CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginOec.frx":05F6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":0616
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPort 
         Height          =   195
         Left            =   2280
         Top             =   645
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
         Caption         =   "frmLoginOec.frx":0632
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginOec.frx":065E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":067E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraLoginInfo 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLoginOec.frx":069A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginOec.frx":06C6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginOec.frx":06E6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveLogin 
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLoginOec.frx":0702
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":073E
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":075E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddLogin 
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLoginOec.frx":077A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginOec.frx":07B0
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":07D0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   2475
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginOec.frx":07EC
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
         Tip             =   "frmLoginOec.frx":080C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":082C
      End
      Begin HexUniControls.ctlUniComboImageXP cboUserName 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   0
         Width           =   1635
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
         Tip             =   "frmLoginOec.frx":0848
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":0868
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPassword 
         Height          =   255
         Left            =   0
         Top             =   420
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
         Caption         =   "frmLoginOec.frx":0884
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginOec.frx":08B8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":08D8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUserName 
         Height          =   255
         Left            =   0
         Top             =   60
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
         Caption         =   "frmLoginOec.frx":08F4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginOec.frx":092A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginOec.frx":094A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   3135
      Left            =   3780
      TabIndex        =   16
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5530
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLoginOec.frx":0966
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
      ScrollBars      =   2
      PasswordChar    =   ""
      TrapTab         =   0   'False
      RaiseChangeEvent=   -1  'True
      RaiseUpdateEvent=   0   'False
      RaiseSelChangeEvent=   -1  'True
      Tip             =   "frmLoginOec.frx":0986
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginOec.frx":09A6
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
Attribute VB_Name = "frmLoginOec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLoginOec.frm
'' Description: Dialog to allow user to enter login information for Open E-Cry
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 10/16/2013   DAJ         Resurrected and added "+" and "-" button for login management
'' 10/24/2013   DAJ         Added Live/Demo/Test server options
'' 07/09/2015   DAJ         Use user answer for Live/Demo question
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK or Cancel?
    
    strIniFile As String                ' INI file
    strConnectIni As String             ' Connection information INI file
    strBrokerName As String             ' Broker name
    
    strUserName As String               ' User name selected by the user
    strPassword As String               ' Password typed in by the user
    strIP As String                     ' IP address to use to connect
    strPort As String                   ' Port to use to connect
End Type
Private m As mPrivate

Public Property Get UserName() As String
    UserName = m.strUserName
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property

Public Property Get IP() As String
    IP = m.strIP
End Property

Public Property Get Port() As String
    Port = m.strPort
End Property

Private Property Get ServerMode() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = -1&
    If optLive.Value = True Then
        lReturn = 0
    ElseIf optDemo.Value = True Then
        lReturn = 1
    ElseIf optTest.Value = True Then
        lReturn = 2
    End If

    ServerMode = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmLoginOec.ServerMode.Get"
    
End Property
Private Property Let ServerMode(ByVal lServerMode As Long)
On Error GoTo ErrSection:
    
    Select Case lServerMode
        Case 0:
            optLive.Value = True
        Case 1:
            optDemo.Value = True
        Case 2:
            optTest.Value = True
    End Select

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmLoginOec.ServerMode.Let"
    
End Property

Private Property Get DefaultIP() As String
    If optLive.Value = True Then
        DefaultIP = GetIniFileProperty("IP", "", "Server", m.strConnectIni)
    ElseIf optDemo.Value = True Then
        DefaultIP = GetIniFileProperty("DemoIP", "", "Server", m.strConnectIni)
    ElseIf optTest.Value = True Then
        DefaultIP = GetIniFileProperty("TestIP", "", "Server", m.strConnectIni)
    End If
End Property

Private Property Get DefaultPort() As String
    If optLive.Value = True Then
        DefaultPort = GetIniFileProperty("Port", "", "Server", m.strConnectIni)
    ElseIf optDemo.Value = True Then
        DefaultPort = GetIniFileProperty("DemoPort", "", "Server", m.strConnectIni)
    ElseIf optTest.Value = True Then
        DefaultPort = GetIniFileProperty("TestPort", "", "Server", m.strConnectIni)
    End If
End Property

Private Property Get OverrideIP() As String
    If optLive.Value = True Then
        OverrideIP = GetIniFileProperty("IP", "", "Overrides", m.strIniFile)
    ElseIf optDemo.Value = True Then
        OverrideIP = GetIniFileProperty("DemoIP", "", "Overrides", m.strIniFile)
    ElseIf optTest.Value = True Then
        OverrideIP = GetIniFileProperty("TestIP", "", "Overrides", m.strIniFile)
    End If
End Property
Private Property Let OverrideIP(ByVal strIP As String)
    If optLive.Value = True Then
        SetIniFileProperty "IP", strIP, "Overrides", m.strIniFile
    ElseIf optDemo.Value = True Then
        SetIniFileProperty "DemoIP", strIP, "Overrides", m.strIniFile
    ElseIf optTest.Value = True Then
        SetIniFileProperty "TestIP", strIP, "Overrides", m.strIniFile
    End If
End Property

Private Property Get OverridePort() As String
    If optLive.Value = True Then
        OverridePort = GetIniFileProperty("Port", "", "Overrides", m.strIniFile)
    ElseIf optDemo.Value = True Then
        OverridePort = GetIniFileProperty("DemoPort", "", "Overrides", m.strIniFile)
    ElseIf optTest.Value = True Then
        OverridePort = GetIniFileProperty("TestPort", "", "Overrides", m.strIniFile)
    End If
End Property
Private Property Let OverridePort(ByVal strPort As String)
    If optLive.Value = True Then
        SetIniFileProperty "Port", strPort, "Overrides", m.strIniFile
    ElseIf optDemo.Value = True Then
        SetIniFileProperty "DemoPort", strPort, "Overrides", m.strIniFile
    ElseIf optTest.Value = True Then
        SetIniFileProperty "TestPort", strPort, "Overrides", m.strIniFile
    End If
End Property

Private Property Get ShowTest() As Boolean
    ShowTest = FileExist("C:\Common\Files.EXE")
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, User Name, Show IP?, Are we switching?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal nBroker As eTT_AccountType, Optional ByVal strUserName As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastUser As String           ' Last user logged into
    Dim Broker As cBroker               ' Broker object
    Dim strIP As String                 ' IP Address for the server
    Dim strPort As String               ' Port for the server

    m.bOK = False
    Set Broker = g.Broker.Broker(nBroker)
    If Not Broker Is Nothing Then
        m.strConnectIni = Broker.ConnectIni
        m.strIniFile = Broker.IniFile
        m.strBrokerName = Broker.BrokerName
        
        Caption = m.strBrokerName & " Login Information"
        txtServerIP.Text = strIP
        txtPort.Text = strPort
        
        strLastUser = UCase(GetIniFileProperty("LastUser", "", "User", m.strIniFile))
        LoadCombo
        If cboUserName.ListCount > 0 Then
            If SetCombo(strUserName) = False Then
                If SetCombo(strLastUser) = False Then
                    strLastUser = GetIniFileProperty("UserName", "", "User", m.strIniFile)
                    If SetCombo(strLastUser) = False Then
                        cboUserName.ListIndex = 0
                    End If
                End If
            End If
        End If
        
        If (cboUserName.ListCount = 0) Or ((cboUserName.ListCount = 1) And (bSwitching = True)) Then
            NewLogin
        End If
        
        If (cboUserName.ListCount > 1) Or ((cboUserName.ListCount = 1) And (bSwitching = False)) Then
            CheckBoxValue(chkShowIP) = bShowIP
            
            MoveFocus txtPassword
    
            ShowForm Me, eForm_Modal, frmMain
            
            If m.bOK = True Then
                m.strUserName = cboUserName.Text
                m.strPassword = Trim(txtPassword.Text)
                m.strIP = Trim(txtServerIP.Text)
                m.strPort = Trim(txtPort.Text)
            
                SetIniFileProperty "LastUser", cboUserName.Text, "User", m.strIniFile
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLoginOec.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboUserName_Click
'' Description: When the user changes the user name give the password the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboUserName_Click()
On Error GoTo ErrSection:

    ServerMode = cboUserName.ItemData(cboUserName.ListIndex)
    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.cboUserName_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowIP_Click
'' Description: Show/Hide the Server IP information as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowIP_Click()
On Error GoTo ErrSection:

    fraServerInfo.Visible = CheckBoxValue(chkShowIP)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.chkShowIP_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddLogin_Click
'' Description: Allow the user to enter in a new login user name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddLogin_Click()
On Error GoTo ErrSection:

    NewLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.cmdAddLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel out of the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLogin_Click
'' Description: Allow the user to login to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLogin_Click()
On Error GoTo ErrSection:

    If Len(Trim(txtServerIP.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtServerIP
        InfBox "Please enter in an IP address for the " & m.strBrokerName & " server", "!", , m.strBrokerName & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPort.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtPort
        InfBox "Please enter in the port for the " & m.strBrokerName & " server", "!", , m.strBrokerName & " Login Error"
        GoTo ErrExit
    End If

    If Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in a password", "!", , m.strBrokerName & " Login Error"
        GoTo ErrExit
    End If
    
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveLogin_Click
'' Description: Allow the user to remove one or more logins
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveLogin_Click()
On Error GoTo ErrSection:

    RemoveLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.cmdRemoveLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRestore_Click
'' Description: Allow the user to restore default server information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRestore_Click()
On Error GoTo ErrSection:

    txtServerIP.Text = DefaultIP
    txtPort.Text = DefaultPort
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.cmdRestore_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Make sure when the form is activated that the password gets focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kBlank")

    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    ' Only show the Test option if this is a "Genesis" user...
    optTest.Visible = ShowTest

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel out of the form if the user clicks on the X
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDemo_Click
'' Description: Select the demo server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDemo_Click()
On Error GoTo ErrSection:

    UpdateServerMode
    SetServerControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.optDemo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLive_Click
'' Description: Select the Live server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLive_Click()
On Error GoTo ErrSection:

    UpdateServerMode
    SetServerControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.optLive_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTest_Click
'' Description: Select the Test server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTest_Click()
On Error GoTo ErrSection:

    UpdateServerMode
    SetServerControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.optTest_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.txtPassword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.txtPort_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_LostFocus
'' Description: If the user is overriding the port, save the override
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPort_LostFocus()
On Error GoTo ErrSection:

    Dim strPort As String               ' Port that is in the text box
    
    strPort = Trim(txtPort.Text)
    
    If Len(strPort) = 0 Then
        txtPort.Text = DefaultPort
    ElseIf strPort = DefaultPort Then
        OverridePort = ""
    Else
        OverridePort = strPort
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.txtPort_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtServerIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtServerIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.txtServerIP_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_LostFocus
'' Description: Check to see if the server needs to change based on the IP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtServerIP_LostFocus()
On Error GoTo ErrSection:

    Dim strIP As String                 ' Server IP that is in the text box
    
    strIP = Trim(txtServerIP.Text)
    
    If Len(strIP) = 0 Then
        txtServerIP.Text = DefaultIP
    ElseIf strIP = DefaultIP Then
        OverrideIP = ""
    Else
        OverrideIP = strIP
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.txtServerIP_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    cmdRestore.Enabled = ((Trim(txtServerIP.Text) <> DefaultIP) Or (Trim(txtPort.Text) <> DefaultPort))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load the user name combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLogin As cGdArray           ' Login information
    
    cboUserName.Clear
    
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        
        For lIndex = 0 To astrLogins.Size - 1
            Set astrLogin = New cGdArray
            astrLogin.SplitFields astrLogins(lIndex), "|"
            
            cboUserName.AddItem astrLogin(0)
            cboUserName.ItemData(cboUserName.NewIndex) = CLng(Val(astrLogin(1)))
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.LoadCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Set the user name combo box to the given user name if possible
'' Inputs:      User Name
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetCombo(ByVal strUserName As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboUserName.ListCount > 0) And (Len(strUserName) > 0) Then
        For lIndex = 0 To cboUserName.ListCount - 1
            If UCase(cboUserName.List(lIndex)) = UCase(strUserName) Then
                bFound = True
                
                cboUserName.ListIndex = lIndex
                ServerMode = cboUserName.ItemData(lIndex)
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginOec.SetCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewLogin
'' Description: Allow the user to give us a new user name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewLogin()
On Error GoTo ErrSection:

    Dim strUserName As String           ' User Name from the user
    Dim strLogins As String             ' Login string from the INI file
    Dim strButtons As String            ' Buttons for the question
    Dim strServer As String             ' Server to login to
    
    strUserName = InfBox("What is your " & m.strBrokerName & " user name?", "?", , m.strBrokerName & " User Name", , , , , , "string")
    If Len(strUserName) > 0 Then
        If SetCombo(strUserName) = False Then
            If ShowTest Then
                strButtons = "+-Live|Demo|Test"
            Else
                strButtons = "+-Live|Demo"
            End If
            strServer = InfBox("Which server would you like to login to?", "?", strButtons, "Login Server")
            
            Select Case strServer
                Case "L"
                    optLive.Value = True
                    optDemo.Value = False
                    optTest.Value = False
                    
                Case "D"
                    optLive.Value = False
                    optDemo.Value = True
                    optTest.Value = False
                    
                Case "T"
                    optLive.Value = False
                    optDemo.Value = False
                    optTest.Value = True
                    
            End Select
            
            cboUserName.AddItem strUserName
            cboUserName.ItemData(cboUserName.NewIndex) = ServerMode
            
            SetCombo strUserName
            MoveFocus txtPassword
            
            SaveLogins
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.NewLogin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveLogin
'' Description: Allow the user to remove login information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLogin()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As cGdArray          ' Array of login information
    Dim astrList As cGdArray            ' List to send to the delete form
    Dim astrToDelete As cGdArray        ' List of logins to delete
    Dim strSelected As String           ' Currently selected login
    Dim lIndex As Long                  ' Index into a for loop

    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        Set astrLogins = New cGdArray
        Set astrList = New cGdArray
        astrList.Create eGDARRAY_Strings
        
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            astrList.Add Parse(astrLogins(lIndex), "|", 1) & vbTab & Str(lIndex)
        Next lIndex
        
        strSelected = cboUserName.Text
        
        Set astrToDelete = frmDelete.ShowMe(astrList, strSelected)
        If Not astrToDelete Is Nothing Then
            For lIndex = astrToDelete.Size - 1 To 0 Step -1
                astrLogins.Remove CLng(Val(astrToDelete(lIndex)))
            Next lIndex
            
            SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
            
            LoadCombo
            If SetCombo(strSelected) = False Then
                If cboUserName.ListCount > 0 Then
                    cboUserName.ListIndex = 0
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.RemoveLogin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateServerMode
'' Description: Update the server mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateServerMode()
On Error GoTo ErrSection:

    Dim lServerMode As Long             ' Server mode
    
    lServerMode = ServerMode
    
    If cboUserName.ListIndex > -1 Then
        If (cboUserName.ItemData(cboUserName.ListIndex) <> lServerMode) And (lServerMode >= 0) Then
            cboUserName.ItemData(cboUserName.ListIndex) = lServerMode
            SaveLogins
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.UpdateServerMode"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveLogins
'' Description: Save the logins
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveLogins()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLogins As cGdArray          ' Array of login information
    
    Set astrLogins = New cGdArray
    astrLogins.Create eGDARRAY_Strings, cboUserName.ListCount
    
    For lIndex = 0 To cboUserName.ListCount - 1
        astrLogins(lIndex) = cboUserName.List(lIndex) & "|" & cboUserName.ItemData(lIndex)
    Next lIndex
    
    SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.SaveLogins"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerControls
'' Description: Set the server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerControls()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP address out of the INI file
    Dim strPort As String               ' Port out of hte INI file

    txtServerIP.Text = DefaultIP
    txtPort.Text = DefaultPort
    
    strIP = OverrideIP
    If Len(strIP) > 0 Then
        If strIP = txtServerIP.Text Then
            OverrideIP = ""
        Else
            txtServerIP.Text = strIP
        End If
    End If

    strPort = OverridePort
    If Len(strPort) > 0 Then
        If strPort = txtPort.Text Then
            OverridePort = ""
        Else
            txtPort.Text = strPort
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginOec.SetServerControls"
    
End Sub

