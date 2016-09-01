VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditAnnot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Annotation"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   Enabled         =   0   'False
   Icon            =   "frmEditAnnot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraGannLines 
      Height          =   1770
      Left            =   5400
      TabIndex        =   87
      Top             =   7320
      Width           =   3870
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmEditAnnot.frx":030A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":032A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":034A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   124
         Top             =   660
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":0366
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0390
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":03B0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   0
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":03CC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":03F2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0412
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   94
         Top             =   300
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":042E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0458
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0478
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   93
         Top             =   1020
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":0494
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":04BE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":04DE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   5
         Left            =   1275
         TabIndex        =   91
         Top             =   300
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":04FA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0524
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0544
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   6
         Left            =   1275
         TabIndex        =   90
         Top             =   660
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":0560
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":058A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":05AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   7
         Left            =   1275
         TabIndex        =   89
         Top             =   1020
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":05C6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":05F0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0610
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   8
         Left            =   1275
         TabIndex        =   88
         Top             =   1380
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":062C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0656
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0676
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrGannColor 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   104
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrGannColor 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   105
         Top             =   630
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrGannColor 
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   106
         Top             =   990
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor clrGannColor 
         Height          =   315
         Index           =   3
         Left            =   2160
         TabIndex        =   107
         Top             =   1350
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniCheckXP chkGannLines 
         Height          =   255
         Index           =   4
         Left            =   420
         TabIndex        =   92
         Top             =   1380
         Width           =   1665
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":0692
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":06BC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":06DC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraRiskReward 
      Height          =   1455
      Left            =   6960
      TabIndex        =   141
      Top             =   6840
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
      Caption         =   "frmEditAnnot.frx":06F8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":0730
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":0750
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtRiskReward 
         Height          =   285
         Left            =   2040
         TabIndex        =   145
         Top             =   315
         Width           =   915
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":076C
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
         Tip             =   "frmEditAnnot.frx":079C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":07BC
      End
      Begin HexUniControls.ctlUniCheckXP chkShowValues 
         Height          =   255
         Left            =   360
         TabIndex        =   143
         Top             =   1080
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
         Caption         =   "frmEditAnnot.frx":07D8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":080E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":082E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkShowProfitLost 
         Height          =   255
         Left            =   360
         TabIndex        =   142
         Top             =   720
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
         Caption         =   "frmEditAnnot.frx":084A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":088A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":08AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRiskReward 
         Height          =   195
         Left            =   360
         Top             =   360
         Width           =   1485
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":08C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmEditAnnot.frx":090E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":092E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraTarget 
      Height          =   2115
      Left            =   3960
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":094A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":0996
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":09B6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkTargetValues 
         Height          =   255
         Left            =   300
         TabIndex        =   31
         Top             =   1770
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
         Caption         =   "frmEditAnnot.frx":09D2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0A30
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0A50
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPullbacks 
         Height          =   225
         Left            =   1920
         TabIndex        =   23
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
         Caption         =   "frmEditAnnot.frx":0A6C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0AA0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0AC0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chk3rd 
         Height          =   255
         Left            =   2340
         TabIndex        =   28
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
         Caption         =   "frmEditAnnot.frx":0ADC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":0B04
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0B24
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chk2nd 
         Height          =   255
         Left            =   1620
         TabIndex        =   27
         Top             =   660
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
         Caption         =   "frmEditAnnot.frx":0B40
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":0B68
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0B88
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chk1st 
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   660
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
         Caption         =   "frmEditAnnot.frx":0BA4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":0BCC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0BEC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTargets 
         Height          =   225
         Left            =   960
         TabIndex        =   24
         Top             =   300
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
         Caption         =   "frmEditAnnot.frx":0C08
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":0C38
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0C58
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboTargetStyle 
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Top             =   1380
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
         Tip             =   "frmEditAnnot.frx":0C74
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0C94
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrTargets 
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   1020
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   300
         Top             =   1080
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
         Caption         =   "frmEditAnnot.frx":0CB0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":0CEA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0D0A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   255
         Left            =   300
         Top             =   660
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
         Caption         =   "frmEditAnnot.frx":0D26
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":0D54
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0D74
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   300
         Top             =   300
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
         Caption         =   "frmEditAnnot.frx":0D90
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":0DBA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0DDA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   300
         Top             =   1440
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
         Caption         =   "frmEditAnnot.frx":0DF6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":0E30
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0E50
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraArrow 
      Height          =   1335
      Left            =   5640
      TabIndex        =   56
      Top             =   240
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
      Caption         =   "frmEditAnnot.frx":0E6C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":0EA4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":0EC4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optArrow 
         Height          =   255
         Index           =   2
         Left            =   1260
         TabIndex        =   63
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":0EE0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0F10
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0F30
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optArrow 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   62
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":0F4C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0F7C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":0F9C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optArrow 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   61
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":0FB8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":0FE0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1000
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboArrowSize 
         Height          =   315
         Left            =   1320
         TabIndex        =   60
         Top             =   930
         Width           =   1815
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
         Tip             =   "frmEditAnnot.frx":101C
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
         MouseIcon       =   "frmEditAnnot.frx":103C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboLineStyle 
         Height          =   315
         Left            =   1320
         TabIndex        =   57
         Top             =   570
         Width           =   1815
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
         Tip             =   "frmEditAnnot.frx":1058
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
         MouseIcon       =   "frmEditAnnot.frx":1078
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label27 
         Height          =   255
         Left            =   240
         Top             =   960
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
         Caption         =   "frmEditAnnot.frx":1094
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":10C8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":10E8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label26 
         Height          =   255
         Left            =   240
         Top             =   600
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
         Caption         =   "frmEditAnnot.frx":1104
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1138
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1158
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1500
         Picture         =   "frmEditAnnot.frx":1174
         Top             =   285
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2520
         Picture         =   "frmEditAnnot.frx":147E
         Top             =   285
         Width           =   480
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGannOptions 
      Height          =   1275
      Left            =   1200
      TabIndex        =   81
      Top             =   3660
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmEditAnnot.frx":1788
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":17B2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":17D2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtRun 
         Height          =   315
         Left            =   2880
         TabIndex        =   113
         Top             =   780
         Width           =   450
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":17EE
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
         Tip             =   "frmEditAnnot.frx":1810
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1830
      End
      Begin HexUniControls.ctlUniRadioXP optRatio 
         Height          =   220
         Index           =   1
         Left            =   1845
         TabIndex        =   112
         Top             =   960
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
         Caption         =   "frmEditAnnot.frx":184C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":1876
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1896
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optRatio 
         Height          =   220
         Index           =   0
         Left            =   1845
         TabIndex        =   111
         Top             =   735
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
         Caption         =   "frmEditAnnot.frx":18B2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":18DE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":18FE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtRise 
         Height          =   315
         Left            =   960
         TabIndex        =   110
         Top             =   780
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":191A
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
         Tip             =   "frmEditAnnot.frx":1940
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1960
      End
      Begin HexUniControls.ctlUniTextBoxXP txtGannToPrice 
         Height          =   315
         Left            =   2880
         TabIndex        =   101
         Top             =   285
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":197C
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
         Tip             =   "frmEditAnnot.frx":19A2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":19C2
      End
      Begin HexUniControls.ctlUniTextBoxXP txtGannPrice 
         Height          =   315
         Left            =   960
         TabIndex        =   82
         Top             =   285
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":19DE
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
         Tip             =   "frmEditAnnot.frx":1A04
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1A24
      End
      Begin HexUniControls.ctlUniLabelXP Label31 
         Height          =   195
         Left            =   2595
         Top             =   840
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":1A40
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1A66
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1A86
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   225
         Left            =   120
         Top             =   810
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
         Caption         =   "frmEditAnnot.frx":1AA2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1AD2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1AF2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label22 
         Height          =   240
         Left            =   3360
         Top             =   840
         Width           =   410
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":1B0E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1B3A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1B5A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label29 
         Height          =   225
         Left            =   2100
         Top             =   360
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
         Caption         =   "frmEditAnnot.frx":1B76
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1BA6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1BC6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGannPrice 
         Height          =   225
         Left            =   105
         Top             =   360
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
         Caption         =   "frmEditAnnot.frx":1BE2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1C16
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1C36
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPatternOnChart 
      Height          =   1995
      Left            =   165
      TabIndex        =   85
      Top             =   5760
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
      Caption         =   "frmEditAnnot.frx":1C52
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":1C72
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":1C92
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgPatternOnChart 
         Height          =   1275
         Left            =   120
         TabIndex        =   86
         Top             =   300
         Width           =   2955
         _cx             =   5212
         _cy             =   2249
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
   Begin HexUniControls.ctlUniFrameWL fraRect 
      Height          =   735
      Left            =   0
      TabIndex        =   32
      Top             =   7920
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":1CAE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":1CD8
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":1CF8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optRounded 
         Height          =   225
         Left            =   600
         TabIndex        =   36
         Top             =   450
         Width           =   1995
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":1D14
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":1D56
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1D76
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optEllipse 
         Height          =   225
         Left            =   1920
         TabIndex        =   35
         Top             =   210
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
         Caption         =   "frmEditAnnot.frx":1D92
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":1DBA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1DDA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optRectangle 
         Height          =   225
         Left            =   600
         TabIndex        =   33
         Top             =   210
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
         Caption         =   "frmEditAnnot.frx":1DF6
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":1E28
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1E48
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   255
         Left            =   180
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
         Caption         =   "frmEditAnnot.frx":1E64
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":1E90
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1EB0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraRegression 
      Height          =   2025
      Left            =   480
      TabIndex        =   48
      Top             =   5040
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
      Caption         =   "frmEditAnnot.frx":1ECC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":1EFA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":1F1A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtRegLineLen 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   210
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":1F36
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
         Tip             =   "frmEditAnnot.frx":1F56
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1F76
      End
      Begin HexUniControls.ctlUniComboImageXP cboChannels 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   180
         Width           =   1215
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
         Tip             =   "frmEditAnnot.frx":1F92
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":1FB2
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraLocation 
         Height          =   375
         Left            =   60
         TabIndex        =   131
         Top             =   1560
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
         Caption         =   "frmEditAnnot.frx":1FCE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnot.frx":1FEE
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":200E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optLocation 
            Height          =   220
            Index           =   0
            Left            =   0
            TabIndex        =   135
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            Caption         =   "frmEditAnnot.frx":202A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":2052
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":2072
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLocation 
            Height          =   220
            Index           =   1
            Left            =   840
            TabIndex        =   134
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            Caption         =   "frmEditAnnot.frx":208E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":20B8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":20D8
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLocation 
            Height          =   220
            Index           =   2
            Left            =   1680
            TabIndex        =   133
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            Caption         =   "frmEditAnnot.frx":20F4
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":211E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":213E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLocation 
            Height          =   220
            Index           =   3
            Left            =   2520
            TabIndex        =   132
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            Caption         =   "frmEditAnnot.frx":215A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":2182
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":21A2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPercent 
         Height          =   315
         Left            =   1320
         TabIndex        =   128
         Top             =   1320
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":21BE
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
         Tip             =   "frmEditAnnot.frx":21DE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":21FE
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPoints 
         Height          =   315
         Left            =   2580
         TabIndex        =   127
         Top             =   780
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":221A
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
         Tip             =   "frmEditAnnot.frx":223A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":225A
      End
      Begin HexUniControls.ctlUniRadioXP optPercent 
         Height          =   195
         Left            =   360
         TabIndex        =   130
         Top             =   1500
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
         Caption         =   "frmEditAnnot.frx":2276
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":22A4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":22C4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPoints 
         Height          =   220
         Left            =   120
         TabIndex        =   129
         Top             =   1260
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "frmEditAnnot.frx":22E0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":230C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":232C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtStdDevVal 
         Height          =   315
         Left            =   2005
         TabIndex        =   126
         Top             =   960
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":2348
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
         Tip             =   "frmEditAnnot.frx":2372
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2392
      End
      Begin HexUniControls.ctlUniComboImageXP cboIndicator 
         Height          =   315
         Left            =   960
         TabIndex        =   52
         Top             =   570
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
         Tip             =   "frmEditAnnot.frx":23AE
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":23CE
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboStdDevStyle 
         Height          =   315
         Left            =   2005
         TabIndex        =   49
         Top             =   1320
         Width           =   1215
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
         Tip             =   "frmEditAnnot.frx":23EA
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":240A
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkStdDevOnOff 
         Height          =   220
         Left            =   240
         TabIndex        =   125
         Top             =   1020
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   "frmEditAnnot.frx":2426
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":246C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":248C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblChannels 
         Height          =   255
         Left            =   1560
         Top             =   360
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
         Caption         =   "frmEditAnnot.frx":24A8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":24EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":250C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStdDevStyle 
         Height          =   255
         Left            =   525
         Top             =   1380
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
         Caption         =   "frmEditAnnot.frx":2528
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":256E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":258E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBars 
         Height          =   255
         Left            =   1800
         Top             =   240
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":25AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":25D2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":25F2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRegressionLength 
         Height          =   255
         Left            =   240
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":260E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":263A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":265A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblIndicator 
         Height          =   255
         Left            =   240
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
         Caption         =   "frmEditAnnot.frx":2676
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":26A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":26C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraDLineText 
      Height          =   2295
      Left            =   2040
      TabIndex        =   119
      Top             =   6840
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
      Caption         =   "frmEditAnnot.frx":26E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":271C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":273C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   7
         Left            =   360
         TabIndex        =   10
         Top             =   1020
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
         Caption         =   "frmEditAnnot.frx":2758
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":27AC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":27CC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   6
         Left            =   360
         TabIndex        =   13
         Top             =   1980
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
         Caption         =   "frmEditAnnot.frx":27E8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2844
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2864
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   1260
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "frmEditAnnot.frx":2880
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":28C4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":28E4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   1
         Left            =   360
         TabIndex        =   140
         Top             =   540
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
         Caption         =   "frmEditAnnot.frx":2900
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2960
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2980
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   0
         Left            =   360
         TabIndex        =   123
         Top             =   300
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
         Caption         =   "frmEditAnnot.frx":299C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":29EC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2A0C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   2
         Left            =   360
         TabIndex        =   122
         Top             =   780
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
         Caption         =   "frmEditAnnot.frx":2A28
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2A7A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2A9A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   4
         Left            =   360
         TabIndex        =   121
         Top             =   1500
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
         Caption         =   "frmEditAnnot.frx":2AB6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2AEC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2B0C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDLineText 
         Height          =   220
         Index           =   5
         Left            =   360
         TabIndex        =   120
         Top             =   1740
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
         Caption         =   "frmEditAnnot.frx":2B28
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2B70
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2B90
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGanncciSquareRange 
      Height          =   2520
      Left            =   5880
      TabIndex        =   21
      Top             =   7320
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":2BAC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":2BEA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":2C0A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkSRangeExtend 
         Height          =   220
         Left            =   240
         TabIndex        =   25
         Top             =   1800
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
         Caption         =   "frmEditAnnot.frx":2C26
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2C6C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2C8C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangeSquare 
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1560
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
         Caption         =   "frmEditAnnot.frx":2CA8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2CEA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2D0A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangeTB 
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1080
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
         Caption         =   "frmEditAnnot.frx":2D26
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2D7A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2D9A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangeCD 
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   840
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
         Caption         =   "frmEditAnnot.frx":2DB6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2E0C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2E2C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangeSameSwing 
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1320
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
         Caption         =   "frmEditAnnot.frx":2E48
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2E86
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2EA6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangeFirstBar 
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   360
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
         Caption         =   "frmEditAnnot.frx":2EC2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2F1E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2F3E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangePrice 
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   600
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
         Caption         =   "frmEditAnnot.frx":2F5A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":2FAE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":2FCE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSRangeOrigin 
         Height          =   255
         Left            =   1800
         Top             =   2160
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
         Caption         =   "frmEditAnnot.frx":2FEA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":301E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":303E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label10 
         Height          =   255
         Left            =   240
         Top             =   2160
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
         Caption         =   "frmEditAnnot.frx":305A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":30A0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":30C0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGannacciMultiply 
      Height          =   1305
      Left            =   3840
      TabIndex        =   53
      Top             =   2400
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":30DC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":30FC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":311C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtGannacciMultiply 
         Height          =   285
         Left            =   1560
         TabIndex        =   54
         Top             =   345
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":3138
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
         Tip             =   "frmEditAnnot.frx":3162
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3182
      End
      Begin HexUniControls.ctlUniCheckXP chkSRangeMultiply 
         Height          =   220
         Left            =   210
         TabIndex        =   58
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "frmEditAnnot.frx":319E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":31F2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3212
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAdjPriceTo 
         Height          =   255
         Left            =   210
         Top             =   1020
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
         Caption         =   "frmEditAnnot.frx":322E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":3272
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3292
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAdjPriceFrom 
         Height          =   255
         Left            =   210
         Top             =   750
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
         Caption         =   "frmEditAnnot.frx":32AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":32F4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3314
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   255
         Left            =   840
         Top             =   375
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
         Caption         =   "frmEditAnnot.frx":3330
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":3366
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3386
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraText 
      Height          =   1155
      Left            =   60
      TabIndex        =   3
      Top             =   2460
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":33A2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":33E0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":3400
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboBoxXP cboTextJustify 
         Height          =   315
         Left            =   1545
         TabIndex        =   59
         Top             =   285
         Width           =   1815
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
         Tip             =   "frmEditAnnot.frx":341C
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
         MouseIcon       =   "frmEditAnnot.frx":343C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkTextBorder 
         Height          =   220
         Left            =   1680
         TabIndex        =   65
         Top             =   120
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
         Caption         =   "frmEditAnnot.frx":3458
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3498
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":34B8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboAnchor 
         Height          =   315
         Left            =   1440
         TabIndex        =   64
         Top             =   600
         Width           =   1815
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
         Tip             =   "frmEditAnnot.frx":34D4
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
         MouseIcon       =   "frmEditAnnot.frx":34F4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniRichTextBoxXP rtfText 
         Height          =   375
         Left            =   960
         TabIndex        =   55
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":3510
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
         Tip             =   "frmEditAnnot.frx":3530
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3550
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
      Begin HexUniControls.ctlUniRadioXP optRight 
         Height          =   195
         Left            =   2460
         TabIndex        =   7
         Top             =   660
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":356C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3596
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":35B6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCenter 
         Height          =   195
         Left            =   1620
         TabIndex        =   6
         Top             =   660
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
         Caption         =   "frmEditAnnot.frx":35D2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":35FE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":361E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLeft 
         Height          =   220
         Left            =   930
         TabIndex        =   5
         Top             =   660
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
         Caption         =   "frmEditAnnot.frx":363A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3662
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3682
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAuto 
         Height          =   220
         Left            =   180
         TabIndex        =   4
         Top             =   660
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
         Caption         =   "frmEditAnnot.frx":369E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":36C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":36E6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtText 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3135
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":3702
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
         Tip             =   "frmEditAnnot.frx":372C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":374C
      End
      Begin HexUniControls.ctlUniLabelXP lblTextJustify 
         Height          =   225
         Left            =   1365
         Top             =   900
         Width           =   1845
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":3768
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":37A2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":37C2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTextAlign 
         Height          =   225
         Left            =   120
         Top             =   855
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
         Caption         =   "frmEditAnnot.frx":37DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":3856
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3876
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkShowValueInAxis 
      Height          =   220
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "frmEditAnnot.frx":3892
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":38D6
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":38F6
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboGreenBlattZones 
      Height          =   315
      Left            =   9015
      TabIndex        =   70
      Top             =   2565
      Width           =   1740
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
      Tip             =   "frmEditAnnot.frx":3912
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":3932
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraExt 
      Height          =   1335
      Left            =   3660
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":394E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":3996
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":39B6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   3
         Left            =   2520
         TabIndex        =   39
         Top             =   270
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
         Caption         =   "frmEditAnnot.frx":39D2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":39FA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3A1A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   2
         Left            =   1830
         TabIndex        =   40
         Top             =   270
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
         Caption         =   "frmEditAnnot.frx":3A36
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3A5E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3A7E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   1
         Left            =   1050
         TabIndex        =   41
         Top             =   270
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
         Caption         =   "frmEditAnnot.frx":3A9A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3AC4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3AE4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optExt 
         Height          =   220
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   270
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
         Caption         =   "frmEditAnnot.frx":3B00
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":3B28
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3B48
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor clrExt 
         Height          =   315
         Left            =   1500
         TabIndex        =   12
         Top             =   540
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniComboImageXP cboExtStyle 
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   900
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
         Tip             =   "frmEditAnnot.frx":3B64
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3B84
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblExtColor 
         Height          =   255
         Left            =   240
         Top             =   600
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
         Caption         =   "frmEditAnnot.frx":3BA0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":3BE0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3C00
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblExtStyle 
         Height          =   255
         Left            =   240
         Top             =   960
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
         Caption         =   "frmEditAnnot.frx":3C1C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":3C5C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3C7C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraBracket 
      Height          =   615
      Left            =   7680
      TabIndex        =   74
      Top             =   8040
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":3C98
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":3CC2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":3CE2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optSquare 
         Height          =   225
         Left            =   1680
         TabIndex        =   75
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":3CFE
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3D3A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3D5A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCurly 
         Height          =   225
         Left            =   360
         TabIndex        =   83
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":3D76
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmEditAnnot.frx":3DAC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3DCC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   180
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
         Caption         =   "frmEditAnnot.frx":3DE8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":3E14
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3E34
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraValue 
      Height          =   1260
      Left            =   120
      TabIndex        =   15
      Top             =   1350
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
      Caption         =   "frmEditAnnot.frx":3E50
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":3E7E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":3E9E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkSRLineDot 
         Height          =   220
         Left            =   120
         TabIndex        =   102
         Top             =   480
         Visible         =   0   'False
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
         Caption         =   "frmEditAnnot.frx":3EBA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3EEA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3F0A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSRRight 
         Height          =   220
         Left            =   1590
         TabIndex        =   114
         Top             =   195
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
         Caption         =   "frmEditAnnot.frx":3F26
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3F50
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3F70
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSRLeft 
         Height          =   220
         Left            =   1005
         TabIndex        =   115
         Top             =   255
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
         Caption         =   "frmEditAnnot.frx":3F8C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":3FB4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":3FD4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDisplaySRValue 
         Height          =   220
         Left            =   1680
         TabIndex        =   118
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "frmEditAnnot.frx":3FF0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":402A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":404A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkExtendSRLine 
         Height          =   220
         Left            =   180
         TabIndex        =   117
         Top             =   900
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
         Caption         =   "frmEditAnnot.frx":4066
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":40A4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":40C4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtFromY 
         Height          =   285
         Left            =   600
         TabIndex        =   68
         Top             =   570
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":40E0
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
         Tip             =   "frmEditAnnot.frx":4110
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4130
      End
      Begin HexUniControls.ctlUniTextBoxXP txtToY 
         Height          =   285
         Left            =   2190
         TabIndex        =   67
         Top             =   570
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":414C
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
         Tip             =   "frmEditAnnot.frx":4178
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4198
      End
      Begin HexUniControls.ctlUniTextBoxXP txtValue 
         Height          =   285
         Left            =   2190
         TabIndex        =   16
         Top             =   195
         Width           =   1035
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":41B4
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
         Tip             =   "frmEditAnnot.frx":41E4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4204
      End
      Begin HexUniControls.ctlUniLabelXP lblFrom 
         Height          =   195
         Left            =   135
         Top             =   615
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":4220
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":424A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":426A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTo 
         Height          =   195
         Left            =   1875
         Top             =   615
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
         Caption         =   "frmEditAnnot.frx":4286
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":42AA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":42CA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblValue 
         Height          =   195
         Left            =   135
         Top             =   240
         Width           =   1290
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":42E6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmEditAnnot.frx":4328
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4348
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdAlert 
      Height          =   330
      Left            =   9420
      TabIndex        =   116
      Top             =   8640
      Width           =   750
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
      Caption         =   "frmEditAnnot.frx":4364
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":438E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":43AE
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdSwitchSides 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   9900
      TabIndex        =   139
      Top             =   600
      Width           =   855
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
      Caption         =   "frmEditAnnot.frx":43CA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":4402
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":4422
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   615
      Left            =   60
      TabIndex        =   14
      Top             =   4140
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
      Caption         =   "frmEditAnnot.frx":443E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":445E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":447E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   330
         Left            =   2575
         TabIndex        =   44
         Top             =   180
         Width           =   750
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
         Caption         =   "frmEditAnnot.frx":449A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":44C8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":44E8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   60
         TabIndex        =   43
         Top             =   180
         Width           =   750
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
         Caption         =   "frmEditAnnot.frx":4504
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":452A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":454A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSaveDefaults 
         Height          =   330
         Left            =   965
         TabIndex        =   42
         Top             =   180
         Width           =   1455
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
         Caption         =   "frmEditAnnot.frx":4566
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":45A8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":45C8
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPattern 
      Height          =   1185
      Left            =   6300
      TabIndex        =   136
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
      Caption         =   "frmEditAnnot.frx":45E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":4604
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":4624
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPatternName 
         Height          =   285
         Left            =   180
         TabIndex        =   146
         Top             =   180
         Width           =   2835
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":4640
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
         Tip             =   "frmEditAnnot.frx":4660
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4680
      End
      Begin HexUniControls.ctlUniCheckXP chkPatternName 
         Height          =   195
         Left            =   180
         TabIndex        =   138
         Top             =   555
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
         Caption         =   "frmEditAnnot.frx":469C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":46F0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4710
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtForecastBars 
         Height          =   285
         Left            =   1500
         TabIndex        =   137
         Top             =   780
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":472C
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
         Tip             =   "frmEditAnnot.frx":474C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":476C
      End
      Begin HexUniControls.ctlUniLabelXP Label20 
         Height          =   255
         Left            =   180
         Top             =   840
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
         Caption         =   "frmEditAnnot.frx":4788
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":47C4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":47E4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkMultiChart 
      Height          =   390
      Left            =   180
      TabIndex        =   66
      Top             =   1260
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":4800
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":4870
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":4890
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraEllipse 
      Height          =   1365
      Left            =   7080
      TabIndex        =   72
      Top             =   2160
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
      Caption         =   "frmEditAnnot.frx":48AC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":48DA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":48FA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtAxisLen 
         Height          =   285
         Left            =   1440
         TabIndex        =   80
         Top             =   210
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmEditAnnot.frx":4916
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
         Tip             =   "frmEditAnnot.frx":4936
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4956
      End
      Begin HexUniControls.ctlUniCheckXP chkQtrLines 
         Height          =   220
         Left            =   240
         TabIndex        =   79
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "frmEditAnnot.frx":4972
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":49B6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":49D6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkAxes 
         Height          =   220
         Left            =   240
         TabIndex        =   78
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "frmEditAnnot.frx":49F2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4A24
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4A44
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAxisLenData 
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   77
         Top             =   480
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
         Caption         =   "frmEditAnnot.frx":4A60
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4A8A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4AAA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAxisLenData 
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   76
         Top             =   480
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
         Caption         =   "frmEditAnnot.frx":4AC6
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4AF2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4B12
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAxisLenData 
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   73
         Top             =   480
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
         Caption         =   "frmEditAnnot.frx":4B2E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4B58
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4B78
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label28 
         Height          =   255
         Left            =   240
         Top             =   480
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
         Caption         =   "frmEditAnnot.frx":4B94
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":4BB8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4BD8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label25 
         Height          =   255
         Left            =   120
         Top             =   240
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
         Caption         =   "frmEditAnnot.frx":4BF4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditAnnot.frx":4C3A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4C5A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkDynamic 
      Height          =   220
      Left            =   960
      TabIndex        =   71
      Top             =   840
      Width           =   2955
      _ExtentX        =   5212
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
      Caption         =   "frmEditAnnot.frx":4C76
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":4CE0
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":4D00
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraQuadrants 
      Height          =   885
      Left            =   480
      TabIndex        =   96
      Top             =   9120
      Width           =   3870
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmEditAnnot.frx":4D1C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":4D4E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":4D6E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   220
         Index           =   0
         Left            =   2040
         TabIndex        =   100
         Top             =   270
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
         Caption         =   "frmEditAnnot.frx":4D8A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4DC0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4DE0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   220
         Index           =   1
         Left            =   2040
         TabIndex        =   99
         Top             =   570
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "frmEditAnnot.frx":4DFC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4E32
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4E52
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   220
         Index           =   2
         Left            =   480
         TabIndex        =   98
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "frmEditAnnot.frx":4E6E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4EA2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4EC2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkQuadrant 
         Height          =   220
         Index           =   3
         Left            =   480
         TabIndex        =   97
         Top             =   570
         Width           =   1530
         _ExtentX        =   2699
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
         Caption         =   "frmEditAnnot.frx":4EDE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":4F12
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":4F32
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkUseFillColor 
      Height          =   195
      Left            =   7005
      TabIndex        =   109
      Top             =   375
      Visible         =   0   'False
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
      Caption         =   "frmEditAnnot.frx":4F4E
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":4FA6
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":4FC6
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin gdOCX.gdSelectColor clrFillColor 
      Height          =   315
      Left            =   9510
      TabIndex        =   108
      Top             =   30
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin HexUniControls.ctlUniCheckXP chkAllPanes 
      Height          =   255
      Left            =   2340
      TabIndex        =   84
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
      Caption         =   "frmEditAnnot.frx":4FE2
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":5024
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":5044
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdFont 
      Height          =   330
      Left            =   2640
      TabIndex        =   45
      Top             =   105
      Width           =   750
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
      Caption         =   "frmEditAnnot.frx":5060
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":508A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":50AA
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkPreIndicator 
      Height          =   220
      Left            =   180
      TabIndex        =   47
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "frmEditAnnot.frx":50C6
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":5112
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":5132
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin VB.Timer tmrEditAnnot 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   540
   End
   Begin HexUniControls.ctlUniComboImageXP cboStyle 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   540
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
      Tip             =   "frmEditAnnot.frx":514E
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":516E
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin gdOCX.gdSelectColor clrColor 
      Height          =   315
      Left            =   720
      TabIndex        =   46
      Top             =   135
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin gdOCX.gdSelectIcon gdSelectIcon 
      Height          =   340
      Left            =   9255
      TabIndex        =   103
      Top             =   1080
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   609
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      Top             =   60
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
      Caption         =   "frmEditAnnot.frx":518A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmEditAnnot.frx":51B8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":51D8
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraWaveLabels 
      Height          =   5235
      Left            =   4680
      TabIndex        =   147
      Top             =   3180
      Width           =   6315
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmEditAnnot.frx":51F4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditAnnot.frx":527C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":529C
      RightToLeft     =   0   'False
      Begin VSFlex7LCtl.VSFlexGrid fgWaveLabels 
         Height          =   4095
         Left            =   150
         TabIndex        =   148
         Top             =   285
         Width           =   2955
         _cx             =   5212
         _cy             =   7223
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
      Begin HexUniControls.ctlUniButtonImageXP cmdDelCustomLabel 
         Height          =   315
         Left            =   855
         TabIndex        =   144
         Top             =   4830
         Width           =   1770
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
         Caption         =   "frmEditAnnot.frx":52B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":52FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":531E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkPointArc 
         Height          =   220
         Left            =   3210
         TabIndex        =   154
         Top             =   285
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "frmEditAnnot.frx":533A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":5392
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":53B2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraWaveContinue 
         Height          =   2055
         Left            =   3210
         TabIndex        =   157
         Top             =   750
         Width           =   2955
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":53CE
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnot.frx":540A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":542A
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optWaveContinue 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   155
            Top             =   780
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
            Caption         =   "frmEditAnnot.frx":5446
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":5484
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":54A4
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optWaveContinue 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   156
            Top             =   1050
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
            Caption         =   "frmEditAnnot.frx":54C0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":54FE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":551E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optWaveContinue 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   158
            Top             =   1320
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
            Caption         =   "frmEditAnnot.frx":553A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":5578
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5598
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label37 
            Height          =   435
            Left            =   120
            Top             =   300
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
            Caption         =   "frmEditAnnot.frx":55B4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":564C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":566C
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP lblRepeatLabel 
            Height          =   255
            Left            =   600
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
            Caption         =   "frmEditAnnot.frx":5688
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":56C2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":56E2
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP lblNoLabelPastEnd 
            Height          =   195
            Left            =   600
            Top             =   1080
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
            Caption         =   "frmEditAnnot.frx":56FE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":5734
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5754
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP lblContinueLabel 
            Height          =   615
            Left            =   600
            Top             =   1320
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
            Caption         =   "frmEditAnnot.frx":5770
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":5814
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5834
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraWaveConnect 
         Height          =   2055
         Left            =   3210
         TabIndex        =   150
         Top             =   3075
         Width           =   2955
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmEditAnnot.frx":5850
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditAnnot.frx":588A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":58AA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optWaveConnect 
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   153
            Top             =   1320
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
            Caption         =   "frmEditAnnot.frx":58C6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":5902
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5922
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optWaveConnect 
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   152
            Top             =   780
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
            Caption         =   "frmEditAnnot.frx":593E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":597A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":599A
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optWaveConnect 
            Height          =   245
            Index           =   0
            Left            =   300
            TabIndex        =   151
            Top             =   240
            Width           =   315
            _ExtentX        =   556
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
            Caption         =   "frmEditAnnot.frx":59B6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmEditAnnot.frx":59F2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5A12
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label24 
            Height          =   615
            Left            =   600
            Top             =   1320
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
            Caption         =   "frmEditAnnot.frx":5A2E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":5AF2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5B12
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP Label23 
            Height          =   495
            Left            =   600
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
            Caption         =   "frmEditAnnot.frx":5B2E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":5BB2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5BD2
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
         Begin HexUniControls.ctlUniLabelXP Label19 
            Height          =   495
            Left            =   600
            Top             =   240
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
            Caption         =   "frmEditAnnot.frx":5BEE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmEditAnnot.frx":5C7C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmEditAnnot.frx":5C9C
            RightToLeft     =   0   'False
            WordWrap        =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdWaveCustomLabels 
         Height          =   315
         Left            =   855
         TabIndex        =   149
         Top             =   4455
         Width           =   1770
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
         Caption         =   "frmEditAnnot.frx":5CB8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditAnnot.frx":5CF8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditAnnot.frx":5D18
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblStyle 
      Height          =   255
      Left            =   180
      Top             =   600
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
      Caption         =   "frmEditAnnot.frx":5D34
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmEditAnnot.frx":5D60
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":5D80
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblColor 
      Height          =   255
      Left            =   180
      Top             =   180
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
      Caption         =   "frmEditAnnot.frx":5D9C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmEditAnnot.frx":5DC8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditAnnot.frx":5DE8
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmEditAnnot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmEditAnnot.frm
'' Description: Editor for annotations
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'' History:
''  09-01-2008 - could not add new controls to implement new request from EWI to allow left extension
''                  for Time Extension tool (received message that this form has reached max-allowed controls)
''             - decision made to move a number of tools to new editor form (frmEditAnnotExt.frm)
''             - decision made to duplicate controls & code to new form in one step then gradually
''                  remove the duplicated code from this form one at a time
''
''  09-09-2008 - duplicated controls & code for various tools to frmEditAnnotEx.frm
''  09-22-2008 - removed controls & code for Dinapoli Expansion (i.e. all controls within fraDNE)
''  11-22-2010 - removed controls & code for Dinapoli Retracement (i.e. all controls within fraDNR)
''               removed controls & code for Speed Resistance Fan (i.e. all controls within fraSpResistFan)
''               removed controls & code for Time Cycles (i.e. all controls within fraTimeCycle)
''               removed controls & code for fib tools (i.e. all contrals within fraFib)
''               removed code for Andrews Fork
''               renamed fraTimeLines, fgTimeLines to fraPatternOnChart, fgPatternOnChart (no longer shared with fib time zone/andrews fork tools)
''               removed code for icon annotations (editing for icons was moved to frmIconAnnot a long time ago)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kFraRegressionHt = 2385
Private Const kCustomWaveLabels = "\custom\PointLabels.cfg"

Private Type mPrivate
    Chart As cChart
    Annot As cAnnotation
    nAnnotIdx As Long
    bTextChanged As Boolean
    bMultiChartOption As Boolean
    bIgnoreUnload As Boolean    '(so MDI activate won't unload this form when a modal form called from here)
    bCenterColorStyle As Boolean
    bWasMultiChart As Boolean
    bReturnOptions As Boolean
End Type
Private m As mPrivate

Private Sub cboAnchor_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboAnchor.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboArrowSize_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboArrowSize.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboChannels_Click()
    Repaint
End Sub

Private Sub cboIndicator_Click()
On Error GoTo ErrSection:

    Repaint
    
    If Not m.Annot Is Nothing Then
        If m.Annot.eType = eANNOT_RegressionLine Then SetRegressionControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboIndicator.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboExtStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboExtStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboLineStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboLineStyle.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cboStdDevStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboStdDevStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboTargetStyle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboTargetStyle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cboTextJustify_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cboTextJustify.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chk1st_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chk1st.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chk2nd_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chk2nd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chk3rd_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chk3rd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkAllPanes_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkAxes.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkAxes_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkAxes.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkDisplaySRValue_Click()
    Repaint
End Sub

Private Sub chkDLineText_Click(Index As Integer)
On Error GoTo ErrSection:
    
    Dim i&
    
    If Index = 5 Then
        i = chkDLineText(5).Value
        chkDLineText(6).Enabled = i
        If Not i Then chkDLineText(6).Value = 0
        
'JM 06-05-2015: original code - leave awhile then remove if all ok
'   Original implementation may have shared Dollar Line frame, but looks like that changed.
'   GannacciSwingSquare has its own frame. Don't think this check is needed
'        If Not m.Annot Is Nothing Then
'            If m.Annot.eType <> eANNOT_GannacciSwingSquare Then
'                i = chkDLineText(5).Value
'                chkDLineText(6).Enabled = i
'                If Not i Then chkDLineText(6).Value = 0
'            End If
'        End If
    End If
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkDLineText_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkDynamic_Click()
On Error GoTo ErrSection:
    
    With m.Annot
        If .eType = eANNOT_RegressionLine Then
            SetRegressionControls
        Else
            Repaint
            If .eType = eANNOT_DollarLine Or .eType = eANNOT_DollarLine2 Or _
               .eType = eANNOT_DollarLine3 Or .eType = eANNOT_DollarLine4 Then
                If chkDynamic.Value = 1 Then
                    txtToY.Enabled = False
                Else
                    txtToY.Enabled = True
                End If
                With m.Annot
                    If .X(2) >= .X(1) Then
                        txtFromY.Text = CStr(.Y(1))
                        txtToY.Text = CStr(.Y(2))
                    Else
                        txtFromY.Text = CStr(.Y(2))
                        txtToY.Text = CStr(.Y(1))
                    End If
                End With
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkDynamic.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkExtendSRLine_Click()
    Repaint
End Sub

Private Sub chkSRangeCD_Click()
On Error GoTo ErrSection:
    
    Static bInProgress As Boolean
    
    If bInProgress Then Exit Sub
    
    bInProgress = True
    If Me.Visible Then
        If GannSROptionsOk(chkSRangeCD) Then Repaint
    End If
    bInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkSRangeCD_Click", eGDRaiseError_Show

End Sub

Private Sub chkSRangeExtend_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkSRangeFirstBar_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkSRangeMultiply_Click()
    
    If Me.Visible Then
        txtGannacciMultiply.Enabled = chkSRangeMultiply.Value
        Repaint
    End If
    
End Sub

Private Sub chkPatternName_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkPatternName.Click", eGDRaiseError_Show

End Sub

Private Sub chkPointArc_Click()
    Repaint
End Sub

Private Sub chkQuadrant_Click(Index As Integer)
On Error GoTo ErrSection:
            
    Dim bRepaint As Boolean
                
    bRepaint = True
        
    If chkQuadrant(Index).Value = vbUnchecked Then
        If chkQuadrant(0).Value = vbUnchecked And _
           chkQuadrant(1).Value = vbUnchecked And _
           chkQuadrant(2).Value = vbUnchecked And _
           chkQuadrant(3).Value = vbUnchecked Then
                
            bRepaint = False
            chkQuadrant(Index).Value = vbChecked
        End If
    End If
                
    If bRepaint Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkQuadrant.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkGannLines_Click(Index As Integer)
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkGannLines.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkPreIndicator_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkPreIndicator.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkQtrLines_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkQtrLines.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkShowProfitLost_Click()
    Repaint
End Sub

Private Sub chkShowValueInAxis_Click()
On Error GoTo ErrSection:
    
    If Me.Visible Then Repaint
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnot.chkShowValueInAxis_Click"

End Sub

Private Sub chkShowValues_Click()
    Repaint
End Sub

Private Sub chkSRangePrice_Click()
    If Me.Visible Then Repaint
End Sub

Private Sub chkSRangeSameSwing_Click()
On Error GoTo ErrSection:
    
    Static bInProgress As Boolean
    
    If bInProgress Then Exit Sub
    
    bInProgress = True
    If Me.Visible Then
        If GannSROptionsOk(chkSRangeSameSwing) Then Repaint
    End If
    bInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkSRangeSameSwing_Click", eGDRaiseError_Show

End Sub

Private Sub chkSRangeSquare_Click()
On Error GoTo ErrSection:
    
    Static bInProgress As Boolean
    
    If bInProgress Then Exit Sub
    
    bInProgress = True
    If Me.Visible Then
        If GannSROptionsOk(chkSRangeSquare) Then Repaint
    End If
    bInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkSRangeSquare_Click", eGDRaiseError_Show

End Sub

Private Sub chkSRangeTB_Click()
On Error GoTo ErrSection:
    
    Static bInProgress As Boolean
    
    If bInProgress Then Exit Sub
    
    bInProgress = True
    If Me.Visible Then
        If GannSROptionsOk(chkSRangeTB) Then Repaint
    End If
    bInProgress = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkSRangeTB_Click", eGDRaiseError_Show

End Sub

Private Sub chkSRLineDot_Click()
    Repaint
End Sub

Private Sub chkStdDevOnOff_Click()
On Error GoTo ErrSection:

    If Me.Visible And Not m.Annot Is Nothing Then
        If m.Annot.eType = eANNOT_RegressionLine Then
            If chkStdDevOnOff.Value = 0 Then
                txtStdDevVal.Text = "0"
                cboChannels.ListIndex = 0
            Else
                With m.Annot
                    If .Prop("StdDevVal") = "0" Then
                        txtStdDevVal.Text = "1"
                    Else
                        txtStdDevVal.Text = Val(.Prop("StdDevVal"))
                    End If
                    If .Prop("ChannelCount") = 0 Then .Prop("ChannelCount") = 1
                    cboChannels.ListIndex = Val(.Prop("ChannelCount")) - 1
                End With
            End If
        End If

        Repaint
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnot.chkStdDevOnOff_Click"

End Sub

Private Sub chkTargetValues_Click()
On Error GoTo ErrSection:

    If chkTargetValues.Value = 1 Then
        Enable cmdFont
    Else
        Disable cmdFont
    End If
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkTargetValues.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub chkTextBorder_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.chkTextBorder.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub chkUseFillColor_Click()
    Repaint
End Sub

Private Sub clrColor_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.clrColor.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrExt_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.clrExt.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrFillColor_Changed()
    Repaint
End Sub

Private Sub clrGannColor_Changed(Index As Integer)
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.clrGannColor.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub clrTargets_Changed()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.clrTargets.Changed", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdAlert_Click()

    Dim Alert As cAlert, i&
    
    If Not m.Annot Is Nothing Then
        If m.Annot.MultiChartFlag Then
            MultiChartAlert m.Annot.AnnotChart, m.Annot
        Else
            Set Alert = m.Annot.AlertObject()
            If Alert Is Nothing Then
                Set Alert = m.Annot.AlertObject(True)
                i = 1
            End If
            If Not frmAlerts.ShowMe(Alert, eGDAlertType_Annot) And i = 1 Then
                'a new alert object was created but user cancelled or
                'user does not have minimum required level/module -- remove it
                m.Annot.UpdateAlert 0
            End If
        End If
    End If
    
ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cmdAlert_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDelCustomLabel_Click()
On Error GoTo ErrSection:

    Dim i&, j&
    Dim aLabels As cGdArray
    
    With fgWaveLabels
        If .Row > 5 And .Row < .Rows Then
            .Redraw = flexRDNone
            .RemoveItem .Row
            .Redraw = flexRDDirect
            
            .Row = .Row
            If .Row < 6 Then cmdDelCustomLabel.Enabled = False
            
            If .Rows > 6 Then
                Set aLabels = New cGdArray
                For i = 6 To .Rows - 1
                    aLabels.Add .TextMatrix(i, 0)
                Next
                aLabels.ToFile g.strAppPath & kCustomWaveLabels
            Else
                KillFile g.strAppPath & kCustomWaveLabels, True
            End If
            
            Repaint
        End If
    End With

    Set aLabels = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cmdDelCustomLabel_Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim Annot As cAnnotation
    Dim aCopies As New cGdArray
    Dim i&, j&, idx&
                    
    If m.bReturnOptions Then
        If Not m.Annot Is Nothing Then m.Annot.Text = ""
    ElseIf Not m.Chart Is Nothing Then
        If m.nAnnotIdx <= m.Chart.Annots.Count And m.nAnnotIdx > 0 Then
            Set Annot = m.Chart.Annots(m.nAnnotIdx)
            If Not Annot Is Nothing Then
                'this is quicker than calling the chart's object remove annot routine
                Annot.geRemoveAnnotation (m.Chart.geChartObj)
                m.Chart.Annots.Remove m.nAnnotIdx
                m.Chart.SyncGlobalAnnots Annot, m.bWasMultiChart
            End If
        End If
    End If
    
    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection:
    
    Dim nStyle&
    
    'set font currently in use
    Me.Font.Name = m.Annot.Prop("FontName")
    Me.Font.Size = Val(m.Annot.Prop("FontSize"))
    Me.Font.Underline = Val(m.Annot.Prop("FontUnderline"))
    nStyle = Val(m.Annot.Prop("FontStyle"))
    Select Case nStyle
        Case 0:
            Me.Font.Italic = False
            Me.Font.Bold = False
        Case 1:
            Me.Font.Italic = False
            Me.Font.Bold = True
        Case 2:
            Me.Font.Italic = True
            Me.Font.Bold = False
        Case 3:
            Me.Font.Italic = True
            Me.Font.Bold = True
    End Select
    
    m.bIgnoreUnload = True
    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        m.Annot.Prop("FontName") = Me.Font.Name
        m.Annot.Prop("FontSize") = Me.Font.Size
        m.Annot.Prop("FontUnderline") = Me.Font.Underline
        
        'style - 0=reg,1=bold,2=italic,3=bold italic
        nStyle = 0
        If Me.Font.Bold = True Then
            If Me.Font.Italic = True Then
                nStyle = 3
            Else
                nStyle = 1
            End If
        ElseIf Me.Font.Italic = True Then
            nStyle = 2
        End If
        
        m.Annot.Prop("FontStyle") = nStyle
        Repaint
    End If
    
ErrExit:
    DoEvents
    m.bIgnoreUnload = False
    Exit Sub
    
ErrSection:
    m.bIgnoreUnload = False
    RaiseError "frmEditAnnot.cmdFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim strMsg$
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdOK
    
    If m.bReturnOptions Then
        SaveWaveLabels
    ElseIf Not m.Annot Is Nothing Then
        ' check if invalid
        If m.Annot.eType = eANNOT_TextEdit Or m.Annot.eType = eANNOT_TextEdit2 Or _
           m.Annot.eType = eANNOT_TextEdit3 Or m.Annot.eType = eANNOT_TextEdit4 Then
            If Len(Trim(rtfText.Text)) = 0 And optArrow(0) = True Then
                cmdDelete_Click
                Exit Sub
            End If
        End If
    
        Repaint '(still need this in order to save the changes)
        
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart
    
        If m.Annot.eType = eANNOT_DollarLine Or m.Annot.eType = eANNOT_DollarLine2 Or _
           m.Annot.eType = eANNOT_DollarLine3 Or m.Annot.eType = eANNOT_DollarLine4 Or _
           m.Annot.eType = eANNOT_GannacciSwingSquare Then
            
            If Not m.bWasMultiChart Then m.Chart.GenerateChart eRedo2_ReloadAnnots        '5217
        End If
        
    End If

    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSaveDefaults_Click()
On Error GoTo ErrSection:
    
    Repaint '(still need this in order to save the changes)
    
    If m.bReturnOptions Then
        SaveWaveLabels
        m.Annot.SaveDefaults
    ElseIf Not m.Annot Is Nothing Then
        
        With m.Annot
            If .eType = eANNOT_DollarLine Or .eType = eANNOT_DollarLine2 Or .eType = eANNOT_DollarLine3 Or _
               .eType = eANNOT_DollarLine4 Or .eType = eANNOT_GannacciSwingSquare Then
                .SaveDefaults txtValue.Tag
            Else
                .SaveDefaults
            End If
        End With
        
        m.Chart.SyncGlobalAnnots m.Annot, m.bWasMultiChart      '6334
        
        If m.Annot.eType = eANNOT_DollarLine Or m.Annot.eType = eANNOT_DollarLine2 Or m.Annot.eType = eANNOT_DollarLine3 Or _
           m.Annot.eType = eANNOT_DollarLine4 Or m.Annot.eType = eANNOT_GannacciSwingSquare Then
            If Not m.bWasMultiChart Then m.Chart.GenerateChart eRedo2_ReloadAnnots       '5217
        End If
        
    End If
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.cmdSaveDefaults.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSwitchSides_Click()

    If Not m.Annot Is Nothing Then
        With m.Annot
            If Val(.Prop("LenRatio")) = 0 Then
                .Prop("LenRatio") = 1
            Else
                .Prop("LenRatio") = 0
            End If
        End With
        Repaint
    End If
    
End Sub

Private Sub cmdWaveCustomLabels_Click()
    
    Dim rtrn$
    Dim aLabels As New cGdArray
    
    rtrn = AskBox("i=? ; Header=Ascii ; get=str ; default=A ; msg=Please enter labels separated by commas")
    
    If Len(rtrn) > 0 Then
        aLabels.Add rtrn$
        aLabels.ToFile g.strAppPath & kCustomWaveLabels, True
        With fgWaveLabels
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rtrn
        End With
    End If

End Sub

Private Sub fgPatternOnChart_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Repaint
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.fgPatternOnChart.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgPatternOnChart_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = 0 Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.fgPatternOnChart.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgPatternOnChart_DblClick()
On Error GoTo ErrSection:

    Dim dDate#
    Dim SelAnnot As cAnnotation
       
    If m.Annot.eType = eANNOT_Pattern Then
        With fgPatternOnChart
            If .TextMatrix(.Row, 0) = "Original" Then
                dDate = m.Annot.DateFromArray(1)
            ElseIf .TextMatrix(.Row, 0) = "Copy" Then
                dDate = m.Annot.dDate(2)
            End If
            m.Chart.Form.CenterTheDate dDate
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.fgPatternOnChart.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgWaveLabels_Click()
    
    Dim strText$, nLines&, nLabelFirstPoint&
    Dim nSelR1&, nSelCol1&, nSelR2&, nSelCol2&
    
    With fgWaveLabels
        .GetSelection nSelR1, nSelCol1, nSelR2, nSelCol2
        If nSelR1 >= 0 And nSelR1 < .Rows Then
            strText = .TextMatrix(nSelR1, 0)
        ElseIf Not m.Annot Is Nothing Then
            strText = m.Annot.Text
        End If
        If optWaveConnect(2).Value = True Then
            nLines = 0
        Else
            nLines = 1
            If optWaveConnect(0).Value = True Then
                nLabelFirstPoint = 0
            ElseIf optWaveConnect(1).Value = True Then
                nLabelFirstPoint = 1
            End If
        End If
        HandleWaveConnect strText, nLines, nLabelFirstPoint
        If .Row > 5 Then
            cmdDelCustomLabel.Enabled = True
        Else
            cmdDelCustomLabel.Enabled = False
        End If
    End With
    
    Repaint

End Sub

Private Sub fgWaveLabels_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then fgWaveLabels_Click

End Sub

Private Sub Form_Activate()

    DoEvents
    m.bIgnoreUnload = False

End Sub

Private Sub Form_Click()
On Error GoTo ErrSection:

    ' acts like "apply"
    If m.bTextChanged Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Form.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    'Set m.Annot = Nothing
    'Set m.Chart = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Form.Deactivate", eGDRaiseError_Show
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
    RaiseError "frmEditAnnot.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me
    
    cmdCancel.Top = -cmdCancel.Height * 2
    gdSelectIcon.AllowCustom = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If m.bIgnoreUnload Then
        Cancel = True
    ElseIf UnloadMode = 0 Then
        Cancel = True
        'cmdOK_Click
        tmrEditAnnot.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
    If m.bCenterColorStyle Then CenterColorStyle
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    If Not m.Annot Is Nothing Then
        If m.Annot.eType = eANNOT_TextEdit Or m.Annot.eType = eANNOT_TextEdit2 Or _
           m.Annot.eType = eANNOT_TextEdit3 Or m.Annot.eType = eANNOT_TextEdit4 Then
            If Not m.Chart Is Nothing Then
                If Not m.Chart.Form Is Nothing Then
                    m.Chart.Form.SyncDrawTools          '5141
                End If
            End If
        End If
    End If

    Set m.Annot = Nothing
    Set m.Chart = Nothing
        
    'frmMain.DockPro.RemoveForm Me.Name

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Sub Edit(Chart As cChart, ByVal nAnnotIdx&)
On Error GoTo ErrSection:
    
    Dim strText$, i&, eUsage As eAnnotUsage
    
    m.bIgnoreUnload = False
    m.bCenterColorStyle = False
    m.bReturnOptions = False
    If FormIsLoaded("frmChartCfg") Then
        If Not frmChartCfg.bNowAdding Then
            Unload frmChartCfg
        End If
    End If
   
    If nAnnotIdx <= 0 Or nAnnotIdx > Chart.Annots.Count Then
        Exit Sub
    End If
    eUsage = Chart.Annots(nAnnotIdx).eUsage
    If eUsage <> eANNOT_UserAdded Then
        Exit Sub
    End If
    
    DoEvents
    
    Set m.Chart = Chart
    Set m.Annot = Chart.Annots(nAnnotIdx)
    If m.Annot Is Nothing Then Exit Sub
     
    m.nAnnotIdx = m.Annot.geAnnId
'    m.nAnnotIdx = nAnnotIdx
    If m.Chart.SymbolID > 0 Then
        chkMultiChart.Caption = "Show for " & m.Chart.Symbol & " in all chart windows"
    ElseIf Len(m.Chart.SpreadSymbols) > 0 Then
        chkMultiChart.Caption = "Show in all chart windows for this spread"
    Else
        chkMultiChart.Caption = "Show in all chart windows for this symbol"
    End If
    HideOnInitialShow
    
    'Developer Note: there are 2 general types of subroutines: Init(xxx)Controls and
    '   Set(xxx)Controls. The "Init" type subroutines are intended to handle setting
    '   control values that do not need to be changed regardless of users input.
    '   The "Set" type subroutines are called to update controls values as users
    '   add/change/modify options.
    
    With m.Annot
        m.bWasMultiChart = .MultiChartFlag
        'set values for controls common to all annotations
        clrColor.Color = .Color
        chkPreIndicator.Value = .PreIndicator
        'show multichart option only if annotation is in price pane AND does not have alert
        If Chart.Tree.Key(m.Annot.gePaneId) = "PRICE PANE" And .AlertObject Is Nothing Then
            If .eType = eANNOT_RegressionLine Then
                'allow multi chart option only if attached to price bar (aardvark 1736)
                If UCase(.Prop("IndicatorKey")) = "PRICE" Then
                    chkMultiChart.Value = Abs(.MultiChartFlag)
                    m.bMultiChartOption = True
                Else
                    m.bMultiChartOption = False
                End If
            Else
                chkMultiChart.Value = Abs(.MultiChartFlag)
                m.bMultiChartOption = True
            End If
        Else
            m.bMultiChartOption = False
        End If
        If .eType = eANNOT_RegressionLine Then
            chkMultiChart.Visible = True
        Else
            chkMultiChart.Visible = m.bMultiChartOption
        End If
        
        LoadPenStyles cboStyle
        cboStyle.Width = clrColor.Width
        SetCombo cboStyle, .Style
        
        ' other properties
        Select Case .eType
            Case eANNOT_VertLine
                Me.Caption = "Vertical Line"
                Me.Icon = Picture16(ToolbarIcon("ID_VertLine"), , True)
                chkShowValueInAxis.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
                chkShowValueInAxis.Value = Val(m.Annot.Prop("ShowInAxis"))
                chkShowValueInAxis.Visible = True
                SetBottom chkShowValueInAxis
                m.bCenterColorStyle = True
            Case eANNOT_Mirror
                InitMirrorControls
            Case eANNOT_Pattern
                InitPatternControls
            Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
                InitHorzlineControls
            
            Case eANNOT_Trendline, eANNOT_Trendline2, _
                 eANNOT_Trendline3, eANNOT_Trendline4, _
                 eANNOT_TrendChannel
                
                InitTrendlineControls
            
            Case eANNOT_RegressionLine
                InitRegressionlineControls
            Case eANNOT_TargetShooter
                InitTargetshooterControls
            Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, eANNOT_DollarLine4
                InitDollarlineControls Chart
            Case eANNOT_Rectangle, eANNOT_TriangleWedge, eANNOT_ChannelHighlight
                InitRectangleControls
            Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4, eANNOT_ArrowLine
                InitTextEditControls
            Case eANNOT_Ellipse
                InitEllipseControls
            Case eANNOT_GannLines
                InitGannLinesControls
            Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
                InitSRLineControls
            Case eANNOT_Bracket
                InitBracketControls
            Case eANNOT_RiskReward
                InitRiskReward Chart
            Case eANNOT_WaveLabels
                ShowWaveLabels Nothing
            Case eANNOT_GannacciSwingSquare
                InitGancciSRangeCtrls Chart
        End Select
    End With
        
    If m.Annot.eType = eANNOT_GannLines Then
        Me.Width = fraGannOptions.Width + 350
        fraButtons.Left = fraGannOptions.Left + 350
    ElseIf m.Annot.eType <> eANNOT_WaveLabels Then
        Me.Width = fraButtons.Left * 2 + fraButtons.Width + Me.Width - Me.ScaleWidth
    End If
    
    'frmMain.DockPro.AddForm Me, DPUndocked
    'frmMain.DockPro.Dockable(Me.Name) = False
    'ShowUndocked Me, Me.Left, Me.Top, Me.Width, Me.Height
    'ShowForm Me, True
    
    CenterFormOnChart Me, m.Chart       '6434
    ShowForm Me
    m.bTextChanged = False          '4234
    
    If m.Annot.eType = eANNOT_TextEdit Or m.Annot.eType = eANNOT_TextEdit2 Or _
       m.Annot.eType = eANNOT_TextEdit3 Or m.Annot.eType = eANNOT_TextEdit4 Then
        MoveFocus rtfText
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Edit", eGDRaiseError_Raise
    
End Sub

Private Sub Repaint()
On Error GoTo ErrSection:
    
    Dim strText$
    Dim i&, j&
    Dim d#, dY1#, dY2#
    
    Dim aValues As New cGdArray
    Dim aDates As New cGdArray
    
    Dim strRatios$, strShow$, strColor$
    
    Dim bAnnotTextEdit As Boolean
    
    If Not Me.Visible Then Exit Sub
    If m.Chart Is Nothing Then Exit Sub
    If m.Annot Is Nothing Then Exit Sub
    
    With m.Annot
        If .eType = eANNOT_TextEdit Or .eType = eANNOT_TextEdit2 Or .eType = eANNOT_TextEdit3 Or .eType = eANNOT_TextEdit4 Then
            bAnnotTextEdit = True
        End If
        ' see if text has changed
        If m.bTextChanged Then
            If rtfText.Visible = True Then
                strText = rtfText.Text
            ElseIf txtText.Visible Then
                strText = txtText
            ElseIf .eType = eANNOT_HorzLine Or .eType = eANNOT_HorzLine2 Or _
                   .eType = eANNOT_HorzLine3 Or .eType = eANNOT_HorzLine4 Then
                strText = ""
            End If
            If Not bAnnotTextEdit Then
                If optLeft Then
                    .geTextAlign = 3      'e_btmLeft
                ElseIf optRight Then
                    .geTextAlign = 4      'e_btmRight
                ElseIf optCenter Then
                    .geTextAlign = 5      'e_btmCenter
                ElseIf optAuto Then
                    .geTextAlign = 9      'e_autoAlign
                End If
            End If
            .Text = strText
        End If
        
        ' main color, style & pre-indicator flag
        .Color = clrColor.Color
        .Style = cboStyle.ItemData(cboStyle.ListIndex)
        .PreIndicator = chkPreIndicator.Value
        If chkMultiChart.Value = 1 Then
            .MultiChartFlag = True
        Else
            .MultiChartFlag = False
        End If
        
        ' other properties
        Select Case m.Annot.eType
            Case eANNOT_VertLine
                .Prop("ShowInAxis") = chkShowValueInAxis.Value
        
            Case eANNOT_HorzLine, eANNOT_HorzLine2, eANNOT_HorzLine3, eANNOT_HorzLine4
                .Prop("ShowInAxis") = chkShowValueInAxis.Value
                If InStr(txtValue, "^") > 0 Then
                    .Y(1) = m.Chart.Bars.PriceFromString(txtValue)
                Else
                    .Y(1) = ValOfText(txtValue)
                End If
                .UpdateAlert 2, True
            
            Case eANNOT_Trendline, eANNOT_Trendline2, _
                 eANNOT_Trendline3, eANNOT_Trendline4, _
                 eANNOT_TrendChannel
                 
                If optExt(1) Then
                    .Prop("Ext") = 1
                ElseIf optExt(2) Then
                    .Prop("Ext") = 2
                ElseIf optExt(3) Then
                    .Prop("Ext") = 3
                Else
                    .Prop("Ext") = 0
                End If
                
                If chkShowValueInAxis.Value = vbChecked Then
                    If chkStdDevOnOff.Value = vbChecked Then
                        .Prop("ShowInAxis") = 2 '>1 --> label main line & channels
                    Else
                        .Prop("ShowInAxis") = 1 'main line only
                    End If
                ElseIf chkStdDevOnOff.Value = vbChecked Then
                    .Prop("ShowInAxis") = -1    'label channels only
                Else
                    .Prop("ShowInAxis") = 0
                End If
                
                .Prop("ExtColor") = clrExt.Color
                .Prop("ExtStyle") = CboItem(cboExtStyle)
                'only save channels properties channels option is not set to none
                'and other channels controls are visible
                If lblIndicator.Visible And optLocation(i).Value = False Then
                    .Prop("ChannelCount") = cboChannels
                    .Prop("ChannelStyle") = CboItem(cboIndicator)
                    .Prop("ChannelPoints") = Str(ValOfText(txtPoints.Text))
                    .Prop("ChannelPercent") = Str(ValOfText(txtPercent.Text))
                    If optPercent.Value = True Then
                        .Prop("ChannelType") = 0
                    Else
                        .Prop("ChannelType") = 1
                    End If
                End If
                cmdAlert.Enabled = m.Annot.CanHaveAlert     'aardvark 3298
                If m.Annot.eType = eANNOT_TrendChannel Then
                    .Prop("HideMainLine") = chkAllPanes.Value
                End If
                    
            Case eANNOT_TargetShooter
                .Prop("Pullbacks") = CLng(optPullbacks)
                .Prop("Pt1") = chk1st
                .Prop("Pt2") = chk2nd
                .Prop("Pt3") = chk3rd
                .Prop("TargetColor") = clrTargets.Color
                .Prop("TargetStyle") = CboItem(cboTargetStyle)
                .Prop("ShowValues") = chkTargetValues
                    
            Case eANNOT_DollarLine, eANNOT_DollarLine2, eANNOT_DollarLine3, _
                 eANNOT_DollarLine4, eANNOT_GannacciSwingSquare
                If IsAlpha(txtFromY) Then
                    'restore value
                    If .X(2) >= .X(1) Then
                        txtFromY = m.Chart.Bars.PriceDisplay(.Y(1))
                    Else
                        txtFromY = m.Chart.Bars.PriceDisplay(.Y(2))
                    End If
                ElseIf IsAlpha(txtToY) Then
                    'restore value
                    If .X(2) >= .X(1) Then
                        txtToY = m.Chart.Bars.PriceDisplay(.Y(2))
                    Else
                        txtToY = m.Chart.Bars.PriceDisplay(.Y(1))
                    End If
                Else
                    d = ValOfText(txtValue)
                    .Prop(txtValue.Tag) = d     'the tag contains NumShares, NumContracts or NumForexContracts accordingly
                    
                    If InStr(txtFromY, "^") > 0 Then
                        dY1 = m.Chart.Bars.PriceFromString(txtFromY)
                    Else
                        dY1 = ValOfText(txtFromY)
                    End If
                    
                    If InStr(txtToY, "^") > 0 Then
                        dY2 = m.Chart.Bars.PriceFromString(txtToY)
                    Else
                        dY2 = ValOfText(txtToY)
                    End If
                    
                    If .X(2) >= .X(1) Then
                        .Y(1) = dY1
                        .Y(2) = dY2
                    Else
                        .Y(2) = dY1
                        .Y(1) = dY2
                    End If
                    
                    If m.Annot.eType = eANNOT_GannacciSwingSquare Then
                        .Prop("IncludeBarOne") = chkSRangeFirstBar
                        .Prop("PriceMove") = chkSRangePrice
                        .Prop("CalendarDays") = chkSRangeCD
                        .Prop("LenBars") = chkSRangeTB
                        .Prop("SameSwing") = chkSRangeSameSwing
                        .Prop("SquareRange") = chkSRangeSquare
                        .Prop("Ext") = chkSRangeExtend
                        .Prop("UseMutiplier") = chkSRangeMultiply
                        .Prop("MultiplierVal") = ValOfText(txtGannacciMultiply.Text)
                        
                        If .Prop("UseMutiplier") = 1 Then
                            d = ValOfText(txtGannacciMultiply.Text)
                            If d <> 0 Then
                                dY1 = dY1 * d
                                dY2 = dY2 * d
                            End If
                        End If
                        
                        lblAdjPriceFrom.Caption = "Adjusted price from:  " & Format(dY1, "#0.00###")
                        lblAdjPriceTo.Caption = "Adjusted price to:      " & Format(dY2, "#0.00###")
                    Else
                        .Prop("KeepAtEnd") = chkDynamic
                        .Prop("ProfitLoss") = chkDLineText(0).Value
                        .Prop("ProfitLossPercent") = chkDLineText(1).Value
                        .Prop("PriceMove") = chkDLineText(2).Value
                        .Prop("SlopeLine") = chkDLineText(3).Value
                        .Prop("Volume") = chkDLineText(4).Value
                        .Prop("LenBars") = chkDLineText(5).Value
                        .Prop("LenHzLine") = chkDLineText(6).Value
                        .Prop("PriceMovePoints") = chkDLineText(7).Value
                    End If
                                    
                End If
                
            Case eANNOT_Rectangle
                If optEllipse Then
                    .Prop("Shape") = 2
                ElseIf optRounded Then
                    .Prop("Shape") = 1
                Else
                    .Prop("Shape") = 0
                End If
                .Prop("FillColor") = clrFillColor.Color
                .Prop("FillPattern") = chkUseFillColor.Value
            
            Case eANNOT_TriangleWedge, eANNOT_ChannelHighlight
                .Prop("FillColor") = clrFillColor.Color
                .Prop("FillPattern") = chkUseFillColor.Value
                
            Case eANNOT_TextEdit, eANNOT_TextEdit2, eANNOT_TextEdit3, eANNOT_TextEdit4, eANNOT_ArrowLine
                If bAnnotTextEdit Then
                    .Prop("Border") = chkTextBorder.Value
                    .SetTextEditSize m.Chart, rtfText.Text
                    .geTextAlign = cboAnchor.ListIndex
                    .Prop("TextJustify") = cboTextJustify.ListIndex
                End If
                .Prop("ArrowSize") = cboArrowSize.ListIndex
                .Prop("ArrowLine") = cboLineStyle.ItemData(cboLineStyle.ListIndex)
                If .Prop("ArrowStyle") = 0 Then
                    .dDate(2) = .dDate(1)      'aardvark 1225 fix
                    .Y(2) = .Y(1)
                End If
                If optArrow(1) = True Then
                    .Prop("ArrowStyle") = 1
                ElseIf optArrow(2) = True Then
                    .Prop("ArrowStyle") = 2
                Else
                    .Prop("ArrowStyle") = 0
                End If
            
            Case eANNOT_Ellipse
                If optAxisLenData(1) = True Then
                    .geMinorAxisLen = m.Chart.Bars.PriceFromString(txtAxisLen.Text)
                    .geMinorAxisLenData = 1     'points
                ElseIf optAxisLenData(2) = True Then
                    .geMinorAxisLen = ValOfText(txtAxisLen.Text)
                    .geMinorAxisLenData = 2     'ticks
                    .geMinorAxisLen = .geMinorAxisLen * m.Chart.Bars.Prop(eBARS_TickMove)
                Else
                    .geMinorAxisLen = ValOfText(txtAxisLen.Text)
                    .geMinorAxisLenData = 0     'ratio
                End If
                .Prop("ShowAxes") = chkAxes
                .Prop("ShowQtrLines") = chkQtrLines
                            
            Case eANNOT_RegressionLine
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'04-27-2007: original code, leave awhile then remove
'                'indicator & price field property
'                strText = cboIndicator.Text
'                m.Annot.geIndId = CboItem(cboIndicator)
'                m.Annot.Prop("IndicatorKey") = m.Chart.Tree.Key(m.Annot.geIndId)
'                If InStr(strText, "(close)") Then
'                    .Prop("PriceField") = 0
'                ElseIf InStr(strText, "(open)") Then
'                    .Prop("PriceField") = 1
'                ElseIf InStr(strText, "(high)") Then
'                    .Prop("PriceField") = 2
'                ElseIf InStr(strText, "(low)") Then
'                    .Prop("PriceField") = 3
'                ElseIf InStr(strText, "(avg high low)") Then
'                    .Prop("PriceField") = 4
'                End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                .Prop("ShowInAxis") = chkShowValueInAxis.Value
                IndicatorsCboToAnnot cboIndicator, m.Chart, m.Annot
                If UCase(m.Annot.Prop("IndicatorKey")) = "PRICE" Then
                    m.bMultiChartOption = True
                Else
                    m.bMultiChartOption = False
                End If
                .Prop("LenInBars") = Int(ValOfText(txtRegLineLen.Text))
                'fixed or variable length
                .Prop("FixLength") = optPoints.Value
                If .Prop("FixLength") Then
                    i = Abs(.X(2) - .X(1))
                    If i <> ValOfText(txtRegLineLen.Text) Then
                        i = i - ValOfText(txtRegLineLen.Text)
                        If .X(2) > .X(1) Then
                            i = .X(1) + i
                            .dDate(1) = m.Chart.Bars(eBARS_DateTime, i)
                        Else
                            i = .X(2) + i
                            .dDate(2) = m.Chart.Bars(eBARS_DateTime, i)
                        End If
                    End If
                End If
                'extensions and stdDev properties
                d = ValOfText(txtStdDevVal.Text)
                .Prop("KeepAtEnd") = chkDynamic
                .Prop("StdDevVal") = d
                .Prop("StdDevStyle") = CboItem(cboStdDevStyle)
                .Prop("ExtStyle") = CboItem(cboExtStyle)
                .Prop("ExtColor") = clrExt.Color
                If chkStdDevOnOff.Value = 0 Or d <= 0 Then
                    .Prop("ChannelCount") = 0
                    cboChannels.ListIndex = 0
                Else
                    .Prop("ChannelCount") = cboChannels
                End If
                If optExt(1) Then
                    .Prop("Ext") = 1
                ElseIf optExt(2) Then
                    .Prop("Ext") = 2
                ElseIf optExt(3) Then
                    .Prop("Ext") = 3
                Else
                    .Prop("Ext") = 0
                End If
                
            Case eANNOT_GannLines
                'only change quadrant property if at least one check box is on
                If chkQuadrant(0).Value = 1 Or _
                   chkQuadrant(1).Value = 1 Or _
                   chkQuadrant(2).Value = 1 Or _
                   chkQuadrant(3).Value = 1 Then
                   
                    .Prop("DirNE") = chkQuadrant(0).Value
                    .Prop("DirSE") = chkQuadrant(1).Value
                    .Prop("DirNW") = chkQuadrant(2).Value
                    .Prop("DirSW") = chkQuadrant(3).Value
                End If
                'angle/fan (lines to show)
                If chkGannLines(0) = 0 Then
                    .Prop("GannFan") = 0
                    For i = 0 To 3
                        clrGannColor(i).Enabled = False
                    Next
                    For i = 1 To 8
                        chkGannLines(i).Enabled = False
                    Next
                Else
                    For i = 0 To 3
                        clrGannColor(i).Enabled = True
                    Next
                    For i = 1 To 8
                        chkGannLines(i).Enabled = True
                    Next
                    .Prop("GannFan") = 1
                    'lines to show
                    .Prop("1x2") = chkGannLines(1)
                    .Prop("1x3") = chkGannLines(2)
                    .Prop("1x4") = chkGannLines(3)
                    .Prop("1x8") = chkGannLines(4)
                    .Prop("2x1") = chkGannLines(5)
                    .Prop("3x1") = chkGannLines(6)
                    .Prop("4x1") = chkGannLines(7)
                    .Prop("8x1") = chkGannLines(8)
                    'color of each lines-pair
                    .Prop("Color1x2") = clrGannColor(0).Color
                    .Prop("Color1x3") = clrGannColor(1).Color
                    .Prop("Color1x4") = clrGannColor(2).Color
                    .Prop("Color1x8") = clrGannColor(3).Color
                End If
            
            Case eANNOT_SRLine, eANNOT_SRLine2, eANNOT_SRLine3, eANNOT_SRLine4
                .Prop("ShowInAxis") = chkShowValueInAxis.Value
                If m.bTextChanged Then
                    If InStr(txtValue, "^") > 0 Then
                        .Y(1) = m.Chart.Bars.PriceFromString(txtValue)
                        .Y(2) = .Y(1)
                    Else
                        .Y(1) = ValOfText(txtValue)
                        .Y(2) = .Y(1)
                    End If
                End If
                If optSRRight.Value = True Then
                    .Prop("TextAlignment") = 6
                Else
                    .Prop("TextAlignment") = 7
                End If
                .SetSRLineExt m.Chart, chkExtendSRLine.Value
                .Prop("ShowValues") = Str(chkDisplaySRValue.Value)
                .Prop("HideSRLineDot") = Str(chkSRLineDot.Value)
                cmdAlert.Enabled = .CanHaveAlert        'aardvark 3298, 4325
                If (.Prop("Ext") = 0) Then
                    .UpdateAlert 0
                Else
                    .UpdateAlert 2, True
                End If
                                
            Case eANNOT_Pattern
                .Text = txtPatternName
                .Prop("ShowPatternName") = chkPatternName.Value
                If ValOfText(txtForecastBars.Text) >= 0 Then
                    .ForecastBarsChange m.Chart, 1, Int(ValOfText(txtForecastBars.Text))
                Else
                    txtForecastBars.Text = "0"
                End If
                
            Case eANNOT_RiskReward
                d = Abs(Int(ValOfText(txtRiskReward)))
                If d = 0 Then d = 1
                .Prop(txtRiskReward.Tag) = d
                .Prop("ShowProfitLoss") = chkShowProfitLost.Value
                .Prop("ShowValues") = chkShowValues.Value
                
            Case eANNOT_WaveLabels
                SaveWaveLabels
            
            Case eANNOT_Bracket
                If optSquare.Value = True Then
                    .Prop("BracketStyle") = 1
                Else
                    .Prop("BracketStyle") = 0
                End If
                            
        End Select
        m.bTextChanged = False
        
        '.AssignDateTime
        ' Do this since # of points could have changed
        ' (e.g. extensions being toggled)
        m.Chart.GenerateChart eRedo1_Scrolled
    
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.Repaint", eGDRaiseError_Raise
    
End Sub

Private Sub gdSelectIcon_Changed()
    SetImageInfo
    Repaint
End Sub

Private Sub Image1_Click()

    optArrow(1).Value = True

End Sub

Private Sub Image2_Click()

    optArrow(2).Value = True

End Sub

Private Sub mnuCopy_Click()
On Error GoTo ErrSection:
    
    Clipboard.Clear
    If rtfText.SelLength > 0 Then
        Clipboard.SetText rtfText.SelText, rtfCFText
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.mnuCopy_Click"
    Resume ErrExit
End Sub

Private Sub mnuCut_Click()
On Error GoTo ErrSection:
    
    If rtfText.SelLength > 0 Then
        rtfText.SelText = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.mnuCut_Click"
    Resume ErrExit
End Sub

Private Sub mnuPaste_Click()
On Error GoTo ErrSection:
    
    If Clipboard.GetFormat(rtfCFText) Then
        rtfText.SelText = Clipboard.GetText
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.mnuPaste_Click"
    Resume ErrExit
End Sub

Private Sub mnuSelectAll_Click()
On Error GoTo ErrSection:
    
    rtfText.SelStart = 0
    rtfText.SelLength = Len(rtfText.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.mnuSelectAll_Click"
    Resume ErrExit
End Sub

Private Sub optArrow_Click(Index As Integer)
On Error GoTo ErrSection:
    
    SetTextEditControls
    Repaint

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnot.optArrow.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optAuto_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optAuto.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optAxisLenData_Click(Index As Integer)
On Error GoTo ErrSection:

    With m.Annot
        If optAxisLenData(1) = True Then
            .geMinorAxisLenData = 1     'points
            txtAxisLen.Text = m.Chart.Bars.PriceDisplay(.geMinorAxisLen)
        ElseIf optAxisLenData(2) = True Then
            .geMinorAxisLenData = 2     'ticks
             txtAxisLen.Text = Format(.geMinorAxisLen / m.Chart.Bars.Prop(eBARS_TickMove), "0")
        Else
            .geMinorAxisLenData = 0     'ratio
            txtAxisLen.Text = Format(.geMinorAxisLen, "0.0##")
        End If
    End With
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optAxisLenData.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optCenter_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optCenter.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optCurly_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optCurly.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optEllipse_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optEllipse.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optExt_Click(Index As Integer)
On Error GoTo ErrSection:

    With m.Annot
        If .eType = eANNOT_Trendline Or .eType = eANNOT_Trendline2 Or _
           .eType = eANNOT_Trendline3 Or .eType = eANNOT_Trendline4 Or _
           .eType = eANNOT_TrendChannel Then
            
            SetTrendlineControls
            If Index = 0 Or Index = 2 Then
                .UpdateAlert 0
            End If
        
        ElseIf .eType = eANNOT_RegressionLine Then
            SetRegressionControls
        End If
    End With
    
    m.bTextChanged = True '(since only text or extensions)
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optExt.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optGannFan_Click(Index As Integer)
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optGannFan.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optLocation_Click(Index As Integer)
    m.Annot.Prop("ChannelLocation") = Index
    SetTrendlineControls
    Repaint
End Sub

Private Sub optPercent_Click()
    Repaint
End Sub

Private Sub optPoints_Click()
    Repaint
End Sub

Private Sub optRatio_Click(Index As Integer)
On Error GoTo ErrSection:

    Dim dRise#

    With m.Annot
        dRise = Val(.Prop("Rise"))
        If optRatio(1) = True Then
            .Prop("GannRatioType") = 1         'ticks
            txtRise.Text = Format(dRise / m.Chart.Bars.Prop(eBARS_TickMove), "0")
        Else
            .Prop("GannRatioType") = 0         'points
            txtRise.Text = m.Chart.Bars.PriceDisplay(dRise)
        End If
    End With
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optRatio.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optLeft_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optLeft.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optPullbacks_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optPullbacks.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optRectangle_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optRectangle.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optRight_Click()
On Error GoTo ErrSection:

    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optRight.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optRounded_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optRounded.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optSquare_Click()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optSquare.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optSRLeft_Click()
    Repaint
End Sub

Private Sub optSRRight_Click()
    Repaint
End Sub

Private Sub optTargets_Click()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.optTargets.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub optWaveConnect_Click(Index As Integer)
    Repaint
End Sub

Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    If Button = 2 Then
        ' prior to popup, enable/disable appropriatly
        mnuSelectAll.Enabled = (Len(rtfText.Text) > 0)
        mnuCut.Enabled = (rtfText.SelLength > 0)
        mnuCopy.Enabled = (rtfText.SelLength > 0)
        mnuPaste.Enabled = (Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(rtfCFText))
        ' now popup the menu
        PopupMenu mnuEdit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.rtfText_MouseUp"
    Resume ErrExit
End Sub

Private Sub tmrEditAnnot_Timer()
On Error GoTo ErrSection:
    
    tmrEditAnnot.Enabled = False
    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.tmrEditAnnot.Timer", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtAxisLen_LostFocus()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtAxisLen.LostFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtDneLabelA_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneLabelA.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelB_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneLabelB.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelC_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneLabelC.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelCOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneLabelCOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneLabelOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneLabelXOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneLabelXOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneRatioCOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneRatioCOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneRatioOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneRatioOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtDneRatioXOP_Change()
On Error GoTo ErrSection:

    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtDneRatioXOP.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtForecastBars_Change()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtForecastBars.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtFromY_Change()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtFromY.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub rtfText_Change()
On Error GoTo ErrSection:
    
    SetTextEditControls
    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.rtfText.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtGannacciMultiply_Change()
    
    If Me.Visible Then
        If Len(txtGannacciMultiply.Text) > 0 And Val(txtGannacciMultiply.Text) <> 0 Then
            Repaint
        End If
    End If

End Sub

Private Sub txtGannToPrice_LostFocus()
On Error GoTo ErrSection:

    SetGannValues 2, ValOfText(txtGannToPrice.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtGannToPrice.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtGannPrice_LostFocus()
On Error GoTo ErrSection:

    SetGannValues 1, ValOfText(txtGannPrice.Text)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtGannPrice.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtPercent_Change()
    Repaint
End Sub

Private Sub txtPoints_Change()
    Repaint
End Sub

Private Sub txtRise_Change()
On Error GoTo ErrSection:

    Dim strPrice$

    'when user changes base or to price, the rise also changes
    'don't want to process this change
    If m.bTextChanged = True Then Exit Sub

    If Left(Trim(txtRise.Text), 1) = "-" Then
        InfBox "Please use only positive numbers for rise value."
        txtRise.Text = Mid(Trim(txtRise.Text), 2)
    Else
        m.bTextChanged = True
        'rise/run
        If Len(Trim(txtRise.Text)) > 0 Then
            If optRatio(1) = True Then      'ticks
                m.Annot.GannRiseChange ValOfText(txtRise.Text) * m.Chart.Bars.Prop(eBARS_TickMove)
            Else
                m.Annot.GannRiseChange m.Chart.Bars.PriceFromString(txtRise.Text)
            End If
        End If
        Repaint
        're-show in case value has changed
        ShowValue "", m.Annot.Y(2), strPrice, False
        txtGannToPrice.Text = strPrice
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtRise.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtRise_LostFocus()
On Error GoTo ErrSection:

    If Len(txtRise.Text) = 0 Then
        If optRatio(0) Then
            optRatio_Click 0
        Else
            optRatio_Click 0
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtRise.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtRiskReward_Change()
    Repaint
End Sub

Private Sub txtRun_Change()
On Error GoTo ErrSection:

    Dim i&, s$
    
    s = Trim(txtRun)
    For i = 1 To Len(s)
        If Not IsDigit(s, i) Then
            InfBox "Please use only positive integers for run value."
            txtRun = StripStr(s, Mid(s, i, 1))
            Exit Sub
        End If
    Next
    
    If Len(s) > 0 Then
        '.Prop("Run") = Int(ValOfText(txtRun.Text))
        m.bTextChanged = True
        m.Annot.GannRunChange m.Chart, Int(ValOfText(txtRun.Text))
        Repaint
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtRun.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtRun_LostFocus()
On Error GoTo ErrSection:

    If Len(txtRun.Text) = 0 Then
        txtRun.Text = Val(m.Annot.Prop("Run"))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtRun.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtRegLineLen_Change()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtRegLineLen.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtStdDevVal_Change()
On Error GoTo ErrSection:
    
    If IsAlpha(txtStdDevVal) Then
        txtStdDevVal.Text = Val(m.Annot.Prop("StdDevVal"))    'aardvark 3096
    Else
        SetRegressionControls
        Repaint
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtStdDevVal.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtText_Change()
On Error GoTo ErrSection:

    SetTextControls False
    m.bTextChanged = True
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtText.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtText_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtText

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtText.GotFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtText_LostFocus()
On Error GoTo ErrSection:

    If m.bTextChanged Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtText.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtToY_Change()
On Error GoTo ErrSection:
    
    Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtToY.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtValue_Change()
On Error GoTo ErrSection:

    m.bTextChanged = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtValue.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub txtValue_GotFocus()
On Error GoTo ErrSection:

    Dim i&
    
    i = Len(txtValue.Text)          '6189
    txtValue.SelStart = i
    txtValue.SelLength = 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtValue_GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtValue_LostFocus()
On Error GoTo ErrSection:

    If m.bTextChanged Then Repaint

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.txtValue.LostFocus", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'Private Sub LoadStyles(cbo As ComboBox)
'On Error GoTo ErrSection:
'
'   This is original code. Replaced with LoadPenStyles.
'   Save awhile then remove - 01/21/2003
'
'    With cbo
'        .Clear
'        If m.Annot.eType = eANNOT_HorzLine Or m.Annot.eType = eANNOT_VertLine Then
'            .AddItem "Thin"
'            .ItemData(.ListCount - 1) = PELT_THINSOLID
'            .AddItem "Medium"
'            .ItemData(.ListCount - 1) = PELT_MEDIUMSOLID
'            .AddItem "Thick"
'            .ItemData(.ListCount - 1) = PELT_THICKSOLID
'            .AddItem "Dashed (Large)"
'            .ItemData(.ListCount - 1) = PELT_DASH
'            .AddItem "Dashed (Small)"'
'            .ItemData(.ListCount - 1) = PELT_DOT
'            .AddItem "Dash Dot"
'            .ItemData(.ListCount - 1) = PELT_DASHDOT
'        Else
'            .AddItem "Thin"
'            .ItemData(.ListCount - 1) = PEGAT_THINSOLIDLINE
'            .AddItem "Medium"
'            .ItemData(.ListCount - 1) = PEGAT_MEDIUMSOLIDLINE
'            .AddItem "Thick"
'            .ItemData(.ListCount - 1) = PEGAT_THICKSOLIDLINE
'            .AddItem "Dashed (Large)"
'            .ItemData(.ListCount - 1) = PEGAT_DASHLINE
'            .AddItem "Dashed (Small)"
'            .ItemData(.ListCount - 1) = PEGAT_DOTLINE
'            .AddItem "Dash Dot"
'            .ItemData(.ListCount - 1) = PEGAT_DASHDOTLINE
'        End If
'        .ListIndex = 0
'    End With
'
'ErrExit:
'    Exit Sub
'
'ErrSection:
'    RaiseError "frmEditAnnot.LoadStyles", eGDRaiseError_Raise
'
'End Sub

Private Sub LoadPenStyles(cbo As ctlUniComboImageXP)
On Error GoTo ErrSection:
    
    With cbo
        .AddItem "Default"
        .ItemData(.ListCount - 1) = eANNOT_Default
        .AddItem "Thin"
        .ItemData(.ListCount - 1) = eANNOT_Thin
        .AddItem "Medium"
        .ItemData(.ListCount - 1) = eANNOT_Medium
        .AddItem "Thick"
        .ItemData(.ListCount - 1) = eANNOT_Thick
        .AddItem "Dashed (Large)"
        .ItemData(.ListCount - 1) = eANNOT_DashLg
        .AddItem "Dashed (Small)"
        .ItemData(.ListCount - 1) = eANNOT_DashSm
        .AddItem "Dash Dot"
        .ItemData(.ListCount - 1) = eANNOT_DashDot
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.LoadPenStyles", eGDRaiseError_Raise
    
End Sub

'Private Sub LoadAnnotIcons()
'On Error GoTo ErrSection:
'
    'Commented out on 09-25-2003
    'Original code. Save for reference and for backwards compatibility conversion.
    'Remove when backwards compatibility is no longer a concern.
    'AddIcon " 0 (none)"
    'AddIcon " 0 (Text)"
    'AddIcon "37 Pointer for text"
'
'    cboStyle.Clear
'
    'Remember to consider backwards compatibility if changing icon numbers.
'    AddIcon "92 Arrow Up"
'    AddIcon "96 Arrow Down"
'    AddIcon "94 Arrow Right"
'    AddIcon "98 Arrow Left"
'    AddIcon "93 Arrow NE"
'    AddIcon "95 Arrow SE"
'    AddIcon "97 Arrow SW"
'    AddIcon "99 Arrow NW"
'    AddIcon "13 Small Plus"
'    AddIcon "14 Small Cross"
'    AddIcon "15 Small Circle"
'    AddIcon "16 Small Solid Circle"
'    AddIcon "17 Small Square"
'    AddIcon "18 Small Solid Square"
'    AddIcon "19 Small Diamond"
'    AddIcon "20 Small Solid Diamond"
'    AddIcon "21 Small Upward Triangle"
'    AddIcon "22 Small Solid Upward Triangle"
'    AddIcon "23 Small Downward Triangle"
'    AddIcon "24 Small Solid Downward Triangle"
'    AddIcon " 1 Medium Plus"
'    AddIcon " 2 Medium Cross"
'    AddIcon " 3 Medium Circle"
'    AddIcon " 4 Medium Solid Circle"
'    AddIcon " 5 Medium Square"
'    AddIcon " 6 Medium Solid Square"
'    AddIcon " 7 Medium Diamond"
'    AddIcon " 8 Medium Solid Diamond"
'    AddIcon " 9 Medium Upward Triangle"
'    AddIcon "10 Medium Solid Upward Triangle"
'    AddIcon "11 Medium Downward Triangle"
'    AddIcon "12 Medium Solid Downward Triangle"
'    AddIcon "25 Large Plus"
'    AddIcon "26 Large Cross"
'    AddIcon "27 Large Circle"
'    AddIcon "28 Large Solid Circle"
'    AddIcon "29 Large Square"
'    AddIcon "30 Large Solid Square"
'    AddIcon "31 Large Diamond"
'    AddIcon "32 Large Solid Diamond"
'    AddIcon "33 Large Upward Triangle"
'    AddIcon "34 Large Solid Upward Triangle"
'    AddIcon "35 Large Downward Triangle"
'    AddIcon "36 Large Solid Downward Triangle"
'
'ErrExit:
'    Exit Sub
'
'ErrSection:
'    RaiseError "frmEditAnnot.LoadAnnotIcons", eGDRaiseError_Raise
'
'End Sub

'Private Sub AddIcon(ByVal strText$)
'On Error GoTo ErrSection:
'
    'Commented out on 09-25-2003 (used only by LoadAnnotIcons)
'    Dim nIcon&
'
'    nIcon = Val(Left(strText, 2))
'    strText = Trim(Mid(strText, 3))
'    With cboStyle
'        .AddItem strText
'        .ItemData(.ListCount - 1) = nIcon
'    End With
'
'ErrExit:
'    Exit Sub
'
'ErrSection:
'    RaiseError "frmEditAnnot.AddIcon", eGDRaiseError_Raise
'End Sub

Private Sub SetCombo(cbo As ctlUniComboImageXP, nValue&)
On Error GoTo ErrSection:

    Dim i&, nMatch&

    For i = 0 To cbo.ListCount - 1
        If nValue = cbo.ItemData(i) Then
            nMatch = i
            Exit For
        End If
    Next
    If nMatch >= 0 And nMatch < cbo.ListCount Then
        cbo.ListIndex = nMatch
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnot.SetCombo", eGDRaiseError_Raise

End Sub

Private Sub SetRegressionControls()
On Error GoTo ErrSection:

    Dim dStdDevVal#, i&

    dStdDevVal = ValOfText(txtStdDevVal.Text)
    If dStdDevVal > 0# Then
        chkStdDevOnOff.Value = 1
        SetCombo cboStdDevStyle, Val(m.Annot.Prop("StdDevStyle"))
        i = Int(Val(m.Annot.Prop("ChannelCount")))
        If i <= 0 Then
            m.Annot.Prop("ChannelCount") = 1
            i = 1
        ElseIf i > 9 Then
            m.Annot.Prop("ChannelCount") = 9
            i = 9
        
        End If
        cboChannels.ListIndex = i - 1
    Else
        'chkStdDevOnOff.Value = 0
        cboChannels.ListIndex = 0
    End If
    'enable /disable std deviation controls
    lblChannels.Enabled = chkStdDevOnOff.Value
    cboChannels.Enabled = chkStdDevOnOff.Value
    lblStdDevStyle.Enabled = chkStdDevOnOff.Value
    cboStdDevStyle.Enabled = chkStdDevOnOff.Value
    
    'show radio buttons for fixed / grow length (aardvark 2914)
    If chkDynamic.Value = 1 Then
        fraRegression.Height = kFraRegressionHt + optPoints.Height + 30
        optPoints.Move chkDynamic.Left + 180, fraRegression.Height - optPoints.Height - 85
        optPercent.Move optPoints.Left + optPoints.Width + 160, optPoints.Top
        optPoints.Visible = True
        optPercent.Visible = True
    Else
        fraRegression.Height = kFraRegressionHt
        optPoints.Visible = False
        optPercent.Visible = False
    End If
    
    'set multichart option
    chkMultiChart.Enabled = m.bMultiChartOption
    
    'set extension frame
    fraExt.Move fraButtons.Left, fraRegression.Top + fraRegression.Height + 120
    If optExt(0) = True Then
        fraExt.Height = clrExt.Top
    Else
        fraExt.Height = 1335
    End If
    SetBottom fraExt

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnot.SetCombo", eGDRaiseError_Raise

End Sub

Private Sub SetTextControls(ByVal bSetTextCtl As Boolean)
On Error GoTo ErrSection:

    With m.Annot
        If bSetTextCtl Then txtText = .Text
        Select Case m.Annot.geTextAlign
        Case 0, 3, 6    'e_topLeft, e_btmLeft, e_ctrLeft
            optLeft = True
        Case 1, 4, 7    'e_topRight, e_btmRight, e_ctrRight
            optRight = True
        Case 2, 5, 8    'e_topCtr, e_btmCtr, e_ctrCtr
            optCenter = True
        Case Else       'e_autoAlign
            optAuto = True
        End Select
    End With

    ' text alignment
    If Len(Trim(txtText)) > 0 Then
        Enable optAuto
        Enable optLeft
        Enable optCenter
        Enable optRight
        Enable lblTextAlign
        Enable cmdFont
    Else
        Disable optAuto
        Disable optLeft
        Disable optCenter
        Disable optRight
        Disable lblTextAlign
        Disable cmdFont
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditAnnot.SetTextControls", eGDRaiseError_Raise

End Sub

Private Sub SetTextEditControls()
On Error GoTo ErrSection:
    
    If Len(Trim(rtfText.Text)) = 0 Then
        Disable cmdFont
        Disable lblTextAlign
        Disable cboAnchor
        Disable lblTextJustify
        Disable cboTextJustify
        Disable chkTextBorder
    Else
        Enable cmdFont
        Enable lblTextAlign
        Enable cboAnchor
        Enable lblTextJustify
        Enable cboTextJustify
        Enable chkTextBorder
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.SetTextEditControls", eGDRaiseError_Raise
    
End Sub

Private Sub SetTrendlineControls()
On Error GoTo ErrSection:

    Dim i&

    chkDynamic.Visible = False
    
    If optLocation(0).Value = True Then
        fraRegression.Height = fraLocation.Height + 300
        lblChannels.Visible = False
        cboChannels.Visible = False
        lblIndicator.Visible = False
        cboIndicator.Visible = False
    Else
        fraRegression.Height = 2295     '2025
        lblChannels.Visible = True
        cboChannels.Visible = True
        lblIndicator.Visible = True
        cboIndicator.Visible = True
    End If
    
    If m.bMultiChartOption = True Then
        chkMultiChart.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 15
        chkShowValueInAxis.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height - 15
    Else
        chkShowValueInAxis.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 15
    End If
    
    If m.Annot.eType = eANNOT_TrendChannel Then
        chkAllPanes.Move chkShowValueInAxis.Left, chkShowValueInAxis.Top + chkShowValueInAxis.Height + 60, chkShowValueInAxis.Width
        fraRegression.Move fraButtons.Left, chkAllPanes.Top + chkAllPanes.Height + 120
    Else
        fraRegression.Move fraButtons.Left, chkShowValueInAxis.Top + chkShowValueInAxis.Height + 120
    End If
    
    fraExt.Move fraButtons.Left, fraRegression.Top + fraRegression.Height + 120
    optPoints.Visible = True
    optPercent.Visible = True
    chkShowValueInAxis.Visible = True
    i = Val(m.Annot.Prop("ShowInAxis"))
    
    'JM 08-10-2015: mod to show value in y-axis:
    '-1=label channels only, 0=no labels, 1=label main line only, >1=label main line & channels
    Select Case i
        Case -1:
            chkShowValueInAxis.Value = vbUnchecked
            chkStdDevOnOff.Value = vbChecked
        Case 0:
            chkShowValueInAxis.Value = vbUnchecked
            chkStdDevOnOff.Value = vbUnchecked
        Case 1:
            chkShowValueInAxis.Value = vbChecked
            chkStdDevOnOff.Value = vbUnchecked
        Case Else:
            chkShowValueInAxis.Value = vbChecked
            chkStdDevOnOff.Value = vbChecked
    End Select
            
    ' see if show text or extensions
    If optExt(0) = True Then
        If fraExt.Height >= 1335 Then    'full height
            fraExt.Height = fraExt.Height - (clrExt.Height + cboExtStyle.Height + 100)
        End If
        fraText.Move fraText.Left, fraExt.Top + fraExt.Height + 100
        SetVisible cboExtStyle, False
        SetVisible lblExtStyle, False
        SetVisible clrExt, False
        SetVisible lblExtColor, False
        SetVisible fraText, True
    Else
        fraExt.Height = 1335            'full height
        SetVisible cboExtStyle, True
        SetVisible lblExtStyle, True
        SetVisible clrExt, True
        SetVisible lblExtColor, True
        SetVisible fraText, False
    End If
            
    If optExt(0) = True Then
        SetBottom fraText
    Else
        SetBottom fraExt
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.SetTrendlineControls", eGDRaiseError_Raise

End Sub

Private Sub ShowValue(ByVal strLabel$, ByVal vValue As Variant, _
    Optional strReturnVal$, Optional ByVal bSetControl As Boolean = True)
On Error GoTo ErrSection:

    Dim i&, iPane&, nAxisLenData&, strValue$
    
    strValue = CStr(vValue)
    
    With m.Annot
        If .Pane = "PRICE PANE" Then
            i = 0
            If .eType = eANNOT_DollarLine Or .eType = eANNOT_DollarLine2 Or .eType = eANNOT_DollarLine3 Or _
               .eType = eANNOT_DollarLine4 Then
                
                If bSetControl = False Then i = 1
            Else
                i = 1
            End If
            
            If i = 1 Then
                    iPane = m.Chart.Tree.Index(.Pane)
                    If iPane > 0 Then
                        strValue = m.Chart.PriceDisplay(iPane, vValue)
                    End If
            End If
        Else
            strValue = Format(vValue, "0.00#")
        End If
    End With

    If bSetControl = True Then
        lblValue.Caption = strLabel
        If m.Annot.eType = eANNOT_HorzLine Or m.Annot.eType = eANNOT_HorzLine2 Or _
           m.Annot.eType = eANNOT_HorzLine3 Or m.Annot.eType = eANNOT_HorzLine4 Then
            txtValue.Left = Me.TextWidth(strLabel) + lblValue.Left + 120
        Else
            lblValue.Left = txtFromY.Left
        End If
        txtValue.Text = strValue
        fraValue.Visible = True
        If m.Annot.eType <> eANNOT_SRLine And m.Annot.eType <> eANNOT_SRLine2 And m.Annot.eType <> eANNOT_SRLine3 And m.Annot.eType <> eANNOT_SRLine4 Then
            optSRLeft.Visible = False
            optSRRight.Visible = False
        End If
    Else
        strReturnVal = strValue
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.ShowValue", eGDRaiseError_Raise
    
End Sub

Private Sub SetBottom(ctlBottom As Control)
On Error GoTo ErrSection:

    If ctlBottom.Visible = False Then ctlBottom.Visible = True
    fraButtons.Top = ctlBottom.Top + ctlBottom.Height
    Me.Height = fraButtons.Top + fraButtons.Height + Me.Height - Me.ScaleHeight
    If Me.Visible Then Me.Refresh
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.SetBottom", eGDRaiseError_Raise
    
End Sub

Private Sub InitGannLinesControls()
On Error GoTo ErrSection:

    Dim strPrice$, i&, bEnable As Boolean
    
    Me.Caption = "Gann Angle/Fan"
    Me.Icon = Picture16(ToolbarIcon("ID_GannLines"), , True)

    'show needed frame(s)
    fraGannOptions.Visible = True
    fraQuadrants.Visible = True
    fraGannLines.Visible = True
    
    fraGannOptions.Move chkMultiChart.Left - 50, chkMultiChart.Top + chkMultiChart.Height + 60
    fraQuadrants.Move chkMultiChart.Left - 40, fraGannOptions.Top + fraGannOptions.Height + 60
    fraGannLines.Move chkMultiChart.Left - 40, fraQuadrants.Top + fraQuadrants.Height + 60
    
    SetBottom fraGannLines
    
    'set controls values
    With m.Annot
        ShowValue "", .Y(1), strPrice, False
        txtGannPrice.Text = strPrice
        ShowValue "", .Y(2), strPrice, False
        txtGannToPrice.Text = strPrice
        'lines directions
        SetQuadrantChkBoxes
        'rise/run ratio
        txtRun.Text = Val(.Prop("Run"))
        optRatio(.Prop("GannRatioType")) = True
        'fan check box
        If .Prop("GannFan") = 0 Then
            bEnable = False
            chkGannLines(0) = 0
        Else
            bEnable = True
            chkGannLines(0) = 1
        End If
        'lines to show
        If .Prop("1x2") = 1 Then chkGannLines(1) = 1
        If .Prop("1x3") = 1 Then chkGannLines(2) = 1
        If .Prop("1x4") = 1 Then chkGannLines(3) = 1
        If .Prop("1x8") = 1 Then chkGannLines(4) = 1
        If .Prop("2x1") = 1 Then chkGannLines(5) = 1
        If .Prop("3x1") = 1 Then chkGannLines(6) = 1
        If .Prop("4x1") = 1 Then chkGannLines(7) = 1
        If .Prop("8x1") = 1 Then chkGannLines(8) = 1
        'color of each line pairs
        clrGannColor(0).Color = Int(Val(.Prop("Color1x2")))
        clrGannColor(1).Color = Int(Val(.Prop("Color1x3")))
        clrGannColor(2).Color = Int(Val(.Prop("Color1x4")))
        clrGannColor(3).Color = Int(Val(.Prop("Color1x8")))
    End With
    
    For i = 0 To 3
        clrGannColor(i).Enabled = bEnable
    Next
    
    For i = 1 To 8
        chkGannLines(i).Enabled = bEnable
    Next
    
    m.bCenterColorStyle = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitGannLinesControls", eGDRaiseError_Raise

End Sub

Private Sub InitHorzlineControls()
On Error GoTo ErrSection:
   
    If m.Annot Is Nothing Then
        Me.Caption = "Horizontal Line"
        Me.Icon = Picture16(ToolbarIcon("ID_HorzLine"), , True)
    Else
        Select Case m.Annot.eType
            Case eANNOT_HorzLine2
                Me.Caption = "Horizontal Line 2"
                Me.Icon = Picture16(ToolbarIcon("ID_HorzLine2"), , True)
            Case eANNOT_HorzLine3
                Me.Caption = "Horizontal Line 3"
                Me.Icon = Picture16(ToolbarIcon("ID_HorzLine3"), , True)
            Case eANNOT_HorzLine4
                Me.Caption = "Horizontal Line 4"
                Me.Icon = Picture16(ToolbarIcon("ID_HorzLine4"), , True)
            Case Else
                Me.Caption = "Horizontal Line"
                Me.Icon = Picture16(ToolbarIcon("ID_HorzLine"), , True)
        End Select
    End If
    
    'show needed frame(s)
    fraValue.Height = fraValue.Height - txtValue.Height - txtToY.Height - 120
    If m.bMultiChartOption = True Then
        chkShowValueInAxis.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
    Else
        chkShowValueInAxis.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    End If
    chkShowValueInAxis.Value = Val(m.Annot.Prop("ShowInAxis"))
    chkShowValueInAxis.Visible = True
    fraValue.Top = chkShowValueInAxis.Top + chkShowValueInAxis.Height
    SetBottom fraValue
    
    'set controls values
    'RH commented out fraValue.BorderStyle = 0
    ShowValue "Value:", m.Annot.Y(1)
    
    'show alert button
    If ExtremeCharts <> 1 Then
        cmdAlert.Move fraValue.Width - cmdAlert.Width, fraValue.Top + txtValue.Top
        cmdAlert.Visible = True
        cmdAlert.Enabled = m.Annot.CanHaveAlert     'aardvark 3532
        cmdAlert.ZOrder
    End If
        
    m.bCenterColorStyle = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitHorzlineControls", eGDRaiseError_Raise

End Sub

Private Sub InitDollarlineControls(Chart As cChart)
On Error GoTo ErrSection:

    Dim strText$, i&, d#
    
    If m.Annot Is Nothing Then Exit Sub
    
    If m.Annot.eType = eANNOT_DollarLine2 Then
        Me.Caption = "Dollar Difference Line 2"
        Me.Icon = Picture16(ToolbarIcon("ID_DollarLine2"), , True)
    ElseIf m.Annot.eType = eANNOT_DollarLine3 Then
        Me.Caption = "Dollar Difference Line 3"
        Me.Icon = Picture16(ToolbarIcon("ID_DollarLine3"), , True)
    ElseIf m.Annot.eType = eANNOT_DollarLine4 Then
        Me.Caption = "Dollar Difference Line 4"
        Me.Icon = Picture16(ToolbarIcon("ID_DollarLine4"), , True)
    Else
        Me.Caption = "Dollar Difference Line"
        Me.Icon = Picture16(ToolbarIcon("ID_DollarLine"), , True)
    End If
    
    'show needed frame(s)
    fraValue.Visible = True
    optSRLeft.Visible = False
    optSRRight.Visible = False
    fraDLineText.Visible = True
    
    'RH commented out fraValue.BorderStyle = 1
    If m.bMultiChartOption = True Then
        fraValue.Top = chkMultiChart.Top + chkMultiChart.Height + 120
    Else
        fraValue.Top = chkPreIndicator.Top + chkPreIndicator.Height + 120
    End If
    
    fraValue.Width = fraValue.Width - 100
    fraDLineText.Width = fraValue.Width
    
    'show needed controls
    lblFrom.Visible = True
    lblTo.Visible = True
    txtFromY.Visible = True
    txtToY.Visible = True
    
    'hide not needed controls
    chkExtendSRLine.Visible = False
    chkDisplaySRValue.Visible = False
    
    'show controls not within any frame
    cmdFont.Visible = True
    chkDynamic.Visible = True
    chkDynamic.Caption = "Dynamic (stay at end of data)"
    
    'set controls values
    lblValue.Visible = True
    txtValue.Visible = True
    txtValue.Enabled = True
    Select Case Chart.Bars.Prop(eBARS_SecurityType)
    Case Asc("S")
        strText = "Number of shares"
        txtValue.Tag = "NumShares"
    Case Asc("F")
        strText = "Number of contracts"
        txtValue.Tag = "NumContracts"
    Case Else
        If IsForex(Chart.Bars.Prop(eBARS_Symbol)) Then
            strText = "Number of contracts"
            txtValue.Tag = "NumForexContracts"
        Else
            strText = "Multiplier"
            txtValue.Tag = "IndexMult"
        End If
    End Select
    ShowValue strText, Trim(m.Annot.Prop(txtValue.Tag))
    fraDLineText.Move fraValue.Left, fraValue.Top + fraValue.Height + 100
    SetBottom fraDLineText
    
    With m.Annot
        If .X(2) >= .X(1) Then
            ShowValue "", .Y(1), strText, False
            txtFromY.Text = strText
            ShowValue "", .Y(2), strText, False
            txtToY.Text = strText
        Else
            ShowValue "", .Y(2), strText, False
            txtFromY.Text = strText
            ShowValue "", .Y(1), strText, False
            txtToY.Text = strText
        End If
        
        chkDynamic.Visible = True
        chkDynamic.Enabled = True
        chkDynamic.Value = Val(.Prop("KeepAtEnd"))
        If chkDynamic.Value = 1 Then
            txtToY.Enabled = False
        Else
            txtToY.Enabled = True
        End If
        
        For i = 0 To 5
            chkDLineText(i).Visible = True
            chkDLineText(i).Enabled = True
        Next
        
        chkDLineText(0).Value = Val(.Prop("ProfitLoss"))
        chkDLineText(1).Value = Val(.Prop("ProfitLossPercent"))
        chkDLineText(2).Value = Val(.Prop("PriceMove"))
        chkDLineText(3).Value = Val(.Prop("SlopeLine"))       '6986
        chkDLineText(4).Value = Val(.Prop("Volume"))
        chkDLineText(5).Value = Val(.Prop("LenBars"))

        If chkDLineText(5).Value = vbChecked Then
            chkDLineText(6).Enabled = True
            chkDLineText(6).Value = Val(.Prop("LenHzLine"))
        Else
            chkDLineText(6).Enabled = False
        End If
        
        'JM 09-01-2015: client request from Heath's/Tim's email (�Show price move in points� checkbox)
        chkDLineText(7).Visible = True
        chkDLineText(7).Enabled = True
        chkDLineText(7).Value = Val(.Prop("PriceMovePoints"))
        
        'reposition/resize controls as needed
        chkDynamic.Move lblFrom.Left + 100, (fraValue.Top + fraValue.Height) - chkDynamic.Height - 80
        chkDynamic.ZOrder
    End With
    
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitDollarlineControls", eGDRaiseError_Raise

End Sub

Private Sub InitBracketControls()
On Error GoTo ErrSection:
            
    Me.Caption = "Bracket"
    Me.Icon = Picture16(ToolbarIcon("ID_Bracket"), , True)
    
    'show needed frame(s)
    fraBracket.Visible = True
    If m.bMultiChartOption = True Then
        fraBracket.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 120
    Else
        fraBracket.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 120
    End If
    SetBottom fraBracket
                
    'set controls values
    Select Case Val(m.Annot.Prop("BracketStyle"))
    Case 1
        optSquare = True
    Case Else
        optCurly = True
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitBracketControls", eGDRaiseError_Raise

End Sub

Private Sub InitRegressionlineControls()
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    Dim nSelIndId&, nSelPaneId&, nSelPriceField&
    Dim nCboIdx&        ', i&, k&
    
    Me.Caption = "Regression Line/Channel"
    Me.Icon = Picture16(ToolbarIcon("ID_RegressionLine"), , True)
              
    'show neeed frame(s)
    fraRegression.Visible = True
    fraExt.Visible = True
    fraExt.Caption = "Extensions"
     
    'show controls for regression line
    chkDynamic.Visible = True
    lblRegressionLength.Visible = True
    lblBars.Visible = True
    txtRegLineLen.Visible = True
    txtStdDevVal.Visible = True
    chkStdDevOnOff.Visible = True
    cboStdDevStyle.Visible = True
    lblIndicator.Caption = "Indicator"
    chkShowValueInAxis.Visible = True
    
    'hide controls for trend lines
    fraLocation.Visible = False
    txtPoints.Visible = False
    txtPercent.Visible = False
    
    'these radio buttons are used by both trend & regression lines
    optPoints.Visible = False
    optPercent.Visible = False
    optPoints.Caption = "Fixed length"
    optPercent.Caption = "Grow length"
    optPoints.Width = optPoints.Width * 1.5
    optPercent.Width = optPoints.Width
                    
    'number of channels
    cboChannels.Clear
    cboChannels.AddItem "1"
    cboChannels.AddItem "2"
    cboChannels.AddItem "3"
    cboChannels.AddItem "4"
    cboChannels.AddItem "5"
    cboChannels.AddItem "6"
    cboChannels.AddItem "7"
    cboChannels.AddItem "8"
    cboChannels.AddItem "9"
    
    'set controls values
    With m.Annot
        optExt(Val(.Prop("Ext"))) = True
        clrExt.Color = Val(.Prop("ExtColor"))
        LoadPenStyles cboExtStyle
        SetCombo cboExtStyle, Val(.Prop("ExtStyle"))
        chkDynamic.Caption = "Dynamic (stay at end of visible data)"
        chkDynamic.Value = Val(.Prop("KeepAtEnd"))
        optPoints.Value = Val(.Prop("FixLength"))
        optPercent.Value = Not Val(.Prop("FixLength"))
        txtStdDevVal.Text = Val(.Prop("StdDevVal"))
        txtRegLineLen.Text = CStr(Abs(.X(2) - .X(1)))
        chkShowValueInAxis.Value = Val(m.Annot.Prop("ShowInAxis"))
    End With
    If Val(txtStdDevVal.Text) = 0 Then
        chkStdDevOnOff.Value = 0
    Else
        chkStdDevOnOff.Value = 1
    End If
    LoadPenStyles cboStdDevStyle
    
    cboIndicator.Clear
    Set Ind = m.Chart.Tree(m.Annot.geIndId)
    If Not Ind Is Nothing Then
        nSelIndId = Ind.geIndId
        nSelPaneId = Ind.geIndpaneId
    End If
    
    nSelPriceField = Val(m.Annot.Prop("PriceField"))
    nCboIdx = PopulateIndicatorsCbo(cboIndicator, m.Chart, nSelIndId, nSelPaneId, nSelPriceField, True)
    
    Set Ind = Nothing
    cboIndicator.ListIndex = nCboIdx
    m.bCenterColorStyle = True

    'position controls
    fraRegression.Height = kFraRegressionHt
    
    chkShowValueInAxis.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
    fraRegression.Move fraButtons.Left, chkShowValueInAxis.Top + chkShowValueInAxis.Height + 120
    lblChannels.Move lblStdDevStyle.Left, lblStdDevStyle.Top + lblStdDevStyle.Height + 80 'number of channels
    cboChannels.Move cboStdDevStyle.Left, lblChannels.Top
    chkDynamic.Move lblRegressionLength.Left + fraRegression.Left, fraRegression.Top + kFraRegressionHt - chkDynamic.Height - 120
    chkDynamic.ZOrder
        
    SetRegressionControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitRegressionlineControls", eGDRaiseError_Raise

End Sub

Private Sub InitTargetshooterControls()
On Error GoTo ErrSection:

    Me.Caption = "Target Shooter"
    Me.Icon = Picture16(ToolbarIcon("ID_TargetShooter"), , True)
    
    'show needed frame(s)
    fraTarget.Visible = True
    If m.bMultiChartOption = True Then
        fraTarget.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 120
    Else
        fraTarget.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 120
    End If
    SetBottom fraTarget
    
    'show controls not within any frame
    cmdFont.Visible = True
    
    'set control values
    If Val(m.Annot.Prop("Pullbacks")) <> 0 Then
        optPullbacks = True
    Else
        optTargets = True
    End If
    clrTargets.Color = Val(m.Annot.Prop("TargetColor"))
    LoadPenStyles cboTargetStyle
    SetCombo cboTargetStyle, Val(m.Annot.Prop("TargetStyle"))
    
    SetCtl chk1st, Val(m.Annot.Prop("Pt1"))
    SetCtl chk2nd, Val(m.Annot.Prop("Pt2"))
    SetCtl chk3rd, Val(m.Annot.Prop("Pt3"))
    SetCtl chkTargetValues, Val(m.Annot.Prop("ShowValues"))
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitTargetshooterControls", eGDRaiseError_Raise

End Sub

Private Sub InitTextEditControls()
On Error GoTo ErrSection:

    Dim i&

    Select Case m.Annot.eType
        Case eANNOT_ArrowLine
            Me.Caption = "Arrow"
            Me.Icon = Picture16(ToolbarIcon("ID_ArrowLine"), , True)
        Case eANNOT_TextEdit
            Me.Caption = "Text"
            Me.Icon = Picture16(ToolbarIcon("ID_Text"), , True)
        Case eANNOT_TextEdit2
            Me.Caption = "Text 2"
            Me.Icon = Picture16(ToolbarIcon("ID_Text2"), , True)
        Case eANNOT_TextEdit3
            Me.Caption = "Text 3"
            Me.Icon = Picture16(ToolbarIcon("ID_Text3"), , True)
        Case eANNOT_TextEdit4
            Me.Caption = "Text 4"
            Me.Icon = Picture16(ToolbarIcon("ID_Text4"), , True)
    End Select
    
    'hide controls within a frame not used by this annotation
    optLeft.Visible = False
    optCenter.Visible = False
    optRight.Visible = False
    optAuto.Visible = False
    txtText.Visible = False
    
    'show needed frame(s)
    If m.Annot.eType = eANNOT_ArrowLine Then
        lblStyle.Visible = False
        cboStyle.Visible = False
        cboStyle.Enabled = False
        optArrow(0).Visible = False
        optArrow(0).Enabled = False
        chkPreIndicator.Move chkPreIndicator.Left, cboStyle.Top + 50
        If m.bMultiChartOption = True Then
            chkMultiChart.Move chkMultiChart.Left, chkPreIndicator.Top + chkPreIndicator.Height + 120
            fraArrow.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 140
        Else
            fraArrow.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 140
        End If
        fraArrow.Caption = "Arrow Options"
        m.bCenterColorStyle = True
        SetBottom fraArrow
        cboLineStyle.Clear
        LoadPenStyles cboLineStyle
        SetCombo cboLineStyle, Val(m.Annot.Prop("ArrowLine"))
        optArrow(Val(m.Annot.Prop("ArrowStyle"))) = True
        With cboArrowSize
            .Clear
            .AddItem "Small"
            .AddItem "Medium"
            .AddItem "Large"
            i = Val(m.Annot.Prop("ArrowSize"))
            If i >= 0 And i < .ListCount Then
                .ListIndex = i
            Else
                .ListIndex = 0
            End If
        End With
        Exit Sub
    End If
    
    fraText.Visible = True
    fraText.Caption = "Text  (hit 'Enter' to start a new line)"
    If m.bMultiChartOption = True Then
        fraText.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 120
    Else
        fraText.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 120
    End If
    fraText.Height = fraText.Height * 2
    fraArrow.Move fraText.Left, fraText.Top + fraText.Height + 120
    SetBottom fraArrow
    
    'show controls not within any frame
    cmdFont.Visible = True
    
    'set controls values
    chkTextBorder.Value = Val(m.Annot.Prop("Border"))
    rtfText.Text = m.Annot.Text
    lblTextAlign.Alignment = 0
    lblTextAlign.Caption = "Anchor Text"
    cboLineStyle.Clear
    LoadPenStyles cboLineStyle
    SetCombo cboLineStyle, Val(m.Annot.Prop("ArrowLine"))
    optArrow(Val(m.Annot.Prop("ArrowStyle"))) = True
    With cboArrowSize
        .Clear
        .AddItem "Small"
        .AddItem "Medium"
        .AddItem "Large"
        i = Val(m.Annot.Prop("ArrowSize"))
        If i >= 0 And i < .ListCount Then
            .ListIndex = i
        Else
            .ListIndex = 0
        End If
    End With
    With cboAnchor
        .Clear
        .AddItem "Top left corner"
        .AddItem "Top right corner"
        .AddItem "Top side center"
        .AddItem "Bottom left corner"
        .AddItem "Bottom right corner"
        .AddItem "Bottom side center"
        .AddItem "Left side center"
        .AddItem "Right side center"
        .AddItem "Center"
        i = m.Annot.geTextAlign
        If i >= 0 And i < .ListCount Then
            .ListIndex = i
        Else
            .ListIndex = 0
        End If
    End With
    
    With cboTextJustify
        .Clear
        .AddItem "Left"
        .AddItem "Right"
        .AddItem "Center"
        i = Val(m.Annot.Prop("TextJustify"))
        If i >= 0 And i < .ListCount Then
            .ListIndex = i
        Else
            .ListIndex = 0
        End If
    End With
    
    SetTextEditControls

    'reposition/resize controls as needed
    rtfText.Move txtText.Left, txtText.Top, txtText.Width, txtText.Height * 3
    lblTextAlign.Move rtfText.Left + 50, rtfText.Height + 360
    lblTextJustify.Move lblTextAlign.Left, lblTextAlign.Top + lblTextAlign.Height + 100
    cboAnchor.Move cboAnchor.Left, rtfText.Height + 320
    cboTextJustify.Move cboAnchor.Left, cboAnchor.Top + cboAnchor.Height + 30
    chkTextBorder.Move lblTextAlign.Left, lblTextJustify.Top + lblTextJustify.Height + 100

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitTextEditControls", eGDRaiseError_Raise

End Sub

Private Sub InitEllipseControls()
On Error GoTo ErrSection:
   
    Dim nLenDataType&
   
    Me.Caption = "Ellipse"
    Me.Icon = Picture16(ToolbarIcon("ID_Ellipse"), , True)

    'show needed frame(s)
    fraEllipse.Visible = True
    fraEllipse.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 120
    SetBottom fraEllipse

    'set controls values
    With m.Annot
        nLenDataType = Int(Val(m.Annot.Prop("MinorAxisLenData")))
        optAxisLenData(nLenDataType) = True
        chkAxes = Val(.Prop("ShowAxes"))
        chkQtrLines = Val(.Prop("ShowQtrLines"))
        If nLenDataType = 2 Then        'ticks
            txtAxisLen.Text = Format(.geMinorAxisLen / m.Chart.Bars.Prop(eBARS_TickMove), "0")
        ElseIf nLenDataType = 1 Then    'points
            txtAxisLen.Text = m.Chart.Bars.PriceDisplay(.geMinorAxisLen)
        Else                            'ratio
            txtAxisLen.Text = Format(.geMinorAxisLen, "0.0##")
        End If
    End With
    
    m.bCenterColorStyle = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitEllipseControls", eGDRaiseError_Raise
End Sub

Private Sub InitTrendlineControls()
On Error GoTo ErrSection:

    Dim i&
    
    If m.Annot Is Nothing Then
        Me.Caption = "Trendline"
        Me.Icon = Picture16(ToolbarIcon("ID_Trendline"), , True)
    Else
        Select Case m.Annot.eType
            Case eANNOT_TrendChannel
                Me.Caption = "Trend Channel"
                Me.Icon = Picture16(ToolbarIcon("ID_TrendChannel"), , True)
            Case eANNOT_Trendline
                Me.Caption = "Trendline"
                Me.Icon = Picture16(ToolbarIcon("ID_Trendline"), , True)
            Case eANNOT_Trendline2
                Me.Caption = "Trendline 2"
                Me.Icon = Picture16(ToolbarIcon("ID_Trendline2"), , True)
            Case eANNOT_Trendline3
                Me.Caption = "Trendline 3"
                Me.Icon = Picture16(ToolbarIcon("ID_Trendline3"), , True)
            Case eANNOT_Trendline4
                Me.Caption = "Trendline 4"
                Me.Icon = Picture16(ToolbarIcon("ID_Trendline4"), , True)
        End Select
    End If
    
    'hide controls within a frame not used by this annotation
    chkTextBorder.Visible = False
    cboAnchor.Visible = False
    cboTextJustify.Visible = False
    lblTextJustify.Visible = False
    rtfText.Visible = False
    
    'show needed frame(s)
    fraRegression.Caption = "Channels"
    fraRegression.Visible = True
    fraExt.Caption = "Extensions"
    fraExt.Visible = True
    
    'these items are in the regression line frame
    'RH commented out fraLocation.BorderStyle = 0
    fraLocation.Visible = True
    lblIndicator.Caption = "Style"
    cboStdDevStyle.Visible = False
    lblStdDevStyle.Visible = False
    cboChannels.Visible = True
    lblChannels.Visible = True
    lblRegressionLength.Visible = False
    lblBars.Visible = False
    txtRegLineLen.Visible = False
    txtStdDevVal.Visible = False
    chkStdDevOnOff.Visible = True
    
    'adjust captions for a couple of check boxes
    chkShowValueInAxis.Caption = "Show value in axis (main line)"
    chkStdDevOnOff.Caption = "Show value in axis (channels)"
    
    'these radio buttons are used by both trend & regression lines
    optPoints.Visible = True
    optPercent.Visible = True
    optPoints.Caption = "Points"
    optPercent.Caption = "% Price"
    txtPercent.Visible = True
    txtPoints.Visible = True
                                             
    'show controls not within any frame
    cmdFont.Visible = True
    If m.Annot.eType = eANNOT_Trendline Or m.Annot.eType = eANNOT_Trendline2 Or _
       m.Annot.eType = eANNOT_Trendline3 Or m.Annot.eType = eANNOT_Trendline4 Or _
       m.Annot.eType = eANNOT_TrendChannel Then
        
        If ExtremeCharts <> 1 Then
            cmdAlert.Visible = True
            cmdAlert.Move cmdFont.Left, cboStyle.Top
            cmdAlert.Enabled = m.Annot.CanHaveAlert     'aardvark 3298
        End If
            
        If m.Annot.eType = eANNOT_TrendChannel Then
            chkAllPanes.Caption = "Hide main trend line"        '6717
            chkAllPanes.Visible = True
            chkAllPanes.Value = Val(m.Annot.Prop("HideMainLine"))
        End If
    End If
    
    'set controls values
    SetTextControls True
    optExt(Val(m.Annot.Prop("Ext"))) = True
    optLocation(Val(m.Annot.Prop("ChannelLocation"))) = True
    clrExt.Color = Val(m.Annot.Prop("ExtColor"))
    LoadPenStyles cboExtStyle
    SetCombo cboExtStyle, Val(m.Annot.Prop("ExtStyle"))

    'number of channels
    cboChannels.Clear
    cboChannels.AddItem "1"
    cboChannels.AddItem "2"
    cboChannels.AddItem "3"
    cboChannels.AddItem "4"
    cboChannels.AddItem "5"
    cboChannels.AddItem "6"
    cboChannels.AddItem "7"
    cboChannels.AddItem "8"
    cboChannels.AddItem "9"
            
    LoadPenStyles cboIndicator
    SetCombo cboIndicator, Val(m.Annot.Prop("ChannelStyle"))       'aardvark 4443
        
    With m.Annot
        txtPercent.Text = Format(m.Annot.Prop("ChannelPercent"), "##0.00")
        txtPoints.Text = Format(m.Annot.Prop("ChannelPoints"), "###0.00###")
        optLocation(Val(.Prop("ChannelLocation"))).Value = True
        If Val(.Prop("ChannelType")) = 0 Then
            optPercent.Value = True
            optPoints.Value = False
        Else
            optPercent.Value = False
            optPoints.Value = True
        End If
        i = Int(Val(.Prop("ChannelCount")))
        If i < 0 Then
            .Prop("ChannelCount") = 0
            i = 0
        ElseIf i > 0 Then
            i = i - 1
        End If
        cboChannels.ListIndex = i
    End With
    
    chkShowValueInAxis.Width = chkShowValueInAxis.Width + 500
    'reposition controls in the regression line frame
    i = 150
    lblChannels.Move lblIndicator.Left - 90, lblIndicator.Top + i 'number of channels
    cboChannels.Move cboStdDevStyle.Left, cboIndicator.Top + i
    lblIndicator.Move lblIndicator.Left - 90, txtStdDevVal.Top + i   'channels line style
    cboIndicator.Move cboIndicator.Left, txtStdDevVal.Top + i
    optPercent.Move lblIndicator.Left - 30, lblIndicator.Top + lblIndicator.Height + 220
    txtPercent.Move cboIndicator.Left, optPercent.Top - 50
    optPoints.Move txtPercent.Left + txtPercent.Width + 225, optPercent.Top
    txtPoints.Move optPoints.Left + optPoints.Width, txtPercent.Top
    chkStdDevOnOff.Move optPercent.Left, txtPercent.Top + txtPercent.Height + i, chkStdDevOnOff.Width + 500
    
    lblChannels.Enabled = True
    cboChannels.Enabled = True
                
    fraLocation.Left = lblIndicator.Left - 120
    fraLocation.Top = txtRegLineLen.Top

    SetTrendlineControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitTrendlineControls", eGDRaiseError_Raise

End Sub

Private Function CboItem(cbo As ctlUniComboImageXP) As Long
On Error GoTo ErrSection:

    If cbo.ListIndex >= 0 And cbo.ListIndex < cbo.ListCount Then    '4346
        CboItem = cbo.ItemData(cbo.ListIndex)
    Else
        CboItem = 0
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditAnnot.CboItem", eGDRaiseError_Raise
    
End Function

Private Sub SetGannValues(ByVal iPoint%, ByVal dValue)
On Error GoTo ErrSection:

    Dim dY#

    With m.Annot
        dY = RoundNum(.Y(iPoint), 2)
        If dY <> dValue Then
            m.bTextChanged = True
            .MovePoint m.Chart, iPoint, .gePaneId, .X(iPoint), dValue
            .geDrawAnn m.Chart
            .geMoveFlag = 0
            'lines directions (redo in case direction changed)
            SetQuadrantChkBoxes
            optRatio_Click 0
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.SetGannValues", eGDRaiseError_Raise
    
End Sub

Private Sub SetQuadrantChkBoxes()
On Error GoTo ErrSection:

    With m.Annot
        If Val(.Prop("DirNE")) = 0 Then
            chkQuadrant(0) = 0
        Else
            chkQuadrant(0) = 1
        End If
        If Val(.Prop("DirSE")) = 0 Then
            chkQuadrant(1) = 0
        Else
            chkQuadrant(1) = 1
        End If
        If Val(.Prop("DirNW")) = 0 Then
            chkQuadrant(2) = 0
        Else
            chkQuadrant(2) = 1
        End If
        If Val(.Prop("DirSW")) = 0 Then
            chkQuadrant(3) = 0
        Else
            chkQuadrant(3) = 1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.SetQuadrantChkBoxes", eGDRaiseError_Raise
    
End Sub

Private Sub SetImageInfo()
    Dim nIcon&
    
    nIcon = gdSelectIcon.Icon
    'TODO: Look into combining with code in chart config form
    With m.Annot
        If nIcon < 8 Then
            .Prop("ImageType") = eCNI_Arrow
            Select Case nIcon
                Case 0
                    .Prop("ImageDir") = eCNI_North
                Case 1
                    .Prop("ImageDir") = eCNI_South
                Case 2
                    .Prop("ImageDir") = eCNI_East
                Case 3
                    .Prop("ImageDir") = eCNI_West
                Case 4
                    .Prop("ImageDir") = eCNI_NorthEast
                Case 5
                    .Prop("ImageDir") = eCNI_SouthWest
                Case 6
                    .Prop("ImageDir") = eCNI_SouthEast
                Case 7
                    .Prop("ImageDir") = eCNI_NorthWest
            End Select
        ElseIf nIcon = 8 Then
            .Prop("ImageType") = eCNI_Plus
        ElseIf nIcon = 9 Then
            .Prop("ImageType") = eCNI_Cross
        ElseIf nIcon = 12 Or nIcon = 17 Then
            .Prop("ImageType") = eCNI_Circle
            If nIcon = 12 Then
                .Prop("ImageStyle") = 1
            Else
                .Prop("ImageStyle") = 0
            End If
        ElseIf nIcon = 13 Or nIcon = 18 Then
            .Prop("ImageType") = eCNI_Square
            If nIcon = 13 Then
                .Prop("ImageStyle") = 1
            Else
                .Prop("ImageStyle") = 0
            End If
        ElseIf nIcon = 14 Or nIcon = 19 Then
            .Prop("ImageType") = eCNI_Diamond
            If nIcon = 14 Then
                .Prop("ImageStyle") = 1
            Else
                .Prop("ImageStyle") = 0
            End If
        Else
            .Prop("ImageType") = eCNI_Triangle
            If nIcon = 10 Or nIcon = 11 Then
                .Prop("ImageStyle") = 1
            Else
                .Prop("ImageStyle") = 0
            End If
            If nIcon = 10 Or nIcon = 15 Then
                .Prop("ImageDir") = eCNI_North
            Else
                .Prop("ImageDir") = eCNI_South
            End If
        End If
    End With

End Sub

Private Sub InitSRLineControls()
On Error GoTo ErrSection:

    Dim nSaveY&, nX&
    Dim strValue$

    If m.Annot Is Nothing Then Exit Sub
    
    If m.Annot.eType = eANNOT_SRLine2 Then
        Me.Caption = "Support/Resistance Line 2"
        Me.Icon = Picture16(ToolbarIcon("ID_SRLine2"), , True)
    ElseIf m.Annot.eType = eANNOT_SRLine3 Then
        Me.Caption = "Support/Resistance Line 3"
        Me.Icon = Picture16(ToolbarIcon("ID_SRLine3"), , True)
    ElseIf m.Annot.eType = eANNOT_SRLine4 Then
        Me.Caption = "Support/Resistance Line 4"
        Me.Icon = Picture16(ToolbarIcon("ID_SRLine4"), , True)
    Else
        Me.Caption = "Support/Resistance Line"
        Me.Icon = Picture16(ToolbarIcon("ID_SRLine"), , True)
    End If
    
    'show needed frame(s)
    fraValue.Width = fraValue.Width - 110
    fraValue.Height = fraValue.Height + 110
    If m.bMultiChartOption = True Then
        chkShowValueInAxis.Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
    Else
        chkShowValueInAxis.Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    End If
    chkShowValueInAxis.Value = Val(m.Annot.Prop("ShowInAxis"))
    chkShowValueInAxis.Visible = True
    fraValue.Top = chkShowValueInAxis.Top + chkShowValueInAxis.Height + 135
    SetBottom fraValue
    
    'show needed controls
    chkExtendSRLine.Visible = True
    chkDisplaySRValue.Visible = True
    optSRLeft.Visible = True
    optSRRight.Visible = True
    chkSRLineDot.Visible = True
    chkSRLineDot.Enabled = True
   
    'hide not needed controls
    lblFrom.Visible = False
    lblTo.Visible = False
    txtFromY.Visible = False
    txtToY.Visible = False
    
    'show controls not within any frame
    cmdFont.Visible = True
    
    'position controls
    If ExtremeCharts <> 1 Then
        cmdAlert.Visible = True
        cmdAlert.Move cmdFont.Left, cboStyle.Top
    End If
    nSaveY = chkExtendSRLine.Top
    chkExtendSRLine.Move lblValue.Left, lblValue.Top + 80
    chkDisplaySRValue.Move lblFrom.Left, lblFrom.Top + 60
    optSRLeft.Move chkDisplaySRValue.Left + chkDisplaySRValue.Width + 50, chkDisplaySRValue.Top
    optSRRight.Move optSRLeft.Left + optSRLeft.Width + 50, optSRLeft.Top
    chkSRLineDot.Move optSRLeft.Left, chkExtendSRLine.Top
    lblValue.Move lblValue.Left, nSaveY + 120
    nX = Me.TextWidth("Value:") + lblValue.Left + 120
    txtValue.Move nX, nSaveY + 90
    
    'set controls values
    lblValue.Caption = "Value:"
    ShowValue "Value:", m.Annot.Y(1), strValue, False
    txtValue.Text = strValue
    With m.Annot
        chkDisplaySRValue.Value = Val(.Prop("ShowValues"))
        chkExtendSRLine.Value = Val(.Prop("Ext"))
        chkSRLineDot.Value = Val(.Prop("HideSRLineDot"))
        cmdAlert.Enabled = .CanHaveAlert
        If .Prop("TextAlignment") = "6" Then
            optSRRight.Value = True
        Else
            optSRLeft.Value = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitSRLineControls", eGDRaiseError_Raise
    
End Sub

Private Sub CenterColorStyle()
On Error Resume Next

    Dim nTotalWidth&, nLeft&
    
    lblColor.Width = Me.TextWidth(lblColor.Caption) + 100
        
    nTotalWidth = lblColor.Width + clrColor.Width
    If m.Annot.eType = eANNOT_WaveLabels Then
        nTotalWidth = nTotalWidth + cmdFont.Width + 100
    End If
    
    nLeft = Me.ScaleLeft + Me.ScaleWidth / 2 - nTotalWidth / 2
    
    'center color controls
    lblColor.Move nLeft
    clrColor.Move lblColor.Left + lblColor.Width
    
    'center style controls
    lblStyle.Move lblColor.Left
    cboStyle.Move clrColor.Left
    
    If cmdFont.Visible Then
        If m.Annot.eType = eANNOT_WaveLabels Then
            cmdFont.Move cboStyle.Left + cboStyle.Width + 100, cboStyle.Top
        Else
            cmdFont.Move Me.ScaleLeft + Me.ScaleWidth - cmdFont.Width - 150
        End If
    End If

End Sub

Private Sub InitMirrorControls()
On Error GoTo ErrSection:

    Me.Caption = "Price Mirror"
    Me.Icon = Picture16(ToolbarIcon("ID_Mirror"), , True)
        
    SetBottom chkMultiChart
    m.bCenterColorStyle = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitMirrorControls", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitPatternControls()
On Error GoTo ErrSection:

    Me.Caption = "Price Pattern"
    Me.Icon = Picture16(ToolbarIcon("ID_Pattern"), , True)
    
    chkMultiChart.Visible = False
    cmdFont.Visible = True
        
    With fraPattern
        .Visible = True
        .Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
    End With
    
    With fraPatternOnChart
        .Height = fgPatternOnChart.RowHeight(1) * 5.2
        .Caption = "Patterns on chart (dblclick to locate)"
        .Visible = True
    End With
    
    txtPatternName.Text = m.Annot.Text
    chkPatternName.Value = Val(m.Annot.Prop("ShowPatternName"))
    txtForecastBars.Text = Val(m.Annot.Prop("ForecastBars"))
    
    With fgPatternOnChart
        .Height = fraPatternOnChart.Height - 500
        .Editable = flexEDNone
        .FixedCols = 0
        .FixedRows = 1
        .AutoSizeMode = flexAutoSizeColWidth
        .ScrollBars = flexScrollBarNone
        .AutoSize 0, 0, True
        .Cols = 3
        .Rows = 3
        
        .TextMatrix(0, 0) = "Pattern"
        .TextMatrix(0, 1) = "From"
        .TextMatrix(0, 2) = "To"
        
        .TextMatrix(1, 0) = "Original"
        .TextMatrix(1, 1) = DateFormat(m.Annot.DateFromArray(0))
        .TextMatrix(1, 2) = DateFormat(m.Annot.DateFromArray(1))
        
        .TextMatrix(2, 0) = "Copy"
        .TextMatrix(2, 1) = DateFormat(m.Annot.dDate(1))
        .TextMatrix(2, 2) = DateFormat(m.Annot.dDate(2))
        
        .ColWidth(0) = .ClientWidth - .ColWidth(1) * 2
        .Move .Left, .Top
    End With
    
    fraPatternOnChart.Move fraPatternOnChart.Left, fraPattern.Top + fraPattern.Height + 150
    fraPatternOnChart.ZOrder
    
    SetBottom fraPatternOnChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitPatternControls", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitRiskReward(Chart As cChart)
On Error GoTo ErrSection:

    Me.Caption = "Risk Reward"
    Me.Icon = Picture16(ToolbarIcon("ID_RiskReward"), , True)
    
    'show needed frame(s)
    fraRiskReward.Visible = True
    'show controls not within any frame
    cmdFont.Visible = True
    
    'set controls values
    lblRiskReward = ""
    
    With m.Annot
        chkShowProfitLost.Value = Val(.Prop("ShowProfitLoss"))
        chkShowValues.Value = Val(.Prop("ShowValues"))
        Select Case Chart.Bars.Prop(eBARS_SecurityType)
            Case Asc("S")
                lblRiskReward = "Number of shares"
                txtRiskReward = Val(.Prop("NumShares"))
                txtRiskReward.Tag = "NumShares"
            Case Asc("F")
                lblRiskReward = "Number of contracts"
                txtRiskReward = Val(.Prop("NumContracts"))
                txtRiskReward.Tag = "NumContracts"
            Case Else
                If IsForex(Chart.Bars.Prop(eBARS_Symbol)) Then
                    lblRiskReward = "Number of contracts"
                    txtRiskReward = Val(.Prop("NumForexContracts"))
                    txtRiskReward.Tag = "NumForexContracts"
                Else
                    lblRiskReward = "Multiplier"
                    txtRiskReward = Val(.Prop("IndexMult"))
                    txtRiskReward.Tag = "IndexMult"
                End If
        End Select
    End With

    If m.bMultiChartOption Then
        fraRiskReward.Move chkMultiChart.Left - 20, chkMultiChart.Top + chkMultiChart.Height + 150
    Else
        fraRiskReward.Move chkPreIndicator.Left - 20, chkPreIndicator.Top + chkPreIndicator.Height + 150
    End If
    SetBottom fraRiskReward

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitRiskReward", eGDRaiseError_Show
    Resume ErrExit

End Sub

Public Sub ShowWaveLabels(AnnotOptions As cAnnotation)
On Error GoTo ErrSection:

    Dim i&
    Dim aCustomLabels As New cGdArray
       
    Me.Caption = "Wave Labels"
    Me.Icon = Picture16(ToolbarIcon("ID_WaveLabels"), , True)
       
    If AnnotOptions Is Nothing Then
        m.bReturnOptions = False
    Else
        HideOnInitialShow
        LoadPenStyles cboStyle
        cboStyle.ListIndex = 0
        Set m.Annot = AnnotOptions
        m.bReturnOptions = True
    End If
        
    cmdFont.Move cmdFont.Left, cboStyle.Top
    cmdFont.Visible = True
    aCustomLabels.FromFile g.strAppPath & kCustomWaveLabels
    
    'set grid values
    With fgWaveLabels
        SetupGrid fgWaveLabels, eGridMode_Grid
        .FixedCols = 0
        .FixedRows = 0
        .Cols = 1
        .Rows = 6
        .ColAlignment(0) = flexAlignCenterCenter
        .Font.Bold = True
        .TextMatrix(0, 0) = "No labels"
        .TextMatrix(1, 0) = "A, B, C"
        .TextMatrix(2, 0) = "a, b, c"
        .TextMatrix(3, 0) = "I, II, III, IV, V"
        .TextMatrix(4, 0) = "i, ii, iii, iv, v"
        .TextMatrix(5, 0) = "1, 2, 3, 4, 5"
        
        If aCustomLabels.Size > 0 Then
            For i = 0 To aCustomLabels.Size - 1
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = aCustomLabels(i)
            Next
        End If
        
        For i = 0 To 5 + aCustomLabels.Size
            If m.Annot.Text = .TextMatrix(i, 0) Then
                .Row = i
                Exit For
            End If
        Next
        If .Row > 5 Then
            cmdDelCustomLabel.Enabled = True
        Else
            cmdDelCustomLabel.Enabled = False
        End If
    End With
    
    With fraWaveLabels
        If m.bReturnOptions Then cmdDelete.Caption = "&Cancel"
        If m.bMultiChartOption Then
            chkMultiChart.Visible = True            '4441
            .Move chkMultiChart.Left, chkMultiChart.Top + chkMultiChart.Height + 100
        Else
            chkMultiChart.Visible = False
            .Move chkPreIndicator.Left, chkPreIndicator.Top + chkPreIndicator.Height + 100
        End If
        .Visible = True
    End With
    
    'set values for controls
    With m.Annot
        clrColor.Color = .Color
        SetCombo cboStyle, .Style
        chkPointArc.Value = Val(.Prop("PointArc"))
        HandleWaveConnect .Text, Val(.Prop("Lines")), Val(.Prop("LabelFirstPoint"))
    End With
    
    SetBottom fraWaveLabels
    m.bCenterColorStyle = True
    Me.Width = fraWaveLabels.Width + 400
    fraButtons.Left = Me.Width / 2 - fraButtons.Width / 2
       
    If m.bReturnOptions Then
        CenterTheForm Me
        ShowForm Me, eForm_Modal
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.ShowWaveLabels", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub HideOnInitialShow()
On Error GoTo ErrSection:

    'hide controls not within a frame
    cmdAlert.Visible = False
    cmdFont.Visible = False
    cmdSwitchSides.Visible = False
    chkDynamic.Visible = False
    chkAllPanes.Visible = False
    gdSelectIcon.Visible = False
    clrFillColor.Visible = False
    chkUseFillColor.Visible = False
    chkShowValueInAxis.Value = False
    
    'hide all frames except buttons frame
    fraArrow.Visible = False
    fraExt.Visible = False
    fraRect.Visible = False
    fraRegression.Visible = False
    fraTarget.Visible = False
    fraText.Visible = False
    fraValue.Visible = False
    fraEllipse.Visible = False
    fraGannOptions.Visible = False
    fraQuadrants.Visible = False
    fraGannLines.Visible = False
    fraPatternOnChart.Visible = False
    fraDLineText.Visible = False
    fraPattern.Visible = False
    fraRiskReward.Visible = False
    fraWaveLabels.Visible = False
    fraGannacciMultiply.Visible = False
    fraGanncciSquareRange.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.HideOnInitialShow", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SaveWaveLabels()
On Error GoTo ErrSection:

    Dim nSelR1&, nSelCol1&, nSelR2&, nSelCol2&
    
    If Not m.Annot Is Nothing Then
        If m.Annot.eType = eANNOT_WaveLabels Then
            With fgWaveLabels
                .GetSelection nSelR1, nSelCol1, nSelR2, nSelCol2
                If nSelR1 >= 0 And nSelR1 < .Rows Then
                    m.Annot.Text = .TextMatrix(nSelR1, 0)
                    m.Annot.Color = clrColor.Color
                    m.Annot.Style = cboStyle.ItemData(cboStyle.ListIndex)
                    m.Annot.Prop("PointArc") = chkPointArc.Value
                    If optWaveConnect(2).Value = True Then
                        m.Annot.Prop("Lines") = 0
                    Else
                        m.Annot.Prop("Lines") = 1
                        If optWaveConnect(0).Value = True Then
                            m.Annot.Prop("LabelFirstPoint") = 0
                        ElseIf optWaveConnect(1).Value = True Then
                            m.Annot.Prop("LabelFirstPoint") = 1
                        End If
                    End If
                    If optWaveContinue(0).Value = True Then
                        m.Annot.Prop("LabelPastEnd") = "repeat"
                    ElseIf optWaveContinue(2).Value = True Then
                        m.Annot.Prop("LabelPastEnd") = "continue"
                    Else
                        m.Annot.Prop("LabelPastEnd") = "none"
                    End If
                End If
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.SetWaveLabels", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub HandleWaveConnect(ByVal strText$, ByVal nConnect&, _
    ByVal nLabelFirstPoint&)

    Dim bContinueLabel As Boolean

    If InStr(strText, ",") > 0 Then
        optWaveConnect(2).Enabled = True
        If nConnect = 0 Then
            optWaveConnect(0).Value = False
            optWaveConnect(1).Value = False
            optWaveConnect(2).Value = True
        Else
            optWaveConnect(2).Value = False
            If nLabelFirstPoint = 0 Then
                optWaveConnect(0).Value = True
                optWaveConnect(1).Value = False
            Else
                optWaveConnect(0).Value = False
                optWaveConnect(1).Value = True
            End If
        End If
        If Not m.Annot Is Nothing Then
            With m.Annot
                bContinueLabel = .CanContinueLabel(strText)
            End With
        End If
        optWaveContinue(0).Enabled = True
        optWaveContinue(1).Enabled = True
        optWaveContinue(2).Enabled = bContinueLabel
        lblContinueLabel.Enabled = bContinueLabel
    Else
        optWaveConnect(2).Enabled = False
        optWaveConnect(2).Value = False
        If nLabelFirstPoint = 0 Then
            optWaveConnect(0).Value = True
            optWaveConnect(1).Value = False
        Else
            optWaveConnect(0).Value = False
            optWaveConnect(1).Value = True
        End If
        
        optWaveContinue(0).Enabled = False
        optWaveContinue(1).Enabled = False
        optWaveContinue(2).Enabled = False
        lblRepeatLabel.Enabled = False
        lblNoLabelPastEnd.Enabled = False
        lblContinueLabel.Enabled = False
    End If
    
    If Not m.Annot Is Nothing Then
        With m.Annot
            If .Prop("LabelPastEnd") = "repeat" Then
                optWaveContinue(0).Value = True
                optWaveContinue(1).Value = False
                optWaveContinue(2).Value = False
            ElseIf .Prop("LabelPastEnd") = "none" Then
                optWaveContinue(0).Value = False
                optWaveContinue(1).Value = True
                optWaveContinue(2).Value = False
            ElseIf .Prop("LabelPastEnd") = "continue" Then
                optWaveContinue(0).Value = False
                optWaveContinue(1).Value = False
                optWaveContinue(2).Value = True
            End If
        End With
    End If

End Sub

Private Sub InitRectangleControls()
On Error GoTo ErrSection:
        
    'hide controls within a frame not used by this annotation
    chkTextBorder.Visible = False
    cboAnchor.Visible = False
    cboTextJustify.Visible = False
    lblTextJustify.Visible = False
    rtfText.Visible = False
    
    clrFillColor.Visible = True
    chkUseFillColor.Visible = True
    
    'set controls used by both triangle and rectangle tools
    chkUseFillColor.Move lblColor.Left, cboStyle.Top + cboStyle.Height + 120
    clrFillColor.Move cmdFont.Left - 130, chkUseFillColor.Top - 60
    chkPreIndicator.Move lblColor.Left, chkUseFillColor.Top + chkUseFillColor.Height + 120
    chkMultiChart.Move lblColor.Left, chkPreIndicator.Top + chkPreIndicator.Height + 120
    
    chkUseFillColor.Value = Val(m.Annot.Prop("FillPattern"))
    clrFillColor.Color = Val(m.Annot.Prop("FillColor"))
    
    'if triangle then set controls for triangle tool and exit
    If m.Annot.eType = eANNOT_TriangleWedge Then
        Me.Caption = "Triangle"
        Me.Icon = Picture16(ToolbarIcon("ID_Triangle"), , True)
        SetBottom chkMultiChart
        m.bCenterColorStyle = True
        Exit Sub
    ElseIf m.Annot.eType = eANNOT_ChannelHighlight Then
        Me.Caption = "Channel Highlight"
        Me.Icon = Picture16(ToolbarIcon("ID_ChannelHighlight"), , True)
        cmdSwitchSides.Move clrFillColor.Left, clrColor.Top
        cmdFont.Visible = False
        cmdSwitchSides.Visible = True
        SetBottom chkMultiChart
        Exit Sub
    End If
    
    Me.Caption = "Rectangle"
    Me.Icon = Picture16(ToolbarIcon("ID_Rectangle"), , True)
    cmdFont.Visible = True
    'show needed frame(s)
    fraRect.Visible = True
    If m.bMultiChartOption = True Then
        fraRect.Move fraButtons.Left, chkMultiChart.Top + chkMultiChart.Height + 120
    Else
        fraRect.Move fraButtons.Left, chkPreIndicator.Top + chkPreIndicator.Height + 120
    End If
    fraText.Move fraText.Left, fraRect.Top + fraRect.Height + 120
    SetBottom fraText
                
    'set controls values
    SetTextControls True
    Select Case Val(m.Annot.Prop("Shape"))
    Case 2
        optEllipse = True
    Case 1
        optRounded = True
    Case Else
        optRectangle = True
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitRectangleControls", eGDRaiseError_Raise

End Sub

Private Sub InitGancciSRangeCtrls(Chart As cChart)
On Error GoTo ErrSection:

    Dim strText$, d#, dY1#, dY2#
    
    If m.Annot Is Nothing Then Exit Sub

    Me.Caption = "GANNacci Square Range"
    Me.Icon = Picture16(ToolbarIcon("ID_GannacciSwingSquare"), , True)

    'show needed frame(s)
    fraGanncciSquareRange.Visible = True
    fraValue.Visible = True
    optSRLeft.Visible = False
    optSRRight.Visible = False

    'RH commented out fraValue.BorderStyle = 0
    If m.bMultiChartOption = True Then
        fraValue.Top = chkMultiChart.Top + chkMultiChart.Height
    Else
        fraValue.Top = chkPreIndicator.Top + chkPreIndicator.Height
    End If

    fraValue.Width = fraValue.Width - 100
    fraGanncciSquareRange.Width = fraValue.Width
    
    'show needed controls
    lblFrom.Visible = True
    lblTo.Visible = True
    txtFromY.Visible = True
    txtToY.Visible = True

    'show controls not within any frame
    cmdFont.Visible = True
    chkDynamic.Visible = False

    fraGannacciMultiply.Visible = True
    lblValue.Visible = False
    txtValue.Visible = False
    txtValue.Enabled = False
    
    lblFrom.Left = 75
    lblFrom.Top = lblValue.Top
    lblTo.Top = lblFrom.Top
    txtFromY.Top = lblFrom.Top - 30
    txtToY.Top = txtFromY.Top
    
    fraValue.Top = fraValue.Top - 105
    fraValue.Height = txtFromY.Height * 2
    
    fraGanncciSquareRange.Move fraValue.Left, fraValue.Top + fraValue.Height + 60
    fraGannacciMultiply.Move fraGanncciSquareRange.Left, fraGanncciSquareRange.Top + fraGanncciSquareRange.Height + 90
    SetBottom fraGannacciMultiply

    With m.Annot
        dY1 = .Y(1)
        dY2 = .Y(2)
        
        d = Val(.Prop("MultiplierVal"))
        If d = 0 Then d = 1#
        txtGannacciMultiply.Text = d
        
        'multiplier frame
        If Val(.Prop("UseMutiplier")) = 0 Then
            txtGannacciMultiply.Enabled = False
            chkSRangeMultiply.Value = vbUnchecked
        Else
            txtGannacciMultiply.Enabled = True
            chkSRangeMultiply.Value = vbChecked
            dY1 = dY1 * d
            dY2 = dY2 * d
        End If
        
        If .X(2) >= .X(1) Then
            ShowValue "", .Y(1), strText, False
            txtFromY.Text = strText
            ShowValue "", .Y(2), strText, False
            txtToY.Text = strText
        Else
            ShowValue "", .Y(2), strText, False
            txtFromY.Text = strText
            ShowValue "", .Y(1), strText, False
            txtToY.Text = strText
        End If
        
        lblAdjPriceFrom.Caption = "Adjusted price from:  " & Format(dY1, "#0.00###")
        lblAdjPriceTo.Caption = "Adjusted price to:      " & Format(dY2, "#0.00###")
    
        chkSRangeFirstBar.Value = Val(m.Annot.Prop("IncludeBarOne"))
        chkSRangePrice.Value = Val(m.Annot.Prop("PriceMove"))
        chkSRangeTB.Value = Val(m.Annot.Prop("LenBars"))
        chkSRangeSameSwing = Val(m.Annot.Prop("SameSwing"))
        chkSRangeSquare = Val(m.Annot.Prop("SquareRange"))
        chkSRangeExtend = Val(m.Annot.Prop("Ext"))
        
        If Not .AnnotChart Is Nothing Then
            If .AnnotChart.Bars.Prop(eBARS_PeriodicityStr) = "Daily" Then
                chkSRangeCD.Enabled = True
                chkSRangeCD.Value = Val(.Prop("CalendarDays"))
            Else
                chkSRangeCD.Enabled = False
            End If
        End If
        
        'origin date for square range
        If .dDate(1) < .dDate(2) Then
            lblSRangeOrigin.Caption = DateFormat(.dDate(2), MM_DD_YYYY)
        Else
            lblSRangeOrigin.Caption = DateFormat(.dDate(1), MM_DD_YYYY)
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditAnnot.InitGancciSRangeCtrls", eGDRaiseError_Raise

End Sub

Private Function GannSROptionsOk(chkCtrl As ctlUniCheckXP) As Boolean  'RH was Checkbox
On Error GoTo ErrSection:

    Dim strErr As String

    If Me.chkSRangeSameSwing.Value = vbUnchecked Then
        If chkSRangeSquare.Value = vbUnchecked Then
            strErr = "One of the Same Swing or Square Range options must be on."
        ElseIf Me.chkSRangeCD.Value = vbUnchecked And Me.chkSRangeTB.Value = vbUnchecked Then
            strErr = "One of the Calendar Days or Trading Bars options must be on for Square Range to show."
        End If
    ElseIf Me.chkSRangeSquare.Value = vbUnchecked Then
        If Me.chkSRangeSameSwing.Value = vbUnchecked Then
            strErr = "One of the Same Swing or Square Range options must be on."
        End If
    End If
    
    If Len(strErr) > 0 Then
        chkCtrl.Value = vbChecked
        InfBox strErr, "I", , "GANNacci Square Range"
    Else
        GannSROptionsOk = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditAnnot.GannSROptionsOk", eGDRaiseError_Raise

End Function


























