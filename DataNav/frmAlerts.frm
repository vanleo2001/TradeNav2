VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAlerts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alerts"
   ClientHeight    =   12360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "frmAlerts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12360
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraConditionChart 
      Height          =   2235
      Left            =   1080
      TabIndex        =   50
      Top             =   7560
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":0462
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":0482
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboIndicator 
         Height          =   315
         Left            =   3690
         TabIndex        =   8
         Top             =   1740
         Width           =   2775
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
         Tip             =   "frmAlerts.frx":049E
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":04BE
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtAlert 
         Height          =   315
         Left            =   3690
         TabIndex        =   11
         Top             =   1500
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":04DA
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
         Tip             =   "frmAlerts.frx":04FC
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":051C
      End
      Begin HexUniControls.ctlUniRadioXP optSpread 
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1770
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
         Caption         =   "frmAlerts.frx":0538
         Enabled         =   0   'False
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0578
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0598
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditChartCond 
         Height          =   435
         Left            =   6660
         TabIndex        =   52
         Top             =   1680
         Width           =   1815
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
         Caption         =   "frmAlerts.frx":05B4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":05FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":061E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSpread 
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   23
         Top             =   1960
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":063A
         Enabled         =   0   'False
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":067E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":069E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblIndicator 
         Height          =   240
         Left            =   2340
         Top             =   1770
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
         Caption         =   "frmAlerts.frx":06BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":06FC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":071C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblChartCondition 
         Height          =   1275
         Left            =   180
         Top             =   300
         Width           =   8295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0738
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":0758
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0778
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraConditionTime 
      Height          =   2535
      Left            =   600
      TabIndex        =   54
      Top             =   7140
      Width           =   9615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":0794
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":07C6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":07E6
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectDate gdAtTime 
         Height          =   315
         Left            =   3660
         TabIndex        =   71
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         ShowDayOfWeek   =   0   'False
         ShowPM          =   1
         ShowDate        =   0
         ShowTime        =   2
         MinDate         =   0
         MaxDate         =   0.99999
         Value           =   0
      End
      Begin HexUniControls.ctlUniCheckXP chkMonday 
         Height          =   255
         Left            =   1500
         TabIndex        =   59
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0802
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0834
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0854
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkTuesday 
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0870
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":08A4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":08C4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkWednesday 
         Height          =   255
         Left            =   2100
         TabIndex        =   63
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":08E0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0918
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0938
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkThursday 
         Height          =   255
         Left            =   2400
         TabIndex        =   65
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0954
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":098A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":09AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFriday 
         Height          =   255
         Left            =   2700
         TabIndex        =   67
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":09C6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":09F8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0A18
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSaturday 
         Height          =   255
         Left            =   3000
         TabIndex        =   69
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0A34
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0A6A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0A8A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSunday 
         Height          =   255
         Left            =   1200
         TabIndex        =   57
         Top             =   480
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0AA6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0AD8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0AF8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSessionEnd 
         Height          =   315
         Left            =   4440
         TabIndex        =   81
         Top             =   1920
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
         Caption         =   "frmAlerts.frx":0B14
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0B3A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0B5A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSessionStart 
         Height          =   315
         Left            =   3360
         TabIndex        =   79
         Top             =   1920
         Width           =   735
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
         Caption         =   "frmAlerts.frx":0B76
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0BA0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0BC0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraTimeOptions 
         Height          =   255
         Left            =   2400
         TabIndex        =   74
         Top             =   900
         Width           =   4395
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0BDC
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAlerts.frx":0BFC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0C1C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optNY 
            Height          =   255
            Left            =   3300
            TabIndex        =   77
            Top             =   0
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
            Caption         =   "frmAlerts.frx":0C38
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":0C68
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":0C88
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optGMT 
            Height          =   255
            Left            =   2340
            TabIndex        =   76
            Top             =   0
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
            Caption         =   "frmAlerts.frx":0CA4
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":0CCA
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":0CEA
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLocal 
            Height          =   255
            Left            =   1380
            TabIndex        =   75
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
            Caption         =   "frmAlerts.frx":0D06
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":0D30
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":0D50
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblDisplayTimes 
            Height          =   195
            Left            =   60
            Top             =   30
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
            Caption         =   "frmAlerts.frx":0D6C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAlerts.frx":0DAE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":0DCE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin gdOCX.gdSelectDate gdAtDateTime 
         Height          =   315
         Left            =   6060
         TabIndex        =   73
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         ShowTime        =   2
      End
      Begin HexUniControls.ctlUniRadioXP optAt 
         Height          =   255
         Left            =   5460
         TabIndex        =   72
         Top             =   390
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
         Caption         =   "frmAlerts.frx":0DEA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0E10
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0E30
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optEvery 
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   390
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
         Caption         =   "frmAlerts.frx":0E4C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0E78
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0E98
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSessionLookup 
         Height          =   255
         Left            =   6930
         TabIndex        =   84
         Top             =   1950
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
         Caption         =   "frmAlerts.frx":0EB4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":0EF4
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0F14
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSessionSymbol 
         Height          =   315
         Left            =   5760
         TabIndex        =   83
         Top             =   1920
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":0F30
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
         Tip             =   "frmAlerts.frx":0F64
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0F84
      End
      Begin HexUniControls.ctlUniLabelXP lblSunday 
         Height          =   255
         Left            =   1200
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":0FA0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":0FC4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":0FE4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSaturday 
         Height          =   255
         Left            =   3000
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":1000
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":1024
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1044
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFriday 
         Height          =   255
         Left            =   2700
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":1060
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":1082
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":10A2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblThursday 
         Height          =   255
         Left            =   2400
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":10BE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":10E2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1102
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblWednesday 
         Height          =   255
         Left            =   2100
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":111E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":1140
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1160
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTuesday 
         Height          =   255
         Left            =   1800
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":117C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":11A0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":11C0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblMonday 
         Height          =   255
         Left            =   1500
         Top             =   300
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":11DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":11FE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":121E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAt 
         Height          =   255
         Left            =   3360
         Top             =   390
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":123A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":125E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":127E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFor 
         Height          =   315
         Left            =   5460
         Top             =   1920
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
         Caption         =   "frmAlerts.frx":129A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":12C0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":12E0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOr 
         Height          =   255
         Left            =   4200
         Top             =   1950
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":12FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":1320
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1340
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSetToSession 
         Height          =   255
         Left            =   1680
         Top             =   1950
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
         Caption         =   "frmAlerts.frx":135C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":13A4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":13C4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraConditionStatus 
      Height          =   2535
      Left            =   240
      TabIndex        =   25
      Top             =   6720
      Width           =   9615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":13E0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":1412
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":1432
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniListBoxXP lstStatusItems 
         Height          =   1815
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   3135
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
         TrapTab         =   0   'False
         Tip             =   "frmAlerts.frx":144E
         MultiSelect     =   0
         Sorted          =   0   'False
         HScroll         =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         RoundedBorders  =   0   'False
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":146E
         ManualStart     =   0   'False
         Columns         =   0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStatusDesc 
         Height          =   1815
         Left            =   3600
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
         Caption         =   "frmAlerts.frx":148A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":14B6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":14D6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblStatusItem 
         Height          =   255
         Left            =   240
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
         Caption         =   "frmAlerts.frx":14F2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":152A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":154A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkAfterBarComplete 
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":1566
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmAlerts.frx":15C0
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":15E0
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraConditionPrice 
      Height          =   2535
      Left            =   180
      TabIndex        =   28
      Top             =   9660
      Width           =   9615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":15FC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":162E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":164E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraNumDays 
         Height          =   615
         Left            =   720
         TabIndex        =   42
         Top             =   1140
         Width           =   5235
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":166A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAlerts.frx":1696
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":16B6
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optAutoDetect 
            Height          =   255
            Left            =   120
            TabIndex        =   44
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
            Caption         =   "frmAlerts.frx":16D2
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   -1  'True
            Tip             =   "frmAlerts.frx":1724
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":1744
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optOverride 
            Height          =   255
            Left            =   2400
            TabIndex        =   45
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
            Caption         =   "frmAlerts.frx":1760
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":179E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":17BE
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtOverride 
            Height          =   315
            Left            =   4140
            TabIndex        =   47
            Top             =   180
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmAlerts.frx":17DA
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
            Tip             =   "frmAlerts.frx":17FC
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":181C
         End
         Begin HexUniControls.ctlUniTextBoxXP txtNumBars 
            Height          =   315
            Left            =   4080
            TabIndex        =   46
            Top             =   60
            Width           =   1095
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmAlerts.frx":1838
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
            Tip             =   "frmAlerts.frx":185A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":187A
         End
         Begin HexUniControls.ctlUniLabelXP lblNumBars1 
            Height          =   195
            Left            =   0
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
            Caption         =   "frmAlerts.frx":1896
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAlerts.frx":192A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":194A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniCheckXP chkShowOnCharts 
         Height          =   255
         Left            =   6060
         TabIndex        =   34
         Top             =   420
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
         Caption         =   "frmAlerts.frx":1966
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":19A2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":19C2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdVerify 
         Height          =   315
         Left            =   8040
         TabIndex        =   48
         Top             =   390
         Width           =   1395
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
         Caption         =   "frmAlerts.frx":19DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":1A0A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1A2A
         RightToLeft     =   0   'False
      End
      Begin NavTradeSenseV3.Editor tsCondition 
         Height          =   555
         Left            =   300
         TabIndex        =   49
         Top             =   1800
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   979
      End
      Begin HexUniControls.ctlUniFrameWL fraPriceExtremes 
         Height          =   375
         Left            =   720
         TabIndex        =   35
         Top             =   780
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
         Caption         =   "frmAlerts.frx":1A46
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAlerts.frx":1A72
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1A92
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkGetsUpTo 
            Height          =   255
            Left            =   3300
            TabIndex        =   39
            Top             =   60
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
            Caption         =   "frmAlerts.frx":1AAE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":1AF0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":1B10
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkGetsDownTo 
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   60
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
            Caption         =   "frmAlerts.frx":1B2C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":1B72
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":1B92
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtDownTo 
            Height          =   285
            Left            =   1860
            TabIndex        =   37
            Top             =   30
            Width           =   1020
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmAlerts.frx":1BAE
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
            Tip             =   "frmAlerts.frx":1BDE
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":1BFE
         End
         Begin HexUniControls.ctlUniTextBoxXP txtUpTo 
            Height          =   285
            Left            =   4920
            TabIndex        =   40
            Top             =   45
            Width           =   1020
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmAlerts.frx":1C1A
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
            Tip             =   "frmAlerts.frx":1C4A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":1C6A
         End
         Begin gdOCX.gdScrollBar sbDownTo 
            Height          =   360
            Left            =   2880
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   -15
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
         Begin gdOCX.gdScrollBar sbUpTo 
            Height          =   360
            Left            =   5940
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   635
         End
      End
      Begin HexUniControls.ctlUniComboBoxXP cboPricePeriod 
         Height          =   315
         Left            =   3780
         TabIndex        =   33
         Top             =   390
         Width           =   1635
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
         Tip             =   "frmAlerts.frx":1C86
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
         MouseIcon       =   "frmAlerts.frx":1CA6
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPriceLookup 
         Height          =   255
         Left            =   2070
         TabIndex        =   31
         Top             =   420
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
         Caption         =   "frmAlerts.frx":1CC2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":1CFE
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1D1E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPriceSymbol 
         Height          =   315
         Left            =   900
         TabIndex        =   30
         Top             =   390
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":1D3A
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
         Tip             =   "frmAlerts.frx":1D6E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1D8E
      End
      Begin HexUniControls.ctlUniLabelXP lblPricePeriod 
         Height          =   255
         Left            =   2880
         Top             =   420
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
         Caption         =   "frmAlerts.frx":1DAA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":1DE0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1E00
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPriceSymbol 
         Height          =   255
         Left            =   240
         Top             =   420
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
         Caption         =   "frmAlerts.frx":1E1C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":1E4A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":1E6A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniCheckXP chkKeepActive 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5580
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":1E86
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmAlerts.frx":1F44
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":1F64
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkActive 
      Height          =   255
      Left            =   8640
      TabIndex        =   0
      Top             =   3240
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
      Caption         =   "frmAlerts.frx":1F80
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmAlerts.frx":1FBE
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":1FDE
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraConditionQB 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":1FFA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":202C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":204C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboTabs 
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1995
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
         Tip             =   "frmAlerts.frx":2068
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2088
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTab 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   900
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
         Caption         =   "frmAlerts.frx":20A4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":20E4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2104
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboSymbols 
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1995
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
         Tip             =   "frmAlerts.frx":2120
         Sorted          =   -1  'True
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2140
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSymbol 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
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
         Caption         =   "frmAlerts.frx":215C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":218A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":21AA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboBoolean 
         Height          =   315
         Left            =   7560
         TabIndex        =   15
         Top             =   900
         Width           =   1935
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
         Tip             =   "frmAlerts.frx":21C6
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":21E6
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboOperators 
         Height          =   315
         Left            =   5580
         TabIndex        =   12
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
         Tip             =   "frmAlerts.frx":2202
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2222
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboFields 
         Height          =   315
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1995
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
         Tip             =   "frmAlerts.frx":223E
         Sorted          =   -1  'True
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":225E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtValue 
         Height          =   315
         Left            =   7560
         TabIndex        =   14
         Top             =   480
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":227A
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
         Tip             =   "frmAlerts.frx":229A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":22BA
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddField 
         Height          =   360
         Left            =   5040
         TabIndex        =   10
         Top             =   457
         Width           =   435
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
         Caption         =   "frmAlerts.frx":22D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":22FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":231C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddSymbol 
         Height          =   360
         Left            =   2400
         TabIndex        =   5
         Top             =   457
         Width           =   435
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
         Caption         =   "frmAlerts.frx":2338
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":235E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":237E
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdBackColor 
         Height          =   315
         Left            =   4800
         TabIndex        =   29
         Top             =   1830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniCheckXP chkChangeCellColor 
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   1800
         Width           =   4575
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":239A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":242A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":244A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   3000
         Top             =   240
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
         Caption         =   "frmAlerts.frx":2466
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2492
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":24B2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   5580
         Top             =   240
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
         Caption         =   "frmAlerts.frx":24CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2502
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2522
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   255
         Left            =   7560
         Top             =   240
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
         Caption         =   "frmAlerts.frx":253E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":256A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":258A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1095
      Left            =   8760
      TabIndex        =   43
      Top             =   3720
      Width           =   915
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":25A6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":25C6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":25E6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   51
         Top             =   600
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
         Caption         =   "frmAlerts.frx":2602
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":2630
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2650
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   56
         Top             =   0
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
         Caption         =   "frmAlerts.frx":266C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":2692
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":26B2
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraAction 
      Height          =   2895
      Left            =   120
      TabIndex        =   53
      Top             =   3180
      Width           =   8475
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":26CE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":26FA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":271A
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectColor gdMessageColor 
         Height          =   315
         Left            =   5385
         TabIndex        =   58
         Top             =   780
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniCheckXP chkConfirmOrder 
         Height          =   255
         Left            =   495
         TabIndex        =   60
         Top             =   2430
         Visible         =   0   'False
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
         Caption         =   "frmAlerts.frx":2736
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":278E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":27AE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkRepeatPlay 
         Height          =   255
         Left            =   3240
         TabIndex        =   62
         Top             =   1920
         Visible         =   0   'False
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
         Caption         =   "frmAlerts.frx":27CA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":282E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":284E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fgActions 
         Height          =   1755
         Left            =   240
         TabIndex        =   86
         Top             =   420
         Width           =   2715
         _cx             =   4789
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
      Begin HexUniControls.ctlUniFrameWL fraOrder 
         Height          =   855
         Left            =   3480
         TabIndex        =   64
         Top             =   2520
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
         Caption         =   "frmAlerts.frx":286A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmAlerts.frx":288A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":28AA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
            Height          =   375
            Left            =   3720
            TabIndex        =   66
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
            Caption         =   "frmAlerts.frx":28C6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmAlerts.frx":28FC
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":291C
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblOrder 
            Height          =   855
            Left            =   0
            Top             =   0
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
            Caption         =   "frmAlerts.frx":2938
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmAlerts.frx":2964
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmAlerts.frx":2984
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniTextBoxXP txtFileName 
         Height          =   285
         Left            =   3180
         TabIndex        =   68
         Top             =   1160
         Width           =   3000
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":29A0
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
         Tip             =   "frmAlerts.frx":29C0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":29E0
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPlaySound 
         Height          =   285
         Left            =   7140
         TabIndex        =   70
         Top             =   1440
         Width           =   820
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
         Caption         =   "frmAlerts.frx":29FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":2A26
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2A46
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtEmailFrom 
         Height          =   285
         Left            =   6180
         TabIndex        =   78
         Top             =   1840
         Width           =   825
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":2A62
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
         Tip             =   "frmAlerts.frx":2A82
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2AA2
      End
      Begin HexUniControls.ctlUniTextBoxXP txtMailServer 
         Height          =   285
         Left            =   3180
         TabIndex        =   80
         Top             =   2520
         Width           =   825
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":2ABE
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
         Tip             =   "frmAlerts.frx":2ADE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2AFE
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
         Height          =   285
         Left            =   6285
         TabIndex        =   82
         Top             =   1440
         Width           =   820
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
         Caption         =   "frmAlerts.frx":2B1A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":2B48
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2B68
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPagerSettings 
         Height          =   435
         Left            =   6240
         TabIndex        =   85
         Top             =   1125
         Visible         =   0   'False
         Width           =   1335
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
         Caption         =   "frmAlerts.frx":2B84
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":2BC2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2BE2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtMessage 
         Height          =   285
         Left            =   3180
         TabIndex        =   88
         Top             =   480
         Width           =   4575
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAlerts.frx":2BFE
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
         Tip             =   "frmAlerts.frx":2C1E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2C3E
      End
      Begin HexUniControls.ctlUniCheckXP chkMessageColor 
         Height          =   375
         Left            =   3240
         TabIndex        =   87
         Top             =   750
         Width           =   2205
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAlerts.frx":2C5A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":2CA0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2CC0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblEmailFrom 
         Height          =   255
         Left            =   3180
         Top             =   1605
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
         Caption         =   "frmAlerts.frx":2CDC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2D28
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2D48
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblMailServer 
         Height          =   255
         Left            =   3180
         Top             =   2280
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
         Caption         =   "frmAlerts.frx":2D64
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2DDC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2DFC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAction 
         Height          =   255
         Left            =   240
         Top             =   600
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
         Caption         =   "frmAlerts.frx":2E18
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2E44
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2E64
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFilename 
         Height          =   255
         Left            =   3180
         Top             =   945
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
         Caption         =   "frmAlerts.frx":2E80
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2EAA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2ECA
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblMessage 
         Height          =   255
         Left            =   3180
         Top             =   240
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
         Caption         =   "frmAlerts.frx":2EE6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2F3A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2F5A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   240
         Top             =   210
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
         Caption         =   "frmAlerts.frx":2F76
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":2FA4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":2FC4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraConditionAT 
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   9615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmAlerts.frx":2FE0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlerts.frx":3012
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlerts.frx":3032
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   2400
         TabIndex        =   18
         Top             =   390
         Width           =   2835
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
         Tip             =   "frmAlerts.frx":304E
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":306E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAllOrders 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   420
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
         Caption         =   "frmAlerts.frx":308A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":30C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":30E6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboManualOrders 
         Height          =   315
         Left            =   2400
         TabIndex        =   20
         Top             =   870
         Width           =   2835
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
         Tip             =   "frmAlerts.frx":3102
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":3122
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optManualOrders 
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   900
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
         Caption         =   "frmAlerts.frx":313E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":317A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":319A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAutoTradingItems 
         Height          =   315
         Left            =   2400
         TabIndex        =   22
         Top             =   1350
         Width           =   2835
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
         Tip             =   "frmAlerts.frx":31B6
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":31D6
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAutoTrade 
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1380
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
         Caption         =   "frmAlerts.frx":31F2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAlerts.frx":3240
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":3260
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniListBoxXP lstAutoTradeAlerts 
         Height          =   1815
         Left            =   5520
         TabIndex        =   24
         Top             =   480
         Width           =   3855
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
         TrapTab         =   0   'False
         Tip             =   "frmAlerts.frx":327C
         MultiSelect     =   0
         Sorted          =   0   'False
         HScroll         =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         RoundedBorders  =   0   'False
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":329C
         ManualStart     =   0   'False
         Columns         =   0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAutoTradeAlert 
         Height          =   255
         Left            =   5520
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
         Caption         =   "frmAlerts.frx":32B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAlerts.frx":32EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAlerts.frx":330C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAlerts.frm
'' Description: Alerts the user if one of the symbols on the quote board hits
''              a user alert
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/12/2010   DAJ         Added Number Bars Required for TradeSense alert
'' 07/22/2010   DAJ         Fixed bug when manual override chosen (#5812)
'' 08/06/2010   DAJ         Fixed bug with keep alive check box with TradeSense alerts (#5849)
'' 12/10/2010   DAJ         Changed over to the IsBrokerUser function
'' 03/07/2011   DAJ         Broker Disconnect Alerts
'' 05/11/2011   DAJ         Utilize IsLiveAccount function
'' 10/11/2011   DAJ         Changed label on the keep alive check box
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 11/26/2012   DAJ         Fix for auto detect issue for RSI with 2000 trades per bar timeframe
'' 01/14/2014   DAJ         Added 'Order Rejected' alert
'' 10/07/2015   DAJ         Improvements to dialogs when cannot auto detect
'' 12/02/2015   DAJ         Added Account filter for "Any Orders" alert
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kActionFraExt = 2895
Private Const kActionFraStd = 2295
Private Const kFormExtHeight = 6675 ' 6315
Private Const kFormStdHeight = 6075 ' 5715

Private Const kstrPositionMismatch = "Position Mismatch"
Private Const kstrAutoTradeItemDisabled = "Automated Trading Item Disabled"

Private Enum eGDTradeItemCols
    eGDTradeItemCol_TradeItemID = 0
    eGDTradeItemCol_Name
    eGDTradeItemCol_NumCols
End Enum

Private Enum eGDActionCols
    eGDActionCol_On = 0
    eGDActionCol_Name
    eGDActionCol_Action
    eGDActionCol_NumCols
End Enum

Private Type mPrivate
    Alert As cAlert                     ' Alert object needed for editing chart condition
    bOK As Boolean                      ' Did the user click on the OK button?
    bIsBoolean As Boolean               ' Does this alert use a boolean field?
    bRemove As Boolean                  ' Remove the alert?
    
    strOrderText As String              ' Order text
        
    strNewField As String               ' Name of the newly created field
    astrFieldInfo As New cGdArray       ' Field information array
    nAlertType As eGDAlertType          ' Alert Type
    strTimeZone As String               ' Time zone of previous selection
    dAtTime As Double                   ' Previous value of the AtTime control
    
    UpTo As cPriceEditor                ' If price gets up to... control
    DownTo As cPriceEditor              ' If price gets down to... control
    PriceBars As cGdBars                ' Bars to hold symbol properties

    ListLoading As cListLoading         ' Lists of stuff for TradeSense
    bCboInitialized As Boolean          ' Flag for whether to process indicator combobox
    
    iOrderActionBeforeEdit As Long
    bKeepActiveOff As Boolean
    bNeedEmailInfo As Boolean
    bEnableControlsInProg As Boolean
    
    bEnableKeepActAlways As Boolean
    
    strCodedText As String              ' Coded text for the Trade Sense expression
End Type
Private m As mPrivate

Private Function TradeItemCol(ByVal Col As eGDTradeItemCols) As Long
    TradeItemCol = Col
End Function

Private Function ActionCol(ByVal Col As eGDActionCols) As Long
    ActionCol = Col
End Function

Private Function ActionRow(ByVal Row As eGDAlertAction) As Long
    ActionRow = Row
End Function

Public Property Get NewField() As String
    NewField = m.strNewField
End Property
Public Property Let NewField(ByVal strNewField As String)
    m.strNewField = strNewField
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: When the form is called from someone else, do some initialization,
''              show the form, then do some afterwards processing
'' Inputs:      Alert, Alert Type
'' Returns:     True if user clicked on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Alert As cAlert, ByVal nAlertType As eGDAlertType, _
    Optional ByVal bFromBoxQB As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strSave$
    Dim eBarFlag As eGDAlertBarOption

    If Alert Is Nothing Then Exit Function          '4325
    
    If Not HasLevelForAlert(nAlertType, True) Then Exit Function        '4610

    m.bNeedEmailInfo = True
    m.bKeepActiveOff = False

    Set m.Alert = Alert
    m.nAlertType = nAlertType
    m.bRemove = False
    m.bCboInitialized = False
    strSave = Alert.EnglishString
    Alert.EditInProg = True
    
    chkAfterBarComplete.Visible = False
    chkAfterBarComplete.Enabled = False
    
    'JM 12-18-2015: need to call this here because the grids are getting loaded before showing the form
    FixFormControls Me, ALT_GRID_ROW_COLOR
        
    InitActionsGrid
    LoadActionsGrid Alert
        
    Select Case m.nAlertType
        Case eGDAlertType_AutoTrade
            SetUpAutoTradeAlert Alert
            
        Case eGDAlertType_Price
            SetUpPriceAlert Alert
            
        Case eGDAlertType_QuoteBoard
            SetUpQuoteBoardAlert Alert, bFromBoxQB
            
        Case eGDAlertType_Status
            SetUpStatusAlert Alert
            
        Case eGDAlertType_Time
            SetUpTimeAlert Alert
            
        Case eGDAlertType_TradeSense
            SetUpTradeSenseAlert Alert
                    
        Case eGDAlertType_Annot, eGDAlertType_Chart
            SetUpChartAlert Alert
    End Select
       
    EnableControls bFromBoxQB
    
    ' Default the actions grid to the message box row...
    fgActions.Row = 0
    fgActions.RowSel = 0
    
    ' Show the form
    ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR
    
    ' If the user pressed OK, save the alert
    If m.bOK = True Then
        Select Case m.nAlertType
            Case eGDAlertType_QuoteBoard
                FillInQuoteBoardAlert Alert
                
            Case eGDAlertType_AutoTrade
                FillInAutoTradeAlert Alert
                
            Case eGDAlertType_Status
                FillInStatusAlert Alert
                
            Case eGDAlertType_Time
                FillInTimeAlert Alert
                
            Case eGDAlertType_Price
                FillInPriceAlert Alert
                Alert.UpdateChartObject False
            
            Case eGDAlertType_TradeSense
                FillInTradeSenseAlert Alert
                
            Case eGDAlertType_Annot, eGDAlertType_Chart
                FillInChartAlert Alert
        End Select
        
'JM:10-05-2007: Leave awhile then remove.
'This code causes alert to immediately go inactive if user unchecks previously checked keepactive box.
'        If chkKeepActive.Enabled And chkKeepActive.Value = vbUnchecked Then
'            Alert.Deactivate = True
'        Else
'            Alert.Deactivate = False
'        End If
                
        If Not Alert.Annotation Is Nothing Then
            eBarFlag = Alert.AlertBarFlag
            If chkAfterBarComplete.Value = vbChecked Then
                If eBarFlag <> eGDAlertBar4_HiLowCompleteBar And eBarFlag <> eGDAlertBar5_CloseCompleteBar Then
                    eBarFlag = eGDAlertBar4_HiLowCompleteBar
                End If
            ElseIf eBarFlag <> eGDAlertBar_HiLowThisBar And eBarFlag <> eGDAlertBar3_CloseThisBar Then
                eBarFlag = eGDAlertBar_HiLowThisBar
            End If
            Alert.Annotation.UpdateAlert 2, True, eBarFlag
        ElseIf Not Alert.Indicator Is Nothing Then
            Alert.Indicator.UpdateAlert 2
        'ElseIf Alert.EnglishString <> strSave Then
        '    'reset the last-checked flag
        '    Alert.ResetLastChecked strSave
        End If
        
        ' As per Tim, when the user saves the alert, make sure that the action will happen
        ' if the condition is true regardless if we have already alerted on this bar...
        Alert.ResetLastChecked ""
        Alert.ResetLastBarAlerted
        
    ElseIf m.bRemove Then
        If Not Alert.Annotation Is Nothing Then
            Alert.Annotation.UpdateAlert 0
        ElseIf Not Alert.Indicator Is Nothing Then
            Alert.Indicator.UpdateAlert 0
        End If
    End If
    
    ' Return the OK status and unload the form
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmAlerts.ShowMe", eGDRaiseError_Raise

End Function

Private Sub cboIndicator_Click()
On Error Resume Next
    
    Dim Annot As cAnnotation, i&
    
    If Not m.bCboInitialized Then Exit Sub
    
    If Not m.Alert Is Nothing Then
        If m.Alert.AlertType = eGDAlertType_Annot Then
            Set Annot = m.Alert.Annotation
            If Not Annot Is Nothing Then
                IndicatorsCboToAnnot cboIndicator, Annot.AnnotChart, Annot
                Annot.UpdateAlert 2, True
                lblChartCondition.Caption = m.Alert.ChartCondition
                i = cboIndicator.ItemData(cboIndicator.ListIndex)
                chkAfterBarComplete.Enabled = Annot.EnableBarComplete(i)
            End If
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboPricePeriod_KeyPress
'' Description: Move focus when the user hits Enter
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboPricePeriod_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If chkGetsDownTo.Visible Then
            MoveFocus chkGetsDownTo
        Else
            MoveFocus tsCondition
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cboPricePeriod_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboPricePeriod_LostFocus
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboPricePeriod_LostFocus()
On Error GoTo ErrSection:

    Dim strPeriod As String             ' Adjusted periodicity string
    
    strPeriod = GetPeriodStr(cboPricePeriod.Text)
    If strPeriod <> cboPricePeriod.Text Then
        cboPricePeriod.Text = strPeriod
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cboPricePeriod_LostFocus"
    
End Sub

Private Sub chkChangeCellColor_Click()

    If chkChangeCellColor.Value = 0 Then
        CheckedCell(fgActions, eAA_ChangeBackColor, ActionCol(eGDActionCol_On)) = False
    Else
        CheckedCell(fgActions, eAA_ChangeBackColor, ActionCol(eGDActionCol_On)) = True
    End If
    fgActions.Row = eAA_ChangeBackColor
    BuildActionString

End Sub

Private Sub chkConfirmOrder_Click()
    BuildActionString eAA_PlaceOrder
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkGetsDownTo_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkGetsDownTo_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.chkGetsDownTo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkGetsUpTo_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkGetsUpTo_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.chkGetsUpTo_Click"
    
End Sub

Private Sub chkKeepActive_Click()
On Error GoTo ErrSection:

    If chkKeepActive.Enabled Then
        If Not m.bEnableControlsInProg Then
            If Not m.Alert Is Nothing Then
                If chkKeepActive.Value = vbChecked Then
                    m.Alert.Deactivate = False
                Else
                    m.Alert.Deactivate = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.chkKeepActive_Click"
    
End Sub

Private Sub chkMessageColor_Click()
    
    If chkMessageColor.Value = vbChecked Then
        gdMessageColor.Visible = True
    Else
        gdMessageColor.Visible = False
    End If
    
    BuildActionString eAA_MessageBox

End Sub

Private Sub chkRepeatPlay_Click()
    BuildActionString eAA_PlaySound
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddField_Click
'' Description: Allow the user to add a field to the quote board and alerts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddField_Click()
On Error GoTo ErrSection:

    Dim strSelected As String           ' Item currently selected in the list
    Dim lIndex As Long                  ' Index into a for loop

    m.strNewField = ""
    If frmQuotes.EditFields = True Then
        strSelected = cboFields.List(cboFields.ListIndex)
        FillFieldList
        
        If Len(m.strNewField) > 0 Then
            SelectComboItem cboFields, m.strNewField
        Else
            SelectComboItem cboFields, strSelected
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdAddField.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddSymbol_Click
'' Description: Allow the user to add a symbol to the quote board and alerts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddSymbol_Click()
On Error GoTo ErrSection:
    
    Dim astrSymbols As New cGdArray     ' Array of symbols the user want to add
    Dim lPoolRec As Long                ' Record of the new symbol in the pool
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrSymbols = frmSymbolSelector.ShowMe
    For lIndex = 0 To astrSymbols.Size - 1
        lPoolRec = g.SymbolPool.PoolRecForSymbol(astrSymbols(0))
        frmQuotes.AddSymbol lPoolRec, "Daily"
    Next lIndex
    If astrSymbols.Size > 0 Then
        FillSymbolList
        SelectComboItem cboSymbols, astrSymbols(0) & " (Daily)"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdAddSymbol.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to select a sound file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    Dim strFile As String            ' Wav or log file chosen (or last one used)

    If fgActions.Row = eAA_PlaySound Then
        strFile = Parse(fgActions.TextMatrix(ActionRow(eAA_PlaySound), ActionCol(eGDActionCol_Action)), ",", 2)
        If Len(strFile) = 0 Then
            strFile = GetIniFileProperty("WavFileLastUsed", "", "QuoteBoard", g.strIniFile)
            If Len(strFile) = 0 Then
                strFile = AddSlash(WindowsPath) & "Media"
            End If
        End If
        txtFileName.Text = CommonDialogFile(frmMain.CommonDialog1, False, "Wave Files (*.wav)|*.wav", strFile)
    ElseIf fgActions.Row = eAA_LogToFile Then
        strFile = txtFileName.Text
        If Len(strFile) = 0 Then
            strFile = GetIniFileProperty("LogToFileLastUsed", g.strAppPath & "\AlertMsgs.txt", "QuoteBoard", g.strIniFile)
        End If
        txtFileName.Text = CommonDialogFile(frmMain.CommonDialog1, False, "Text Files (*.txt)|*.txt", strFile)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdBrowse.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit the order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to send to the edit order form
    Dim strSymbol As String             ' Symbol to send to the order
    Dim astrOrder As New cGdArray       ' Array of parameters to the order
    
    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            strSymbol = RollSymbolForDate(Parse(cboSymbols.Text, "(", 1), Date)
        
        Case eGDAlertType_Annot, eGDAlertType_Chart, eGDAlertType_Price, eGDAlertType_TradeSense
            strSymbol = txtPriceSymbol.Text
    End Select
        
    With Order
        .OrderID = -1&
        
        If Len(m.strOrderText) = 0 Then
            .AccountID = AccountForOrder
            .Buy = True
            .Expiration = -1&
            .LimitPrice = 0
            .OrderType = eTT_OrderType_Market
            If SecurityType(strSymbol, True) = "S" Then
                .Quantity = 100
            Else
                .Quantity = 1
            End If
            .StopPrice = 0
            .SymbolOrSymbolID = strSymbol
        Else
            astrOrder.SplitFields m.strOrderText, ","
            
            .AccountID = CLng(ValOfText(astrOrder(5)))
            .Buy = (UCase(astrOrder(0)) = "BUY")
            .Expiration = CLng(ValOfText(astrOrder(6)))
            .LimitPrice = ValOfText(astrOrder(4))
            .OrderType = ValOfText(astrOrder(2))
            .Quantity = ValOfText(astrOrder(1))
            .StopPrice = ValOfText(astrOrder(3))
            .SymbolOrSymbolID = strSymbol
        End If
    End With
    
    If frmTTEditOrder.ShowMe(Order, Abs(Order.Buy), eGDTTEditOrderMode_FromAlert) = eGDEditOrderReturn_Submit Then
        m.strOrderText = ""
        With Order
            If .Buy Then
                astrOrder(0) = "Buy"
            Else
                astrOrder(0) = "Sell"
            End If
            astrOrder(1) = Str(.Quantity)
            astrOrder(2) = Str(.OrderType)
            astrOrder(3) = Str(.StopPrice)
            astrOrder(4) = Str(.LimitPrice)
            astrOrder(5) = Str(.AccountID)
            astrOrder(6) = Str(.Expiration)
            If Order.SymbolID = m.Alert.SymbolID Then
                astrOrder(7) = ""
            Else
                astrOrder(7) = Str(Order.SymbolID)
            End If
            
            m.strOrderText = astrOrder.JoinFields(",")
            lblOrder.Caption = OrderToCaption(m.strOrderText)
        End With
    ElseIf m.iOrderActionBeforeEdit = 2 Then
        fgActions.Cell(flexcpChecked, ActionRow(eAA_PlaceOrder), ActionCol(eGDActionCol_On)) = 2        'aardvark 4328
        lblOrder.Caption = ""
    End If
        
    BuildActionString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPagerSettings_Click
'' Description: Allow the user to configure pager settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPagerSettings_Click()
On Error GoTo ErrSection:

    frmPagerSettings.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdPagerSettings.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPlaySound_Click
'' Description: Allow the user to play back the chosen sound file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPlaySound_Click()
On Error GoTo ErrSection:

    PlaySoundFile Trim(txtFileName.Text)
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdPlaySound.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPriceLookup_Click
'' Description: Lookup the symbol for the price symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPriceLookup_Click()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Return from the lookup symbol routine
    
    strSymbol = LookupSymbol(txtPriceSymbol.Text)
    If Len(strSymbol) > 0 Then
        If txtPriceSymbol.Text <> strSymbol Then
            txtPriceSymbol.Text = strSymbol
            
            Set m.PriceBars = New cGdBars
            DM_GetBars m.PriceBars, txtPriceSymbol.Text, cboPricePeriod.Text, LastDailyDownload - 5
            
            Set m.UpTo = New cPriceEditor
            m.UpTo.Init sbUpTo, txtUpTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1)
            
            Set m.DownTo = New cPriceEditor
            m.DownTo.Init sbDownTo, txtDownTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdPriceLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSessionEnd_Click
'' Description: Change the times to the session end of the chosen symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSessionEnd_Click()
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars object
    Dim dTime As Double                 ' Session Start Time from bars
    Dim dNow As Double                  ' Now in the appropriate time zone
    
    If SetBarProperties(Bars, txtSessionSymbol.Text) = True Then
        dTime = Bars.Prop(eBARS_EndTime) / 1440#
        
        Select Case True
            Case optLocal.Value = True
                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                dNow = Now
            Case optGMT.Value = True
                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "GMT")
                dNow = ConvertTimeZone(Now, "", "GMT")
            Case optNY.Value = True
                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "NY")
                dNow = ConvertTimeZone(Now, "", "NY")
        End Select
        
        gdAtTime.Value = dTime
        If dNow < Int(dNow) + dTime Then
            gdAtDateTime.Value = Int(dNow) + dTime
        Else
            gdAtDateTime.Value = Int(dNow) + 1 + dTime
            Do While Not IsWeekday(gdAtDateTime.Value)
                gdAtDateTime.Value = gdAtDateTime.Value + 1
            Loop
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdSessionEnd_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSessionLookup_Click
'' Description: Lookup the symbol to get session information for
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSessionLookup_Click()
On Error GoTo ErrSection:
    
    Dim strSymbol As String             ' Symbol that the user selected

    strSymbol = LookupSymbol(txtSessionSymbol.Text)
    If Len(strSymbol) > 0 Then
        txtSessionSymbol.Text = strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdSessionLookup_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSessionStart_Click
'' Description: Change the times to the session start of the chosen symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSessionStart_Click()
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary bars object
    Dim dTime As Double                 ' Session Start Time from bars
    Dim dNow As Double                  ' Now in the appropriate time zone
    
    If SetBarProperties(Bars, txtSessionSymbol.Text) = True Then
        dTime = Bars.Prop(eBARS_StartTime) / 1440#
        
        Select Case True
            Case optLocal.Value = True
                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                dNow = Now
            Case optGMT.Value = True
                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "GMT")
                dNow = ConvertTimeZone(Now, "", "GMT")
            Case optNY.Value = True
                dTime = ConvertTimeZone(dTime, Bars.Prop(eBARS_ExchangeTimeZoneInf), "NY")
                dNow = ConvertTimeZone(Now, "", "NY")
        End Select
        
        gdAtTime.Value = dTime
        If dNow < Int(dNow) + dTime Then
            gdAtDateTime.Value = Int(dNow) + dTime
        Else
            gdAtDateTime.Value = Int(dNow) + 1 + dTime
            Do While Not IsWeekday(gdAtDateTime.Value)
                gdAtDateTime.Value = gdAtDateTime.Value + 1
            Loop
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.cmdSessionStart_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdVerify_Click
'' Description: Verify the expression that the user typed in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdVerify_Click()
On Error GoTo ErrSection:

    Verify True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdVerify_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgActions_AfterEdit
'' Description: Rebuild the action string after user change
'' Inputs:      Row and Column of Cell to Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgActions_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim strMsg$
    
    If Row = ActionRow(eAA_PlaceOrder) Then
        'aardvark 4328
        If fgActions.Cell(flexcpChecked, Row, Col) = 1 Then
            m.iOrderActionBeforeEdit = 2
            cmdEdit_Click
            m.iOrderActionBeforeEdit = 0
            Exit Sub
        Else
            lblOrder.Caption = ""
        End If
    ElseIf Row = ActionRow(eAA_SendEmail) Then
        If fgActions.Cell(flexcpChecked, Row, Col) = 1 Then
            If ExtremeCharts >= 1 Then
                strMsg = "Since these notifications rely upon external technologies (e.g. your ISP), Extreme Charts cannot guarantee timely email delivery. Please plan your trading accordingly."
            Else
                strMsg = "Since these notifications rely upon external technologies (e.g. your ISP), Trade Navigator cannot guarantee timely email delivery. Please plan your trading accordingly."
            End If
            InfBox strMsg, "I", "Ok", "Alert email notification"
            If chkKeepActive.Value = vbChecked Then
                chkKeepActive.Value = vbUnchecked       'JM 12-14-2009: from Tim - turn off so won't get multiple emails
                m.bKeepActiveOff = True
            End If
        End If
    End If
    
    BuildActionString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.fgActions_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgActions_AfterRowColChange
'' Description: Enable controls after the row changes
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgActions_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If NewRow >= 0 Then
        If NewRow <> OldRow Then
            fgActions.Row = NewRow
            fgActions.RowSel = NewRow
            
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.fgActions_RowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgActions_BeforeEdit
'' Description: Only allow the user to edit the first column
'' Inputs:      Row and Column of Cell to Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgActions_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    fgActions.Row = Row
    fgActions.RowSel = Row

    If Col <> ActionCol(eGDActionCol_On) Then
        Cancel = True
    End If
        
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.fgActions_BeforeEdit"
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgActions_GotFocus
'' Description: When the control gets the focus, enable/disable controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgActions_GotFocus()
On Error GoTo ErrSection:

    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.fgActions_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Display help if the user presses F1
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up after ourselves upon unload of the form
'' Inputs:      Whether or not to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.astrFieldInfo = Nothing
    
    If Not m.Alert Is Nothing Then
        m.Alert.EditInProg = False
        Set m.Alert = Nothing
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.Form.Unload", eGDRaiseError_Show
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
    If cmdCancel.Caption = "&Remove" Then m.bRemove = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, print out the alert info
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrAction As New cGdArray      ' Action properties
    Dim bActionOn As Boolean            ' Is at least one action turned on?
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdOK

    Select Case m.nAlertType
        Case eGDAlertType_QuoteBoard
            ' Make sure the user entered in a valid value
            If txtValue.Text = "" Then
                MoveFocus txtValue
                Err.Raise vbObjectError + 1000, , "You must enter in a valid value"
            End If
            
            ' If they are not doing an alert on the date or the time, do not allow
            ' a '/' or a ':' anywhere, or a '-' after the first character
            If cboFields.Text <> "Session" Then
                If InStr(Trim(txtValue.Text), "/") Then
                    MoveFocus txtValue
                    Err.Raise vbObjectError + 1000, , "'/' Only allowed in date fields"
                ElseIf InStr(2, Trim(txtValue.Text), "-") Then
                    MoveFocus txtValue
                    Err.Raise vbObjectError + 1000, , "'-' invalid in this context"
                End If
            End If
            If cboFields.Text <> "Last Tick" Then
                If InStr(Trim(txtValue.Text), ":") Then
                    MoveFocus txtValue
                    Err.Raise vbObjectError + 1000, , "':' Only allowed in time fields"
                End If
            End If
            
        Case eGDAlertType_Time
            If optEvery.Value = True Then
                If chkSunday = vbUnchecked And chkMonday = vbUnchecked And chkTuesday = vbUnchecked And chkWednesday = vbUnchecked And chkThursday = vbUnchecked And chkFriday = vbUnchecked And chkSaturday = vbUnchecked Then
                    Err.Raise vbObjectError + 1000, , "You must select at least one day to trigger the alert on"
                End If
            End If
            
        Case eGDAlertType_Price
            If Len(txtPriceSymbol.Text) = 0 Then
                InfBox "Please specify a symbol.", "I", "Ok", "Price Alert"
                GoTo ErrExit
            ElseIf ((chkGetsDownTo.Value = vbUnchecked) And (chkGetsUpTo.Value = vbUnchecked)) Then
                Err.Raise vbObjectError + 1000, , "You must enable either the 'Price Gets Down To' or 'Price Gets Up To' value"
            End If
        
        Case eGDAlertType_TradeSense
            If Len(Trim(tsCondition.Text)) = 0 Then
                Err.Raise vbObjectError + 1000, , "You must specify a TradeSense condition"
            Else
                If Verify = False Then
                    GoTo ErrExit
                End If
            End If

    End Select
    
    bActionOn = False
    
    ' check for required fields
    For lIndex = 0 To fgActions.Rows - 1
        If CheckedCell(fgActions, lIndex, ActionCol(eGDActionCol_On)) = True Then
            bActionOn = True
            astrAction.SplitFields fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action)), ","
            Select Case lIndex
                'Case eAA_MessageBox
                 
                Case eAA_LogToFile
                    If Len(astrAction(2)) = 0 Then              'aardvark 3712
                        fgActions.Row = lIndex
                        MoveFocus txtFileName
                        Err.Raise vbObjectError + 1000, , "You must enter a filename"
                    Else
                        'Mike requested a check be put in to validate file name entered by user
                        Dim fhLogFile As Long
                        
                        fhLogFile = FreeFile
                        If Len(txtFileName.Text) = 0 Then
                            txtFileName.Text = GetIniFileProperty("LogToFileLastUsed", g.strAppPath & "\AlertMsgs.txt", "QuoteBoard", g.strIniFile)
                        End If
                        
                        If Len(txtFileName.Text) > 0 Then
                            Open txtFileName.Text For Append As #fhLogFile
                            'invalid file name will result in error remaining code below will not execute
                            Close #fhLogFile
                            If FileLen(txtFileName.Text) = 0 Then
                                KillFile txtFileName.Text, True
                            End If
                            SetIniFileProperty "LogToFileLastUsed", astrAction(2), "QuoteBoard", g.strIniFile
                        End If
                    End If
                 
                'Case eAA_ChangeBackColor
                 
                Case eAA_SendPage
                    If Len(astrAction(1)) = 0 Then
                        fgActions.Row = lIndex
                        MoveFocus txtMessage
                        Err.Raise vbObjectError + 1000, , "You must enter a numeric message"
                    End If
                    If Len(GetIniFileProperty("Pager", "", "PagerSettings", g.strIniFile)) = 0 Then
                        fgActions.Row = lIndex
                        ' need to first fill in the pager settings
                        frmPagerSettings.ShowMe
                        Exit Sub
                    End If
                     
                Case eAA_SendEmail
                    If Len(astrAction(2)) = 0 Then
                        fgActions.Row = lIndex
                        MoveFocus txtFileName
                        Err.Raise vbObjectError + 1000, , "You must enter an email address to send mail to."
                    ElseIf Len(Parse(astrAction(2), ";", 4)) > 0 Then
                        InfBox "You can only enter a max of 3 email addresses.", "I", "Ok", "Email Alerts"
                        GoTo ErrExit
                    End If
                    If Len(astrAction(4)) = 0 Then
                        fgActions.Row = lIndex
                        MoveFocus txtMailServer
'                        Err.Raise vbObjectError + 1000, , "You must enter an email server name or IP address."
                    ElseIf Len(astrAction(3)) = 0 Then
                        astrAction(3) = Parse(astrAction(2), ";", 1)
                        fgActions.TextMatrix(fgActions.Row, ActionCol(eGDActionCol_Action)) = astrAction.JoinFields(",")
                    End If
                     
                    'save last used email addresses and mail server name to INI file
                    SetIniFileProperty "EmailToLastUsed", astrAction(2), "QuoteBoard", g.strIniFile
                    SetIniFileProperty "EmailFromLastUsed", astrAction(3), "QuoteBoard", g.strIniFile
                    SetIniFileProperty "MailServerLastUsed", astrAction(4), "QuoteBoard", g.strIniFile
                
                Case eAA_PlaySound
                    If Len(astrAction(1)) = 0 Then
                        fgActions.Row = lIndex
                        MoveFocus txtFileName
                        Err.Raise vbObjectError + 1000, , "You must enter a filename for a sound file."
                    ElseIf Not m.Alert Is Nothing Then
                        m.Alert.RepeatPlay = chkRepeatPlay.Value * (-1)
                    End If
                
            End Select
        End If
    Next lIndex
    
    If bActionOn = False Then
        Err.Raise vbObjectError + 1000, , "You must turn on at least one action"
    End If
    
    ' Hide the form and set the OK flag
    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center the form and initialize some
''              combo boxes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    ' Center the form
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    m.bEnableKeepActAlways = FileExist("NeverDisable.flg")
    
    Me.Icon = Picture16(ToolbarIcon("ID_Alerts"), , True)
    
    ' Fill in the operator combo box
'    cboOperators.AddItem "Less Than"
'    cboOperators.AddItem "Less Than or Equal"
'    cboOperators.AddItem "Greater Than"
'    cboOperators.AddItem "Greater Than or Equal"
'    cboOperators.AddItem "Equal"
    
    cboOperators.AddItem "<  (less than)"
    cboOperators.AddItem "<= (less or equal)"
    cboOperators.AddItem ">  (greater than)"
    cboOperators.AddItem ">= (greater or equal)"
    cboOperators.AddItem "=  (equal)"
    
    'Fill in the boolean list box...
    cboBoolean.AddItem "True"
    cboBoolean.AddItem "False"

    gdBackColor.Color = 65535
    gdMessageColor.Color = 65535

    chkMessageColor.Top = txtFileName.Top
    gdMessageColor.Top = txtFileName.Top
    
    Set m.astrFieldInfo = New cGdArray
    m.astrFieldInfo.Create eGDARRAY_Strings

    'Load internally generated TradeSense lists (Symbols, etc.)
    ' (when activate, in case list has changed)
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    With tsCondition
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .Usage = 8                      ' Only allow criteria functions
        .TurnOnEditing
        .Refresh
    End With
    
    txtNumBars.Locked = True
    txtNumBars.Enabled = False
    txtOverride.Move txtNumBars.Left, txtNumBars.Top
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the form is unloaded by some way other than code, cancel
''              the unload and hide the form to allow the ShowMe to finish
'' Inputs:      Whether or not to Cancel the unload, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Me.Hide
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdAtTime_Changed
'' Description: As the AtDateTime gets changed, keep the AtTime in sync
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdAtDateTime_Changed()
On Error GoTo ErrSection:
    
    If Visible Then
        gdAtTime.Value = gdAtDateTime.Value - Int(gdAtDateTime.Value)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.gdAtDateTime_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdAtTime_Changed
'' Description: As the AtTime gets changed, keep the AtDateTime in sync
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdAtTime_Changed()
On Error GoTo ErrSection:

    Dim dDateTime As Double             ' Value of the date/time control
    
    If Visible Then
        dDateTime = gdAtDateTime.Value
        If gdAtTime.Value <> m.dAtTime Then
            gdAtDateTime.Value = gdAtDateTime.Value - (m.dAtTime - gdAtTime.Value)
            m.dAtTime = gdAtTime.Value
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.gdAtTime_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdBackColor_Changed
'' Description: Set the back color as it changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdBackColor_Changed()
On Error GoTo ErrSection:

    BuildActionString eAA_ChangeBackColor       '5679

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.gdBackColor_Changed"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboBoolean_Click
'' Description: When the user changes the field, update the caption above the
''              list box to the entry the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboBoolean_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cboBoolean_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboFields_Click
'' Description: When the user changes the field, update the caption above the
''              list box to the entry the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboFields_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into the array
    Dim strID As String                 ' ID of the criteria (if applicable)
    Dim Criteria As New cCriteria       ' Temporary criteria object

    With cboFields
        lIndex = .ItemData(.ListIndex)
        If lIndex > -1& Then
            strID = Parse(m.astrFieldInfo(lIndex), ";", 3)
            If Len(strID) > 0 Then
                Set Criteria = g.SymbolPool.Criterias(strID)
                If Not Criteria Is Nothing Then
                    If Criteria.IsBoolean = True Then
                        txtValue.Visible = False
                        cboBoolean.Top = txtValue.Top
                        cboBoolean.Visible = True
                        
                        SetOperator "="
                        cboOperators.Enabled = False
                        m.bIsBoolean = True
                    Else
                        txtValue.Visible = True
                        cboBoolean.Visible = False
                        
                        cboOperators.Enabled = True
                        m.bIsBoolean = False
                    End If
                End If
            Else
                txtValue.Visible = True
                cboBoolean.Visible = False
                
                cboOperators.Enabled = True
                m.bIsBoolean = False
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cboFields_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboOperators_Click
'' Description: When the user changes the operator, update the caption above the
''              list box to the entry the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOperators_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cboOperators_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbols_Click
'' Description: When the user changes the symbol, update the caption above the
''              list box to the entry the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbols_Click()
On Error GoTo ErrSection:
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.cboSymbols_Click"

End Sub

Private Sub gdMessageColor_Changed()
    BuildActionString eAA_MessageBox
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lstStatusItems_Click
'' Description: As the status item changes, change the description label
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lstStatusItems_Click()
On Error GoTo ErrSection:

    Select Case UCase(lstStatusItems.Text)
        Case "ORDER STATUS CHANGE"
            lblStatusDesc.Caption = "Will trigger an alert whenever a status changes on any order."
        Case "GENESIS STREAMING"
            lblStatusDesc.Caption = "Will trigger an alert whenever Genesis streaming gets disconnected."
        Case "E-SIGNAL STREAMING"
            lblStatusDesc.Caption = "Will trigger an alert whenever E-Signal streaming gets disconnected."
        Case UCase(kstrPositionMismatch)
            lblStatusDesc.Caption = "Will trigger an alert whenever a position mismatch occurs."
        Case UCase(kstrAutoTradeItemDisabled)
            lblStatusDesc.Caption = "Will trigger an alert whenever an automated trading item gets disabled."
        Case Else
            lblStatusDesc.Caption = "Will trigger an alert whenever the connection to " & g.Broker.BrokerName(lstStatusItems.ItemData(lstStatusItems.ListIndex)) & " is dropped."
    End Select
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.lstStatusItems_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAllOrders_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAllOrders_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optAllOrders_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAt_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAt_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optAt_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAutoDetect_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAutoDetect_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optAutoDetect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAutoTrade_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAutoTrade_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optAutoTrade_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optEvery_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optEvery_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optEvery_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optGMT_Click
'' Description: Convert the times to the GMT time zone
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optGMT_Click()
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Current time

    If Visible Then
        dTime = ConvertTimeZone(gdAtDateTime.Value, m.strTimeZone, "GMT")
        gdAtDateTime.Value = dTime
        gdAtTime.Value = dTime - Int(dTime)
        
        m.strTimeZone = "GMT"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optGMT_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLocal_Click
'' Description: If the user clicks on local, convert the times to local
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLocal_Click()
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Current time

    If Visible Then
        dTime = ConvertTimeZone(gdAtDateTime.Value, m.strTimeZone, "")
        gdAtDateTime.Value = dTime
        gdAtTime.Value = dTime - Int(dTime)
        
        m.strTimeZone = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optLocal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optManualOrders_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optManualOrders_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optManualOrders_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optNY_Click
'' Description: Convert the times to the New York time zone
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optNY_Click()
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Current time

    If Visible Then
        dTime = ConvertTimeZone(gdAtDateTime.Value, m.strTimeZone, "NY")
        gdAtDateTime.Value = dTime
        gdAtTime.Value = dTime - Int(dTime)
        
        m.strTimeZone = "NY"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optNY_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optOverride_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optOverride_Click()
On Error GoTo ErrSection:

    If Visible Then
        If (Len(Trim(txtOverride.Text)) = 0) And (Len(Trim(txtNumBars.Text)) > 0) Then
            txtOverride.Text = Trim(txtNumBars.Text)
        End If
        
        EnableControls
        MoveFocus txtOverride
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optOverride_Click"

End Sub

Private Sub optSpread_Click(Index As Integer)
On Error GoTo ErrSection:

    If Not optSpread(Index).Visible Then Exit Sub
    
    If Index = 0 And m.Alert.field = "Low" Then
        m.Alert.field = "High"
        lblChartCondition.Caption = Replace(lblChartCondition.Caption, "down to", "up to")
        m.Alert.ChartCondition = lblChartCondition.Caption
    ElseIf m.Alert.field = "High" Then
        m.Alert.field = "Low"
        lblChartCondition.Caption = Replace(lblChartCondition.Caption, "up to", "down to")
        m.Alert.ChartCondition = lblChartCondition.Caption
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optSpread_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSymbol_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSymbol_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optTab_Click
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optTab_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.optTab_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_Change
'' Description: Enable/Disable controls as the expression changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_Change()
On Error GoTo ErrSection:

    EnableControls
    
    ' Don't allow the user to use an assignment in the expression for now...
    If InStr(tsCondition.Text, ":=") <> 0 Then
        InfBox "You cannot have an assignment operator in this expression.", "!", , "Expression Error"
        tsCondition.Text = Replace(tsCondition.Text, ":=", "")
        If Len(tsCondition.Text) > 0 Then
            tsCondition.SelStart = Len(tsCondition.Text)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.tsCondition_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_GotFocus
'' Description: Reinitialize the control when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_GotFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = tsCondition
    With tsCondition
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = True
        .ShowNewFunction = False
        .Usage = 8
        .TurnOnEditing
        .Refresh
    End With
    
    If Len(Trim(tsCondition.Text)) = 0 Then
        tsCondition.Text = ""
        SendKeys " "
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.tsCondition_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tsCondition_LostFocus
'' Description: Clean up after the control loses the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tsCondition_LostFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Nothing
    tsCondition.RemoveTradeSense

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.tsCondition_LostFocus"
    
End Sub

Private Sub txtAlert_Change()
On Error GoTo ErrSection:

    Dim Bars As cGdBars
    Dim strPrice$, strCaption$, d#, i&
    
    If Not txtAlert.Visible Then Exit Sub
    
    d = -999999
    If Not m.Alert Is Nothing Then
        If Not m.Alert.Annotation Is Nothing Then
            If Not m.Alert.Annotation.AnnotChart Is Nothing Then
                If Not m.Alert.Annotation.AnnotChart.Bars Is Nothing Then
                    Set Bars = m.Alert.Annotation.AnnotChart.Bars
                    d = Bars.PriceFromString(txtAlert.Text)
                    m.Alert.Value = d
                End If
            End If
        End If
    End If
    
    If Not Bars Is Nothing And d <> -999999 Then
        i = InStr(Me.lblChartCondition.Caption, " to ")
        If i > 0 Then
            strPrice = Bars.PriceDisplay(d)
            strCaption = Left(lblChartCondition.Caption, i) & "to " & strPrice
            lblChartCondition.Caption = strCaption
            If d <> ValOfText(txtAlert.Text) Then
                txtAlert.BackColor = vbYellow
                txtAlert.Font.Bold = True
            Else
                txtAlert.BackColor = vbWhite
                txtAlert.Font.Bold = False
                m.Alert.Annotation.Y(1) = d
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtAlert_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtEmailFrom_Change
'' Description: Set the from e-mail as it changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtEmailFrom_Change()
On Error GoTo ErrSection:

    BuildActionString
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtEmailFrom.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFileName_Change
'' Description: Set the filename as it changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFileName_Change()
On Error GoTo ErrSection:

    BuildActionString

    If fgActions.Row = eAA_PlaySound Then
        Enable cmdPlaySound, (Len(Trim(txtFileName.Text)) > 0)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtFileName_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtMailServer_Change
'' Description: Set the mail server as it changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtMailServer_Change()
On Error GoTo ErrSection:

    BuildActionString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtMailServer_Change"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtMessage_Change
'' Description: Set the message as it changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtMessage_Change()
On Error GoTo ErrSection:

    BuildActionString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtMessage_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceSymbol_Click 0
'' Description: If the user clicks in the symbol text box, lookup symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    
    cmdPriceLookup_Click        'aardvark 6045
    
'JM 11-30-2010: original code, leave awhile then remove if all ok
'    Dim strSymbol As String             ' Return from the lookup symbol routine
'
'    strSymbol = LookupSymbol(txtPriceSymbol.Text)
'    If Len(strSymbol) > 0 Then
'        If txtPriceSymbol.Text <> strSymbol Then
'            txtPriceSymbol.Text = strSymbol
'
'            DM_GetBars m.PriceBars, txtPriceSymbol.Text, cboPricePeriod.Text, LastDailyDownload - 5
'
'            m.UpTo.Init sbUpTo, txtUpTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1)
'            m.DownTo.Init sbDownTo, txtDownTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1)
'        End If
'    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtPriceSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceSymbol_GotFocus
'' Description: If the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceSymbol_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPriceSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtPriceSymbol_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceSymbol_KeyPress,  0
'' Description: Bring up the symbol selector with the key pressed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol returned from lookup symbol routine
    
    strSymbol = LookupSymbol(txtPriceSymbol.Text, KeyAscii)
    If Len(strSymbol) > 0 Then
        If txtPriceSymbol.Text <> strSymbol Then
            txtPriceSymbol.Text = strSymbol
            
            DM_GetBars m.PriceBars, txtPriceSymbol.Text, cboPricePeriod.Text, LastDailyDownload - 5
            m.UpTo.Init sbUpTo, txtUpTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1)
            m.DownTo.Init sbDownTo, txtDownTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1)
        End If
    End If
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtPriceSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSessionSymbol_Click 0
'' Description: If the user clicks in the symbol text box, lookup symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSessionSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Return from the lookup symbol routine
    
    strSymbol = LookupSymbol(txtSessionSymbol.Text)
    If Len(strSymbol) > 0 Then
        txtSessionSymbol.Text = strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtSessionSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSessionSymbol_GotFocus
'' Description: If the text box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSessionSymbol_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSessionSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtSessionSymbol_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSessionSymbol_KeyPress,  0
'' Description: Bring up the symbol selector with the key pressed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSessionSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol returned from lookup symbol routine
    
    strSymbol = LookupSymbol(txtSessionSymbol.Text, KeyAscii)
    If Len(strSymbol) > 0 Then
        txtSessionSymbol.Text = strSymbol
    End If
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtSessionSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    lstValue_Click
'' Description: When the user changes the value, update the caption above the
''              list box to the entry the user chose
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtValue_Change()
On Error GoTo ErrSection:
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.txtValue.Change", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable/Hide/Show controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls(Optional ByVal bFromBoxQB As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index from the action list box
    Dim strSymbol As String             ' Symbol from the symbol caption
    
    m.bEnableControlsInProg = True
    
    If m.nAlertType = eGDAlertType_Annot Or m.nAlertType = eGDAlertType_Chart Then
        fraConditionChart.Visible = True
    Else
        fraConditionChart.Visible = False
    End If
    
    Enable cboSymbols, optSymbol.Value
    
    If bFromBoxQB Then
        cboTabs.Visible = False
        cboTabs.Enabled = False
        optTab.Visible = False
        optTab.Enabled = False
    Else
        cboTabs.Visible = True
        optTab.Visible = True
        Enable cboTabs, optTab.Value
    End If
    
    Enable optAutoTrade, (cboAutoTradingItems.ListCount > 0)
    Enable optManualOrders, (cboManualOrders.ListCount > 0)
    Enable cboAutoTradingItems, optAutoTrade.Value
    Enable cboManualOrders, optManualOrders
    
    Enable lblSunday, optEvery.Value
    Enable chkSunday, optEvery.Value
    Enable lblMonday, optEvery.Value
    Enable chkMonday, optEvery.Value
    Enable lblTuesday, optEvery.Value
    Enable chkTuesday, optEvery.Value
    Enable lblWednesday, optEvery.Value
    Enable chkWednesday, optEvery.Value
    Enable lblThursday, optEvery.Value
    Enable chkThursday, optEvery.Value
    Enable lblFriday, optEvery.Value
    Enable chkFriday, optEvery.Value
    Enable lblSaturday, optEvery.Value
    Enable chkSaturday, optEvery.Value
    Enable gdAtTime, optEvery.Value
    Enable gdAtDateTime, optAt.Value
    
    Enable txtDownTo, (chkGetsDownTo.Value = vbChecked)
    Enable sbDownTo, (chkGetsDownTo.Value = vbChecked)
    Enable txtUpTo, (chkGetsUpTo.Value = vbChecked)
    Enable sbUpTo, (chkGetsUpTo.Value = vbChecked)
    
    strSymbol = Parse(cboSymbols.Text, "(", 1)
    If Left(strSymbol, 1) = "$" And IsForex(strSymbol) = False Then
        If CheckedCell(fgActions, eAA_PlaceOrder, ActionCol(eGDActionCol_On)) = True Then
            InfBox "You cannot place an order on an index", "!", , "Error"
            CheckedCell(fgActions, eAA_PlaceOrder, ActionCol(eGDActionCol_On)) = False
        End If
    ElseIf Len(m.strOrderText) = 0 Then
        If SecurityType(strSymbol, True) = "S" Then
            lblOrder.Caption = OrderToCaption("Buy,100,0,0,0," & AccountForOrder & ",-1")
        Else
            lblOrder.Caption = OrderToCaption("Buy,1,0,0,0," & AccountForOrder & ",-1")
        End If
    Else
        lblOrder.Caption = OrderToCaption(m.strOrderText)
    End If
        
    If m.nAlertType = eGDAlertType_QuoteBoard Then
        chkChangeCellColor.Visible = Not bFromBoxQB
        gdBackColor.Visible = Not bFromBoxQB
        chkShowOnCharts.Visible = False
    Else
        chkChangeCellColor.Visible = False
        gdBackColor.Visible = False
        If m.nAlertType = eGDAlertType_Price Then
            chkShowOnCharts.Visible = True
        Else
            chkShowOnCharts.Visible = False
        End If
    End If
    
    Select Case m.nAlertType
        Case eGDAlertType_AutoTrade
            If (optAutoTrade.Value = True) Or (optAllOrders.Value = True) Then
                chkKeepActive.Value = vbChecked
            Else
                chkKeepActive.Value = vbUnchecked
            End If
            
            If optAutoTrade.Value = True Or optManualOrders.Value = True Then
                With lstAutoTradeAlerts
                    If .List(1) <> "Order Price Hit" Then .AddItem "Order Price Hit", 1
                End With
            Else
                With lstAutoTradeAlerts
                    If .List(1) = "Order Price Hit" Then .RemoveItem 1
                End With
            End If
            If Not m.bEnableKeepActAlways Then Disable chkKeepActive
            
        Case eGDAlertType_Status
'original code: leave awhile then remove 01-02-2007
'            If UCase(lstStatusItems.Text) = "ORDER STATUS CHANGE" Then
'                chkDeactivate.Value = vbUnchecked
'            Else
'                chkDeactivate.Value = vbChecked
'            End If
'            Disable chkDeactivate
'modified code: leave awhile then remove 04-05-2007
'            If UCase(lstStatusItems.Text) = "GENESIS STREAMING" Or UCase(lstStatusItems.Text) = "E-SIGNAL STREAMING" Then
'                If m.Alert Is Nothing Then
'                    chkKeepActive.Value = vbUnchecked
'                ElseIf m.Alert.Deactivate Then
'                    chkKeepActive.Value = vbUnchecked
'                Else
'                    chkKeepActive.Value = vbChecked
'                End If
'                Enable chkKeepActive
'            Else
'                chkKeepActive.Value = vbChecked
'                Disable chkKeepActive
'            End If
'Dave thinks user should not be given access to the keep active check box for this type of alert 04-05-2007
            chkKeepActive.Value = vbChecked
            If Not m.bEnableKeepActAlways Then Disable chkKeepActive
        
        Case eGDAlertType_Time
            If optEvery.Value = True Then
                chkKeepActive.Value = vbChecked
            Else
                chkKeepActive.Value = vbUnchecked
            End If
            If Not m.bEnableKeepActAlways Then Disable chkKeepActive
            
        Case Else
            If m.Alert Is Nothing Then
                chkKeepActive.Value = vbUnchecked
                If Not m.bEnableKeepActAlways Then Disable chkKeepActive
            ElseIf CheckedCell(fgActions, eAA_PlaceOrder, ActionCol(eGDActionCol_On)) = True Then
                If m.bEnableKeepActAlways Then
                    If m.Alert.Deactivate Then
                        chkKeepActive.Value = vbUnchecked
                    Else
                        chkKeepActive.Value = vbChecked
                    End If
                    Enable chkKeepActive
                Else
                    chkKeepActive.Value = vbUnchecked
                    Disable chkKeepActive
                End If
            ElseIf m.Alert.Deactivate Then
                chkKeepActive.Value = vbUnchecked
                Enable chkKeepActive
            Else
                If Not m.bKeepActiveOff Then chkKeepActive.Value = vbChecked
                Enable chkKeepActive
            End If
            
            ' DAJ 08/06/2010: Moved this down into here from being its own case so that the
            ' preceding keep alive stuff happens for TradeSense alerts as well...
            If m.nAlertType = eGDAlertType_TradeSense Then
                Enable cmdOK, (Len(Trim(tsCondition.Text)) > 0)
                Enable cmdVerify, (Len(Trim(tsCondition.Text)) > 0)
                
                If optOverride.Value = True Then
                    txtNumBars.Visible = False
                    txtOverride.Visible = True
                Else
                    txtNumBars.Visible = True
                    txtOverride.Visible = False
                End If
            End If
            
    End Select
    
    lIndex = fgActions.Row
    Select Case lIndex
        Case eAA_MessageBox: ' Message Box
            If (m.nAlertType = eGDAlertType_AutoTrade) And (optAllOrders.Value = True) Then
                lblMessage.Caption = "Display order information"
                lblMessage.Width = 4575
                lblMessage.Visible = True
                txtMessage.Visible = False
                txtMessage.Text = ""
            Else
                lblMessage = "Custom Message (optional):"
                lblMessage.Visible = True
                txtMessage.Visible = True
                txtMessage.Width = 4575
            End If
                        
            lblFilename.Visible = False
            txtFileName.Visible = False
            
            lblMailServer.Visible = False
            txtMailServer.Visible = False
            lblEmailFrom.Visible = False
            txtEmailFrom.Visible = False
                        
            cmdPagerSettings.Visible = False
            cmdBrowse.Visible = False
            cmdPlaySound.Visible = False
            chkRepeatPlay.Visible = False
            chkConfirmOrder.Visible = False
            
            chkMessageColor.Visible = True
            If chkMessageColor.Value = vbChecked Then
                gdMessageColor.Visible = True
            Else
                gdMessageColor.Visible = False
            End If
                        
            fraOrder.Visible = False
                        
        Case eAA_LogToFile: ' Log to File
            lblMessage = "Custom Message (optional):"
            lblMessage.Visible = True
            txtMessage.Visible = True
            
            lblFilename.Move lblFilename.Left, 940
            lblFilename = "File:"
            lblFilename.Visible = True
            txtFileName.Move txtFileName.Left, 1160, 3700, txtMessage.Height
            txtFileName.Visible = True
            
            lblMailServer.Visible = False
            txtMailServer.Visible = False
            lblEmailFrom.Visible = False
            txtEmailFrom.Visible = False
                        
            cmdPagerSettings.Visible = False
            cmdPlaySound.Visible = False
            cmdBrowse.Move txtFileName.Left + txtFileName.Width + 50, txtFileName.Top
            cmdBrowse.Visible = True
            chkRepeatPlay.Visible = False
            chkConfirmOrder.Visible = False
            
            chkMessageColor.Visible = False
            gdMessageColor.Visible = False
            
            fraOrder.Visible = False
                
'Change cell color has been moved out of action grid
'        Case eAA_ChangeBackColor: ' Change Cell Color
'            lblMessage.Visible = False
'            txtMessage.Visible = False
'
'            lblFilename.Visible = False
'            txtFileName.Visible = False
'
'            lblMailServer.Visible = False
'            txtMailServer.Visible = False
'            lblEmailFrom.Visible = False
'            txtEmailFrom.Visible = False
'
'            'lblBackColor.Move lblMessage.Left, lblMessage.Top
'            'gdBackColor.Move txtMessage.Left, txtMessage.Top
'
'            cmdPagerSettings.Visible = False
'            cmdBrowse.Visible = False
'            cmdPlaySound.Visible = False
'
'            fraOrder.Visible = False
                
        Case eAA_SendPage: ' Send Page
            lblMessage = "Numeric Message to send (required):"
            lblMessage.Visible = True
            txtMessage.Visible = True
            
            lblFilename.Visible = False
            txtFileName.Visible = False
            
            lblMailServer.Visible = False
            txtMailServer.Visible = False
            lblEmailFrom.Visible = False
            txtEmailFrom.Visible = False
            
            cmdPagerSettings.Left = txtMessage.Left
            cmdPagerSettings.Visible = True
            cmdBrowse.Visible = False
            cmdPlaySound.Visible = False
            chkRepeatPlay.Visible = False
            chkConfirmOrder.Visible = False
            
            chkMessageColor.Visible = False
            gdMessageColor.Visible = False
            
            fraOrder.Visible = False
                
        Case eAA_SendEmail: 'send email
            lblMessage = "Custom Message (optional):"
            lblMessage.Visible = True
            txtMessage.Visible = True
            
            lblFilename.Move lblFilename.Left, 940
            lblFilename = "Send To  (Email address):"
            lblFilename.Visible = True
            txtFileName.Move txtMessage.Left, 1160, 4575, txtMessage.Height
            txtFileName.Visible = True
            
            lblEmailFrom.Visible = True
            txtEmailFrom.Visible = True
            txtEmailFrom.Move txtMessage.Left, txtEmailFrom.Top, 4575, txtMessage.Height
            
            lblMailServer.Visible = True
            txtMailServer.Visible = True
            txtMailServer.Move txtMessage.Left, txtMailServer.Top, 4575, txtFileName.Height
            
            cmdPagerSettings.Visible = False
            cmdBrowse.Visible = False
            cmdPlaySound.Visible = False
            chkRepeatPlay.Visible = False
            chkConfirmOrder.Visible = False
        
            chkMessageColor.Visible = False
            gdMessageColor.Visible = False
            
            fraOrder.Visible = False
    
        Case eAA_PlaceOrder
            lblMessage.Visible = False
            txtMessage.Visible = False
            lblFilename.Visible = False
            txtFileName.Visible = False
            lblEmailFrom.Visible = False
            txtEmailFrom.Visible = False
            lblMailServer.Visible = False
            txtMailServer.Visible = False
            cmdPagerSettings.Visible = False
            cmdBrowse.Visible = False
            cmdPlaySound.Visible = False
            chkRepeatPlay.Visible = False
            
            chkConfirmOrder.Move fgActions.Left + fgActions.Width + 100, (fgActions.Top + fgActions.Height) - chkConfirmOrder.Height
            chkConfirmOrder.Visible = True
            
            chkMessageColor.Visible = False
            gdMessageColor.Visible = False
            
            With fraOrder
                .Move lblMessage.Left, fgActions.Top 'lstAction.Top
                .Visible = True
            End With
            
        Case eAA_PlaySound
            lblMessage.Visible = False
            txtMessage.Visible = False
            
            lblFilename = "Sound File:"
            lblFilename.Move lblMessage.Left, lblMessage.Top
            lblFilename.Visible = True
            txtFileName.Move txtMessage.Left, txtMessage.Top, 3000
            txtFileName.Visible = True
            
            lblMailServer.Visible = False
            txtMailServer.Visible = False
            lblEmailFrom.Visible = False
            txtEmailFrom.Visible = False
            
            cmdPagerSettings.Visible = False
            cmdBrowse.Move txtFileName.Left + txtFileName.Width + 50, txtFileName.Top
            cmdBrowse.Visible = True
            cmdPlaySound.Top = txtFileName.Top
            cmdPlaySound.Visible = True
            chkRepeatPlay.Move txtFileName.Left, txtFileName.Top + txtFileName.Height + 100
            chkRepeatPlay.Visible = True
            chkConfirmOrder.Visible = False
            
            chkMessageColor.Visible = False
            gdMessageColor.Visible = False
            
            fraOrder.Visible = False
            
        Case eAA_MsgHistory
            lblMessage = "Custom Message (optional):"
            lblMessage.Visible = True
            txtMessage.Visible = True
            txtMessage.Width = 4575
            
            lblFilename.Visible = False
            txtFileName.Visible = False
            
            lblMailServer.Visible = False
            txtMailServer.Visible = False
            lblEmailFrom.Visible = False
            txtEmailFrom.Visible = False
            
            cmdPagerSettings.Visible = False
            cmdBrowse.Visible = False
            cmdPlaySound.Visible = False
            chkRepeatPlay.Visible = False
            chkConfirmOrder.Visible = False
                        
            chkMessageColor.Visible = False
            gdMessageColor.Visible = False
            
            fraOrder.Visible = False
                        
    End Select
    
    If m.nAlertType = eGDAlertType_Annot Or m.nAlertType = eGDAlertType_Chart Then
        Me.Height = kActionFraExt * 2 + cmdEditChartCond.Height
        If lIndex = 4 Then
            fraAction.Height = kActionFraExt
        Else
            fraAction.Height = kActionFraStd
            Me.Height = Me.Height - 500
        End If
    ElseIf lIndex = 4 Then
        fraAction.Height = kActionFraExt
        Me.Height = kFormExtHeight
    Else
        fraAction.Height = kActionFraStd
        Me.Height = kFormStdHeight
    End If
    
    FillActionControls

ErrExit:
    m.bEnableControlsInProg = False
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.EnableControls", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOperator
'' Description: Select the correct item in list box based on operator passed in
'' Inputs:      Operator to Select
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOperator(ByVal strOperator As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    strOperator = Parse(strOperator, " ", 1)
    With cboOperators
        For lIndex = 0 To .ListCount - 1
            If Parse(.List(lIndex), " ", 1) = strOperator Then
                .ListIndex = lIndex
                Exit For
            End If
        Next
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetOperator", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSymbolList
'' Description: Fill the symbol list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillSymbolList()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of Symbol/Period pairs from Quote Board
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol to put into the list

    Set astrSymbols = frmQuotes.AlertSymbols

    cboSymbols.Clear

    ' Fill in the symbols combo box
    For lIndex = 0 To astrSymbols.Size - 1
        If Len(astrSymbols(lIndex)) > 0 Then
            If IsAlpha(Parse(astrSymbols(lIndex), vbTab, 1)) Then
                strSymbol = Parse(astrSymbols(lIndex), vbTab, 1)
            Else
                strSymbol = GetSymbol(CLng(Val(Parse(astrSymbols(lIndex), vbTab, 1))))
            End If
            
            If InStr(astrSymbols(lIndex), vbTab) = 0 Then
                cboSymbols.AddItem strSymbol
            Else
                cboSymbols.AddItem strSymbol & " (" & Parse(astrSymbols(lIndex), vbTab, 2) & ")"
            End If
        End If
    Next lIndex

ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbols = Nothing
    RaiseError "frmAlerts.FillSymbolList", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFieldList
'' Description: Fill the field list with the available fields from the quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillFieldList(Optional ByVal bFromBoxQB As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strName As String               ' Name of the quote board field
    
    m.astrFieldInfo.SplitFields frmQuotes.AlertFields2, "|"
    
    cboFields.Clear

    ' Fill in the fields combo box
    For lIndex = 0 To m.astrFieldInfo.Size - 1
        strName = Parse(m.astrFieldInfo(lIndex), ";", 1)
        
        If bFromBoxQB Then
            If strName = "Open" Or strName = "High" Or strName = "Low" Or strName = "Last" Then
                cboFields.AddItem strName
                cboFields.ItemData(cboFields.NewIndex) = lIndex
            End If
        Else
            Select Case strName
                Case "Symbol", "SymbolID", "SecType", "Dirty", "Period", "Feed Symbol", "Exchange", "Description", "T"
                
                Case Else
                    cboFields.AddItem strName
                    cboFields.ItemData(cboFields.NewIndex) = lIndex
            End Select
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillFieldList", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderToCaption
'' Description: Make an english string out of the given order string
'' Inputs:      Order Text
'' Returns:     English String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderToCaption(ByVal strOrderText As String) As String
On Error GoTo ErrSection:

    Dim astrOrder As New cGdArray       ' Order broken out into an array
    Dim strReturn As String             ' String to return from the function
    Dim strSymbol As String             ' Symbol to work with
    
    If fgActions.Cell(flexcpChecked, ActionRow(eAA_PlaceOrder), ActionCol(eGDActionCol_On)) = 2 Then
        Exit Function       'aardvark 4328
    End If
    
    astrOrder.Create eGDARRAY_Strings
    astrOrder.SplitFields strOrderText, ","
    
    If astrOrder.Size >= 8 Then
        strSymbol = RollSymbolForDate(g.SymbolPool.SymbolForID(astrOrder(7)), Date)
    Else
        strSymbol = RollSymbolForDate(Parse(cboSymbols.Text, "(", 1), Date)
    End If
    
    strReturn = astrOrder(0) & " " & Format(astrOrder(1), "#,##0") & " "
    strReturn = strReturn & strSymbol & vbCrLf
    Select Case ValOfText(astrOrder(2))
        Case eTT_OrderType_Market
            strReturn = strReturn & "At MARKET" & vbCrLf
        Case eTT_OrderType_Stop
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(3)), strSymbol) & " STOP" & vbCrLf
        Case eTT_OrderType_Limit
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(4)), strSymbol) & " LIMIT" & vbCrLf
        Case eTT_OrderType_StopWithLimit
            strReturn = strReturn & "At " & PriceDisplay(ValOfText(astrOrder(3)), strSymbol) & " STOP with " & PriceDisplay(ValOfText(astrOrder(4)), strSymbol) & " LIMIT" & vbCrLf
    End Select
    strReturn = strReturn & "In Account: " & g.Broker.AccountNameForID(CLng(Val(astrOrder(5))))
    If ValOfText(astrOrder(6)) = 0 Then
        strReturn = strReturn & " GTC"
    ElseIf ValOfText(astrOrder(6)) > 0 Then
        strReturn = strReturn & " GTD: " & Format(ValOfText(astrOrder(6)), DateFormat("Format", MM_DD_YYYY))
    End If
    
    OrderToCaption = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlerts.OrderToCaption", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpQuoteBoardAlert
'' Description: Set up the user interface for a quote board alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpQuoteBoardAlert(ByVal Alert As cAlert, ByVal bFromBoxQB As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As New cGdBars             ' Temporary Bars object
    Dim bFound As Boolean               ' Did we find the field?
    
    fraConditionAT.Visible = False
    fraConditionPrice.Visible = False
    fraConditionQB.Visible = True
    fraConditionStatus.Visible = False
    fraConditionTime.Visible = False
    
    fraConditionQB.Move 120, 480  ' 120
    
    FillSymbolList
    FillTabList
    FillFieldList bFromBoxQB
    
    ' Fill in the controls appropriately
    If Len(Alert.Symbol) = 0 And Len(Alert.TabName) = 0 Then
        chkActive.Value = vbChecked
        optSymbol.Value = True
        optTab.Value = False
        cboSymbols.ListIndex = 0
        cboFields.Text = cboFields.List(0)
        SetOperator "="
        txtValue.Text = "0"
        cboBoolean.Text = "True"
    Else
        If Alert.Active Then chkActive.Value = vbChecked Else chkActive.Value = vbUnchecked
        optSymbol.Value = Alert.IsSymbol
        optTab.Value = Not Alert.IsSymbol
        SelectComboItem cboSymbols, Alert.Symbol & " (" & Alert.Period & ")"
        If Len(Alert.TabName) > 0 Then SelectComboItem cboTabs, Alert.TabName
        If Len(Alert.CriteriaID) = 0 Then
            SelectComboItem cboFields, Alert.field
        Else
            bFound = False
            For lIndex = 0 To m.astrFieldInfo.Size - 1
                If Parse(m.astrFieldInfo(lIndex), ";", 3) = Alert.CriteriaID Then
                    bFound = True
                    cboFields.Text = Parse(m.astrFieldInfo(lIndex), ";", 1)
                    Exit For
                End If
            Next lIndex
            If bFound = False Then cboFields.ListIndex = 1
        End If
        SetOperator Alert.Operator
        cboBoolean.Text = "True"
        If cboFields.Text = "Session" Then
            txtValue.Text = DateFormat(Alert.Value)
        ElseIf cboFields.Text = "Last Tick" Then
            txtValue.Text = Format(CVDate(Alert.Value))
        ElseIf Alert.AsTradingUnits Then
            SetBarProperties Bars, Alert.Symbol
            txtValue.Text = Bars.PriceDisplay(Alert.Value, True)
        Else
            txtValue.Text = Str(Alert.Value)
        End If
        If Alert.Value = 0# Then
            cboBoolean.Text = "False"
        Else
            cboBoolean.Text = "True"
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetUpQuoteBoardAlert", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInQuoteBoardAlert
'' Description: Fill in the alert with the user interface information
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInQuoteBoardAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim bAsTU As Boolean                ' As trading units?
    Dim dValue As Double                ' Value for the alert
    Dim Bars As New cGdBars             ' Temporary Bars object
    Dim strSymbol As String             ' Symbol for the alert
    Dim strPeriod As String             ' Bar period for the alert
    Dim strOrderText As String          ' Order text for the alert
    Dim strCriteriaID As String         ' Criteria ID for the alert
    Dim lIndex As Long                  ' Index into a for loop
    
    bAsTU = False
    If m.bIsBoolean = False Then
        If cboFields.Text = "Session" Or cboFields.Text = "Last Tick" Then
            dValue = DateOf(txtValue.Text)
        ElseIf cboFields.Text = "% Change" Then
            If Right(txtValue.Text, 1) = "%" Then txtValue = Left(txtValue.Text, Len(txtValue.Text) - 1)
            dValue = ValOfText(txtValue.Text)
        ElseIf InStr(txtValue.Text, "^") > 0 Then
            SetBarProperties Bars, Parse(cboSymbols.Text, "(", 1)
            dValue = Bars.PriceFromString(txtValue.Text)
            bAsTU = True
        Else
            dValue = ValOfText(txtValue.Text)
        End If
    Else
        If UCase(cboBoolean.Text) = "TRUE" Then
            dValue = 1
        Else
            dValue = 0
        End If
    End If
    
    If InStr(cboSymbols.Text, "(") = 0 Then
        strSymbol = cboSymbols.Text
        strPeriod = "Daily"
    Else
        strSymbol = Parse(cboSymbols.Text, "(", 1)
        strPeriod = Parse(cboSymbols.Text, "(", 2)
        strPeriod = Replace(strPeriod, ")", "")
    End If
    
    ' Need to do this to turn off an existing alert before applying the change...
'JM 02-17-2010: original code; does not work, leave awhile then remove when all okay
'    If Len(Alert.Symbol) > 0 Then
'        Alert.Active = False
'        Alert.CheckAlert
'    End If
    If Alert.IsSymbol Then
        frmQuotes.DisplayAlert Alert, True          '5574
    End If
    
    strCriteriaID = Parse(m.astrFieldInfo(cboFields.ItemData(cboFields.ListIndex)), ";", 3)

    With Alert
        .AlertType = eGDAlertType_QuoteBoard
        
        .IsSymbol = optSymbol.Value
        If .IsSymbol Then
            .Symbol = strSymbol
            .SymbolID = GetSymbolID(strSymbol)
            .Period = strPeriod
            .TabName = ""
        Else
            .Symbol = ""
            .SymbolID = 0&
            .Period = ""
            .TabName = cboTabs.Text
        End If
        .field = cboFields.Text
        .Operator = Parse(cboOperators.Text, " ", 1)
        .Value = dValue
        .CriteriaID = strCriteriaID
        
        .Active = (chkActive.Value = vbChecked)
        '.Deactivate = (chkdeactivate.Value = vbchecked)    'original code: leave awhile then remove 01-25-2007
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        .AsTradingUnits = bAsTU
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillInQuoteBoardAlert", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillAutoTradingItems
'' Description: Fill the automated trading item combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillAutoTradingItems()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    With cboAutoTradingItems
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblAutoTradingItem];", dbOpenDynaset)
        Do While Not rs.EOF
            .AddItem rs!Name
            .ItemData(.NewIndex) = rs!TradingItemID
            
            rs.MoveNext
        Loop
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillAutoTradingItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpAutoTradeAlert
'' Description: Set up the user interface with the auto trade alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpAutoTradeAlert(ByVal Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    fraConditionAT.Visible = True
    fraConditionPrice.Visible = False
    fraConditionQB.Visible = False
    fraConditionStatus.Visible = False
    fraConditionTime.Visible = False
    
    fraConditionAT.Move 120, 480 ' 120
    
    FillAutoTradingItems
    FillManualOrdersCombo
    LoadAccountsCombo
    
    With lstAutoTradeAlerts
        .AddItem "Order Submitted"
        .AddItem "Order Price Hit"
        .AddItem "Order Filled"
        .AddItem "Order Cancelled"
        .AddItem "Order Rejected"
    End With
    
    If Len(Alert.AutoTradeCondition) = 0 Then
        chkActive.Value = vbChecked
        optAllOrders.Value = True
        If cboManualOrders.ListCount > 0 Then cboManualOrders.ListIndex = 0
        If cboAutoTradingItems.ListCount > 0 Then cboAutoTradingItems.ListIndex = 0
        If cboAccounts.ListCount > 0 Then cboAccounts.ListIndex = 0
        lstAutoTradeAlerts.ListIndex = 0
    Else
        If Alert.Active Then chkActive.Value = vbChecked Else chkActive.Value = vbUnchecked
        Select Case Alert.OrderAlertType
            Case eGDOrderAlertType_ManualOrder
                optAutoTrade.Value = False
                optManualOrders.Value = True
                optAllOrders.Value = False
            
                mGenesis.SelectComboByItemData cboManualOrders, Alert.TradeItemID
                If cboAccounts.ListCount > 0 Then cboAccounts.ListIndex = 0
                If cboAutoTradingItems.ListCount > 0 Then cboAutoTradingItems.ListIndex = 0
            
            Case eGDOrderAlertType_AutoTrade
                optAutoTrade.Value = True
                optManualOrders.Value = False
                optAllOrders.Value = False
            
                mGenesis.SelectComboByItemData cboAutoTradingItems, Alert.TradeItemID
                If cboAccounts.ListCount > 0 Then cboAccounts.ListIndex = 0
                If cboManualOrders.ListCount > 0 Then cboManualOrders.ListIndex = 0
            
            Case eGDOrderAlertType_AllOrders
                optAutoTrade.Value = False
                optManualOrders.Value = False
                optAllOrders.Value = True
        
                mGenesis.SelectComboByItemData cboAccounts, Alert.TradeItemID
                If cboAutoTradingItems.ListCount > 0 Then cboAutoTradingItems.ListIndex = 0
                If cboManualOrders.ListCount > 0 Then cboManualOrders.ListIndex = 0
        End Select
        
        With lstAutoTradeAlerts
            For lIndex = 0 To .ListCount - 1
                If .List(lIndex) = Alert.AutoTradeCondition Then
                    .ListIndex = lIndex
                    Exit For
                End If
            Next lIndex
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetUpAutoTradeAlert", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpStatusAlert
'' Description: Set up the user interface with the status alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpStatusAlert(ByVal Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    fraConditionAT.Visible = False
    fraConditionPrice.Visible = False
    fraConditionQB.Visible = False
    fraConditionStatus.Visible = True
    fraConditionTime.Visible = False
    
    fraConditionStatus.Move 120, 480 ' 120
    
    With lstAutoTradeAlerts
        .AddItem "Order Submitted"
        .AddItem "Order Price Hit"
        .AddItem "Order Filled"
        .AddItem "Order Rejected"
    End With
    
    With lstStatusItems
        .AddItem "Order Status Change"
        .ItemData(.NewIndex) = -1
        
        If HasModule("RTG") Then
            .AddItem "Genesis Streaming"
            .ItemData(.NewIndex) = -1
        End If
        
        If HasModule("RTE") Then
            .AddItem "E-Signal Streaming"
            .ItemData(.NewIndex) = -1
        End If
        
        For lIndex = 1 To kNumBrokers - 1
            If g.Broker.IsLiveAccount(lIndex) Then
                If g.Broker.IsBrokerUser(lIndex) Then
                    .AddItem g.Broker.BrokerName(lIndex) & " Online Brokerage"
                    .ItemData(.NewIndex) = lIndex
                End If
            End If
        Next lIndex
        
        .AddItem kstrPositionMismatch
        .ItemData(.NewIndex) = -1
        
        .AddItem kstrAutoTradeItemDisabled
        .ItemData(.NewIndex) = -1
    End With
    
    If Len(Alert.StatusItem) = 0 Then
        chkActive.Value = vbChecked
        If lstStatusItems.ListCount > 0 Then lstStatusItems.ListIndex = 0
    Else
        If Alert.Active Then chkActive.Value = vbChecked Else chkActive.Value = vbUnchecked
        With lstStatusItems
            For lIndex = 0 To .ListCount - 1
                If .List(lIndex) = Alert.StatusItem Then
                    .ListIndex = lIndex
                    Exit For
                End If
            Next lIndex
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetUpStatusAlert", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInAutoTradeAlert
'' Description: Fill in the automated trading alert from the user interface
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInAutoTradeAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With Alert
        .AlertType = eGDAlertType_AutoTrade
        Select Case True
            Case optAutoTrade
                .OrderAlertType = eGDOrderAlertType_AutoTrade
                .TradeItemID = cboAutoTradingItems.ItemData(cboAutoTradingItems.ListIndex)
            Case optManualOrders
                .OrderAlertType = eGDOrderAlertType_ManualOrder
                .TradeItemID = cboManualOrders.ItemData(cboManualOrders.ListIndex)
            Case optAllOrders
                .OrderAlertType = eGDOrderAlertType_AllOrders
                .TradeItemID = cboAccounts.ItemData(cboAccounts.ListIndex)
        End Select
        .AutoTradeCondition = lstAutoTradeAlerts.Text
        
        .Active = (chkActive.Value = vbChecked)
        '.Deactivate = (chkDeactivate.Value = vbChecked)        'original code: leave awhile then remove 01-25-2007
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.FillInAutoTradeAlert", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInStatusAlert
'' Description: Fill in the status alert from the user interface
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInStatusAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With Alert
        .AlertType = eGDAlertType_Status
        .StatusItem = lstStatusItems.Text
        .Broker = lstStatusItems.ItemData(lstStatusItems.ListIndex)
        
        .Active = (chkActive.Value = vbChecked)
        '.Deactivate = (chkDeactivate.Value = vbChecked)        'original code: leave awhile then remove 01-25-2007
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.FillInStatusAlert", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitActionsGrid
'' Description: Initialize the actions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitActionsGrid()
On Error GoTo ErrSection:

    With fgActions
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = eAA_NumActions
        .FixedRows = 0
        .Cols = ActionCol(eGDActionCol_NumCols)
        .FixedCols = 0
        
        .ColDataType(ActionCol(eGDActionCol_On)) = flexDTBoolean
        
        .ColHidden(ActionCol(eGDActionCol_Action)) = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.InitActionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadActionsGrid
'' Description: Load the actions grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadActionsGrid(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With fgActions
        .Redraw = flexRDNone
        
        .TextMatrix(eAA_MessageBox, ActionCol(eGDActionCol_Name)) = "Pop-Up Message"
        .TextMatrix(eAA_LogToFile, ActionCol(eGDActionCol_Name)) = "Log Message To File"
        .TextMatrix(eAA_ChangeBackColor, ActionCol(eGDActionCol_Name)) = "Change Cell Color"
        .TextMatrix(eAA_SendPage, ActionCol(eGDActionCol_Name)) = "Send Page"
        .TextMatrix(eAA_SendEmail, ActionCol(eGDActionCol_Name)) = "Send E-Mail"
        .TextMatrix(eAA_PlaceOrder, ActionCol(eGDActionCol_Name)) = "Place Order"
        .TextMatrix(eAA_PlaySound, ActionCol(eGDActionCol_Name)) = "Play Sound"
        .TextMatrix(eAA_MsgHistory, ActionCol(eGDActionCol_Name)) = "Message History"
        
        For lIndex = 0 To eAA_NumActions - 1
            .TextMatrix(lIndex, ActionCol(eGDActionCol_Action)) = Alert.ActionString(lIndex)
            If Parse(Alert.ActionString(lIndex), ",", 1) = "0" Then
                CheckedCell(fgActions, lIndex, ActionCol(eGDActionCol_On)) = False
            Else
                CheckedCell(fgActions, lIndex, ActionCol(eGDActionCol_On)) = True
            End If
            If lIndex = eAA_PlaceOrder Then
                m.strOrderText = Parse(Alert.ActionString(eAA_PlaceOrder), Chr(34), 2)
            End If
        Next lIndex
        
        .RowHidden(ActionRow(eAA_MsgHistory)) = True
        Select Case m.nAlertType
            Case eGDAlertType_QuoteBoard
                If ValOfText(Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 1)) = 1 Then
                    chkChangeCellColor.Value = vbChecked
                    gdBackColor.Color = ValOfText(Parse(Alert.ActionString(eAA_ChangeBackColor), ",", 2))
                Else
                    chkChangeCellColor.Value = vbUnchecked
                End If
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = False
                .RowHidden(ActionRow(eAA_PlaySound)) = False
            Case eGDAlertType_AutoTrade
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = True
                .RowHidden(ActionRow(eAA_PlaySound)) = False
            Case eGDAlertType_Status
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = True
                .RowHidden(ActionRow(eAA_PlaySound)) = False
            Case eGDAlertType_Price
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = False
                .RowHidden(ActionRow(eAA_PlaySound)) = False
            Case eGDAlertType_Time
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = True
                .RowHidden(ActionRow(eAA_PlaySound)) = False
            Case eGDAlertType_Chart, eGDAlertType_Annot
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = False
                .RowHidden(ActionRow(eAA_PlaySound)) = False
            Case eGDAlertType_TradeSense
                .RowHidden(ActionRow(eAA_MessageBox)) = False
                .RowHidden(ActionRow(eAA_LogToFile)) = False
                .RowHidden(ActionRow(eAA_ChangeBackColor)) = True
                .RowHidden(ActionRow(eAA_SendPage)) = False
                .RowHidden(ActionRow(eAA_SendEmail)) = False
                .RowHidden(ActionRow(eAA_PlaceOrder)) = False
                .RowHidden(ActionRow(eAA_PlaySound)) = False
        End Select
        
        SetBackColors fgActions
        
        .ColWidth(ActionCol(eGDActionCol_On)) = 400
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.LoadActionsGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillActionControls
'' Description: Fill in the action controls for the currently selected action
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillActionControls()
On Error GoTo ErrSection:

    Dim astrAction As New cGdArray      ' Split out action parameters
    Dim nAction As eGDAlertAction       ' Action selected

    If (fgActions.Row >= fgActions.FixedRows) And (fgActions.Row < fgActions.Rows) Then
        astrAction.SplitFields fgActions.TextMatrix(fgActions.Row, ActionCol(eGDActionCol_Action)), ","
        nAction = fgActions.Row
        Select Case nAction
            Case eAA_MessageBox
                txtMessage.Text = astrAction(1)
                If Len(astrAction(2)) > 0 Then
                    gdMessageColor.Color = ValOfText(astrAction(2))     'must do this before setting checkbox value
                    chkMessageColor.Value = vbChecked
                End If
            Case eAA_LogToFile
                txtMessage.Text = astrAction(1)
                If Len(astrAction(2)) > 0 Then              'aardvark 3712
                    txtFileName.Text = astrAction(2)
                Else
                    txtFileName.Text = GetIniFileProperty("LogToFileLastUsed", g.strAppPath & "\AlertMsgs.txt", "QuoteBoard", g.strIniFile)
                End If
            Case eAA_ChangeBackColor
                If Len(astrAction(1)) > 0 Then
                    gdBackColor.Color = Val(astrAction(1))
                Else
                    gdBackColor.Color = 65535
                End If
            Case eAA_SendPage
                txtMessage.Text = astrAction(1)
            Case eAA_SendEmail
                txtMessage.Text = astrAction(1)
                If Len(astrAction(2)) > 0 Then
                    txtFileName.Text = astrAction(2)
                ElseIf m.bNeedEmailInfo Then
                    txtFileName.Text = GetIniFileProperty("EmailToLastUsed", "", "QuoteBoard", g.strIniFile)
                End If
                If Len(astrAction(3)) > 0 Then
                    txtEmailFrom.Text = astrAction(3)
                ElseIf m.bNeedEmailInfo Then
                    txtEmailFrom.Text = GetIniFileProperty("EmailFromLastUsed", "", "QuoteBoard", g.strIniFile)
                End If
                If Len(astrAction(4)) > 0 Then
                    txtMailServer.Text = astrAction(4)
                ElseIf m.bNeedEmailInfo Then
                    txtMailServer.Text = GetIniFileProperty("MailServerLastUsed", "", "QuoteBoard", g.strIniFile)
                End If
                m.bNeedEmailInfo = False
                
            Case eAA_PlaceOrder
                If Len(astrAction(1)) = 0 Then
                    If SecurityType(Parse(cboSymbols.Text, "(", 1)) = "S" Then
                        lblOrder.Caption = OrderToCaption("Buy,100,0,0,0," & AccountForOrder & ",-1")
                    Else
                        lblOrder.Caption = OrderToCaption("Buy,1,0,0,0," & AccountForOrder & ",-1")
                    End If
                Else
                    lblOrder.Caption = OrderToCaption(astrAction(1))
                End If
                If astrAction.Size > 2 Then
                    chkConfirmOrder.Value = astrAction(2)
                Else
                    chkConfirmOrder.Value = 1
                End If
            Case eAA_PlaySound
                If Len(astrAction(1)) > 0 Then
                    txtFileName.Text = astrAction(1)
                Else
                    txtFileName.Text = GetIniFileProperty("WavFileLastUsed", "", "QuoteBoard", g.strIniFile)
                End If
                If Len(astrAction(2)) > 0 Then
                    chkRepeatPlay = astrAction(2)           'Abs(m.Alert.RepeatPlay)
                End If
                
            Case eAA_MsgHistory
                txtMessage.Text = astrAction(1)
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillActionControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildActionString
'' Description: Build the action string from the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildActionString(Optional ByVal iRow = -1)
On Error GoTo ErrSection:

    Dim strActive As String             ' Is current action active?
    Dim strOrderText As String          ' Order text
    
    Dim lRow As Long                    '5679 - use local variable to prevent triggering FillActionControls prematurely

    If CheckedCell(fgActions, fgActions.Row, ActionCol(eGDActionCol_On)) = True Then
        strActive = "1"
    Else
        strActive = "0"
    End If
        
    If Len(m.strOrderText) = 0 Then
        If SecurityType(Parse(cboSymbols.Text, "(", 1)) = "S" Then
            strOrderText = "Buy,100,0,0,0," & AccountForOrder & ",-1"
        Else
            strOrderText = "Buy,1,0,0,0," & AccountForOrder & ",-1"
        End If
    Else
        strOrderText = m.strOrderText
    End If
    
    With fgActions
        If iRow >= .FixedRows And iRow < .Rows Then
            lRow = iRow
        Else
            lRow = .Row
        End If
        
        Select Case lRow
            Case eAA_MessageBox
                If chkMessageColor.Value = vbChecked Then
                    .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & Trim(txtMessage.Text) & Chr(34) & "," & Str(gdMessageColor.Color)
                Else
                    .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & Trim(txtMessage.Text) & Chr(34)
                End If
            Case eAA_LogToFile
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & Trim(txtMessage.Text) & Chr(34) & "," & Trim(txtFileName.Text)
            Case eAA_ChangeBackColor
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Str(gdBackColor.Color)
            Case eAA_SendPage
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & Trim(txtMessage.Text) & Chr(34)
            Case eAA_SendEmail
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & Trim(txtMessage.Text) & Chr(34) & "," & Trim(txtFileName.Text) & "," & Trim(txtEmailFrom.Text) & "," & Trim(txtMailServer.Text)
            Case eAA_PlaceOrder
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & strOrderText & Chr(34) & "," & Str(chkConfirmOrder.Value)
                'must disable the keep-active feature
                If strActive = "1" Then
                    chkKeepActive.Value = 0
                    If Not m.bEnableKeepActAlways Then Disable chkKeepActive
                Else
                    Enable chkKeepActive
                End If
            Case eAA_PlaySound
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Trim(txtFileName.Text) & "," & chkRepeatPlay.Value
            Case eAA_MsgHistory
                .TextMatrix(lRow, ActionCol(eGDActionCol_Action)) = strActive & "," & Chr(34) & Trim(txtMessage.Text) & Chr(34)
        End Select
        
        If lRow <> .Row Then .Row = lRow            'to trigger FillActionControls
        
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.BuildActionString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillTabList
'' Description: Fill the quote board tab combo list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillTabList()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lCurrTab As Long
    
    lCurrTab = -1
    For lIndex = 0 To frmQuotes.TabRecords - 1
        If frmQuotes.TabStr(eGDTabSettings_Style, lIndex) = eGDQuoteStyle_Grid Then
            cboTabs.AddItem frmQuotes.TabStr(eGDTabSettings_Name, lIndex)
            If lIndex = frmQuotes.vsTab.CurrTab Then lCurrTab = cboTabs.ListCount - 1
        End If
    Next lIndex
    
    If lCurrTab >= 0 And lCurrTab < cboTabs.ListCount Then
        cboTabs.ListIndex = lCurrTab
    Else
        cboTabs.ListIndex = 0
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillTabList"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectComboItem
'' Description: Select a combo item from a combo box
'' Inputs:      Combo Box, Selection
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectComboItem(cbo As ctlUniComboImageXP, ByVal strSelection As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With cbo
        .ListIndex = 0
        For lIndex = 0 To .ListCount - 1
            If .List(lIndex) = strSelection Then
                .ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.SelectComboItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillManualOrdersCombo
'' Description: Fill the manual orders combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillManualOrdersCombo()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As New cPtOrder           ' Temporary order object
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rs.EOF
        If IsOpenOrder(rs!Status) = True Then
            If Order.Load(rs!OrderID) Then
                cboManualOrders.AddItem Order.OrderText & " - " & Order.BrokerID
                cboManualOrders.ItemData(cboManualOrders.NewIndex) = Order.OrderID
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillManualOrdersCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpTimeAlert
'' Description: Set up the controls from the time alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpTimeAlert(ByVal Alert As cAlert)
On Error GoTo ErrSection:

    fraConditionAT.Visible = False
    fraConditionPrice.Visible = False
    fraConditionQB.Visible = False
    fraConditionStatus.Visible = False
    fraConditionTime.Visible = True
    
    fraConditionTime.Move 120, 480  ' 120
    
    With Alert
        If .TriggerTime = 0 Then
            chkActive.Value = vbChecked
            optAt.Value = True
            
            chkSunday.Value = vbUnchecked
            chkMonday.Value = vbChecked
            chkTuesday.Value = vbChecked
            chkWednesday.Value = vbChecked
            chkThursday.Value = vbChecked
            chkFriday.Value = vbChecked
            chkSaturday.Value = vbUnchecked
            
            gdAtTime.Value = 0.3125 ' 7:30 am
            gdAtDateTime.Value = Date + 0.3125
            optLocal.Value = True
            m.strTimeZone = ""
            m.dAtTime = gdAtTime.Value
        Else
            If .Active Then chkActive.Value = vbChecked Else chkActive.Value = vbUnchecked
            optEvery.Value = .Weekday
            optAt.Value = Not .Weekday
            
            If Mid(.WeekdayMask, vbSunday, 1) = "1" Then chkSunday.Value = vbChecked Else chkSunday.Value = vbUnchecked
            If Mid(.WeekdayMask, vbMonday, 1) = "1" Then chkMonday.Value = vbChecked Else chkMonday.Value = vbUnchecked
            If Mid(.WeekdayMask, vbTuesday, 1) = "1" Then chkTuesday.Value = vbChecked Else chkTuesday.Value = vbUnchecked
            If Mid(.WeekdayMask, vbWednesday, 1) = "1" Then chkWednesday.Value = vbChecked Else chkWednesday.Value = vbUnchecked
            If Mid(.WeekdayMask, vbThursday, 1) = "1" Then chkThursday.Value = vbChecked Else chkThursday.Value = vbUnchecked
            If Mid(.WeekdayMask, vbFriday, 1) = "1" Then chkFriday.Value = vbChecked Else chkFriday.Value = vbUnchecked
            If Mid(.WeekdayMask, vbSaturday, 1) = "1" Then chkSaturday.Value = vbChecked Else chkSaturday.Value = vbUnchecked
            
            If .Weekday Then
                gdAtTime.Value = .TriggerTime
                gdAtDateTime.Value = Date + .TriggerTime
            Else
                gdAtTime.Value = .TriggerTime
                gdAtDateTime.Value = .TriggerTime
            End If
            m.dAtTime = gdAtTime.Value
            
            Select Case .TimeZone
                Case ""
                    optLocal.Value = True
                Case "GMT"
                    optGMT.Value = True
                Case "NY"
                    optNY.Value = True
            End Select
            m.strTimeZone = .TimeZone
        End If
    End With
    
    txtSessionSymbol.Text = GetSymbol(ActiveChart.SymbolID)


ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetUpTimeAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInTimeAlert
'' Description: Fill in the time alert from the controls
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInTimeAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With Alert
        .AlertType = eGDAlertType_Time
        .Weekday = optEvery.Value
        
'4258 - need to clear this out else will end up concatenating to existing setting & everything after initial 0's & 1's are ignored
        .WeekdayMask = ""
        
        If chkSunday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        If chkMonday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        If chkTuesday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        If chkWednesday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        If chkThursday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        If chkFriday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        If chkSaturday = vbChecked Then .WeekdayMask = .WeekdayMask & "1" Else .WeekdayMask = .WeekdayMask & "0"
        
        If optEvery.Value = True Then
            .TriggerTime = gdAtTime.Value
        Else
            .TriggerTime = gdAtDateTime.Value
        End If
        
        Select Case True
            Case optLocal.Value = True
                .TimeZone = ""
            Case optGMT.Value = True
                .TimeZone = "GMT"
            Case optNY.Value = True
                .TimeZone = "NY"
        End Select
    
        .Active = (chkActive.Value = vbChecked)
        '.Deactivate = (chkDeactivate.Value = vbChecked)        'original code: leave awhile then remove 01-25-2007
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillInTimeAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Lookup a symbol for the user to trade
'' Inputs:      Starting Symbol, Key Pressed
'' Returns:     Symbol selected (blank if Cancel)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LookupSymbol(Optional ByVal strSymbol As String = "", Optional ByVal KeyAscii As Long = 0&) As String
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(strSymbol, False, True, "Symbol to get Session Times for")
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol to get Session Times for", False, False)
    End If
    
    If astrSymbol.Size > 0 Then
        LookupSymbol = astrSymbol(0)
    Else
        LookupSymbol = ""
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlerts.LookupSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpPriceAlert
'' Description: Set up the controls from the price alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpPriceAlert(ByVal Alert As cAlert)
On Error GoTo ErrSection:

    fraConditionAT.Visible = False
    fraConditionPrice.Visible = True
    fraConditionQB.Visible = False
    fraConditionStatus.Visible = False
    fraConditionTime.Visible = False
    
    fraConditionPrice.Move 120, 480  ' 120
    
    fraPriceExtremes.Visible = True
    tsCondition.Visible = False
    cmdVerify.Visible = False
    fraNumDays.Visible = False
    
    With cboPricePeriod
        .AddItem "Daily"
        .AddItem "60 Minute"
        .AddItem "30 Minute"
        .AddItem "5 Minute"
        .AddItem "1 Minute"
        
        .ListIndex = 0
    End With
    
    Set m.PriceBars = New cGdBars
    Set m.UpTo = New cPriceEditor
    Set m.DownTo = New cPriceEditor
    
    If Len(Alert.Symbol) = 0 Then
        chkActive.Value = vbChecked
        txtPriceSymbol.Text = ""
        If Len(ActiveChart.Chart.SpreadSymbols) > 0 Then
            'give user option to select a different symbol      -5194
            InfBox "The active chart is a spread chart. Price alerts are not valid for spread charts. You will need to specify a symbol.", "I", "Ok", "Price Alert"
            m.UpTo.Init sbUpTo, txtUpTo, Nothing
            m.DownTo.Init sbDownTo, txtDownTo, Nothing
        Else
            txtPriceSymbol.Text = GetSymbol(ActiveChart.SymbolID)
            DM_GetBars m.PriceBars, txtPriceSymbol.Text, cboPricePeriod.Text, LastDailyDownload - 5
            m.UpTo.Init sbUpTo, txtUpTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1), -99999
            m.DownTo.Init sbDownTo, txtDownTo, m.PriceBars, m.PriceBars(eBARS_Close, m.PriceBars.Size - 1), -999999         '5509
        End If
    Else
        If Alert.Active Then chkActive.Value = vbChecked Else chkActive.Value = vbUnchecked
        txtPriceSymbol.Text = Alert.Symbol
        cboPricePeriod.Text = Alert.Period
        
        DM_GetBars m.PriceBars, txtPriceSymbol.Text, cboPricePeriod.Text, LastDailyDownload - 5
        m.UpTo.Init sbUpTo, txtUpTo, m.PriceBars, Alert.GetsUpToPrice, -99999
        m.DownTo.Init sbDownTo, txtDownTo, m.PriceBars, Alert.GetsDownToPrice, -99999
        
        If Alert.UseGetsDownTo Then chkGetsDownTo.Value = vbChecked Else chkGetsDownTo.Value = vbUnchecked
        If Alert.UseGetsUpTo Then chkGetsUpTo.Value = vbChecked Else chkGetsUpTo.Value = vbUnchecked
        If Alert.ShowOnCharts Then chkShowOnCharts.Value = vbChecked Else chkShowOnCharts.Value = vbUnchecked
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetUpPriceAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpTradeSenseAlert
'' Description: Set up the controls from the TradeSense alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpTradeSenseAlert(ByVal Alert As cAlert)
On Error GoTo ErrSection:

    Dim lTop As Long                    ' Top of the control

    fraConditionAT.Visible = False
    fraConditionPrice.Visible = True
    fraConditionQB.Visible = False
    fraConditionStatus.Visible = False
    fraConditionTime.Visible = False
    
    fraConditionPrice.Move 120, 480  ' 120

    fraPriceExtremes.Visible = False
    fraNumDays.Visible = True
    fraNumDays.Move lblPriceSymbol.Left, fraPriceExtremes.Top
    
    lTop = fraNumDays.Top + fraNumDays.Height + 60
    
    tsCondition.Visible = True
    tsCondition.Move 120, lTop, fraConditionPrice.Width - 240, fraConditionPrice.Height - lTop - 60
    cmdVerify.Visible = True

    With cboPricePeriod
        .AddItem "Daily"
        .AddItem "60 Minute"
        .AddItem "30 Minute"
        .AddItem "5 Minute"
        .AddItem "1 Minute"
        
        .ListIndex = 0
    End With
    
    If Len(Alert.Symbol) = 0 Then
        chkActive.Value = vbChecked
        txtPriceSymbol.Text = GetSymbol(ActiveChart.SymbolID)
        tsCondition.Text = ""
        
        optAutoDetect.Value = True
        txtNumBars.Text = ""
        txtOverride.Text = ""
    Else
        If Alert.Active Then chkActive.Value = vbChecked Else chkActive.Value = vbUnchecked
        txtPriceSymbol.Text = Alert.Symbol
        cboPricePeriod.Text = Alert.Period
        
        tsCondition.Text = Alert.PriceCondition
        Verify Me.Visible
        
        If Alert.AutoDetect Then
            optAutoDetect.Value = True
        Else
            optOverride.Value = True
        End If
        If Alert.NumBarsCalc = -1& Then
            txtNumBars.Text = ""
        Else
            txtNumBars.Text = Str(Alert.NumBarsCalc)
        End If
        If Alert.NumBarsOver = -1& Then
            txtOverride.Text = ""
        Else
            txtOverride.Text = Str(Alert.NumBarsOver)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.SetUpTradeSenseAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInPriceAlert
'' Description: Fill in the price alert from the controls
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInPriceAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With Alert
        .AlertType = eGDAlertType_Price
        
        .Symbol = txtPriceSymbol.Text
        .SymbolID = GetSymbolID(.Symbol)
        .Period = cboPricePeriod.Text
        .UseGetsDownTo = (chkGetsDownTo.Value = vbChecked)
        .GetsDownToPrice = m.DownTo.Price
        .UseGetsUpTo = (chkGetsUpTo.Value = vbChecked)
        .GetsUpToPrice = m.UpTo.Price
        .ResetUseCloseOfBar
        If chkShowOnCharts.Value = vbChecked Then
            .ShowOnCharts = True
        Else
            .ShowOnCharts = False
        End If
        
        .Active = (chkActive.Value = vbChecked)
        '.Deactivate = (chkDeactivate.Value = vbChecked)        'orignal code: leave awhile then remove 01-25-2007
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillInPriceAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInTradeSenseAlert
'' Description: Fill in the TradeSense alert from the controls
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInTradeSenseAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With Alert
        .AlertType = eGDAlertType_TradeSense
        
        .Symbol = txtPriceSymbol.Text
        .SymbolID = GetSymbolID(.Symbol)
        .Period = cboPricePeriod.Text
        
        .AutoDetect = optAutoDetect.Value
        If Len(Trim(txtNumBars.Text)) = 0 Then
            .NumBarsCalc = -1&
        Else
            .NumBarsCalc = CLng(ValOfText(txtNumBars.Text))
        End If
        If (Len(Trim(txtOverride.Text)) = 0) Or (optOverride = False) Then
            .NumBarsOver = -1&
        Else
            .NumBarsOver = CLng(ValOfText(txtOverride.Text))
        End If
        
        .PriceCondition = tsCondition.Text
        .Active = (chkActive.Value = vbChecked)
        '.Deactivate = (chkDeactivate.Value = vbChecked)        'original code: leave awhile then remove 01-25-2007
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.FillInTradeSenseAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Verify
'' Description: Verify the condition
'' Inputs:      None
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Verify(Optional ByVal bShowMsg As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim Expr As New cExpression         ' Expression to verify condition
    Dim Func As New cFunction           ' Temporary function object
    Dim Inputs As New cInputs           ' Collection of inputs for the expression
    Dim bExtraInputs As Boolean         ' Does the expression have extra inputs?
    Dim strNotKnown As String           ' Inputs that are not known
    Dim lIndex As Long                  ' Index into a for loop
    Dim strParmName As String           ' Parameter Name
    Dim lNumBars As Long                ' Number of bars necessary
    Dim strCodedText As String          ' Coded text expression
    Dim strExpression As String         ' Expression to evaluate
    Dim bReturn As Boolean              ' Return value for the function

    If Len(Trim(tsCondition.Text)) = 0 Then
        If bShowMsg Then InfBox "Must specify an expression", "!", , "TradeSense Alert Error"
    Else
        strExpression = tsCondition.Text
        With Expr
            .PortfolioNavigator = False
            .Functions = g.Functions
            .ValidateFunctionRule strExpression
            strCodedText = .CodedText
            
            If m.ListLoading Is Nothing Then
                Set m.ListLoading = New cListLoading
                m.ListLoading.Load
            End If
            
            tsCondition.TurnOffEditing
            tsCondition.TextRTF = Func.GetRTF(.EditText)
            tsCondition.ExprIsFormatted = True
            
            bExtraInputs = False
            strNotKnown = ""
            If Not Expr.Inputs Is Nothing Then
                Set Inputs = Expr.Inputs
                For lIndex = 1 To Inputs.Count
                    strParmName = Inputs.Item(lIndex).ParmName
                    If Inputs.Item(lIndex).ParmTypeID <> 5 Then
                        strNotKnown = strNotKnown & "|" & strParmName
                        bExtraInputs = True
                    ElseIf UCase(Left(strParmName, 7)) <> "MARKET1" Then
                        If UCase(strParmName) <> "DAILY" And UCase(strParmName) <> "WEEKLY" And _
                                Left(strParmName, 1) <> Chr(34) And Right(strParmName, 1) <> Chr(34) Then
                            strNotKnown = strNotKnown & "|" & strParmName
                            bExtraInputs = True
                        End If
                    End If
                Next lIndex
            End If
            
            If bExtraInputs Then
                If bShowMsg Then InfBox "There are unrecognized inputs in your expression:|" & strNotKnown & "|", "!", , "TradeSense Alert Error"
            ElseIf .FunctionReturnType <> 3 And .FunctionReturnType <> 6 Then
                If bShowMsg Then InfBox "Expression must be a Boolean Expression", "!", , "TradeSense Alert Error"
            Else
                bReturn = EngineVerify(strCodedText)
                If bReturn Then
                    m.strCodedText = strCodedText
                Else
                    m.strCodedText = ""
                End If
            End If
        End With
    End If
    
    If bReturn = True Then
        bReturn = AutoDetect(bShowMsg)
    End If
    
    Verify = bReturn

ErrExit:
    Set Expr = Nothing
    Set Func = Nothing
    Exit Function
    
ErrSection:
    Set Expr = Nothing
    Set Func = Nothing
    RaiseError "frmAlerts.Verify"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditChartCond_Click
'' Description: If the user clicks the edit chart condition button bring up condition builder
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditChartCond_Click()
On Error GoTo ErrSection:

    If Not m.Alert Is Nothing Then
        frmConditionBuilder.ShowMe Nothing, , eType_Alert, , m.Alert
        lblChartCondition.Caption = m.Alert.ChartCondition
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.cmdEditChartCond_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpChartAlert
'' Description: Set up the controls from the chart alert
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpChartAlert(Alert As cAlert)
On Error GoTo ErrSection:
    
    Dim bEdit As Boolean
    
    Dim Chart As cChart
    Dim Annot As cAnnotation
    Dim Ind As cIndicator
    
    Dim cboIdx&, nIndId&, nPaneID&, nPriceField&
    
    fraConditionQB.Visible = False
    fraConditionAT.Visible = False
    fraConditionStatus.Visible = False
                
    If Alert.AlertType = eGDAlertType_Annot Then
        cmdEditChartCond.Visible = False
        chkAfterBarComplete.Visible = True
        Set m.Alert = Alert
        Set Annot = m.Alert.Annotation
        If Annot Is Nothing Then
            cmdCancel.Caption = "&Cancel"
            If Not ActiveChart Is Nothing Then
                Set Chart = ActiveChart.Chart
            End If
        Else
            Set Chart = Annot.AnnotChart
            If Alert.Annotation.AlertAdded Then
                cmdCancel.Caption = "&Remove"
            Else
                cmdCancel.Caption = "&Cancel"
            End If
            If Annot.eType = eANNOT_Trendline Then
                bEdit = True
                If Not Chart Is Nothing Then
                    nIndId = Chart.Tree.Index(Annot.Prop("IndicatorKey"))
                    nPriceField = ValOfText(Annot.Prop("PriceField"))
                    nPaneID = Annot.gePaneId
                    cboIdx = PopulateIndicatorsCbo(cboIndicator, Chart, nIndId, nPaneID, nPriceField, False)
                    If cboIdx >= 0 And cboIdx < cboIndicator.ListCount Then
                        cboIndicator.ListIndex = cboIdx
                        nIndId = cboIndicator.ItemData(cboIdx)
                        If Chart.Tree.NodeLevel(nIndId) > 0 Then Set Ind = Chart.Tree(nIndId)
                    End If
                    m.bCboInitialized = True
                End If
            Else
                lblIndicator.Visible = False
                cboIndicator.Visible = False
                Set Ind = Chart.Tree("PRICE")
                
                If Annot.Prop("FakePriceAlert") = 1 Then
                    bEdit = True
                    txtAlert.Text = Chart.Bars.PriceDisplay(Alert.Value)
                    If Alert.field = "High" Then
                        optSpread(0).Value = True
                    Else
                        optSpread(1).Value = True
                    End If
                    chkAfterBarComplete.Visible = False
                    optSpread(0).Visible = True
                    optSpread(0).Enabled = True
                    optSpread(1).Visible = True
                    optSpread(1).Enabled = True
                    txtAlert.Visible = True
                    txtAlert.Enabled = True
                    optSpread(1).Top = optSpread(0).Top
                    txtAlert.Top = optSpread(0).Top - 30
                End If
            End If
            If Not Ind Is Nothing Then
                If Alert.AlertBarFlag = eGDAlertBar4_HiLowCompleteBar Or Alert.AlertBarFlag = eGDAlertBar5_CloseCompleteBar Then
                    chkAfterBarComplete.Value = vbChecked
                    chkAfterBarComplete.Enabled = True
                Else
                    chkAfterBarComplete.Value = vbUnchecked
                    chkAfterBarComplete.Enabled = Annot.EnableBarComplete(Ind.geIndId)
                End If
            End If
        End If
    ElseIf Alert.AlertType = eGDAlertType_Chart Then
        lblIndicator.Visible = False
        cboIndicator.Visible = False
        Set m.Alert = Alert
        If Alert.Indicator Is Nothing Then
            cmdCancel.Caption = "&Cancel"
        ElseIf Alert.Indicator.AlertAdded Then
            cmdCancel.Caption = "&Remove"
            bEdit = True
        Else
            cmdCancel.Caption = "&Cancel"
        End If
    End If
    
    chkActive.Value = Abs(Alert.Active)
            
    If bEdit Then
        fraConditionChart.Move 120, fraConditionQB.Top, fraConditionChart.Width
    Else
        fraConditionChart.Move 120, fraConditionQB.Top, _
            fraConditionChart.Width, fraConditionChart.Height - cmdEditChartCond.Height - 100
    End If
    
    txtPriceSymbol.Text = Alert.Symbol
    
    fraAction.Move 120, fraConditionChart.Top + fraConditionChart.Height + 50
    chkActive.Move chkActive.Left, fraAction.Top + 50
    fraButtons.Move fraButtons.Left, chkActive.Top + chkActive.Height + 150

    lblChartCondition.Caption = Alert.ChartCondition

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.SetUpChartAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInChartAlert
'' Description: Fill in the annot or indicator alert from the user interface
'' Inputs:      Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInChartAlert(Alert As cAlert)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    With Alert
        .AlertType = m.nAlertType
        .Active = -1 * chkActive.Value
        '.Deactivate = -1 * chkDeactivate.Value         'original code: leave awhile then remove
        If chkKeepActive.Value = vbChecked Then
            .Deactivate = False
        Else
            .Deactivate = True
        End If
        
        For lIndex = 0 To fgActions.Rows - 1
            .ActionString(lIndex) = fgActions.TextMatrix(lIndex, ActionCol(eGDActionCol_Action))
        Next lIndex
    
        If Len(Parse(.ActionString(eAA_PlaySound), ",", 2)) > 0 Then
            SetIniFileProperty "WavFileLastUsed", Parse(.ActionString(eAA_PlaySound), ",", 2), "QuoteBoard", g.strIniFile
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlerts.FillInChartAlert"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EngineVerify
'' Description: Verify the expression with the engine
'' Inputs:      None
'' Returns:     True if verifies through the engine, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EngineVerify(ByVal strCodedText As String) As Boolean
On Error GoTo ErrSection:

    Dim astrParms As New cGdArray       ' Parameters to pass to the engine
    Dim astrBarNames As New cGdArray    ' List of Bar Names to pass to the engine
    Dim aScanExpr As New cGdArray       ' List of expressions to pass to the engine
    Dim strError As String              ' Error message back from the engine
    
    If Len(strCodedText) > 0 Then
        ' Init the expression evaluator with list of scan expressions
        aScanExpr.Add strCodedText
        
        MarketsInExpressions aScanExpr, 0#, False, astrBarNames, Nothing, cboPricePeriod.Text
        
        astrParms(0) = "AlertVerify"
        If SetupExpressions(astrParms, astrBarNames, aScanExpr, strError) Then
            EngineVerify = True
        Else
            InfBox "An error occured with the engine verification:||" & strError & "|", , , "Engine Verification Error", , , , , , , , eGDAlign_Left
        End If
        
        ' Clear the expression evaluator when done with it
        SetupExpressions astrParms
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlerts.EngineVerify"
    
End Function

Private Property Get AccountForOrder() As Long
On Error GoTo ErrSection:

    Dim nID As Long
    
    If m.nAlertType = eGDAlertType_Annot Or m.nAlertType = eGDAlertType_Chart Then
        If Not m.Alert Is Nothing Then
            If Not m.Alert.Annotation Is Nothing Then
                If Not m.Alert.Annotation.AnnotChart Is Nothing Then
                    nID = m.Alert.Annotation.AnnotChart.TradeAccountID
                End If
            ElseIf Not m.Alert.Indicator Is Nothing Then
                If Not m.Alert.Indicator.IndChart Is Nothing Then
                    nID = m.Alert.Indicator.IndChart.TradeAccountID
                End If
            End If
        End If
    End If
    
    If nID = 0 Then nID = DefaultAccount
    
    AccountForOrder = nID
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmAlerts.AccountForOrder"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoDetect
'' Description: Auto detect the number of bars required for a TradeSense alert
'' Inputs:      Show Message?
'' Returns:     True if auto detected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoDetect(ByVal bShowMessage As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim AD As New cAutoDetect           ' Auto detect object
    Dim lNumBars As Long                ' Number of bars required
    Dim strProgramName As String
    
    If Len(m.strCodedText) > 0 Then
        ' DAJ 11/26/2012: The symbols combo box isn't used for TradeSense alerts, the
        ' text box is.  Ran into a situation trying to verify an expression on 2000 trade
        ' bars that couldn't auto detect with the default symbols...
        'lNumBars = AD.AutoDetect(m.strCodedText, cboSymbols.Text, cboPricePeriod.Text)
        lNumBars = AD.AutoDetect(m.strCodedText, txtPriceSymbol.Text, cboPricePeriod.Text)
        txtNumBars.Text = Str(lNumBars)

        If ExtremeCharts >= 1 Then
            strProgramName = "Extreme Charts "
        Else
            strProgramName = "Trade Navigator "
        End If
    
        If (ValOfText(txtOverride.Text) < lNumBars) And (optOverride.Value = True) Then
            If bShowMessage = True Then
                InfBox strProgramName & "has determined that your alert needs at least " & Trim(CStr(lNumBars)) & " bars to run properly.||The value has been set accordingly", "i", , "Alert"
            End If
            optAutoDetect = True
            txtOverride.Text = Str(lNumBars)
        End If
        
        If (lNumBars = -1&) And ((optAutoDetect.Value = True) Or (ValOfText(txtOverride.Text) <= 0)) Then
            If bShowMessage = True Then
                InfBox strProgramName & "could not automatically|determine how many bars are needed to calculate the alert.  Please specify an|override for the number of necessary bars.", "!", , "Alert Error"
            End If
            optOverride = True
            MoveFocus txtOverride
        End If
    End If
    
    AutoDetect = (lNumBars > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlerts.AutoDetect"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsCombo
'' Description: Load the accounts combo box with all the valid accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadAccountsCombo()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Accounts As cPtAccounts         ' Collection of all accounts the user can see
    Dim Account As cPtAccount           ' Account from the collection

    cboAccounts.Clear
    
    cboAccounts.AddItem "All Accounts"
    cboAccounts.ItemData(cboAccounts.NewIndex) = 0
    
    cboAccounts.AddItem "All Live Accounts"
    cboAccounts.ItemData(cboAccounts.NewIndex) = -1
    
    cboAccounts.AddItem "All Simulated Accounts"
    cboAccounts.ItemData(cboAccounts.NewIndex) = -2
    
    Set Accounts = g.Broker.AllAccounts
    For lIndex = 1 To Accounts.Count
        Set Account = Accounts(lIndex)
        cboAccounts.AddItem Account.Name
        cboAccounts.ItemData(cboAccounts.NewIndex) = Account.AccountID
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlerts.LoadAccountsCombo"
    
End Sub

