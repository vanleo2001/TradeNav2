VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmBarPeriod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bar Time Period"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraPF 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
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
      Caption         =   "frmBarPeriod.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBarPeriod.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPF 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   0
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmBarPeriod.frx":005C
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
         Tip             =   "frmBarPeriod.frx":007E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":009E
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   180
         Top             =   30
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
         Caption         =   "frmBarPeriod.frx":00BA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBarPeriod.frx":00EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":010C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   2280
         Top             =   30
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmBarPeriod.frx":0128
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBarPeriod.frx":0152
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0172
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPoints 
      Height          =   420
      Left            =   2400
      TabIndex        =   24
      Top             =   3525
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
      Caption         =   "frmBarPeriod.frx":018E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBarPeriod.frx":01AE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":01CE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optAsPoints 
         Height          =   220
         Left            =   0
         TabIndex        =   25
         Top             =   0
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
         Caption         =   "frmBarPeriod.frx":01EA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0216
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0236
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAsTicks 
         Height          =   220
         Left            =   0
         TabIndex        =   26
         Top             =   210
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
         Caption         =   "frmBarPeriod.frx":0252
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmBarPeriod.frx":027E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":029E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   4440
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
      Caption         =   "frmBarPeriod.frx":02BA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmBarPeriod.frx":02E8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":0308
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4440
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
      Caption         =   "frmBarPeriod.frx":0324
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmBarPeriod.frx":034A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":036A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtNumPeriods 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   3525
      Width           =   795
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmBarPeriod.frx":0386
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
      Tip             =   "frmBarPeriod.frx":03A8
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":03C8
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   3375
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   3315
      _ExtentX        =   5847
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
      Caption         =   "frmBarPeriod.frx":03E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmBarPeriod.frx":0410
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":0430
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optIntBreakout 
         Height          =   220
         Left            =   1680
         TabIndex        =   19
         Top             =   1647
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
         Caption         =   "frmBarPeriod.frx":044C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":048E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":04AE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optBreakout 
         Height          =   225
         Left            =   180
         TabIndex        =   12
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "frmBarPeriod.frx":04CA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":04FC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":051C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optIntRenko 
         Height          =   220
         Left            =   1680
         TabIndex        =   22
         Top             =   2433
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
         Caption         =   "frmBarPeriod.frx":0538
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0562
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0582
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optIntKagi 
         Height          =   220
         Left            =   1680
         TabIndex        =   21
         Top             =   2171
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
         Caption         =   "frmBarPeriod.frx":059E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":05C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":05E6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optIntPF 
         Height          =   220
         Left            =   1680
         TabIndex        =   20
         Top             =   1909
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
         Caption         =   "frmBarPeriod.frx":0602
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0640
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0660
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optRenko 
         Height          =   225
         Left            =   180
         TabIndex        =   15
         Top             =   2719
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
         Caption         =   "frmBarPeriod.frx":067C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":06A6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":06C6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optKagi 
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   2446
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
         Caption         =   "frmBarPeriod.frx":06E2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":070A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":072A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPF 
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Top             =   2173
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
         Caption         =   "frmBarPeriod.frx":0746
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0784
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":07A4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optEachTick 
         Height          =   220
         Left            =   1680
         TabIndex        =   23
         Top             =   2695
         Visible         =   0   'False
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
         Caption         =   "frmBarPeriod.frx":07C0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0802
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0822
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optIntVolume 
         Height          =   220
         Left            =   1680
         TabIndex        =   17
         Top             =   1118
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
         Caption         =   "frmBarPeriod.frx":083E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":087E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":089E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTicks 
         Height          =   225
         Left            =   1680
         TabIndex        =   18
         Top             =   1380
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
         Caption         =   "frmBarPeriod.frx":08BA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":08F8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0918
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optMinutes 
         Height          =   225
         Left            =   1680
         TabIndex        =   16
         Top             =   720
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
         Caption         =   "frmBarPeriod.frx":0934
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0964
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0984
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optVolume 
         Height          =   225
         Left            =   180
         TabIndex        =   11
         Top             =   1900
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
         Caption         =   "frmBarPeriod.frx":09A0
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":09CE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":09EE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optYears 
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   1627
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
         Caption         =   "frmBarPeriod.frx":0A0A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0A38
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0A58
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optQuarters 
         Height          =   225
         Left            =   180
         TabIndex        =   9
         Top             =   1354
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
         Caption         =   "frmBarPeriod.frx":0A74
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0AA8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0AC8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optMonths 
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   1081
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
         Caption         =   "frmBarPeriod.frx":0AE4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0B14
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0B34
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optWeeks 
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   808
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
         Caption         =   "frmBarPeriod.frx":0B50
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmBarPeriod.frx":0B7E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0B9E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDays 
         Height          =   220
         Left            =   180
         TabIndex        =   6
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
         Caption         =   "frmBarPeriod.frx":0BBA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmBarPeriod.frx":0BE6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0C06
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblIntraday2 
         Height          =   195
         Left            =   1740
         Top             =   450
         Width           =   1395
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmBarPeriod.frx":0C22
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBarPeriod.frx":0C6A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0C8A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblIntraday 
         Height          =   195
         Left            =   1725
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
         Caption         =   "frmBarPeriod.frx":0CA6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBarPeriod.frx":0CEE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0D0E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDaily 
         Height          =   195
         Left            =   180
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
         Caption         =   "frmBarPeriod.frx":0D2A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmBarPeriod.frx":0D66
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmBarPeriod.frx":0D86
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblUnits 
      Height          =   255
      Left            =   2340
      Top             =   3525
      Visible         =   0   'False
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
      Caption         =   "frmBarPeriod.frx":0DA2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBarPeriod.frx":0DD2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":0DF2
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPeriods 
      Height          =   255
      Left            =   120
      Top             =   3525
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
      Caption         =   "frmBarPeriod.frx":0E0E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBarPeriod.frx":0E4A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBarPeriod.frx":0E6A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmBarPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    lPeriodicity As Long
    bAsPoints As Boolean
    dTickMove As Double
    Bars As cGdBars
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.lPeriodicity = 0
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim dPeriods#, dRev#
    
    MoveFocus cmdOK
    DoEvents
    
    If optIntBreakout Then
        If Not HasModule("ROCK*", True) Then
            If Not HasLevel(eTN3_Standard, True, "Breakout Bars") Then
                Exit Sub
            End If
        End If
    End If
    
    SetIniFileProperty "AsPoints", optAsPoints.Value, "BarPeriod", g.strIniFile
    
    If optEachTick Then
        dPeriods = 1
    Else
        If optAsPoints.Visible And m.bAsPoints And m.dTickMove > 0 Then
            dPeriods = m.Bars.PriceFromString(txtNumPeriods)
            dPeriods = Int(dPeriods / m.dTickMove + 0.5)
        Else
            dPeriods = ValOfText(txtNumPeriods)
        End If
        
        If m.lPeriodicity = ePRD_EodPF Or m.lPeriodicity = ePRD_IntPF Then
            dRev = ValOfText(txtPF)
            If dRev < 0 Or dRev > 166 Or dRev <> Int(dRev) Then
                Beep
                MoveFocus txtPF
                InfBox "Reversal must be an integer from 0 to 166", "[]", , "Error"
                Exit Sub
            ElseIf dPeriods < 1 Or dPeriods >= 100000 Then
                Beep
                MoveFocus txtNumPeriods
                Exit Sub
            End If
            dPeriods = CLng(dPeriods) + dRev * 100000
        End If
    End If
    
    If dPeriods <= 0 Or dPeriods >= ePRD_EachTick Then
        ' invalid
        Beep
        MoveFocus txtNumPeriods
    Else
        m.lPeriodicity = m.lPeriodicity + CLng(dPeriods)
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.cmdOK.Click", eGDRaiseError_Show
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
    RaiseError "frmBarPeriod.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    Me.Icon = Picture16(ToolbarIcon("kSelect"))

    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.lPeriodicity = 0
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optAsPoints_Click()

    If Me.Visible And m.dTickMove > 0 Then
        m.bAsPoints = True
        txtNumPeriods = m.Bars.PriceDisplay(ValOfText(txtNumPeriods) * m.dTickMove, True)
    End If

End Sub

Private Sub optAsTicks_Click()
    
    If Me.Visible And m.dTickMove > 0 Then
        m.bAsPoints = False
        txtNumPeriods = CStr(Int(m.Bars.PriceFromString(txtNumPeriods) / m.dTickMove + 0.5))
    End If

End Sub

Private Sub optBreakout_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optBreakout.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optDays_Click()
On Error GoTo ErrSection:

    If Me.Visible Then txtNumPeriods = "1"

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optDays.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optEachTick_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optEachTick.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optIntBreakout_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optIntBreakout.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optIntKagi_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optIntKagi.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optIntPF_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optIntPF.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optIntRenko_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optIntRenko.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optIntVolume_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optIntVolume.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optKagi_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optKagi.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optMinutes_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optMinutes.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optMonths_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optMonths.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optPF_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optPF.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optQuarters_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optQuarters.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optRenko_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optRenko.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optTicks_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optTicks.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optVolume_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optVolume.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optWeeks_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optWeeks.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub optYears_Click()
On Error GoTo ErrSection:

    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.optYears.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtNumPeriods_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtNumPeriods

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.txtNumPeriods.GotFocus", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub FixControls()
On Error GoTo ErrSection:
    
    Dim strText$, eType As eBarsPeriodType
        
    If optWeeks Then
        eType = ePRD_Weeks
        strText = "Weeks"
    ElseIf optMonths Then
        eType = ePRD_Months
        strText = "Months"
    ElseIf optQuarters Then
        eType = ePRD_Quarters
        strText = "Quarters"
    ElseIf optYears Then
        eType = ePRD_Years
        strText = "Years"
    ElseIf optMinutes Then
        eType = ePRD_Minutes
        strText = "Minutes"
    ElseIf optTicks Then
        eType = ePRD_Ticks
        strText = "Trades"
    ElseIf optVolume Then
        eType = ePRD_EodVol
        strText = "Volume"
    ElseIf optEachTick Then
        eType = ePRD_EachTick
    ElseIf optIntVolume Then
        eType = ePRD_IntVol
        strText = "Volume"
    ElseIf optPF Then
        eType = ePRD_EodPF
        lblPeriods = "Box si&ze:"
    ElseIf optKagi Then
        eType = ePRD_EodKagi
        lblPeriods = "Reversal:"
    ElseIf optRenko Then
        eType = ePRD_EodRenko
        lblPeriods = "Box si&ze:"
    ElseIf optIntPF Then
        eType = ePRD_IntPF
        lblPeriods = "Box si&ze:"
    ElseIf optIntKagi Then
        eType = ePRD_IntKagi
        lblPeriods = "Reversal:"
    ElseIf optIntRenko Then
        eType = ePRD_IntRenko
        lblPeriods = "Box si&ze:"
    ElseIf optBreakout Then
        eType = ePRD_EodBreakout
        lblPeriods = "Price Range:"
    ElseIf optIntBreakout Then
        eType = ePRD_IntBreakout
        lblPeriods = "Price Range:"
    Else
        eType = ePRD_Days
        strText = "Days"
    End If
    
    If Len(strText) > 0 Then
        lblPeriods = strText & " per &bar:"
        fraPoints.Visible = False
    Else
        fraPoints.Visible = True
    End If
    
    lblUnits.Visible = optVolume.Value
    If optEachTick Then
        txtNumPeriods.Visible = False
        lblPeriods.Visible = False
    ElseIf Not txtNumPeriods.Visible Then
        txtNumPeriods.Visible = True
        lblPeriods.Visible = True
    End If
    
' not ready yet
optEachTick.Enabled = False
    
    m.lPeriodicity = eType
    
    If m.lPeriodicity = ePRD_EodPF Or m.lPeriodicity = ePRD_IntPF Then
        fraPF.Visible = True
        cmdOK.Top = fraPF.Top + fraPF.Height + 120
    Else
        fraPF.Visible = False
        cmdOK.Top = fraPF.Top + 120
    End If
    cmdCancel.Top = cmdOK.Top
    Me.Height = cmdOK.Top + cmdOK.Height + Me.Height - Me.ScaleHeight + 120
    
    MoveFocus txtNumPeriods
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBarPeriod.FixControls", eGDRaiseError_Raise

End Sub

Public Function ShowMe(ByVal nPeriodicity&, Bars As cGdBars, _
    Optional ByRef Chart As cChart = Nothing) As Long
On Error GoTo ErrSection:

    Dim nPeriods&, nBoxes&

    m.bAsPoints = GetIniFileProperty("AsPoints", True, "BarPeriod", g.strIniFile)
    If m.bAsPoints Then
        optAsPoints = True
    Else
        optAsTicks = True
    End If

    Set m.Bars = Bars
    If Not Bars Is Nothing Then
        m.dTickMove = Bars.Prop(eBARS_TickMove)
    Else
        m.dTickMove = 0
    End If

    If m.dTickMove = 0 Then
        optPF.Enabled = False
        optKagi.Enabled = False
        optRenko.Enabled = False
        optBreakout.Enabled = False
        optIntPF.Enabled = False
        optIntKagi.Enabled = False
        optIntRenko.Enabled = False
        optIntBreakout.Enabled = False
    Else
        optPF.Enabled = True
        optKagi.Enabled = True
        optRenko.Enabled = True
        optBreakout.Enabled = True
        optIntPF.Enabled = True
        optIntKagi.Enabled = True
        optIntRenko.Enabled = True
        optIntBreakout.Enabled = True
    End If

    m.lPeriodicity = GetPeriodType(nPeriodicity)
    nPeriods = GetPeriodsPerBar(nPeriodicity)
    Select Case m.lPeriodicity
    Case ePRD_EodPF, ePRD_EodKagi, ePRD_EodRenko, ePRD_EodBreakout, ePRD_IntPF, ePRD_IntKagi, ePRD_IntRenko, ePRD_IntBreakout
        If m.lPeriodicity = ePRD_EodPF Or m.lPeriodicity = ePRD_IntPF Then
            nBoxes = nPeriods \ 100000
            If nBoxes < 0 Or nBoxes > 166 Then
                nBoxes = 3
            End If
            If nBoxes = 0 Then
                txtPF = "" '"Wyckoff"
            Else
                txtPF = CStr(nBoxes)
            End If
            nPeriods = nPeriods Mod 100000
        End If
        If m.bAsPoints And m.dTickMove > 0 Then
            txtNumPeriods = m.Bars.PriceDisplay(nPeriods * m.dTickMove, True)
        Else
            txtNumPeriods = CStr(nPeriods)
        End If
    Case Else
        txtNumPeriods = CStr(nPeriods)
    End Select
    
    Select Case m.lPeriodicity
    ' Intraday types
    Case ePRD_EachTick
        optEachTick = True
    Case ePRD_Ticks
        optTicks = True
    Case ePRD_Minutes
        optMinutes = True
    Case ePRD_IntVol
        optIntVolume = True
    Case ePRD_IntPF
        optIntPF = True
    Case ePRD_IntKagi
        optIntKagi = True
    Case ePRD_IntRenko
        optIntRenko = True
    Case ePRD_IntBreakout
        optIntBreakout = True
    ' End-of-day types
    Case ePRD_Days
        optDays = True
    Case ePRD_Weeks
        optWeeks = True
    Case ePRD_Months
        optMonths = True
    Case ePRD_Quarters
        optQuarters = True
    Case ePRD_Years
        optYears = True
    Case ePRD_EodVol
        optVolume = True
    Case ePRD_EodPF
        optPF = True
    Case ePRD_EodKagi
        optKagi = True
    Case ePRD_EodRenko
        optRenko = True
    Case ePRD_EodBreakout
        optBreakout = True
    Case Else
        optDays = True
    End Select
    
    ' hide intraday for BetterTrades (unless they have intraday data)
    If ExtremeCharts = 1 Then
        If Not HasModule("IT") And Not IsIntraday(nPeriodicity) Then
            optEachTick.Visible = False
            optTicks.Visible = False
            optMinutes.Visible = False
            optIntVolume.Visible = False
            optIntPF.Visible = False
            optIntKagi.Visible = False
            optIntRenko.Visible = False
            optIntBreakout.Visible = False
            lblIntraday.Visible = False
            lblIntraday2.Visible = False
        End If
    End If
    
    FixControls
    If Not Chart Is Nothing Then CenterFormOnChart Me, Chart        '6499
    ShowForm Me, True
    Set m.Bars = Nothing
    
    If m.lPeriodicity = 0 Then
        'cancelled
        'ShowMe = nPeriodicity
        ShowMe = 0
    Else
        ShowMe = m.lPeriodicity
    End If
    Unload Me
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBarPeriod.ShowMe", eGDRaiseError_Raise

End Function

Private Sub txtPF_GotFocus()
    SelectAll txtPF
End Sub

