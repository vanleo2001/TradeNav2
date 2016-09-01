VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTimeSalesCfg 
   Caption         =   "Time & Sales Settings"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraColors 
      Height          =   3300
      Left            =   4410
      TabIndex        =   17
      Top             =   75
      Width           =   4155
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTimeSalesCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSalesCfg.frx":0068
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":0088
      RightToLeft     =   0   'False
      Begin gdOCX.gdSelectColor gdUpColor 
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Color           =   49152
         CustomColor     =   49152
      End
      Begin gdOCX.gdSelectColor gdDownColor 
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   858
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdUpColorBid 
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1356
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Color           =   49152
         CustomColor     =   49152
      End
      Begin gdOCX.gdSelectColor gdDownColorBid 
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   1854
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdUpColorAsk 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   2352
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Color           =   49152
         CustomColor     =   49152
      End
      Begin gdOCX.gdSelectColor gdDownColorAsk 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   2850
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP lblDownColor 
         Height          =   195
         Left            =   1200
         Top             =   918
         Width           =   2520
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":00A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":010A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":012A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUpColor 
         Height          =   195
         Left            =   1200
         Top             =   420
         Width           =   2580
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":0146
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":01A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":01C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUpColorBid 
         Height          =   195
         Left            =   1200
         Top             =   1416
         Width           =   2580
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":01E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0242
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0262
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDownColorBid 
         Height          =   195
         Left            =   1200
         Top             =   1914
         Width           =   2520
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":027E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":02E0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0300
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUpColorAsk 
         Height          =   195
         Left            =   1200
         Top             =   2412
         Width           =   2580
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":031C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":037A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":039A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDownColorAsk 
         Height          =   195
         Left            =   1200
         Top             =   2910
         Width           =   2520
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":03B6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0418
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0438
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSession 
      Height          =   945
      Left            =   135
      TabIndex        =   13
      Top             =   2430
      Width           =   4155
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTimeSalesCfg.frx":0454
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSalesCfg.frx":047C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":049C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optDate 
         Height          =   220
         Left            =   180
         TabIndex        =   15
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
         Caption         =   "frmTimeSalesCfg.frx":04B8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":04EE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":050E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCurrentSession 
         Height          =   220
         Left            =   180
         TabIndex        =   14
         Top             =   300
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
         Caption         =   "frmTimeSalesCfg.frx":052A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0570
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0590
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdSessionDate 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
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
   End
   Begin HexUniControls.ctlUniFrameWL fraVolFilter 
      Height          =   975
      Left            =   135
      TabIndex        =   8
      Top             =   3465
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
      Caption         =   "frmTimeSalesCfg.frx":05AC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSalesCfg.frx":05FC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":061C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtShowVolMax 
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   600
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTimeSalesCfg.frx":0638
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
         Tip             =   "frmTimeSalesCfg.frx":0658
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0678
      End
      Begin HexUniControls.ctlUniTextBoxXP txtShowVolMin 
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         Top             =   270
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmTimeSalesCfg.frx":0694
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
         Tip             =   "frmTimeSalesCfg.frx":06B4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":06D4
      End
      Begin HexUniControls.ctlUniLabelXP lblShowVolMax 
         Height          =   255
         Left            =   120
         Top             =   630
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
         Caption         =   "frmTimeSalesCfg.frx":06F0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0776
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0796
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblShowVolMin 
         Height          =   255
         Left            =   120
         Top             =   300
         Width           =   4020
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":07B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":083E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":085E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Height          =   375
      Left            =   5745
      TabIndex        =   5
      Top             =   3765
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
      Caption         =   "frmTimeSalesCfg.frx":087A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmTimeSalesCfg.frx":089E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":08BE
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   375
      Left            =   6885
      TabIndex        =   4
      Top             =   3765
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
      Caption         =   "frmTimeSalesCfg.frx":08DA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmTimeSalesCfg.frx":0906
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":0926
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdFont 
      Height          =   315
      Left            =   3210
      TabIndex        =   3
      Top             =   75
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
      Caption         =   "frmTimeSalesCfg.frx":0942
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmTimeSalesCfg.frx":096C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":098C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraDisplayStyle 
      Height          =   1935
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   4155
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmTimeSalesCfg.frx":09A8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTimeSalesCfg.frx":09E2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTimeSalesCfg.frx":0A02
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optDisplayStyle 
         Height          =   220
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1605
         Width           =   3795
         _ExtentX        =   6694
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
         Caption         =   "frmTimeSalesCfg.frx":0A1E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0A8C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0AAC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkIncludeBidAsk 
         Height          =   220
         Left            =   480
         TabIndex        =   6
         Top             =   480
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
         Caption         =   "frmTimeSalesCfg.frx":0AC8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0B14
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0B34
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDisplayStyle 
         Height          =   220
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1260
         Width           =   3795
         _ExtentX        =   6694
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
         Caption         =   "frmTimeSalesCfg.frx":0B50
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0BA8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0BC8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDisplayStyle 
         Height          =   220
         Index           =   0
         Left            =   180
         TabIndex        =   1
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
         Caption         =   "frmTimeSalesCfg.frx":0BE4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0C34
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0C54
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBidAskNote 
         Height          =   495
         Left            =   780
         Top             =   720
         Width           =   3075
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmTimeSalesCfg.frx":0C70
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmTimeSalesCfg.frx":0D38
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTimeSalesCfg.frx":0D58
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7755
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTimeSalesCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    frmTSales As frmTimeSales
    nDisplayStyle As Long
End Type

Private m As mPrivate

Public Sub ShowMe(frmCaller As frmTimeSales)
On Error GoTo ErrSection:

    Dim nVol As Long
    Dim bCanFilterVol As Boolean
    
    Set m.frmTSales = frmCaller
    
    If m.frmTSales Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
'    eStyle_None = 0
'    eStyle_TickByTick       '3 columns, Time-Price-Size (no wrap)
'    eStyle_MinByMin         '2 columns, Tim-Price (wrap price)
'    eStyle_TickBidAsk       '4 columns, Time-Price-Type-Size (no wrap)
'    eStyle_Cumulative       '8 columns, Time-Price-Trades-Contracts-AvgTradeSize-LargestTrade-BuyVol-SellVol
    'display style
    m.nDisplayStyle = m.frmTSales.TS_DisplayStyle
    
    Select Case m.nDisplayStyle
        Case 1, 3
            optDisplayStyle(0) = True
            If m.nDisplayStyle = 3 Then
                chkIncludeBidAsk.Value = 1
            Else
                chkIncludeBidAsk.Value = 0
            End If
        Case 2
            optDisplayStyle(1) = True
        Case 4
            optDisplayStyle(2) = True
        Case Else
            optDisplayStyle(0) = True
            chkIncludeBidAsk.Value = 0
    End Select
    
    'colors
    gdUpColor.Color = m.frmTSales.TS_UpColor
    gdDownColor.Color = m.frmTSales.TS_DownColor
    gdUpColorBid.Color = m.frmTSales.TS_UpColorBid
    gdDownColorBid.Color = m.frmTSales.TS_DownColorBid
    gdUpColorAsk.Color = m.frmTSales.TS_UpColorAsk
    gdDownColorAsk.Color = m.frmTSales.TS_DownColorAsk
    'session date
    optCurrentSession.Value = m.frmTSales.TS_SessionCurrent
    optDate.Value = Not m.frmTSales.TS_SessionCurrent
    If m.frmTSales.TS_SessionDate > 0 Then
        gdSessionDate = m.frmTSales.TS_SessionDate
    Else
        gdSessionDate = Date
    End If
    
    bCanFilterVol = m.frmTSales.TS_CanHaveVolFilter
    
    txtShowVolMin.Text = ""
    txtShowVolMax.Text = ""
    
    If bCanFilterVol Then
        nVol = m.frmTSales.TS_VolFilterMin
        If nVol > 0 Then txtShowVolMin.Text = Str(nVol)
    
        nVol = m.frmTSales.TS_VolFilterMax
        If nVol > 0 Then txtShowVolMax.Text = Str(nVol)
    End If
    
    fraVolFilter.Enabled = bCanFilterVol
    lblShowVolMin.Enabled = bCanFilterVol
    lblShowVolMax.Enabled = bCanFilterVol
    txtShowVolMin.Enabled = bCanFilterVol
    txtShowVolMax.Enabled = bCanFilterVol
    
    EnableBidAskControls
        
    CenterTheForm Me
    ShowForm Me, eForm_Modal
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesCfg.ShowMe", eGDRaiseError_Raise

End Sub

Private Sub chkIncludeBidAsk_Click()
On Error GoTo ErrSection:

    If chkIncludeBidAsk.Value = 1 Then
        m.nDisplayStyle = 3
    Else
        m.nDisplayStyle = 1
    End If
    
    EnableBidAskControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesCfg.chkIncludeBidAsk.Click", eGDRaiseError_Raise

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next:

    Unload Me

End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrSection

    'set font currently in use
    Me.Font.Name = m.frmTSales.TS_FontName
    Me.Font.Size = m.frmTSales.TS_FontSize
    Me.FontItalic = m.frmTSales.TS_FontItalic
    Me.Font.Bold = m.frmTSales.TS_FontBold
    If CommonDialogFont(CommonDialog1, Me.Font) Then
        m.frmTSales.TS_FontName = Me.Font.Name
        m.frmTSales.TS_FontSize = Me.Font.Size
        m.frmTSales.TS_FontItalic = Me.FontItalic
        m.frmTSales.TS_FontBold = Me.Font.Bold
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesCfg.cmdFont.Click", eGDRaiseError_Raise

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim nVolMax&, nVolMin&
    
    nVolMax = ValOfText(txtShowVolMax.Text)
    nVolMin = ValOfText(txtShowVolMin.Text)
    
    If nVolMax > 0 And nVolMin > 0 Then
        If nVolMax < nVolMin Then
            InfBox "Invalid volume filter criteria.", "I"
            Exit Sub
        End If
    End If
    
    'display style
    m.frmTSales.TS_DisplayStyle = m.nDisplayStyle
    'colors
    m.frmTSales.TS_UpColor = gdUpColor.Color
    m.frmTSales.TS_DownColor = gdDownColor.Color
    m.frmTSales.TS_UpColorBid = gdUpColorBid.Color
    m.frmTSales.TS_DownColorBid = gdDownColorBid.Color
    m.frmTSales.TS_UpColorAsk = gdUpColorAsk.Color
    m.frmTSales.TS_DownColorAsk = gdDownColorAsk.Color
    
    'settings relating to session data has changed, eg date, volume filter
    If m.frmTSales.TS_SessionCurrent <> optCurrentSession.Value Then
        m.frmTSales.TS_SessionChanged = True
        m.frmTSales.TS_SessionCurrent = optCurrentSession.Value
    ElseIf m.frmTSales.TS_SessionCurrent = False Then
        If m.frmTSales.TS_SessionDate <> gdSessionDate Then
            m.frmTSales.TS_SessionChanged = True
        End If
    End If
    m.frmTSales.TS_SessionDate = gdSessionDate
    
    m.frmTSales.TS_VolFilterMax = ValOfText(txtShowVolMax.Text)
    m.frmTSales.TS_VolFilterMin = ValOfText(txtShowVolMin.Text)
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesCfg.cmdOK.Click", eGDRaiseError_Raise

End Sub

Private Sub Form_Load()
    Me.Icon = Picture16(ToolbarIcon("ID_TimeSales"), , True)
    
    g.Styler.StyleForm Me
End Sub

Private Sub optCurrentSession_Click()
    gdSessionDate.Enabled = False
End Sub

Private Sub optDate_Click()
    gdSessionDate.Enabled = True
End Sub

Private Sub EnableBidAskControls()
On Error GoTo ErrSection:

    Dim bEnableColors As Boolean
    Dim bEnableBidAsk As Boolean
    
    If m.nDisplayStyle = 1 Then
        bEnableColors = True
    ElseIf m.nDisplayStyle = 3 Then
        bEnableColors = True
        bEnableBidAsk = True
    End If
    
    chkIncludeBidAsk.Enabled = bEnableColors
    lblBidAskNote.Enabled = bEnableColors
    
    gdUpColor.Enabled = bEnableColors
    gdDownColor.Enabled = bEnableColors
    lblUpColor.Enabled = bEnableColors
    lblDownColor.Enabled = bEnableColors
    
    gdUpColorBid.Enabled = bEnableBidAsk
    gdDownColorBid.Enabled = bEnableBidAsk
    lblUpColorBid.Enabled = bEnableBidAsk
    lblDownColorBid.Enabled = bEnableBidAsk
    
    gdUpColorAsk.Enabled = bEnableBidAsk
    gdDownColorAsk.Enabled = bEnableBidAsk
    lblUpColorAsk.Enabled = bEnableBidAsk
    lblDownColorAsk.Enabled = bEnableBidAsk
    
    If bEnableColors Or bEnableBidAsk Then
        fraColors.Enabled = True
    Else
        fraColors.Enabled = False
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesCfg.EnableBidAskControls", eGDRaiseError_Raise

End Sub

Private Sub optDisplayStyle_Click(Index As Integer)
On Error GoTo ErrSection:

    If Not Me.Visible Then GoTo ErrExit
    
    'display style
    If optDisplayStyle(0).Value = True Then
        If chkIncludeBidAsk.Value = 0 Then
            m.nDisplayStyle = 1
        Else
            m.nDisplayStyle = 3
        End If
    ElseIf optDisplayStyle(1).Value = True Then
        m.nDisplayStyle = 2
    ElseIf optDisplayStyle(2).Value = True Then
        m.nDisplayStyle = 4
    End If
    
    EnableBidAskControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTimeSalesCfg.optDisplayStyle.Click", eGDRaiseError_Raise

End Sub

