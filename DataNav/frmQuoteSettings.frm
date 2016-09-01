VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmQuoteSettings 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraColors 
      Height          =   2295
      Left            =   60
      TabIndex        =   5
      Top             =   1620
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
      Caption         =   "frmQuoteSettings.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkUpdatedColor 
         Height          =   220
         Left            =   105
         TabIndex        =   2
         Top             =   1005
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
         Caption         =   "frmQuoteSettings.frx":0068
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":00C4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":00E4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkColorSymbol 
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   1440
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":0100
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0162
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0182
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdUnchColor 
         Height          =   315
         Left            =   2445
         TabIndex        =   9
         Top             =   510
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdDownColor 
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   510
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdUpdateColor 
         Height          =   315
         Left            =   2445
         TabIndex        =   10
         Top             =   930
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdUpColor 
         Height          =   315
         Left            =   105
         TabIndex        =   6
         Top             =   510
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   195
         Left            =   105
         Top             =   1995
         Width           =   3000
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":019E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":020E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":022E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUpColor 
         Height          =   255
         Left            =   143
         Top             =   270
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
         Caption         =   "frmQuoteSettings.frx":024A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":027C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":029C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   225
         Left            =   375
         Top             =   1725
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "frmQuoteSettings.frx":02B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":0328
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0348
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUnchColor 
         Height          =   255
         Left            =   2430
         Top             =   270
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
         Caption         =   "frmQuoteSettings.frx":0364
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":0398
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":03B8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDownColor 
         Height          =   255
         Left            =   1200
         Top             =   270
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
         Caption         =   "frmQuoteSettings.frx":03D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":040A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":042A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraBarStyle 
      Height          =   1635
      Left            =   3660
      TabIndex        =   11
      Top             =   1620
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
      Caption         =   "frmQuoteSettings.frx":0446
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":0496
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":04B6
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optOHLC 
         Height          =   220
         Left            =   180
         TabIndex        =   12
         Top             =   360
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
         Caption         =   "frmQuoteSettings.frx":04D2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmQuoteSettings.frx":04FA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":051A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCandlestick 
         Height          =   220
         Left            =   180
         TabIndex        =   13
         Top             =   600
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
         Caption         =   "frmQuoteSettings.frx":0536
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":056C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":058C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optBollinger 
         Height          =   220
         Left            =   180
         TabIndex        =   27
         Top             =   840
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
         Caption         =   "frmQuoteSettings.frx":05A8
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":05E2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0602
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optNone 
         Height          =   220
         Left            =   180
         TabIndex        =   15
         Top             =   1320
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
         Caption         =   "frmQuoteSettings.frx":061E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0646
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0666
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optThermometer 
         Height          =   220
         Left            =   180
         TabIndex        =   14
         Top             =   1080
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
         Caption         =   "frmQuoteSettings.frx":0682
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":06B8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":06D8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSettings 
      Height          =   2520
      Left            =   60
      TabIndex        =   19
      Top             =   4020
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4445
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmQuoteSettings.frx":06F4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":0730
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":0750
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkBidAsk 
         Height          =   260
         Left            =   180
         TabIndex        =   8
         Top             =   1900
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":076C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":07C0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":07E0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkShowButtons 
         Height          =   260
         Left            =   180
         TabIndex        =   21
         Top             =   1350
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":07FC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0834
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0854
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkCompactQB 
         Height          =   260
         Left            =   180
         TabIndex        =   26
         Top             =   2175
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":0870
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":08BE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":08DE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkDateTime 
         Height          =   260
         Left            =   180
         TabIndex        =   28
         Top             =   1625
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":08FA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0952
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0972
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCategories 
         Height          =   435
         Left            =   180
         TabIndex        =   25
         Top             =   840
         Width           =   1215
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
         Caption         =   "frmQuoteSettings.frx":098E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":09BE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":09DE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAutoUpdate 
         Height          =   435
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   1215
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
         Caption         =   "frmQuoteSettings.frx":09FA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0A34
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0A54
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   495
         Left            =   1620
         Top             =   810
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
         Caption         =   "frmQuoteSettings.frx":0A70
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":0B12
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0B32
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAutoUpdate 
         Height          =   495
         Left            =   1620
         Top             =   300
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
         Caption         =   "frmQuoteSettings.frx":0B4E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":0C20
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0C40
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraBarHeight 
      Height          =   555
      Left            =   3660
      TabIndex        =   16
      Top             =   3360
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
      Caption         =   "frmQuoteSettings.frx":0C5C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":0C90
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":0CB0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optRelativeHeight 
         Height          =   220
         Left            =   180
         TabIndex        =   18
         Top             =   240
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
         Caption         =   "frmQuoteSettings.frx":0CCC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmQuoteSettings.frx":0CFE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0D88
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFullHeight 
         Height          =   220
         Left            =   1200
         TabIndex        =   17
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
         Caption         =   "frmQuoteSettings.frx":0DA4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0DCC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0E28
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFonts 
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4275
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmQuoteSettings.frx":0E44
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":0E6E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":0E8E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdBoxFont 
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   840
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
         Caption         =   "frmQuoteSettings.frx":0EAA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0EDC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0EFC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdGridFont 
         Height          =   435
         Left            =   180
         TabIndex        =   1
         Top             =   300
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
         Caption         =   "frmQuoteSettings.frx":0F18
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":0F4C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0F6C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblBoxFontSample 
         Height          =   495
         Left            =   1440
         Top             =   840
         Width           =   2715
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":0F88
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":0FC6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":0FE6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGridFontSample 
         Height          =   495
         Left            =   1440
         Top             =   300
         Width           =   2715
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":1002
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":1042
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":1062
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1275
      Left            =   4500
      TabIndex        =   22
      Top             =   180
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
      Caption         =   "frmQuoteSettings.frx":107E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":10AA
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":10CA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   0
         TabIndex        =   24
         Top             =   540
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
         Caption         =   "frmQuoteSettings.frx":10E6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":1114
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":1134
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   23
         Top             =   0
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
         Caption         =   "frmQuoteSettings.frx":1150
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":1176
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":1196
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraForex 
      Height          =   2295
      Left            =   1065
      TabIndex        =   29
      Top             =   1620
      Width           =   3555
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmQuoteSettings.frx":11B2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteSettings.frx":11DE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteSettings.frx":11FE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkUpDownBk 
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   3165
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmQuoteSettings.frx":121A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmQuoteSettings.frx":12C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":12E6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdForexBk 
         Height          =   315
         Left            =   480
         TabIndex        =   30
         Top             =   1800
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdForexDown 
         Height          =   315
         Left            =   1920
         TabIndex        =   31
         Top             =   1080
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdForexUp 
         Height          =   315
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdForexText 
         Height          =   315
         Left            =   1920
         TabIndex        =   33
         Top             =   1800
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label11 
         Height          =   255
         Left            =   600
         Top             =   840
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
         Caption         =   "frmQuoteSettings.frx":1302
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":1334
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":1354
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label10 
         Height          =   255
         Left            =   1920
         Top             =   840
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
         Caption         =   "frmQuoteSettings.frx":1370
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":13A6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":13C6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   195
         Left            =   2040
         Top             =   1560
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
         Caption         =   "frmQuoteSettings.frx":13E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":1418
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":1438
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   195
         Left            =   360
         Top             =   1560
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
         Caption         =   "frmQuoteSettings.frx":1454
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmQuoteSettings.frx":1496
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmQuoteSettings.frx":14B6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmQuoteSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmQuoteSettings.frm
'' Description: Allow the user to change settings on the quote board
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 01/29/04  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      None
'' Returns:     True if OK, False if Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(strGridFont As String, strBoxFont As String, UpColor As OLE_COLOR, _
                DownColor As OLE_COLOR, UnchColor As OLE_COLOR, _
                UpdateColor As OLE_COLOR, bFullHeight As Boolean, lStyle As Long, _
                lForexBkColor As Long, lForexTextColor As Long, lForexUpDownBk As Long, _
                lUseUpdateColor As Long, lColorSymbol As Long, eShowExtraInfo As eGEQbExtraInfo, _
                lCompactQB As Long, bShowButtons As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bForex As Boolean
    
    ' Set up the font information...
    FontFromString lblGridFontSample.Font, strGridFont
    FontFromString lblBoxFontSample.Font, strBoxFont
    
    ' Set up the color information...
    gdUpColor.Color = UpColor
    gdDownColor.Color = DownColor
    gdUnchColor.Color = UnchColor
    gdUpdateColor.Color = UpdateColor
    
    gdForexUp.Color = UpColor
    gdForexDown.Color = DownColor
    gdForexBk.Color = lForexBkColor
    gdForexText.Color = lForexTextColor
    chkUpDownBk.Value = lForexUpDownBk
    
    Select Case eShowExtraInfo
        Case eGEQbExtraInfo_None:
            chkDateTime.Value = vbUnchecked
            chkBidAsk.Value = vbUnchecked
        Case eGEQbExtraInfo_TimeStamp:
            chkDateTime.Value = vbChecked
            chkBidAsk.Value = vbUnchecked
        Case eGEQbExtraInfo_BidAsk:
            chkDateTime.Value = vbUnchecked
            chkBidAsk.Value = vbChecked
        Case eGEQbExtraInfo_All:
            chkDateTime.Value = vbChecked
            chkBidAsk.Value = vbChecked
    End Select
    
    chkCompactQB.Value = lCompactQB
    
    ' Set up the bar information...
    optFullHeight = bFullHeight
    optRelativeHeight = Not bFullHeight
    Select Case lStyle
        Case eGDQuoteStyle_Thermometer:
            optThermometer.Value = True
        Case eGDQuoteStyle_OHLC:
            optOHLC.Value = True
        Case eGDQuoteStyle_Candlestick:
            optCandlestick.Value = True
        Case eGDQuoteStyle_Bollinger:
            optBollinger.Value = True
        Case eGDQuoteStyle_NoBarPicture:
            optNone.Value = True
        Case eGDQuoteStyle_Forex
            bForex = True
            cmdBoxFont.Caption = "Forex Font"
            lblBoxFontSample.Caption = "Forex font sample"
    End Select
    
    If lUseUpdateColor = 0 Then
        chkUpdatedColor.Value = vbUnchecked
    Else
        chkUpdatedColor.Value = vbChecked
    End If
    
    If lColorSymbol = 0 Then
        chkColorSymbol.Value = vbUnchecked
    Else
        chkColorSymbol.Value = vbChecked
    End If
    
    chkShowButtons.Value = Abs(bShowButtons)
    
    'show different frame for forex style
    fraColors.Visible = Not bForex
    fraBarStyle.Visible = Not bForex
    fraBarHeight.Visible = Not bForex
    fraForex.Visible = bForex
    
    ShowForm Me, eForm_ActModal
    
    If m.bOK Then
        strGridFont = FontToString(lblGridFontSample.Font)
        strBoxFont = FontToString(lblBoxFontSample.Font)
        
        bShowButtons = chkShowButtons.Value * (-1)
        lCompactQB = chkCompactQB.Value
        
        If bForex Then
            UpColor = gdForexUp.Color
            DownColor = gdForexDown.Color
            lForexBkColor = gdForexBk.Color
            lForexTextColor = gdForexText.Color
            lForexUpDownBk = chkUpDownBk.Value
        Else
            UpColor = gdUpColor.Color
            DownColor = gdDownColor.Color
            UnchColor = gdUnchColor.Color
            UpdateColor = gdUpdateColor.Color
            bFullHeight = optFullHeight.Value
            Select Case True
                Case optThermometer
                    lStyle = eGDQuoteStyle_Thermometer
                Case optOHLC
                    lStyle = eGDQuoteStyle_OHLC
                Case optCandlestick
                    lStyle = eGDQuoteStyle_Candlestick
                Case optBollinger
                    lStyle = eGDQuoteStyle_Bollinger
                Case optNone
                    lStyle = eGDQuoteStyle_NoBarPicture
            End Select
            If lStyle <> eGDQuoteStyle_Forex Then
                'only change these values for box or grid style style
                lUseUpdateColor = chkUpdatedColor.Value
                lColorSymbol = chkColorSymbol.Value
                
                If chkDateTime.Value = vbUnchecked And chkBidAsk.Value = vbUnchecked Then
                    eShowExtraInfo = eGEQbExtraInfo_None
                ElseIf chkDateTime.Value = vbChecked And chkBidAsk.Value = vbUnchecked Then
                    eShowExtraInfo = eGEQbExtraInfo_TimeStamp
                ElseIf chkDateTime.Value = vbUnchecked And chkBidAsk.Value = vbChecked Then
                    eShowExtraInfo = eGEQbExtraInfo_BidAsk
                ElseIf chkDateTime.Value = vbChecked And chkBidAsk.Value = vbChecked Then
                    eShowExtraInfo = eGEQbExtraInfo_All
                End If
                
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmQuoteSettings.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAutoUpdate_Click
'' Description: Allow the user to change the automatic refresh times
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAutoUpdate_Click()
On Error GoTo ErrSection:

    frmConfig.ShowMe eSnapQuoteTab

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuoteSettings.cmdAutoUpdate.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBoxFont_Click
'' Description: Allow the user to change the font for the box style quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBoxFont_Click()
On Error GoTo ErrSection:

    CommonDialogFont frmMain.CommonDialog1, lblBoxFontSample.Font

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteSettings.cmdBoxFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: When Cancel is clicked, let ShowMe unload the form
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
    RaiseError "frmQuoteSettings.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCategories_Click
'' Description: Allow the user to change the category information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCategories_Click()
On Error GoTo ErrSection:

    frmQuotes.EditCategories

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cQuoteSettings.cmdCategories.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdGridFont_Click
'' Description: Allow the user to change the font for the grid quote board
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdGridFont_Click()
On Error GoTo ErrSection:

    CommonDialogFont frmMain.CommonDialog1, lblGridFontSample.Font

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteSettings.cmdGridFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: When OK is clicked, let ShowMe save the settings and unload
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteSettings.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Quote Board Settings"
    Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteSettings.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', let ShowMe unload the form
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteSettings.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

