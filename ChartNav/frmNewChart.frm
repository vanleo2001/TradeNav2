VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmNewChart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Chart"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   13350
   Begin HexUniControls.ctlUniFrameWL fraSeasonal 
      Height          =   855
      Left            =   5880
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmNewChart.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNewChart.frx":0040
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNewChart.frx":0060
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtCycleNum 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   315
         Width           =   495
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmNewChart.frx":007C
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
         Tip             =   "frmNewChart.frx":009E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":00BE
      End
      Begin HexUniControls.ctlUniComboImageXP cboCycle 
         Height          =   315
         Left            =   1575
         TabIndex        =   34
         Top             =   300
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
         Tip             =   "frmNewChart.frx":00DA
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":00FA
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboBarType 
         Height          =   315
         Left            =   3960
         TabIndex        =   33
         Top             =   300
         Width           =   1395
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
         Tip             =   "frmNewChart.frx":0116
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0136
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate dtFromDate 
         Height          =   315
         Left            =   2520
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         AllowWeekends   =   0   'False
         MaxDate         =   42605
         MaxDateIsToday  =   -1  'True
         Value           =   2
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   195
         Left            =   1920
         Top             =   660
         Visible         =   0   'False
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
         Caption         =   "frmNewChart.frx":0152
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNewChart.frx":017C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":019C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   360
         Top             =   360
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
         Caption         =   "frmNewChart.frx":01B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNewChart.frx":01E4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0204
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   3240
         Top             =   360
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
         Caption         =   "frmNewChart.frx":0220
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNewChart.frx":0252
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0272
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11880
      Top             =   1320
   End
   Begin HexUniControls.ctlUniFrameWL fraTrades 
      Height          =   1095
      Left            =   5880
      TabIndex        =   16
      Top             =   1320
      Width           =   5595
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmNewChart.frx":028E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNewChart.frx":02E0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNewChart.frx":0300
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Top             =   660
         Width           =   3135
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
         Tip             =   "frmNewChart.frx":031C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":033C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtStrategy 
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   420
         Width           =   2835
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmNewChart.frx":0358
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
         Tip             =   "frmNewChart.frx":0378
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0398
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSelectStrategy 
         Height          =   300
         Left            =   4860
         TabIndex        =   17
         Top             =   420
         Width           =   315
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
         Caption         =   "frmNewChart.frx":03B4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":03DA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":03FA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTradesAccount 
         Height          =   255
         Left            =   420
         TabIndex        =   20
         Top             =   720
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
         Caption         =   "frmNewChart.frx":0416
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":0456
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0476
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTradesNone 
         Height          =   255
         Left            =   420
         TabIndex        =   21
         Top             =   240
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
         Caption         =   "frmNewChart.frx":0492
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":04BA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":04DA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optTradesStrategy 
         Height          =   255
         Left            =   420
         TabIndex        =   22
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
         Caption         =   "frmNewChart.frx":04F6
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":0528
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0548
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraSettings 
      Height          =   1215
      Left            =   150
      TabIndex        =   12
      Top             =   4230
      Width           =   5595
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmNewChart.frx":0564
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNewChart.frx":05B0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNewChart.frx":05D0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkCopyAnnots 
         Height          =   435
         Left            =   3420
         TabIndex        =   15
         Top             =   720
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
         Caption         =   "frmNewChart.frx":05EC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmNewChart.frx":0654
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0674
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   780
         Width           =   1395
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
         Tip             =   "frmNewChart.frx":0690
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
         MouseIcon       =   "frmNewChart.frx":06B0
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBarPeriod 
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   780
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
         Caption         =   "frmNewChart.frx":06CC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":0702
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0722
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboTemplate 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   300
         Width           =   3435
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
         Tip             =   "frmNewChart.frx":073E
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":075E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblTemplate 
         Height          =   255
         Left            =   420
         Top             =   330
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
         Caption         =   "frmNewChart.frx":077A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNewChart.frx":07B8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":07D8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraData 
      Height          =   3950
      Left            =   150
      TabIndex        =   11
      Top             =   120
      Width           =   5595
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmNewChart.frx":07F4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNewChart.frx":0840
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNewChart.frx":0860
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optProfile 
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   1670
         Width           =   4305
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":087C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":08E6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":095C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSeasonal 
         Height          =   255
         Left            =   180
         TabIndex        =   40
         Top             =   1350
         Width           =   4305
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":0978
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":09CA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0A40
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optPercentComp 
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   1030
         Width           =   4305
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":0A5C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":0AAE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0B24
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraSymbols 
         Height          =   1875
         Left            =   90
         TabIndex        =   24
         Top             =   2040
         Width           =   5400
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":0B40
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmNewChart.frx":0B66
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0B86
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectColor gdSelectColor1 
            Height          =   270
            Left            =   4530
            TabIndex        =   31
            Top             =   1410
            Visible         =   0   'False
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   476
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdFgDelete 
            Height          =   375
            Left            =   4320
            TabIndex        =   27
            Top             =   300
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
            Caption         =   "frmNewChart.frx":0BA2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmNewChart.frx":0BD6
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":0BF6
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdSaveAs 
            Height          =   375
            Left            =   4320
            TabIndex        =   26
            Top             =   780
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
            Caption         =   "frmNewChart.frx":0C12
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmNewChart.frx":0C40
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":0C60
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAutoMultiplier 
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   5055
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmNewChart.frx":0C7C
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmNewChart.frx":0D24
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":0D44
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin VSFlex7LCtl.VSFlexGrid fgPercentComp 
            Height          =   570
            Left            =   0
            TabIndex        =   30
            Top             =   1035
            Width           =   4215
            _cx             =   7435
            _cy             =   1005
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
         Begin VSFlex7LCtl.VSFlexGrid fgSpread 
            Height          =   570
            Left            =   0
            TabIndex        =   28
            Top             =   300
            Width           =   4215
            _cx             =   7435
            _cy             =   1005
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
      Begin HexUniControls.ctlUniComboImageXP cboSpread 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   630
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
         Tip             =   "frmNewChart.frx":0D60
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0D80
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSelectSym 
         Height          =   300
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   315
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
         Caption         =   "frmNewChart.frx":0D9C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":0DC2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0DE2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   240
         Width           =   2595
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmNewChart.frx":0DFE
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
         Tip             =   "frmNewChart.frx":0E1E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0E3E
      End
      Begin HexUniControls.ctlUniRadioXP optStandard 
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   300
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
         Caption         =   "frmNewChart.frx":0E5A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmNewChart.frx":0E88
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0EFE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optSpread 
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   650
         Width           =   1300
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":0F1A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":0F48
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":0FBE
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   255
         Left            =   4665
         Top             =   135
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
         Caption         =   "frmNewChart.frx":0FDA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmNewChart.frx":1008
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":1028
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   555
      Left            =   720
      TabIndex        =   10
      Top             =   5880
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
      Caption         =   "frmNewChart.frx":1044
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmNewChart.frx":1078
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmNewChart.frx":1098
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   120
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
         Caption         =   "frmNewChart.frx":10B4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":10DA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":10FA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   120
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
         Caption         =   "frmNewChart.frx":1116
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmNewChart.frx":1144
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":1164
         RightToLeft     =   0   'False
      End
   End
   Begin NavTradeSenseV3.Editor Editor1 
      Height          =   285
      Left            =   11760
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   503
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTabProfile 
      Height          =   6135
      Left            =   6600
      TabIndex        =   42
      Top             =   2880
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   10821
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
      Caption         =   "Data|Profile Display"
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
      Begin HexUniControls.ctlUniFrameWL fraTPO 
         Height          =   5760
         Left            =   6855
         TabIndex        =   57
         Top             =   330
         Width           =   6120
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":1180
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmNewChart.frx":11AC
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":11CC
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtForecastBars 
            Height          =   285
            Left            =   3405
            TabIndex        =   6
            Top             =   785
            Width           =   435
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmNewChart.frx":11E8
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
            Tip             =   "frmNewChart.frx":120C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":122C
         End
         Begin HexUniControls.ctlUniFrameWL fraStats 
            Height          =   1695
            Left            =   420
            TabIndex        =   64
            Top             =   3990
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
            Caption         =   "frmNewChart.frx":1248
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmNewChart.frx":127C
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":129C
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniTextBoxXP txtPercentVolume_VA 
               Height          =   315
               Left            =   2640
               TabIndex        =   80
               Top             =   570
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmNewChart.frx":12B8
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
               Tip             =   "frmNewChart.frx":12DC
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":12FC
            End
            Begin HexUniControls.ctlUniTextBoxXP txtPercentTPO_VA 
               Height          =   315
               Left            =   2640
               TabIndex        =   78
               Top             =   210
               Width           =   615
               _ExtentX        =   0
               _ExtentY        =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frmNewChart.frx":1318
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
               Tip             =   "frmNewChart.frx":133C
               HideSelection   =   -1  'True
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":135C
            End
            Begin gdOCX.gdSelectColor gdColorVolume_POC 
               Height          =   315
               Left            =   3750
               TabIndex        =   72
               Top             =   1290
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin gdOCX.gdSelectColor gdColorTPO_POC 
               Height          =   315
               Left            =   3750
               TabIndex        =   71
               Top             =   930
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin gdOCX.gdSelectColor gdColorVolume_VA 
               Height          =   315
               Left            =   3750
               TabIndex        =   70
               Top             =   570
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin gdOCX.gdSelectColor gdColorTPO_VA 
               Height          =   315
               Left            =   3750
               TabIndex        =   69
               Top             =   210
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniCheckXP chkVolume_POC 
               Height          =   255
               Left            =   270
               TabIndex        =   68
               Top             =   1320
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
               Caption         =   "frmNewChart.frx":1378
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":13D2
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":13F2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkTPO_POC 
               Height          =   255
               Left            =   270
               TabIndex        =   67
               Top             =   960
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
               Caption         =   "frmNewChart.frx":140E
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":1462
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1482
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkVolume_VA 
               Height          =   255
               Left            =   270
               TabIndex        =   66
               Top             =   600
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
               Caption         =   "frmNewChart.frx":149E
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":14E4
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1504
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkTPO_VA 
               Height          =   255
               Left            =   270
               TabIndex        =   65
               Top             =   240
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
               Caption         =   "frmNewChart.frx":1520
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":1560
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1580
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label14 
               Height          =   225
               Left            =   3240
               Top             =   615
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
               Caption         =   "frmNewChart.frx":159C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":15BE
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":15DE
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP Label13 
               Height          =   225
               Left            =   3240
               Top             =   255
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
               Caption         =   "frmNewChart.frx":15FA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":161C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":163C
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniComboBoxXP cboTicksPerRow 
            Height          =   315
            Left            =   2640
            TabIndex        =   58
            Top             =   360
            Width           =   1200
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
            Tip             =   "frmNewChart.frx":1658
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
            MouseIcon       =   "frmNewChart.frx":1678
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
            MaxLength       =   0
            RightToLeft     =   0   'False
            LeftMargin      =   0
            RightMargin     =   0
            SelectOnFocus   =   0   'False
         End
         Begin HexUniControls.ctlUniFrameWL fraLetters 
            Height          =   2745
            Left            =   420
            TabIndex        =   59
            Top             =   1200
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
            Caption         =   "frmNewChart.frx":1694
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmNewChart.frx":16DE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":16FE
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniComboImageXP cboColorScheme 
               Height          =   315
               Left            =   270
               TabIndex        =   23
               Top             =   1560
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
               Tip             =   "frmNewChart.frx":171A
               Sorted          =   0   'False
               HScroll         =   0   'False
               RoundedBorders  =   -1  'True
               IconDim         =   16
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":173A
               DropDownOnTextClick=   -1  'True
               DropDownWidth   =   -1
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkOpenClose 
               Height          =   255
               Left            =   270
               TabIndex        =   37
               Top             =   2370
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
               Caption         =   "frmNewChart.frx":1756
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":17A8
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":17C8
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectColor gdColorVolume 
               Height          =   315
               Left            =   3000
               TabIndex        =   77
               Top             =   1950
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniCheckXP chkVolume 
               Height          =   255
               Left            =   270
               TabIndex        =   76
               Top             =   1980
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
               Caption         =   "frmNewChart.frx":17E4
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":182C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":184C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniCheckXP chkTPO 
               Height          =   255
               Left            =   270
               TabIndex        =   75
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
               Caption         =   "frmNewChart.frx":1868
               Enabled         =   -1  'True
               Align           =   0
               CheckBackColor  =   -2147483643
               CheckForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":1898
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":18B8
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniButtonImageXP cmdFontTPO 
               Cancel          =   -1  'True
               Height          =   315
               Left            =   3405
               TabIndex        =   63
               Top             =   630
               Width           =   1425
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
               Caption         =   "frmNewChart.frx":18D4
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               ShowFocus       =   -1  'True
               Tristate        =   0   'False
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":18FC
               Style           =   -1
               RoundedBorders  =   -1  'True
               xTranspColor    =   0
               yTranspColor    =   0
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":191C
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optBlocks 
               Height          =   255
               Left            =   270
               TabIndex        =   62
               Top             =   960
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
               Caption         =   "frmNewChart.frx":1938
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":1994
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":19B4
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optClasicLettering 
               Height          =   255
               Left            =   270
               TabIndex        =   61
               Top             =   660
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
               Caption         =   "frmNewChart.frx":19D0
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":1A1C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1A3C
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optAlphaNumeric 
               Height          =   255
               Left            =   270
               TabIndex        =   60
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
               Caption         =   "frmNewChart.frx":1A58
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":1AB2
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1AD2
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectColor gdColorOpenClose 
               Height          =   315
               Left            =   3000
               TabIndex        =   38
               Top             =   2340
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin gdOCX.gdSelectColor gdColorFrom 
               Height          =   315
               Left            =   3000
               TabIndex        =   39
               Top             =   1560
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin gdOCX.gdSelectColor gdColorTo 
               Height          =   315
               Left            =   3960
               TabIndex        =   49
               Top             =   1560
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
               CustomColor     =   255
            End
            Begin HexUniControls.ctlUniLabelXP lblColorLabel1 
               Height          =   255
               Left            =   3000
               Top             =   1320
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
               Caption         =   "frmNewChart.frx":1AEE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":1B18
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1B38
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblColorLabel2 
               Height          =   255
               Left            =   4200
               Top             =   1320
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
               Caption         =   "frmNewChart.frx":1B54
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":1B7A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":1B9A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin VB.Image imgGradientHorz 
               Height          =   480
               Left            =   3600
               Picture         =   "frmNewChart.frx":1BB6
               Top             =   1320
               Width           =   480
            End
         End
         Begin HexUniControls.ctlUniLabelXP Label21 
            Height          =   195
            Left            =   805
            Top             =   825
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
            Caption         =   "frmNewChart.frx":1EC0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewChart.frx":1F1A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":1F3A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label12 
            Height          =   210
            Left            =   3960
            Top             =   405
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
            Caption         =   "frmNewChart.frx":1F56
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewChart.frx":1F90
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":1FB0
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label11 
            Height          =   210
            Left            =   805
            Top             =   405
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
            Caption         =   "frmNewChart.frx":1FCC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmNewChart.frx":2014
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":2034
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSessionData 
         Height          =   5760
         Left            =   45
         TabIndex        =   43
         Top             =   330
         Width           =   6120
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmNewChart.frx":2050
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmNewChart.frx":208C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmNewChart.frx":20AC
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniFrameWL fraDates 
            Height          =   1695
            Left            =   120
            TabIndex        =   44
            Top             =   180
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
            Caption         =   "frmNewChart.frx":20C8
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmNewChart.frx":2102
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":2122
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniFrameWL fraIntraday 
               Height          =   1500
               Left            =   3120
               TabIndex        =   51
               Top             =   105
               Width           =   2610
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmNewChart.frx":213E
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmNewChart.frx":2174
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":2194
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtNumProfiles 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   50
                  Top             =   150
                  Width           =   495
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmNewChart.frx":21B0
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
                  Tip             =   "frmNewChart.frx":21D4
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":21F4
               End
               Begin HexUniControls.ctlUniButtonImageXP cmdStartStop 
                  Height          =   315
                  Left            =   1695
                  TabIndex        =   52
                  Top             =   1140
                  Width           =   900
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
                  Caption         =   "frmNewChart.frx":2210
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  ShowFocus       =   -1  'True
                  Tristate        =   0   'False
                  Pressed         =   0   'False
                  Tip             =   "frmNewChart.frx":223C
                  Style           =   -1
                  RoundedBorders  =   -1  'True
                  xTranspColor    =   0
                  yTranspColor    =   0
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":225C
                  RightToLeft     =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblNumSessions 
                  Height          =   195
                  Left            =   0
                  Top             =   195
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
                  Caption         =   "frmNewChart.frx":2278
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   2
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":22BC
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":22DC
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblStartStopInfo2 
                  Height          =   195
                  Left            =   45
                  Top             =   870
                  Width           =   2220
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "frmNewChart.frx":22F8
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":234E
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":236E
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblStartStopTimes 
                  Height          =   255
                  Left            =   45
                  Top             =   1185
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
                  Caption         =   "frmNewChart.frx":238A
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   2
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":23C4
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":23E4
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblStartStopInfo1 
                  Height          =   195
                  Left            =   0
                  Top             =   675
                  Width           =   2625
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "frmNewChart.frx":2400
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":2462
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":2482
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniRadioXP optEndOfData 
               Height          =   240
               Left            =   300
               TabIndex        =   45
               Top             =   1170
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
               Caption         =   "frmNewChart.frx":249E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   -1  'True
               Tip             =   "frmNewChart.frx":24D6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":24F6
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin gdOCX.gdSelectDate dtDateTo 
               Height          =   315
               Left            =   600
               TabIndex        =   47
               Top             =   840
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               AllowWeekends   =   0   'False
               Value           =   37274
            End
            Begin gdOCX.gdSelectDate dtDateFrom 
               Height          =   315
               Left            =   600
               TabIndex        =   48
               Top             =   300
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               AllowWeekends   =   0   'False
               MaxDate         =   42605
               MaxDateIsToday  =   -1  'True
               Value           =   2
            End
            Begin HexUniControls.ctlUniRadioXP optToDate 
               Height          =   240
               Left            =   300
               TabIndex        =   46
               Top             =   900
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
               Caption         =   "frmNewChart.frx":2512
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":2544
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":2564
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblEndOfData 
               Height          =   255
               Left            =   555
               Top             =   1163
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
               Caption         =   "frmNewChart.frx":2580
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":25B8
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":25D8
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblDateFrom 
               Height          =   195
               Left            =   180
               Top             =   360
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
               Caption         =   "frmNewChart.frx":25F4
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":261E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":263E
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
            Begin HexUniControls.ctlUniLabelXP lblDateTo 
               Height          =   195
               Left            =   180
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
               Caption         =   "frmNewChart.frx":265A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":2680
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":26A0
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
         Begin HexUniControls.ctlUniFrameWL fraSessionProfile 
            Height          =   3675
            Left            =   120
            TabIndex        =   56
            Top             =   1875
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
            Caption         =   "frmNewChart.frx":26BC
            Enabled         =   -1  'True
            ForeColor       =   -2147483642
            BackColor       =   -2147483633
            Tip             =   "frmNewChart.frx":26FE
            VistaStyle      =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmNewChart.frx":271E
            RightToLeft     =   0   'False
            Begin HexUniControls.ctlUniRadioXP optMultiDays 
               Height          =   255
               Left            =   360
               TabIndex        =   53
               Top             =   2205
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
               Caption         =   "frmNewChart.frx":273A
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":2788
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":27A8
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniRadioXP optIntraday 
               Height          =   255
               Left            =   360
               TabIndex        =   54
               Top             =   1005
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
               Caption         =   "frmNewChart.frx":27C4
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frmNewChart.frx":2800
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":2820
               ShowFocus       =   -1  'True
               RightToLeft     =   0   'False
            End
            Begin HexUniControls.ctlUniFrameWL fraMultiSessions 
               Height          =   1095
               Left            =   240
               TabIndex        =   55
               Top             =   2205
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
               Caption         =   "frmNewChart.frx":283C
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmNewChart.frx":287C
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":289C
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtSessionsPerProfile 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   73
                  Top             =   600
                  Width           =   735
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmNewChart.frx":28B8
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
                  Tip             =   "frmNewChart.frx":28DA
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":28FA
               End
               Begin HexUniControls.ctlUniLabelXP lblMultidayDesc 
                  Height          =   255
                  Left            =   60
                  Top             =   270
                  Width           =   5325
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "frmNewChart.frx":2916
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   2
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":29B4
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":29D4
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblSessionsPerProfile 
                  Height          =   255
                  Left            =   1200
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
                  Caption         =   "frmNewChart.frx":29F0
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":2A3A
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":2A5A
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniFrameWL fraIntradayiSessions 
               Height          =   975
               Left            =   240
               TabIndex        =   74
               Top             =   1005
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
               Caption         =   "frmNewChart.frx":2A76
               Enabled         =   -1  'True
               ForeColor       =   -2147483642
               BackColor       =   -2147483633
               Tip             =   "frmNewChart.frx":2ABE
               VistaStyle      =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":2ADE
               RightToLeft     =   0   'False
               Begin HexUniControls.ctlUniTextBoxXP txtIntradayMinutes 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   79
                  Top             =   600
                  Width           =   735
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Enabled         =   -1  'True
                  Locked          =   0   'False
                  Text            =   "frmNewChart.frx":2AFA
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
                  Tip             =   "frmNewChart.frx":2B1E
                  HideSelection   =   -1  'True
                  RightToLeft     =   0   'False
                  ManualStart     =   0   'False
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":2B3E
               End
               Begin HexUniControls.ctlUniLabelXP lblIntradayDesc 
                  Height          =   255
                  Left            =   60
                  Top             =   270
                  Width           =   5325
                  _ExtentX        =   0
                  _ExtentY        =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "frmNewChart.frx":2B5A
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   2
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":2BDE
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":2BFE
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
               Begin HexUniControls.ctlUniLabelXP lblIntradayMin 
                  Height          =   255
                  Left            =   1200
                  Top             =   630
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
                  Caption         =   "frmNewChart.frx":2C1A
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Alignment       =   0
                  VAlignment      =   0
                  BackStyle       =   1
                  BorderStyle     =   0
                  AutoSize        =   0   'False
                  Tip             =   "frmNewChart.frx":2C5A
                  Style           =   0
                  Enabled         =   -1  'True
                  Margin          =   0
                  RoundedBorders  =   0   'False
                  MousePointer    =   0
                  MouseIcon       =   "frmNewChart.frx":2C7A
                  RightToLeft     =   0   'False
                  WordWrap        =   0   'False
               End
            End
            Begin HexUniControls.ctlUniLabelXP Label5 
               Height          =   615
               Left            =   480
               Top             =   270
               Width           =   4695
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frmNewChart.frx":2C96
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   2
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frmNewChart.frx":2E0A
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frmNewChart.frx":2E2A
               RightToLeft     =   0   'False
               WordWrap        =   0   'False
            End
         End
      End
   End
End
Attribute VB_Name = "frmNewChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kActiveChart = "< Copy settings from active chart >"
Private Const kCustomSpread = "< Custom >"
Private Const kTopOfSymRadio = 300
Private Const kTopOfSymText = 270
Private Const kGridHeight = 1515
Private Const kFrameDataHeight = 3950

'constants for spread operators
Private Const kPlus = "plus"
Private Const kMinus = "minus"
Private Const kDivide = "divide"
Private Const kOpAll = "plus|minus|divide"
Private Const kOpPlusMinus = "plus|minus|"

'form width / height
Private Const kFormWidth = 5940
Private Const kFormHeight = 7125

Private Type mPrivate
    oFunctionTree As New cGdTree
    Chart As cChart
    strSymbol As String
    strPrevSym As String        'for determining whether to rebuild the percent change grid
    bEditSpread As Boolean
    bEditProfile As Boolean
    eNewChartType As eChartType
    
    nFrameSpace As Long
    nPeriodicity As Long
    nAccountID As Long
    nShowTrades As Long
    nStrategyId As Long         'this is real, not gamemode, strategy ID
    
    nNumProfilesIntraday As Long
    nNumProfilesMultiday As Long
    nTicksPerRow As Long
    nSessionsPerProfile As Long
    dtDateFrom As Double
    dtDtDateTo As Double
    dtEODLastDate As Double
    
    bConvert As Boolean         'flag to indicate whether to re-save function with { ... } (08-25-2005)
    bSkipGridLoad As Boolean
    bInit As Boolean

    bColorChecked As Boolean    'for color selector of percent change grid
    lMouseRow As Long
    lMouseCol As Long
    
    bFirstActivate As Boolean
End Type

Private m As mPrivate

Private Sub cboBarPeriod_Click()
On Error GoTo ErrSection:

    Dim i&
    
    i = GetPeriodicity(cboBarPeriod.Text)
    If i <> m.nPeriodicity Then
        m.nPeriodicity = i
        FixControls
    End If
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.cboBarPeriod_Click"
End Sub

Private Sub cboBarPeriod_LostFocus()
On Error GoTo ErrSection:

    Dim i&
    
    i = GetPeriodicity(cboBarPeriod.Text)
    If i <> m.nPeriodicity Then
        m.nPeriodicity = i
        FixControls
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.cboBarPeriod_LostFocus"

End Sub

Private Sub cboColorScheme_Click()
On Error GoTo ErrSection:
    
    Dim Ind As cIndicator
    
    If Not Me.Visible Then Exit Sub
    
    If Not m.Chart Is Nothing Then
        If Not m.Chart.Tree Is Nothing Then
            Set Ind = m.Chart.Tree("PRICE")
            If Not Ind Is Nothing Then
                Ind.ProfileColorScheme = cboColorScheme.ListIndex
            End If
        End If
    End If
    
    SetProfileControls

ErrEixt:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.cboColorScheme_Click"

End Sub

Private Sub cboCycle_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        If cboCycle.ListIndex > 1 Then
            cboBarType.ListIndex = 0
            cboBarType.Enabled = False
        Else
            cboBarType.Enabled = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.cboCycle_Click"

End Sub

Private Sub cboSpread_Click()
On Error GoTo ErrSection:

    If cboSpread.ListIndex > 0 Or m.bInit Then
        LoadGrid
    End If
    
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.cboSpread_Click"
    
End Sub

Private Sub cboTemplate_Click()
On Error GoTo ErrSection:

    Dim i&, strFile$, strText$
    Dim nPeriod&, nShowTrades&, nSystemID&
    
    Dim aStrings As New cGdArray
    
    strFile = cboTemplate.Text
    If strFile <> kActiveChart Then
    
        nPeriod = -1
        nShowTrades = -1
        nSystemID = -1
        aStrings.FromFile App.Path & "\Charts\Templates\" & strFile & ".cht", , "END="
        For i = 0 To aStrings.Size - 1
            strText = Trim(UCase(aStrings(i)))
            If Parse(strText, "=", 1) = "PERIODICITY" Then
                
                nPeriod = Val(Parse(strText, "=", 2))
                m.nPeriodicity = nPeriod
                If nShowTrades <> -1 And nSystemID <> -1 Then Exit For
            
            ElseIf Parse(strText, "=", 1) = "SHOWTRADES" Then
                
                nShowTrades = Val(Parse(strText, "=", 2))
                m.nShowTrades = nShowTrades
                If nPeriod <> -1 And nSystemID <> -1 Then Exit For
            
            ElseIf Parse(strText, "=", 1) = "SYSTEMID" Then
                
                nSystemID = Val(Parse(strText, "=", 2))
                m.nStrategyId = nSystemID
                If nPeriod <> -1 And nSystemID <> -1 Then Exit For
                
            End If
        Next
    ElseIf Not ActiveChart Is Nothing Then
        chkCopyAnnots.Value = 1
        m.nPeriodicity = ActiveChart.Chart.Bars.Prop(eBARS_Periodicity)
    Else
        m.nPeriodicity = 0
    End If
    
    If nShowTrades = -1 Then m.nShowTrades = 0      'did not exist in template
    If nSystemID = -1 Then m.nStrategyId = 0
    If nPeriod = -1 Then m.nPeriodicity = GetPeriodicity("Daily")
    
    FixControls
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.cboTemplate_Click"
End Sub

Private Sub chkAutoMultiplier_Click()
On Error GoTo ErrSection:

    If Not m.bInit Then
        FixSpreadGrid
        If cboSpread.ListIndex > 0 Then
            cboSpread.ListIndex = 0
        End If
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.chkAutoMultiplier_Click"
    
End Sub

Private Sub cmdBarPeriod_Click()
On Error GoTo ErrSection:

    Dim i&, Bars As New cGdBars
    
    SetBarProperties Bars, m.strSymbol
    i = frmBarPeriod.ShowMe(m.nPeriodicity, Bars)
    If i > 0 Then m.nPeriodicity = i
    Set Bars = Nothing
    FixControls
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.cmdBarPeriod_Click"
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next:

    Unload Me

End Sub

Private Sub NewGridRow()
On Error GoTo ErrSection:
    
    With fgSpread
        If .MergeRow(.Rows - 1) = True Then
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = ""
            .MergeRow(.Rows - 1) = False
            If .Rows > .FixedRows + 2 Then
                .ColComboList(0) = kOpPlusMinus
            Else
                .ColComboList(0) = kOpAll
            End If
            .Row = .Rows - 1
            .Col = 0
            If .Row = .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                .Col = 2
                fgSpread_CellButtonClick .FixedRows, 2
            Else
                .EditCell
                SendKeys "{F4}" '(to dropdown the combo list)
            End If
        End If
    End With
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".NewGridRow"

End Sub

Private Sub cmdFgDelete_Click()
On Error GoTo ErrSection:

    Dim fg As VSFlexGrid
    
    If optSpread.Value = True Then
        Set fg = fgSpread
    ElseIf optPercentComp.Value = True Then
        gdSelectColor1.Visible = False
        Set fg = fgPercentComp
    End If

    If Not fg Is Nothing Then
        With fg
            If .Rows > .FixedRows + 1 Then
                If .Row >= .FixedRows And .Row < .Rows - 1 Then
                    .RemoveItem .Row
                    .TextMatrix(.FixedRows, 0) = ""
                    .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                    .Select .FixedRows, 0
                End If
            End If
        End With
    End If
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdFgDelete_Click"

End Sub

Private Sub cmdFontTPO_Click()
On Error GoTo ErrSection:

    Dim Ind As cIndicator

    If CommonDialogFont(frmMain.CommonDialog1, Me.Font) Then
        If Not m.Chart Is Nothing Then
            If Not m.Chart.Tree Is Nothing Then
                Set Ind = m.Chart.Tree("PRICE")
                
                If Not Ind Is Nothing Then
                    Ind.FontName = Me.Font.Name
                    Ind.FontBold = Me.Font.Bold
                    Ind.FontSize = Me.Font.Size
                    Ind.FontItalic = Me.FontItalic
                End If
                
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdFontTPO_Click"

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    Dim bUnload As Boolean
    Dim bOK As Boolean
    
    SetIniFileProperty "NewChartType", m.eNewChartType, "Charting", g.strIniFile
    
    Me.Hide
    DoEvents
    bUnload = True
    
    Select Case m.eNewChartType
        Case eTypeChart_Standard
            NewStandardChart
        Case eTypeChart_Spread
            If NewSpreadChart() Then
                StatusMsg ""
            Else
                bUnload = False
                ShowForm Me, eForm_Modal
            End If
        Case eTypeChart_PercentComp
            NewPercentChart
        Case eTypeChart_Seasonal
            bOK = False
            If HasModule("CYCLE") Then
                bOK = True
            ElseIf HasPlatinum(True) Then
                bOK = True
            End If
            If bOK Then
                SeasonalChartNew m.strSymbol, CDbl(dtFromDate.Value), Int(ValOfText(txtCycleNum)), cboCycle.Text, cboBarType.Text
            End If
        Case eTypeChart_Profile
            SaveProfileSettingsLastUsed             '6844
            NewProfileChart
        Case Default
            DebugLog "Unknown chart type: " & Str(m.eNewChartType) & " frmNewChart.cmdOK failed."
    End Select
    
ErrExit:
    If bUnload Then Unload Me
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdOK_Click"
        
End Sub

Private Sub cmdSaveAs_Click()
    SaveSpread
End Sub

Private Sub cmdSelectStrategy_Click()
On Error GoTo ErrSection:

    Dim nStrategyId As Long
    
    nStrategyId = ValOfText(frmAddToChart.ShowMe(eAdd_System, "New Chart"))
    If nStrategyId > 0 Then
        m.nStrategyId = nStrategyId
        txtStrategy = SystemNameForID(m.nStrategyId)
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSelectStrategy_Click"
    
End Sub

Private Sub cmdSelectSym_Click()
On Error GoTo ErrSection:
    
    m.strSymbol = txtSymbol.Text
    ShowSymSelector m.strSymbol
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".cmdSelectSym_Click"
        
End Sub

Private Sub cmdStartStop_Click()
On Error GoTo ErrSection:
    
    If Not m.Chart Is Nothing Then
        If Not m.Chart.Bars Is Nothing Then
            If frmStartStopTimes.ShowMe(m.Chart.Bars, m.Chart) = True Then
                StartStopTimeLabel m.Chart, lblStartStopTimes, lblStartStopInfo2        '6867
                m.Chart.RedoMode = eRedo9_ReloadData
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".cmdStartStop_Click"

End Sub

Private Sub fgPercentComp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim nRec&, lRow&, lCol&
    Dim aSymbols As New cGdArray
    Dim lColor&
    
    If tmr.Enabled Then
        Cancel = True
        Exit Sub
    End If
    
    With fgPercentComp
        
        lCol = .MouseCol
        lRow = .MouseRow
    
        If lCol = 1 And lRow = .Rows - 1 Then
            Set aSymbols = frmSymbolSelector.ShowMe("$DJIA", False, True, "Comparison Symbol")
            If aSymbols.Size > 0 Then
                nRec = g.SymbolPool.PoolRecForSymbol(aSymbols(0), True)
                If nRec < 0 Then
                    Beep
                Else
                    lColor = gdSelectColor1.Color
                    If lColor = 0 Then lColor = -1
                    
                    .TextMatrix(lRow, 1) = aSymbols(0)
                    .Cell(flexcpChecked, lRow, 0) = flexChecked
                    .Cell(flexcpPictureAlignment, lRow, 0) = flexAlignCenterCenter
                    .Cell(flexcpBackColor, lRow, 2) = lColor
                    
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = "Click to add..."
                End If
            End If
        End If
    
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.fgPercentComp_BeforeMouseDown"
    
End Sub

Private Sub fgPercentComp_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    gdSelectColor1.Visible = False
End Sub

Private Sub fgPercentComp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lCol&, lRow&, lRowsHeight&
    
    
    If tmr.Enabled Then Exit Sub
    
    With fgPercentComp
        lCol = .MouseCol
        lRow = .MouseRow
        
        If lCol = 2 Then
            If lRow >= .FixedRows And lRow < .Rows - 1 Then
                If InStr(.TextMatrix(lRow, 2), "Click") = 0 Then
                    m.lMouseRow = lRow
                    
                    lRowsHeight = .RowHeight(0) * .Rows
                    If lRowsHeight < .ClientHeight Then
                        gdSelectColor1.Move .Left + .Width - .ColWidth(2), .Top + .RowHeight(0) * .MouseRow, .ColWidth(2)
                    Else
                        'adjust for vertical scroll bar
                        gdSelectColor1.Move .Left + .Width - .ColWidth(2), .Top + .RowHeight(0) * (.MouseRow - .TopRow + 1), .ColWidth(2) - 225
                    End If
                    
                    gdSelectColor1.Color = .Cell(flexcpBackColor, lRow, lCol)
                    gdSelectColor1.Visible = True
                    gdSelectColor1.ZOrder
                End If
            End If
        Else
            gdSelectColor1.Visible = False
        End If
    End With

End Sub

Private Sub fgSpread_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim i&
    
    i = AddBlankRow
    If i > 0 Then
        With fgSpread
            If Col = 0 Then
                If Len(.TextMatrix(.Row, 0)) > 0 And Len(.TextMatrix(.Row, 1)) = 0 Then
                    fgSpread_CellButtonClick .Row, 1
                End If
            ElseIf Col = 1 Then
                If Len(.TextMatrix(.Row, 2)) = 0 Then
                    'new row just got added: default multiplier & contracts to 1
                    .TextMatrix(.Row, 2) = "1"
                    .TextMatrix(.Row, 3) = "1"
                    .TextMatrix(.Row, 4) = "1"             'save initial multiplier to hidden column
                    AddBlankRow
                    .Col = 2
                    .EditCell
                End If
            End If
        End With
    ElseIf Col = 2 Then
        If chkAutoMultiplier.Value = 0 Then
            'save user-entered multiplier to hidden column
            fgSpread.TextMatrix(fgSpread.Row, 4) = fgSpread.TextMatrix(fgSpread.Row, 2)
        End If
    End If
    
    With cboSpread
        If .ListIndex > 0 Then
            m.bSkipGridLoad = True
            .ListIndex = 0
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_AfterEdit"

End Sub

Private Sub fgSpread_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    With fgSpread
        Select Case .Rows
            Case 1, 2, 3
                .ColComboList(0) = kOpAll
            Case 4
                If .MergeRow(3) Then
                    .ColComboList(0) = kOpAll
                Else
                    .ColComboList(0) = kOpPlusMinus
                End If
            Case Else
                .ColComboList(0) = kOpPlusMinus
        End Select
        If Row = .FixedRows And Col = 0 Then Cancel = True
        If Col = 2 And chkAutoMultiplier.Value = 1 Then
            Cancel = True
        End If
    End With
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_BeforeEdit"

End Sub

Private Sub fgSpread_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim strText$, strNext$
    Dim nID&, dTickVM#
    Dim Bars As cGdBars
    
    If frmSymbolSelector.Visible Then Exit Sub
    
    With fgSpread
        If .Row > .FixedRows Then
            strText = .TextMatrix(.Row - 1, 1)
        Else
            strText = .TextMatrix(.Row, 1)
        End If
        strNext = GetNextContract(strText)
        If Len(strNext) > 0 Then strText = strNext
        ShowSymSelector strText
        
        If chkAutoMultiplier.Value = 1 Then
            Set Bars = New cGdBars
            strText = .TextMatrix(.Row, 1)
            nID = GetMarketInfo(strText, Bars)
            If nID > 0 And Bars.Prop(eBARS_TickValue) > 0 And Bars.Prop(eBARS_TickMove) > 0 Then
                dTickVM = Bars.Prop(eBARS_TickValue) / Bars.Prop(eBARS_TickMove)
            End If
            If nID > 0 And dTickVM > 0 Then
                .TextMatrix(.Row, 2) = Str(dTickVM)
            End If
        End If
    End With
        
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".fgSpread_CellButtonClick"

End Sub

Private Sub fgSpread_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error Resume Next:

    FinishEdit = True

End Sub

Private Sub fgSpread_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static nPrevMouseRow&

    With fgSpread
    
        If .Row >= .FixedRows And .Row < .Rows And .Row = .MouseRow Then
            If .MergeRow(.Row) Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                .MergeRow(.Row) = False
                If .Row = .FixedRows Then
                    fgSpread_CellButtonClick .FixedRows, 2
                    .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                Else
                    .Col = 0
                    .EditCell
                    SendKeys "{F4}"         'dropdown the combo list for (+, -, /) column
                End If
            ElseIf .MouseCol = 0 Then
                .EditCell
                SendKeys "{F4}"
            ElseIf .MouseCol = 1 Then
                'prevent symbol selector from coming up a second time from a double-click
                If .MouseRow <> nPrevMouseRow Then
                    fgSpread_CellButtonClick .Row, 2
                End If
            ElseIf .MouseCol = 2 Then
                .EditCell
            End If
        End If
    
        nPrevMouseRow = .MouseRow
    End With
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".fgSpread_MouseUp"

End Sub

Private Sub Form_Activate()
On Error Resume Next

    ' first time only: move focus to template (so can easily arrow down to Default or Standard template)
    If m.bFirstActivate Then
        m.bFirstActivate = False
        If m.eNewChartType = eTypeChart_Standard Then
            MoveFocus cboTemplate
        End If
    End If

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Dim aTemplates As cGdArray, i&
    Dim bIntraday As Boolean
    
    m.nFrameSpace = fraSettings.Top - (fraData.Top + fraData.Height)
    
    Me.Width = kFormWidth
    Me.Height = kFormHeight
    Me.Icon = Picture16(ToolbarIcon("ID_Chart"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    fraTrades.Left = fraSettings.Left
    fraSeasonal.Left = fraSettings.Left
    fraButtons.Left = Me.Width / 2 - fraButtons.Width / 2
    
    bIntraday = HasModule("IT") Or HasModule("FT") Or HasModule("ST")
    
    InitSeaonalComboCtrl cboCycle
    InitSeasonalBartypeCombo cboBarType
    
    With cboBarPeriod
        If bIntraday Then
            .AddItem "5 minute"
            .AddItem "10 minute"
            .AddItem "15 minute"
            .AddItem "30 minute"
            .AddItem "60 minute"
        End If
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Quarterly"
        .AddItem "Yearly"
        .AddItem "2 days"
        .AddItem "3 days"
        .AddItem "4 days"
        .Text = "Daily"
    End With
        
    cboTemplate.Clear
    Set aTemplates = GetAllowedList("T", True)
    If Not aTemplates Is Nothing Then
        If Not ActiveChart Is Nothing Then
            If Not ActiveChart.Chart Is Nothing Then
                If ActiveChart.Chart.SeasonalCycleTypeEnum = eCycleType_Undefined Then
                    aTemplates.Add kActiveChart, 0
                End If
            End If
        End If
        
        For i = 0 To aTemplates.Size - 1
            cboTemplate.AddItem Parse(aTemplates(i), vbTab, 1)
        Next
        cboTemplate.ListIndex = 0
    End If
    
    'show/hide spread chart buttons based on version
    If ExtremeCharts = 1 Then
        optSpread.Visible = False
        cboSpread.Visible = False
        optStandard.Visible = False
        lblSymbol.Visible = True
        
        lblSymbol.Left = optStandard.Left
        lblSymbol.Top = kTopOfSymRadio + 140
        txtSymbol.Top = kTopOfSymText + 140
        cmdSelectSym.Top = txtSymbol.Top
    Else
        optSpread.Visible = True
        cboSpread.Visible = True
        optStandard.Visible = True
        lblSymbol.Visible = False
        
        txtSymbol.Top = kTopOfSymText
        cmdSelectSym.Top = txtSymbol.Top
    End If
    
    fgPercentComp.Top = chkAutoMultiplier.Top
    
    InitGrid
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".Form_Load"
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    gdSelectColor1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.oFunctionTree = Nothing
    Set m.Chart = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.Form_Unload"
    
End Sub

Private Sub gdSelectColor1_Changed()
On Error Resume Next
    
    CheckColorSelect        'JM 06-11-2010 this event fires in the compiled EXE, but not in the IDE

End Sub

Private Sub gdSelectColor1_DropDown()
On Error Resume Next
    
    m.lMouseCol = 2
    m.bColorChecked = False
    tmr.Enabled = True

End Sub

Private Sub lblTemplate_DblClick()
    ' just a quick way to select the Default template
    On Error Resume Next
    If cboTemplate.ListCount > 1 Then
        cboTemplate.ListIndex = 1
    End If
End Sub

Private Sub optEndOfData_Click()

    If Not Me.Visible Then Exit Sub

    EnableProfileIntraday optIntraday.Value
    If dtDateTo.Value <> m.dtEODLastDate Then
        dtDateTo.Value = m.dtEODLastDate
        SyncProfileDateTime -1
    End If
End Sub

Private Sub optIntraday_Click()

    If Not Me.Visible Then Exit Sub

    EnableProfileIntraday True
    If m.nNumProfilesIntraday <= 0 Then m.nNumProfilesIntraday = 10
    SyncProfileDateTime m.nNumProfilesIntraday

End Sub

Private Sub optMultiDays_Click()

    If Not Me.Visible Then Exit Sub
    
    EnableProfileIntraday False
    If m.nNumProfilesMultiday <= 0 Then m.nNumProfilesMultiday = 20
    SyncProfileDateTime m.nNumProfilesMultiday

End Sub

Private Sub optPercentComp_Click()
On Error GoTo ErrSection:
    
    FixControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".optPercentComp_Click"
        
End Sub

Private Sub optProfile_Click()
On Error GoTo ErrSection:

    If Not HasModule("TPRO") Then
'        InfBox "Trade Profile charts require the TPRO enablement.", "i", , "Upgrade Required"
        optStandard.Value = True
        Exit Sub
    End If

    FixControls
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".optProfile_Click"
        
End Sub

Private Sub optSeasonal_Click()
On Error GoTo ErrSection:

    FixControls
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".optSeasonal_Click"
        
End Sub

Private Sub optSpread_Click()
On Error GoTo ErrSection:

    If Not HasGold(True, , False) Then
        optStandard.Value = True
    Else
        FixControls
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".optSpread_Click"
        
End Sub

Private Sub optStandard_Click()
On Error GoTo ErrSection:

    FixControls
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".optStandard_Click"
        
End Sub

Private Sub optToDate_Click()
    EnableProfileIntraday optIntraday.Value
End Sub

Private Sub optTradesAccount_Click()
On Error Resume Next
    m.nShowTrades = 2
    FixTradesControls
End Sub

Private Sub optTradesNone_Click()
On Error Resume Next
    m.nShowTrades = 0
    FixTradesControls
End Sub

Private Sub optTradesStrategy_Click()
On Error Resume Next
    m.nShowTrades = 1
    FixTradesControls
End Sub

Private Sub tmr_Timer()
On Error Resume Next

    Static FocusCheck As Boolean
    Dim i&

    If gdSelectColor1.Visible Then
        If Not gdSelectColor1.DropDownVisible Then
            If m.bColorChecked Then
                gdSelectColor1.Visible = False
                tmr.Enabled = False
            Else
                CheckColorSelect
            End If
        End If
    ElseIf vsTabProfile.Visible Then
        i = Int(ValOfText(txtNumProfiles.Text))
        
        If dtDateFrom.Value <> m.dtDateFrom Or dtDateTo.Value <> m.dtDtDateTo Then
            SyncProfileDateTime -1
        ElseIf optMultiDays.Value = True And i <> m.nNumProfilesMultiday Then
            SyncProfileDateTime i
        End If
        
        If vsTabProfile.CurrTab = 1 Then
            If FocusCheck Then
                If cboTicksPerRow.ListIndex = 0 And m.nTicksPerRow = -1 Then
                    vsTabProfile.SetFocus
                Else
                    i = Int(ValOfText(cboTicksPerRow.Text))
                    If i = m.nTicksPerRow Then
                        vsTabProfile.SetFocus
                    End If
                End If
                FocusCheck = False
            ElseIf cboTicksPerRow.ListIndex = 0 Then
                If m.nTicksPerRow <> -1 Then
                    m.nTicksPerRow = -1
                    FocusCheck = True
                End If
            Else
                i = Int(ValOfText(cboTicksPerRow.Text))
                If i > 0 And m.nTicksPerRow <> i Then
                    m.nTicksPerRow = i
                    FocusCheck = True
                End If
            End If
        End If
        
    Else
        tmr.Enabled = False
    End If

End Sub

Private Sub txtNumProfiles_Change()
    
    If Not Me.Visible Then Exit Sub
    
    SyncProfileDateTime Int(ValOfText(txtNumProfiles.Text))

End Sub

Private Sub txtSessionsPerProfile_LostFocus()
    m.nSessionsPerProfile = Int(ValOfText(txtSessionsPerProfile.Text))
End Sub

Private Sub txtSymbol_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    ShowSymSelector UCase(Chr(KeyCode))
    
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".txtSymbol_KeyDown"
        
End Sub

Private Sub ShowSymSelector(ByVal strChar)
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) back from the symbol selector
    Dim strNewSym$
    
    If Len(strChar) = 0 Then
        Set astrSymbols = frmSymbolSelector.ShowMe("", False)
    Else
        Set astrSymbols = frmSymbolSelector.ShowMe(strChar, False, , , False, True)
    End If
    
    If astrSymbols.Size > 0 Then
        If m.eNewChartType = eTypeChart_Spread Then
            fgSpread.TextMatrix(fgSpread.Row, 1) = astrSymbols(0)
            fgSpread_AfterEdit fgSpread.Row, 1
        Else
            strNewSym = astrSymbols(0)
            txtSymbol.Text = strNewSym
            m.strSymbol = strNewSym
        End If
    End If
    
    Set astrSymbols = Nothing
    
    If Left(m.strSymbol, 1) = "$" Or IsForex(m.strSymbol) Or Not HasModule("TPRO") Then          '6842
        optProfile.Enabled = False
    Else
        optProfile.Enabled = True
        SetProfileControls
    End If
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".ShowSymSelector"
        
End Sub

Public Sub ShowMe(ByVal strSymbol$, _
    Optional ByVal bEditSpread As Boolean = False, _
    Optional ByVal Chart As cChart = Nothing, _
    Optional ByVal bEditProfile As Boolean = False)
On Error GoTo ErrSection:

    Dim aSymbols As New cGdArray, i&

    m.bSkipGridLoad = False
    m.bEditSpread = bEditSpread
    m.bEditProfile = bEditProfile
    m.strPrevSym = ""
    
    If bEditSpread And Chart Is Nothing Then Exit Sub
    
    If bEditProfile Then
        If Chart Is Nothing Then Exit Sub
        If Chart.Tree Is Nothing Then Exit Sub
        If Chart.Bars Is Nothing Then Exit Sub
    End If
    
    If m.bEditSpread Then
        Me.Caption = "Chart Data"
        optSpread = True    'this should trigger load grid via click event
        txtSymbol = ""
    ElseIf m.bEditProfile Then
        Me.Caption = "Settings for Profile Chart"
        Me.optProfile = True
        fraData.Visible = False
    Else
        Me.Caption = "Settings for New Chart"
        If InStr(strSymbol, ";") Then
            'extract first symbol from symbol string (format: funcName~funcId;op,sym,mult)
            strSymbol = Parse(strSymbol, ";", 2)
            strSymbol = Parse(strSymbol, ",", 2)
        End If
        Set aSymbols = frmSymbolSelector.ShowMe(strSymbol, False, , "Symbol for New Chart", True)
        strSymbol = aSymbols(0)
        If Len(strSymbol) = 0 Then Exit Sub
        optStandard = True
        txtSymbol = strSymbol
        
        If Not Chart Is Nothing Then
            m.nAccountID = Chart.TradeAccountID
            m.nShowTrades = Chart.ShowTrades
        ElseIf Not ActiveChart Is Nothing Then
            m.nAccountID = ActiveChart.Chart.TradeAccountID
            m.nShowTrades = ActiveChart.Chart.ShowTrades
            m.nStrategyId = ActiveChart.Chart.SystemID
        Else
            m.nAccountID = 0
            m.nShowTrades = 0
            m.nStrategyId = -1
        End If
        PopulateAccountsCbo cboAccounts, m.nAccountID
        If m.nShowTrades = 1 Then
            optTradesStrategy.Value = True
            If m.nStrategyId > 0 Then
                txtStrategy.Text = SystemNameForID(m.nStrategyId)
            End If
        ElseIf m.nShowTrades = 2 Then
            optTradesAccount.Value = True
        Else
            optTradesNone.Value = True
        End If
    End If

    m.strSymbol = strSymbol
    Set m.Chart = Chart
    m.nPeriodicity = 0
    If Not m.Chart Is Nothing Then
        m.nPeriodicity = m.Chart.Bars.Prop(eBARS_Periodicity)
    ElseIf ActiveChart Is Nothing Then
        m.nPeriodicity = GetPeriodicity("Daily")
    Else
        m.nPeriodicity = ActiveChart.Chart.Bars.Prop(eBARS_Periodicity)
    End If
    
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
    
    'color scheme
    cboColorScheme.Clear
    cboColorScheme.AddItem "Gradient"
    cboColorScheme.AddItem "Rainbow"
    cboColorScheme.AddItem "Up/Down"
'    cboColorScheme.AddItem "Buy/Sell Volume"
    cboColorScheme.ListIndex = MktProf_Color_Gradient
    
    If Left(m.strSymbol, 1) = "$" Or IsForex(m.strSymbol) Or Not HasModule("TPRO") Then          '6842
        optProfile.Enabled = False
    Else
        LoadProfileSettingsLastUsed
        SetProfileControls
    End If
    
    m.bInit = True
    
    LoadFunctions
    LoadCboSpread   'this will trigger the click event which will load the spread grid
        
    'save original multiplier to hidden field
    With fgSpread
        For i = .FixedRows To .Rows - 2
            .TextMatrix(i, 4) = .TextMatrix(i, 2)
        Next
    End With
    
    'need to do this last to make sure grid has been populated
    If m.bEditSpread Then
        chkAutoMultiplier.Value = Abs(Chart.SpreadAsDollar)
    Else
        chkAutoMultiplier.Value = 0
    End If
    
    m.bInit = False
    
    m.bFirstActivate = True
    ShowForm Me, eForm_Modal
    Set m.Chart = Nothing
    
    SyncProfileDateTime m.nNumProfilesIntraday
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".ShowMe"
        
End Sub

Private Sub FixControls()
On Error GoTo ErrSection:

    cboBarPeriod.Text = GetPeriodStr(m.nPeriodicity)
    
    If optProfile Then
        Me.Width = kFormWidth + 690
        vsTabProfile.Visible = True
        tmr.Enabled = True
    Else
        Me.Width = kFormWidth
        vsTabProfile.Visible = False
        tmr.Enabled = False
    End If
    
    If optSpread Then
        m.eNewChartType = eTypeChart_Spread
        cboSpread.Enabled = True
        txtSymbol.Enabled = False
        cmdSelectSym.Enabled = False
        
        chkAutoMultiplier.Visible = True
        fgSpread.Visible = True
        fgPercentComp.Visible = False
        
        cmdFgDelete.Top = fgSpread.Top
        cmdFgDelete.Visible = True
        fraData.Height = kFrameDataHeight
        If optTradesAccount.Value = True Then
            optTradesNone.Value = True
        End If
        
        'fraSymbols.Top = 1770
        fraSymbols.Visible = True
        
        cboTemplate.Enabled = True
        chkCopyAnnots.Visible = True
        chkCopyAnnots.Enabled = True
    
        cmdSaveAs.Visible = True
        If cboSpread.ListIndex = 0 Then
            cmdSaveAs.Enabled = True
        Else
            cmdSaveAs.Enabled = False
        End If
    
    ElseIf optPercentComp Then
        m.eNewChartType = eTypeChart_PercentComp
        cboSpread.Enabled = False
        txtSymbol.Enabled = False
        cmdSelectSym.Enabled = False
        
        fraData.Height = kFrameDataHeight
        
        'fraSymbols.Top = 1770
        fraSymbols.Visible = True
        
        chkAutoMultiplier.Visible = False
        fgSpread.Visible = False
        fgPercentComp.Visible = True
        
        cmdFgDelete.Top = fgPercentComp.Top
        cmdFgDelete.Visible = True
        If optTradesAccount.Value = True Then
            optTradesNone.Value = True
        End If
        vsTabProfile.Visible = False
        
        cboTemplate.Enabled = False
        chkCopyAnnots.Visible = False
        chkCopyAnnots.Enabled = False
        
        cmdSaveAs.Visible = False
        
        If Me.Visible Then
            If m.strSymbol <> m.strPrevSym Then
                InitPercentCompGrid fgPercentComp, Nothing, kGridHeight + chkAutoMultiplier.Height - 15, m.strSymbol
                m.strPrevSym = m.strSymbol
            End If
        End If

    ElseIf optSeasonal Then
        m.eNewChartType = eTypeChart_Seasonal
        cboSpread.Enabled = False
        txtSymbol.Enabled = True
        cmdSelectSym.Enabled = True
        
        fraSymbols.Visible = False
        fgSpread.Visible = False
        cmdFgDelete.Visible = False
        vsTabProfile.Visible = False
        
        fraData.Height = kFrameDataHeight - fraSymbols.Height
        
        cboTemplate.Enabled = False
        chkCopyAnnots.Visible = False
        chkCopyAnnots.Enabled = False
        
        cmdSaveAs.Visible = False
    ElseIf optProfile Then
        m.eNewChartType = eTypeChart_Profile
        
        cboSpread.Enabled = False
        txtSymbol.Enabled = True
        cmdSelectSym.Enabled = True
        
        fraTrades.Visible = False
        fraSeasonal.Visible = False
        fraSymbols.Visible = False
        fraSettings.Visible = False
        fgSpread.Visible = False
        cmdFgDelete.Visible = False
        vsTabProfile.Visible = True
        
        fraData.Height = kFrameDataHeight - fraSymbols.Height
        
        If m.bEditProfile Then
            vsTabProfile.Top = fraData.Top
            vsTabProfile.CurrTab = 1
        Else
            vsTabProfile.Top = fraData.Top + fraData.Height + 150
            vsTabProfile.CurrTab = 0
        End If
        vsTabProfile.Left = fraData.Left
    
        cboTemplate.Enabled = False
        chkCopyAnnots.Visible = False
        chkCopyAnnots.Enabled = False
        
        EnableProfileIntraday optIntraday.Value
    Else
        m.eNewChartType = eTypeChart_Standard
        cboSpread.Enabled = False
        txtSymbol.Enabled = True
        cmdSelectSym.Enabled = True
        
        fraSymbols.Visible = False
        fgSpread.Visible = False
        cmdFgDelete.Visible = False
        vsTabProfile.Visible = False
        
        fraData.Height = kFrameDataHeight - fraSymbols.Height
        vsTabProfile.Top = fraData.Top + fraData.Height + 150
        
        cboTemplate.Enabled = True
        
        If cboTemplate.Text = kActiveChart Then
            chkCopyAnnots.Visible = True
            chkCopyAnnots.Enabled = True
        Else
            chkCopyAnnots.Visible = False
            chkCopyAnnots.Enabled = False
        End If
    
        If m.nShowTrades = 1 Then
            optTradesStrategy.Value = True
            If m.nStrategyId > 0 Then
                txtStrategy.Text = SystemNameForID(m.nStrategyId)
            End If
        ElseIf m.nShowTrades = 2 Then
            optTradesAccount.Value = True
        Else
            optTradesNone.Value = True
        End If
    
    End If
    
    fraSettings.Top = fraData.Top + fraData.Height + m.nFrameSpace
    If m.bEditSpread Then
        optStandard.Enabled = False
        fraSettings.Visible = False
        fraButtons.Top = fraData.Top + fraData.Height
        fraTrades.Visible = False
        fraSeasonal.Visible = False
    ElseIf m.eNewChartType = eTypeChart_Seasonal Then
        optStandard.Enabled = True
        fraSettings.Visible = False
        fraTrades.Visible = False
        
        fraSeasonal.Visible = True
        fraSeasonal.Move fraSettings.Left, fraSettings.Top
        fraButtons.Top = fraSeasonal.Top + fraSeasonal.Height
    ElseIf m.eNewChartType = eTypeChart_Profile Then
        fraButtons.Top = vsTabProfile.Top + vsTabProfile.Height
    Else
        optStandard.Enabled = True
        fraSettings.Visible = True
        fraSeasonal.Visible = False
        If m.eNewChartType = eTypeChart_PercentComp Or m.eNewChartType = eTypeChart_Spread _
           Or m.eNewChartType = eTypeChart_Seasonal Or Not HasGold(False, , False) Then
            fraTrades.Visible = False
            fraButtons.Top = fraSettings.Top + fraSettings.Height
        Else
            'only show options for trades/orders if Gold and standard chart
            fraTrades.Visible = True
            fraTrades.Top = fraSettings.Top + fraSettings.Height + m.nFrameSpace
            fraButtons.Top = fraTrades.Top + fraTrades.Height
        End If
    End If
    Me.Height = fraButtons.Top + fraButtons.Height + (Me.Height - Me.ScaleHeight)
    
    
'    If cboTemplate.Text = "< Copy settings from active chart >" Then
'        chkCopyAnnots.Visible = True
'    Else
'        chkCopyAnnots.Visible = False
'        chkCopyAnnots.Value = 0
'    End If

    Exit Sub

ErrSection:
    RaiseError Me.Name & ".FixControls"
        
End Sub

Private Sub NewPercentChart()
On Error GoTo ErrSection:

    Dim aNewSymbols As New cGdArray
    
    ParsePercentCompGrid fgPercentComp, aNewSymbols, m.strSymbol
    PercentChangeChartNew Nothing, aNewSymbols, m.strSymbol, m.nPeriodicity
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.NewPercentChart"

End Sub

Private Sub NewStandardChart()
On Error GoTo ErrSection:

    Dim i&, strFile$
    Dim aFiles As New cGdArray
    Dim frm As frmChart
    Dim frmActive As Form

    Set frmActive = ActiveChart()
    If Len(m.strSymbol) > 0 Then
        Set frm = New frmChart
        If cboTemplate <> kActiveChart Then
            frm.Chart.TemplateApply cboTemplate
        Else
            frm.Chart.TemplateLoad "", False
            If chkCopyAnnots And Not ActiveChart Is Nothing Then
                'ActiveChart.Chart.TemplateSave
                aFiles.GetMatchingFiles g.ChartGlobals.strCPCRoot & "\Charts\" & ActiveChart.Chart.Template & "^*.ano", False
                For i = 0 To aFiles.Size - 1
                    strFile = frm.Chart.Template & "^" & Parse(aFiles(i), "^", 2)
                    FileCopy g.ChartGlobals.strCPCRoot & "\Charts\" & aFiles(i), g.ChartGlobals.strCPCRoot & "\Charts\" & strFile
                Next
            End If
        End If
        frm.Chart.SetSymbol m.strSymbol
        frm.Chart.ChangeBarPeriod m.nPeriodicity, False
        frm.Chart.ShowTrades = m.nShowTrades
        frm.Chart.ShowToolbar = 0
        If m.nShowTrades = 1 Then
            frm.Chart.SystemID = m.nStrategyId
        ElseIf m.nShowTrades = 2 Then
            frm.Chart.TradeAccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
        End If
        
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            If g.ChartGlobals.frmActiveNonDetached.WindowState = vbNormal Then
                '6741 - SetNormalPlacement is a misnomer, it is a Property Let functionality
                '   does not position or size the form at all
                frm.SetNormalPlacement frmActive.GetNormalPlacement
                
                'SetRatioPlacement positions & sizes the form based on passed in string.
                'If the autosize option is off, this function knows to use the form's normal
                'placement string instead (hence the normal placement string is set right above)
                frm.SetRatioPlacement frmActive.GetRatioPlacement, True
            Else
                frm.WindowState = vbMaximized       '5203
            End If
        End If
        frm.Chart.TemplateSave      '5007
        frm.SkipFocusFix = True
        ActiveChartFormSet frm
        ShowForm frm
        MoveFocus frm.pbChart
        
        Set frm = Nothing
    End If
    
    Exit Sub

ErrSection:
    RaiseError Me.Name & ".NewStandardChart"
        
End Sub

Private Function NewSpreadChart() As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim strOp$, strSym$, strMult$, strContract$, strFunction$
    Dim aSymStrings As New cGdArray
    Dim frm As frmChart
    Dim Ind As cIndicator
    Dim oFunction As cFunction
    
    Dim bMaximized As Boolean

    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        If g.ChartGlobals.frmActiveNonDetached.WindowState = vbMaximized Then
            bMaximized = True
        End If
    End If
    
    'build array of symbol strings
    For i = fgSpread.FixedRows To fgSpread.Rows - 1
        With fgSpread
            If .MergeRow(i) Then
                strOp = ""
                strSym = ""
                strMult = ""
                strContract = ""
            Else
                strOp = .TextMatrix(i, 0)
                strSym = .TextMatrix(i, 1)
                strMult = .TextMatrix(i, 2)
                strContract = .TextMatrix(i, 3)
            End If
        End With
        If Len(strOp) = 0 Or strOp = kPlus Then
            strOp = "+"
        ElseIf strOp = kMinus Then
            strOp = "-"
        ElseIf strOp = kDivide Then
            strOp = "/"
        Else
            aSymStrings.Size = 0       'precaustionary: should never happen
            InfBox "Invalid operator: " & strOp, "!", , "New Chart"
            Exit For
        End If
        
        If Len(strOp) > 0 And Len(strSym) > 0 And Len(strMult) > 0 And Len(strContract) > 0 Then
            aSymStrings.Add strOp & "," & strSym & "," & strMult & "," & strContract
        End If
    Next
    
    If aSymStrings.Size = 0 Or aSymStrings.Size = 1 Then
        Set aSymStrings = Nothing
        InfBox "Spread charts require at least 2 symbols.", "I", , "New Chart"
        Exit Function
    End If
        
    'push all minuses to bottom
    aSymStrings.Sort
    
    Dim TempBars As cGdBars
    Dim strTemp As String
    Dim nTempID As Long
    Dim nSession As Long
    Dim bOK As Boolean
    
    'assume OK
    bOK = True
    If g.RealTime.Active Then
        'fix for aardvark 4025
        For i = 0 To aSymStrings.Size - 1
            Set TempBars = New cGdBars
            strTemp = Parse(aSymStrings(i), ",", 2)
            If Len(strTemp) > 0 Then
                nTempID = g.SymbolPool.SymbolIDforSymbol(strTemp)
                If nTempID > 0 Then
                    GetAvailTickData TempBars, nSession, strTemp, nTempID, 0, 0
                    g.RealTime.RefreshSymbolList
                Else
                    bOK = False
                    Exit For
                End If
            Else
                bOK = False
                Exit For
            End If
        Next
        While frmStatus.IsBusy
            DoEvents
        Wend
    End If
    
    If Not bOK Then
        MsgBox "Unable to retrieve real time data for spread symbols." & vbCrLf & "Try adding symbols to quoteboard first."
        NewSpreadChart = False
        Set aSymStrings = Nothing
        Set TempBars = Nothing
        Exit Function
    End If
    
    m.strSymbol = aSymStrings.JoinFields(";")
    Set aSymStrings = Nothing
    Set TempBars = Nothing
    
    If cboSpread.ItemData(cboSpread.ListIndex) > 0 Then
        strFunction = cboSpread.Text & "~" & cboSpread.ItemData(cboSpread.ListIndex)
    End If
            
    If Len(m.strSymbol) > 0 Then
        If m.bEditSpread And Not m.Chart Is Nothing Then
            m.Chart.SetSymbol strFunction & ";" & m.strSymbol, True
        Else
            strTemp = cboTemplate.Text
            Set frm = New frmChart
            If cboTemplate.Text = kActiveChart Then
                If Not ActiveChart Is Nothing Then
                    If Not ActiveChart.Chart Is Nothing Then
                        strTemp = ActiveChart.Chart.TemplateApplied
                    End If
                End If
            End If
            If Len(strTemp) > 0 Then
                If InStr(UCase(strTemp), "WOODIES") = 0 Then
                    frm.Chart.TemplateApply strTemp
                Else
                    frm.Chart.TemplateApply "Standard"
                End If
            End If
            frm.Chart.SetSymbol strFunction & ";" & m.strSymbol
            frm.Chart.ChangeBarPeriod m.nPeriodicity, False
            frm.Chart.ShowTrades = 0            '4310
            frm.Chart.ShowToolbar = 0
            Set Ind = frm.Chart.Tree("Price")
            If Not Ind Is Nothing Then
                Ind.DisplayType = eINDIC_Line
            End If
            frm.Chart.GenerateChart
            If InStr(UCase(strTemp), "WOODIES") <> 0 Then
                frm.Chart.TemplateApply strTemp, True
                frm.Chart.GenerateChart
            End If
            
            'JM 07-17-2015 - fix per Richard's spec
            '   if user chooses a template that will not work for spread chart TN crashes
            '   example: Richard's VolSpike.CHT template
            '   fix is to initially turn off all indicators except price
            For i = 1 To frm.Chart.Tree.Count
                If TypeOf frm.Chart.Tree(i) Is cIndicator Then
                    If UCase(frm.Chart.Tree(i).Name) <> "PRICE" Then frm.Chart.Tree(i).Display = False
                End If
            Next
            
            If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                frm.SetNormalPlacement g.ChartGlobals.frmActiveNonDetached.GetNormalPlacement   '6741
                frm.SetRatioPlacement g.ChartGlobals.frmActiveNonDetached.GetRatioPlacement, True
            End If
            
            frm.Chart.TemplateSave      '5007
            frm.SkipFocusFix = True
            ActiveChartFormSet frm
            
            If bMaximized Then          '5244
                bMaximized = LockWindowUpdate(frmMain.hWnd)
                frm.WindowState = vbNormal
                frm.WindowState = vbMaximized
                If bMaximized Then LockWindowUpdate 0
            End If
            
            ShowForm frm
            MoveFocus frm.pbChart
            
            
            Set frm = Nothing
        End If
        
        NewSpreadChart = True
    End If
    
    Exit Function

ErrSection:
    RaiseError Me.Name & ".NewSpreadChart"
        
End Function

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgSpread
        .Redraw = flexRDNone
        SetupGrid Me.fgSpread, eGridMode_Grid
        .ExplorerBar = flexExMove
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarVertical
        .FixedRows = 1
        .Rows = 1
        .Cols = 5
        'column headers
        .TextMatrix(0, 0) = "+ - /"
        .TextMatrix(0, 1) = "Symbol"
        .TextMatrix(0, 2) = "Multiplier"
        .TextMatrix(0, 3) = "#Contracts"
        .TextMatrix(0, 4) = "Multiplier User"       'hidden column that saves multiplier input by user
        'button & dropdown for columns
        .ColComboList(0) = kOpAll
        .ColComboList(1) = "..."
        'alignment
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
        'width
        .ColWidth(0) = 700          'plus/minus
        .ColWidth(1) = 1800         'symbol
        .ColWidth(2) = 820          'multiplier
        .ColWidth(3) = 800          'contracts
        .ColWidth(4) = 750          'hidden column that saves multiplier input by user
        'hidden columns
        .ColHidden(4) = True
        
        .Height = kGridHeight

    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError Me.Name & ".InitGrid"
        
End Sub

Private Sub LoadGridFromSymbol()

    Dim aData As New cGdArray
    Dim aFields As New cGdArray
    Dim bDivide As Boolean
    Dim strText$, strOp$, i&

    If Len(m.strSymbol) = 0 Then
        fgSpread.Rows = fgSpread.FixedRows
        Exit Sub
    End If
    
    If m.bEditSpread Then
        aData.SplitFields m.strSymbol, ";"
        If InStr(aData(0), ",") = 0 Then
            aData.Remove 0, 1 ' remove function name
        End If
    Else
        aData(0) = "+," & m.strSymbol & "," & "1" & "," & "1"
        strText = GetNextContract(m.strSymbol)
        If Len(strText) > 0 Then
            aData(1) = "-," & strText & "," & "1" & "," & "1"
        End If
        strText = ""
    End If
    
    If aData.Size < 0 Then Exit Sub
    
    'push minuses to the bottom
    aData.Sort
    
    With fgSpread
        .Rows = .FixedRows
        For i = 0 To aData.Size - 1
            .Rows = .Rows + 1
            strText = aData(i)
            aFields.SplitFields strText, ","
            If aFields(0) = "-" Then
                strOp = kMinus
            ElseIf aFields(0) = "/" Then
                strOp = kDivide
            Else
                strOp = kPlus
            End If
            .TextMatrix(.Rows - 1, 0) = strOp
            .TextMatrix(.Rows - 1, 1) = aFields(1)
            .TextMatrix(.Rows - 1, 2) = aFields(2)
            If aFields.Size > 3 Then
                .TextMatrix(.Rows - 1, 3) = aFields(3)
            Else
                .TextMatrix(.Rows - 1, 3) = "1"
            End If
        Next
        .TextMatrix(.FixedRows, 0) = ""
        .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
    End With

End Sub

Private Sub LoadGridFromTable(tbData As cGdTable, bDivide As Boolean)

    Dim i&
    Dim aTemp As New cGdArray

    If bDivide Then
        If tbData.NumRecords = 2 Then
            'ratio spreads can only have 2 symbols
            With fgSpread
                .Rows = .FixedRows + 2
                .TextMatrix(.FixedRows, 0) = ""
                .TextMatrix(.FixedRows, 1) = tbData(1, 0)
                .TextMatrix(.FixedRows, 2) = tbData(2, 0)
                .TextMatrix(.FixedRows, 3) = tbData(3, 0)
                .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
                .TextMatrix(.FixedRows + 1, 0) = tbData(0, 1)
                .TextMatrix(.FixedRows + 1, 1) = tbData(1, 1)
                .TextMatrix(.FixedRows + 1, 2) = tbData(2, 1)
                .TextMatrix(.FixedRows + 1, 2) = tbData(3, 1)
                
                'save original multiplier to hidden column
                .TextMatrix(.FixedRows, 4) = .TextMatrix(.FixedRows, 2)
                .TextMatrix(.FixedRows + 1, 4) = .TextMatrix(.FixedRows + 1, 2)
                
                .MergeRow(1) = False
                .MergeRow(2) = False
            End With
        End If
        
        Exit Sub
    End If
    
    aTemp.Clear
    Set aTemp = tbData.CreateSortedIndex(0, eGdSort_Descending Or eGdSort_Stable)
    If aTemp.Size < 1 Then
        Set tbData = Nothing
        Set aTemp = Nothing
        Exit Sub
    End If
    
    fgSpread.Rows = fgSpread.FixedRows
    For i = 0 To aTemp.Size - 1
        With fgSpread
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = tbData(0, aTemp(i))
            .TextMatrix(.Rows - 1, 1) = tbData(1, aTemp(i))
            .TextMatrix(.Rows - 1, 2) = tbData(2, aTemp(i))
            .TextMatrix(.Rows - 1, 3) = tbData(3, aTemp(i))
            'save original multiplier to hidden column
            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 2)
        End With
    Next
    
    With fgSpread
        If .Rows >= .FixedRows Then
            .TextMatrix(.FixedRows, 0) = ""
            .Cell(flexcpBackColor, .FixedRows, 0) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
        End If
    End With

    Set aTemp = Nothing

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim oFunction As cFunction
    Dim tbData As cGdTable
    Dim strText$, strOp$, i&
        
    Dim bDivide As Boolean
        
    m.bConvert = False
    With cboSpread
        i = .ItemData(.ListIndex)
    End With
    If i > 0 Then
        m.bSkipGridLoad = False
        Set oFunction = m.oFunctionTree(Str(i))
        
        If Not oFunction Is Nothing Then
            With oFunction
                Editor1.TextRTF = .GetRTF(.Formatted)
                strText = Editor1.Text
                If 0 = InStr(strText, "{") Then m.bConvert = True
                Set tbData = SpreadExprToTable(strText, bDivide)
            End With
            If Not tbData Is Nothing Then
                LoadGridFromTable tbData, bDivide
                If m.bConvert Then
                    SaveSpread
                    m.bConvert = False
                End If
            End If
        End If
    ElseIf Not m.bSkipGridLoad Then
        LoadGridFromSymbol
    End If
    
    If Not m.bInit Then
        strText = ""
        With fgSpread
            For i = .FixedRows To .Rows - 1
                strText = strText & ";" & .TextMatrix(i, 0) & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2)
            Next
        End With
        
        'set init flag to prevent checkbox click event from calling fix spread grid function
        m.bInit = True
        If IsDollarMultiplier(strText) Then
            chkAutoMultiplier.Value = 1
        Else
            chkAutoMultiplier.Value = 0
        End If
        m.bInit = False
        
    End If
    
    AddBlankRow
    FixControls
        
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.LoadGrid"
        
End Sub

Private Function AddBlankRow() As Long
On Error GoTo ErrSection:

    Dim strOp$, strSym$, strMult$, i&
    Dim bNewRow As Boolean
    
    'see if new blank row should be added
    bNewRow = True
    With fgSpread
        For i = .FixedRows To .Rows - 1
            If .MergeRow(i) = True Then
                bNewRow = False
                Exit For
            End If
            If i = .FixedRows Then
                strOp = kPlus
            Else
                strOp = .TextMatrix(i, 0)
            End If
            strSym = .TextMatrix(i, 1)
            strMult = .TextMatrix(i, 2)
            If Len(strOp) = 0 Or Len(strSym) = 0 Or Len(strMult) = 0 Or strOp = kDivide Then
                bNewRow = False
                AddBlankRow = i
                Exit For
            End If
        Next
    End With
    
    If bNewRow Then
        With fgSpread
            .Rows = .Rows + 1
            If .Rows - 1 = .FixedRows Then
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to get started ..."
            Else
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to add another row ..."
            End If
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
        End With
    End If

    Exit Function
    
ErrSection:
    RaiseError Me.Name & ".AddBlankRow"
        
End Function

Private Sub LoadFunctions()
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim oFunction As cFunction
    
    m.oFunctionTree.Clear
    
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblLibrarys.* " & _
                "FROM tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
                "WHERE ( (tblLibrarys.Ignore)=0 AND [FunctionCategoryID] = 27 );", dbOpenDynaset)
    
    Do While Not rs.EOF
        Set oFunction = New cFunction
        
        oFunction.FunctionID = rs!FunctionID
        oFunction.Load
        m.oFunctionTree.Add oFunction, Str(oFunction.FunctionID)
        
        rs.MoveNext
    Loop
    
    Set oFunction = Nothing
    
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.LoadFunctions"

End Sub

Private Sub LoadCboSpread()
On Error GoTo ErrSection:

    Dim i&, nID&, nIdx&, strName$
    Dim aIdx As cGdArray
    Dim tbData As New cGdTable
    Dim oFunction As cFunction

    With cboSpread
        .AddItem kCustomSpread
        .ItemData(.ListCount - 1) = 0
        .ListIndex = 0
    End With
    
    If m.oFunctionTree.Count = 0 Then Exit Sub
    
    'create table fields
    tbData.CreateField eGDARRAY_Strings, 0, "FunctionName"
    tbData.CreateField eGDARRAY_Longs, 1, "FunctionID"
    
    If m.bEditSpread Then
        strName = Parse(m.strSymbol, ";", 1)
        If InStr(strName, "~") <> 0 Then
            nID = Parse(strName, "~", 2)
        End If
    End If
    
    For i = 0 To m.oFunctionTree.Count
        Set oFunction = m.oFunctionTree(i)
        If Not oFunction Is Nothing Then
            tbData.AddRecord ""
            tbData(0, tbData.NumRecords - 1) = oFunction.FunctionName
            tbData(1, tbData.NumRecords - 1) = oFunction.FunctionID
        End If
        Set oFunction = Nothing
    Next
    
    'sort by name
    Set aIdx = tbData.CreateSortedIndex(0)
        
    For i = 0 To aIdx.Size - 1
        With cboSpread
            .AddItem tbData(0, aIdx(i))
            .ItemData(.ListCount - 1) = tbData(1, aIdx(i))
            If nID = tbData(1, aIdx(i)) Then
                nIdx = i + 1        'the first item is the < Custom > label
            End If
        End With
    Next
    
    cboSpread.ListIndex = nIdx
    
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.LoadCboSpread"
    
End Sub

Private Sub SaveSpread()
On Error GoTo ErrSection:
    
    Dim i&, strName$, strNewName$, strText$
    Dim oFunction As cFunction

    With cboSpread
        i = .ItemData(.ListIndex)
        If i > 0 Then
            strName = Trim(.Text)
        Else
            strName = "New Spread Name"
        End If
    End With
    
    If m.bConvert Then
        If i > 0 Then
            Set oFunction = m.oFunctionTree(Str(i))
            If Not oFunction Is Nothing Then
                SaveFunction oFunction, oFunction.FunctionName, False
            End If
        End If
        Exit Sub
    End If

    'prompt for name
    strNewName = strName
    Do While Len(strNewName) > 0
        strText = "Save the current Spread function as..."
        strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & strName & " ; " & strText)
        ' Strip out a colon if it exists in the name...
        If InStr(strNewName, ":") Then strNewName = Replace(strNewName, ":", "")
        strNewName = Trim(strNewName)
        
        If Len(strNewName) = 0 Then
            Exit Do         'user must have cancelled or just gave blank name (ignore request)
        Else
            If strNewName = strName Then
                strText = "Overwrite the current Spread function?"
                If "O" = InfBox(strText, "?", "+Overwrite|Cancel", strName) Then
                    If i > 0 Then
                        Set oFunction = m.oFunctionTree(Str(i))
                        If Not oFunction Is Nothing Then
                            SaveFunction oFunction, oFunction.FunctionName, False
                        End If
                    End If
                    strNewName = ""
                    Exit Do
                End If
            ElseIf DuplicateFuncName(strNewName) Then
                InfBox "'" & strNewName & "' already exists.", "e", , "Error"
            Else
                Set oFunction = New cFunction
                strText = oFunction.ValidName(strNewName)
                If strText <> "" Then
                    InfBox strText, "e", , "Error"
                Else
                    Exit Do
                End If
            End If
        End If
    Loop
    
    Set oFunction = Nothing
    If Len(strNewName) > 0 Then
        Set oFunction = New cFunction
        SaveFunction oFunction, strNewName, True
        With cboSpread
            .AddItem oFunction.FunctionName
            .ItemData(.ListCount - 1) = oFunction.FunctionID
            .ListIndex = .ListCount - 1
        End With
    End If
    
    Set oFunction = Nothing
    
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.SaveSpread"
    
End Sub

Private Sub SaveFunction(oFunction As cFunction, ByVal strName$, ByVal bNew As Boolean)
On Error GoTo ErrSection:
    
    Dim Expr As cExpression
    Dim i&, strExpr$
    
    strExpr = BuildSpreadExpr(fgSpread)
    If Len(strExpr) = 0 Then Exit Sub
    
    Set Expr = New cExpression
    
    'Verify...
    Editor1.Text = strExpr
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule Editor1.Text
        ' Save verify settings...
        oFunction.FunctionIDs = .GetFIDs
        oFunction.Formatted = .EditText
        oFunction.FormattedWithFillWords = .Preview
        oFunction.CodedText = .CodedText
        oFunction.DataTypeID = .FunctionReturnType
        oFunction.ReturnTypeID = .FunctionReturnType
        oFunction.LateCalculating = .LateCondition
    End With

    With oFunction
        .Inputs = Expr.Inputs
        
        If bNew Then
            .SecurityLevel = 0
            .CannotDelete = False
            .LibraryID = kSN_UserLibrary
            .Password = ""
            .FunctionCategoryID = 27
            .Usage = 14
        End If
        
        ' Set values specific to BOTH builtin and user functions
        If Len(strName) > 0 Then .FunctionName = strName
        .CodedName = StripStr(.FunctionName, " ")
        .ImplementationTypeID = kSN_Custom
        
        ' Function rule...
        .TradeSenseUsage = .CodedName
        .Reverify = False
        .LastModified = Now
        .Save
    End With

    g.bDirtyLibrariesMDB = True
    RefreshFunction oFunction
    RefreshReverify oFunction.FunctionID
    
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.SaveFunction"

End Sub

Private Function DuplicateFuncName(ByVal strFuncName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim QryDef As QueryDef
        
    Set QryDef = g.dbNav.QueryDefs("qryFunctionIDFromName")
    QryDef.Parameters(0).Value = strFuncName
    Set rs = QryDef.OpenRecordset
    
    DuplicateFuncName = rs.RecordCount
    
    Set rs = Nothing
    Set QryDef = Nothing
    
    Exit Function

ErrSection:
    RaiseError "frmNewChart.DuplicateFuncName"

End Function

Private Sub FixTradesControls()
On Error GoTo ErrSection:

    If optTradesStrategy.Value Then
        txtStrategy.Visible = True
        cmdSelectStrategy.Visible = True
        cboAccounts.Visible = False
    Else
        txtStrategy.Visible = False
        cmdSelectStrategy.Visible = False
        cboAccounts.Visible = True
        cboAccounts.Enabled = optTradesAccount.Value
    End If

#If 0 Then
    If optTradesStrategy.Value = True Then
        txtStrategy.Enabled = True
        cmdSelectStrategy.Enabled = True
        cboAccounts.Enabled = False
    ElseIf optTradesAccount.Value = True Then
        txtStrategy.Enabled = False
        cmdSelectStrategy.Enabled = False
        cboAccounts.Enabled = True
    Else
        txtStrategy.Enabled = False
        cmdSelectStrategy.Enabled = False
        cboAccounts.Enabled = False
    End If
#End If

    Exit Sub

ErrSection:
    RaiseError "frmNewChart.FixTradesControls"

End Sub

Private Sub FixSpreadGrid()
On Error GoTo ErrSection:

    Dim bError As Boolean
    
    If Not optSpread Or fgSpread.Rows < 3 Or m.bInit Then
        Exit Sub
    End If
    
    bError = ToggleAutoMultiplier(fgSpread, chkAutoMultiplier.Value)
    
    If bError Then
        m.bInit = True
        chkAutoMultiplier.Value = 0
        m.bInit = False
    End If
    
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.FixSpreadGrid"
    
End Sub

Private Sub CheckColorSelect()
On Error GoTo ErrSection
    
    Dim iColor As Long
    
    If gdSelectColor1.Visible Then
        iColor = gdSelectColor1.Color
        If iColor = 0 Then iColor = -1
        With fgPercentComp
            If m.lMouseCol = 2 Then
                If m.lMouseRow >= .FixedRows And m.lMouseRow < .Rows Then
                    .Cell(flexcpBackColor, m.lMouseRow, 2) = iColor
                    .Select m.lMouseRow, m.lMouseCol
                    .Refresh
                    m.bColorChecked = True
                End If
            End If
        End With
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.CheckColorSelect"
    
End Sub

Private Sub SyncProfileDateTime(ByVal nNumProfiles&)
On Error GoTo ErrExit:

    Dim i&, n&, dtDate#
    
    n = nNumProfiles
    If n = -1 And dtDateFrom.Value >= dtDateTo.Value Then
        MsgBox "Date from must be earlier than date to.", vbOKOnly
        dtDateFrom.Value = m.dtDateFrom
        dtDateTo.Value = m.dtDtDateTo
    End If
    
    If n = -1 Then
        If dtDateFrom.Value <> m.dtDateFrom Or dtDateTo.Value <> m.dtDtDateTo Then
            'set number of profiles based on date values
            If 0 = gdIsHoliday(dtDateTo, "") Then i = 1
            dtDate = dtDateTo.Value - 1
            While dtDate >= dtDateFrom.Value
                If IsWeekday(dtDate) And 0 = gdIsHoliday(dtDate, "") Then
                    i = i + 1
                End If
                dtDate = dtDate - 1
            Wend
            txtNumProfiles.Text = Val(i)
            If optIntraday Then
                m.nNumProfilesIntraday = i
            Else
                m.nNumProfilesMultiday = i
            End If
            m.dtDateFrom = dtDateFrom.Value
            m.dtDtDateTo = dtDateTo.Value
        End If
    Else
        'set date from based on number of profiles
        If 0 = gdIsHoliday(dtDateTo, "") Then i = 1
        dtDate = dtDateTo.Value - 1
        While i < nNumProfiles
            If IsWeekday(dtDate) And 0 = gdIsHoliday(dtDate, "") Then
                i = i + 1
                dtDateFrom.Value = dtDate
            End If
            dtDate = dtDate - 1
        Wend
        m.dtDateFrom = dtDateFrom.Value
        m.dtDtDateTo = dtDateTo.Value
        
        i = Int(ValOfText(txtNumProfiles.Text))     'aardvark 6840
        If i > 0 Then
            If optIntraday Then
                m.nNumProfilesIntraday = i
            Else
                m.nNumProfilesMultiday = i
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.SyncProfileDateTime"

End Sub

Private Sub SetProfileControls()
On Error GoTo ErrSection:

    Dim Ind As cIndicator

    If m.dtEODLastDate = 0 Then m.dtEODLastDate = g.SymbolPool.EodLastDate(m.strSymbol)

    If m.Chart Is Nothing Then
        If Not Me.Visible Then LoadProfileSettingsLastUsed         '6844
        
        dtDateTo.Value = m.dtEODLastDate
        optEndOfData.Value = True
        optIntraday.Value = True
        
        'last-used settings
        txtNumProfiles = Str(m.nNumProfilesIntraday)
        
        If m.nTicksPerRow = -1 Then
            cboTicksPerRow.ListIndex = 0
        Else
            cboTicksPerRow.Text = Str(m.nTicksPerRow)
        End If
        
        'default to gradient color scheme and same colors as Trade Profile form (see frmMarketProfile)
        If cboColorScheme.ListIndex = MktProf_Color_Gradient Then
            imgGradientHorz.Visible = True
            lblColorLabel1.Visible = True
            lblColorLabel2.Visible = True
            gdColorFrom.Visible = True
            gdColorTo.Visible = True
            
            lblColorLabel1.Caption = "From"
            lblColorLabel2.Caption = "To"
            gdColorFrom.Color = vbYellow
            gdColorTo.Color = vbBlack
        ElseIf cboColorScheme.ListIndex = MktProf_Color_OpenClose Then
            imgGradientHorz.Visible = False
            lblColorLabel1.Visible = True
            lblColorLabel2.Visible = True
            gdColorFrom.Visible = True
            gdColorTo.Visible = True
            
            lblColorLabel1.Caption = "Up"
            lblColorLabel2 = "Down"
            gdColorFrom.Color = vbGreen
            gdColorTo.Color = vbRed
        Else
            imgGradientHorz.Visible = False
            lblColorLabel1.Visible = False
            lblColorLabel2.Visible = False
            gdColorFrom.Visible = False
            gdColorTo.Visible = False
        End If
        
        SyncProfileDateTime m.nNumProfilesIntraday
        
    Else
        If m.Chart.ToEndOfData Then
            optEndOfData.Value = True
        Else
            optToDate.Value = True
        End If
        dtDateFrom.Value = m.Chart.FromDate
        dtDateTo.Value = m.Chart.ToDate
        
        m.dtDateFrom = 0        'reset
        m.dtDtDateTo = 0
        
        If m.Chart.Bars.IsIntraday Then
            optIntraday.Value = True
        Else
            optMultiDays.Value = True
        End If
        
        SyncProfileDateTime -1
        
        If Not m.Chart.Tree Is Nothing Then
            Set Ind = m.Chart.Tree("PRICE")
        End If
    End If
    
    If Not Ind Is Nothing Then
        txtForecastBars = Str(m.Chart.BlankBars)
        
        m.nTicksPerRow = Ind.TicksPerRow
        m.nSessionsPerProfile = Ind.SessionsPerProfile
        txtSessionsPerProfile.Text = Str(m.nSessionsPerProfile)
        
        'minutes per bar
        If m.Chart.Bars.IsIntraday Then
            txtIntradayMinutes.Text = Str(m.Chart.Bars.Prop(eBARS_PeriodsPerBar))
        End If
        
        'ticks scale
        If m.nTicksPerRow = -1 Then
            cboTicksPerRow.ListIndex = 0
        Else
            cboTicksPerRow.Text = Str(m.nTicksPerRow)
        End If
        
        'letter options
        Select Case Ind.ProfileStyleTPO
            Case 0
                optAlphaNumeric.Value = True
            Case 1
                optClasicLettering.Value = True
            Case 2
                optBlocks.Value = True
        End Select
        
        cboColorScheme.ListIndex = Ind.ProfileColorScheme
        
        If Ind.ProfileColorScheme = MktProf_Color_Gradient Or Ind.ProfileColorScheme = MktProf_Color_OpenClose Then
            gdColorFrom.Color = Ind.ProfileColor(ePCStruct_TPO)
            gdColorTo.Color = Ind.ProfileColor(ePCStruct_TPO_ColorTo)
        
            If Ind.ProfileColorScheme = MktProf_Color_Gradient Then
                imgGradientHorz.Visible = True
                lblColorLabel1.Caption = "From"
                lblColorLabel2.Caption = "To"
            Else
                imgGradientHorz.Visible = False
                lblColorLabel1.Caption = "Up"
                lblColorLabel2.Caption = "Down"
            End If
            
            lblColorLabel1.Visible = True
            lblColorLabel2.Visible = True
            gdColorFrom.Visible = True
            gdColorTo.Visible = True
        Else
            imgGradientHorz.Visible = False
            lblColorLabel1.Visible = False
            lblColorLabel2.Visible = False
            gdColorFrom.Visible = False
            gdColorTo.Visible = False
        End If
        
        'TPO properties
        chkTPO.Value = Ind.ProfileShowHide(ePCStruct_TPO)
        chkTPO_POC.Value = Ind.ProfileShowHide(ePCStruct_TPO_POC)
        chkTPO_VA.Value = Ind.ProfileShowHide(ePCStruct_TPO_VA)
        chkOpenClose.Value = Ind.ProfileShowHide(ePCStruct_Open)
        
        gdColorTPO_POC.Color = Ind.ProfileColor(ePCStruct_TPO_POC)
        gdColorTPO_VA.Color = Ind.ProfileColor(ePCStruct_TPO_VA)
        gdColorOpenClose.Color = Ind.ProfileColor(ePCStruct_Open)
        
        txtPercentTPO_VA.Text = Str(Ind.ProfileParm(ePCStruct_TPO_VA))
        
        'volume properties
        chkVolume.Value = Ind.ProfileShowHide(ePCStruct_Volume)
        chkVolume_POC.Value = Ind.ProfileShowHide(ePCStruct_Volume_POC)
        chkVolume_VA.Value = Ind.ProfileShowHide(ePCStruct_Volume_VA)
        
        gdColorVolume.Color = Ind.ProfileColor(ePCStruct_Volume)
        gdColorVolume_POC.Color = Ind.ProfileColor(ePCStruct_Volume_POC)
        gdColorVolume_VA.Color = Ind.ProfileColor(ePCStruct_Volume_VA)
        
        txtPercentVolume_VA.Text = Str(Ind.ProfileParm(ePCStruct_Volume_VA))
        
        Me.Font = Ind.FontName
        Me.Font.Bold = Ind.FontBold
        Me.Font.Size = Ind.FontSize
        Me.FontItalic = Ind.FontItalic
        
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.SetProfileControls"

End Sub

Private Sub EnableProfileIntraday(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    lblMultidayDesc.Enabled = Not bEnable
    lblSessionsPerProfile.Enabled = Not bEnable
    txtSessionsPerProfile.Enabled = Not bEnable
    
    optClasicLettering.Enabled = bEnable
    lblStartStopInfo1.Enabled = bEnable
    lblStartStopInfo2.Enabled = bEnable
    lblStartStopTimes.Enabled = bEnable
    cmdStartStop.Enabled = bEnable
    
    fraIntradayiSessions.Enabled = bEnable
    lblIntradayDesc.Enabled = bEnable
    lblIntradayMin.Enabled = bEnable
    txtIntradayMinutes.Enabled = bEnable
    
    If optEndOfData.Value = True Then
        dtDateTo.Enabled = False
        lblEndOfData.Enabled = True
    Else
        dtDateTo.Enabled = True
        lblEndOfData.Enabled = False
    End If
    
    If bEnable Then
        If m.nNumProfilesIntraday <= 0 Then m.nNumProfilesIntraday = 10
        txtNumProfiles.Text = m.nNumProfilesIntraday
    Else
        If m.nNumProfilesMultiday <= 0 Then m.nNumProfilesMultiday = 20
        txtNumProfiles.Text = m.nNumProfilesMultiday
    End If
    
    If m.Chart Is Nothing Then
        lblStartStopInfo1.Enabled = False
        lblStartStopInfo2.Enabled = False
        lblStartStopTimes.Enabled = False
        cmdStartStop.Enabled = False
        cmdStartStop.Enabled = False
    Else
        StartStopTimeLabel m.Chart, lblStartStopTimes, lblStartStopInfo2
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.EnableProfileIntraday"

End Sub

Private Sub NewProfileChart()
On Error GoTo ErrSection:

    Dim i&, idx&, strTemp$
    
    Dim frmActive As Form
    Dim frm As Form
    
    Dim Chart As cChart
    Dim Pane As cPane
    Dim PricePane As cPane
    Dim Ind As cIndicator

    If Len(m.strSymbol) <= 0 Then Exit Sub
    If m.bEditProfile And m.Chart Is Nothing Then Exit Sub
    
    If m.bEditProfile Then
        Set Chart = m.Chart
        Set frm = m.Chart.Form
    Else
        Set frmActive = ActiveChart()
        Set frm = New frmChart
        frm.Chart.TemplateLoad "", False
        Set Chart = frm.Chart
    End If
    
    If Chart.Tree Is Nothing Then Exit Sub
        
    'remove all panes & indicators except price
    For i = Chart.Tree.Count To 1 Step -1
        If TypeOf Chart.Tree(i) Is cPane Then
            Set Pane = Chart.Tree(i)
            If Not Pane Is PricePane Then
                Chart.Tree.Remove i
            End If
        ElseIf TypeOf Chart.Tree(i) Is cIndicator Then
            Set Ind = Chart.Tree(i)
            If UCase(Ind.Name) = "PRICE" Then
                Ind.DataType = eINDIC_ProfileBars
                Ind.DisplayType = eINDIC_Profile
                
                Ind.TicksPerRow = m.nTicksPerRow
                Ind.SessionsPerProfile = m.nSessionsPerProfile
                
                If optAlphaNumeric.Value = True Then
                    Ind.ProfileStyleTPO = 0
                ElseIf optClasicLettering.Value = True Then
                    Ind.ProfileStyleTPO = 1
                ElseIf optBlocks.Value = True Then
                    Ind.ProfileStyleTPO = 2
                End If
                
                Ind.ProfileColorScheme = cboColorScheme.ListIndex
                
                'TPO properties
                Ind.ProfileShowHide(ePCStruct_TPO) = Abs(chkTPO.Value)
                Ind.ProfileShowHide(ePCStruct_TPO_POC) = Abs(chkTPO_POC.Value)
                Ind.ProfileShowHide(ePCStruct_TPO_VA) = Abs(chkTPO_VA.Value)
                Ind.ProfileShowHide(ePCStruct_Open) = Abs(chkOpenClose.Value)
                
                Ind.ProfileColor(ePCStruct_TPO) = gdColorFrom.Color
                Ind.ProfileColor(ePCStruct_TPO_ColorTo) = gdColorTo.Color
                
                Ind.ProfileColor(ePCStruct_TPO_POC) = gdColorTPO_POC.Color
                Ind.ProfileColor(ePCStruct_TPO_VA) = gdColorTPO_VA.Color
                Ind.ProfileColor(ePCStruct_Open) = gdColorOpenClose.Color
                
                Ind.ProfileParm(ePCStruct_TPO_VA) = ValOfText(txtPercentTPO_VA.Text)
                
                'volume properties
                Ind.ProfileShowHide(ePCStruct_Volume) = Abs(chkVolume.Value)
                Ind.ProfileShowHide(ePCStruct_Volume_POC) = Abs(chkVolume_POC.Value)
                Ind.ProfileShowHide(ePCStruct_Volume_VA) = Abs(chkVolume_VA.Value)
                
                Ind.ProfileColor(ePCStruct_Volume) = gdColorVolume.Color
                Ind.ProfileColor(ePCStruct_Volume_POC) = gdColorVolume_POC.Color
                Ind.ProfileColor(ePCStruct_Volume_VA) = gdColorVolume_VA.Color
                
                Ind.ProfileParm(ePCStruct_Volume_VA) = ValOfText(txtPercentVolume_VA.Text)
                
                'font
                If Not m.bEditProfile Then
                    Ind.FontName = Me.Font.Name
                    Ind.FontSize = Me.Font.Size
                    Ind.FontBold = Me.Font.Bold
                End If
                
                idx = Chart.Tree.RelativeIndex(i, eTREE_Parent)
                If TypeOf Chart.Tree(idx) Is cPane Then
                    Set PricePane = Chart.Tree(idx)
                End If
            Else
                Chart.Tree.Remove i
            End If
        End If
    Next
        
    Chart.IsProfileChart = True
        
    If optMultiDays.Value = True Then
        strTemp = "Daily"
    Else
        strTemp = txtIntradayMinutes.Text & "m"
        Chart.MaxIntradayDays = m.nNumProfilesIntraday
    End If
    m.nPeriodicity = GetPeriodicity(strTemp)
        
    Chart.ShowTrades = 0
    Chart.ShowToolbar = 0
    
    Chart.SetSymbol m.strSymbol
    Chart.ChangeBarPeriod m.nPeriodicity, False, , True
    
    If Chart.FromDate <> dtDateFrom.Value Or Chart.ToDate <> dtDateTo.Value Then
        Chart.RedoMode = eRedo9_ReloadData
    End If
    
    Chart.FromDate = dtDateFrom.Value
    Chart.ToDate = dtDateTo.Value
    Chart.ToEndOfData = optEndOfData.Value
        
    If m.Chart Is Nothing Then
        'this is new chart
        Set m.Chart = Chart
        txtForecastBars_LostFocus   'to set blank bars
        
        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
            If g.ChartGlobals.frmActiveNonDetached.WindowState = vbNormal Then
                '6741 - SetNormalPlacement is a misnomer, it is a Property Let functionality
                '   does not position or size the form at all
                frm.SetNormalPlacement frmActive.GetNormalPlacement
                
                'SetRatioPlacement positions & sizes the form based on passed in string.
                'If the autosize option is off, this function knows to use the form's normal
                'placement string instead (hence the normal placement string is set right above)
                frm.SetRatioPlacement frmActive.GetRatioPlacement, True
            Else
                frm.WindowState = vbMaximized       '5203
            End If
        End If
        
'        frm.Chart.TemplateSave      '5007
        frm.SkipFocusFix = True
        ActiveChartFormSet frm
        ShowForm frm
        MoveFocus frm.pbChart
    Else
        m.Chart.GenerateChart eRedo3_Settings
    End If
        
    Set frm = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmNewChart.NewProfileChart", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub txtForecastBars_Change()
On Error GoTo ErrSection:
    
    If Not m.Chart Is Nothing Then
        If ValOfText(txtForecastBars) <> m.Chart.BlankBars Then
            m.Chart.ResetLastScreenDate
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.txtForecastBars.Change", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub txtForecastBars_LostFocus()
On Error GoTo ErrSection:
   
    Dim dBars#, iBars&
   
    'aardvark 3295 fix:
    'need a double for users inputing value a long cannot accomodate (e.g. 999999999999)
    dBars = ValOfText(txtForecastBars)
    If dBars <= 0 Then
        iBars = 1
        txtForecastBars = iBars
    ElseIf dBars > 500 Then
        iBars = 500
        txtForecastBars = iBars
    Else
        iBars = Int(dBars)
    End If
    
    If Not m.Chart Is Nothing Then
        If iBars <> m.Chart.BlankBars Then
            m.Chart.BlankBars(Me) = iBars
            m.Chart.RedoMode = eRedo9_ReloadData
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.txtForecastBars.LostFocus", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub SaveProfileSettingsLastUsed()
On Error GoTo ErrSection:

    Dim i As Long
    
    SetIniFileProperty "TicksPerRow", m.nTicksPerRow, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "SessionPerProfile", m.nSessionsPerProfile, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "NumProfilesIntraDay", m.nNumProfilesIntraday, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "NumProfilesMultiDay", m.nNumProfilesMultiday, "ProfileOnChart", g.strIniFile
    
    SetIniFileProperty "FontName_TPO", Me.Font.Name, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "FontSize_TPO", Me.Font.Size, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "FontBold_TPO", Me.Font.Bold, "ProfileOnChart", g.strIniFile
    
    SetIniFileProperty "TPO_ColorScheme", cboColorScheme.ListIndex, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileColor_TPO", gdColorTPO_POC.Color, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileColor_TPO_VA", gdColorTPO_VA.Color, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileColor_OpenClose", gdColorOpenClose.Color, "ProfileOnChart", g.strIniFile
    
    i = Int(ValOfText(txtIntradayMinutes.Text))
    If i > 0 Then
        SetIniFileProperty "IntradayProfileMinutes", i, "ProfileOnChart", g.strIniFile
    End If
    
    'TPO letter style
    If optClasicLettering.Value = True Then
        SetIniFileProperty "ProfileTPOStyle", 1, "ProfileOnChart", g.strIniFile
    ElseIf optBlocks.Value = True Then
        SetIniFileProperty "ProfileTPOStyle", 2, "ProfileOnChart", g.strIniFile
    Else
        SetIniFileProperty "ProfileTPOStyle", 0, "ProfileOnChart", g.strIniFile
    End If
    
    SetIniFileProperty "ProfileShow_TPO", chkTPO.Value, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileShow_TPO_POC", chkTPO_POC.Value, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileShow_TPO_VA", chkTPO_VA.Value, "ProfileOnChart", g.strIniFile
    
    i = Int(ValOfText(txtPercentTPO_VA.Text))
    If i >= 0 And i <= 100 Then
        SetIniFileProperty "Profile_TPO_VA_PERCENT", i, "ProfileOnChart", g.strIniFile      '6866
    End If

    'volume
    SetIniFileProperty "ProfileColor_Volume", gdColorVolume.Color, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileColor_Volume_POC", gdColorVolume_POC.Color, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileColor_Volume_VA", gdColorVolume_VA.Color, "ProfileOnChart", g.strIniFile
    
    SetIniFileProperty "ProfileShow_Volume", chkVolume.Value, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileShow_Volume_POC", chkVolume_POC.Value, "ProfileOnChart", g.strIniFile
    SetIniFileProperty "ProfileShow_Volume_VA", chkVolume_VA.Value, "ProfileOnChart", g.strIniFile
    
    i = Int(ValOfText(txtPercentVolume_VA.Text))
    If i >= 0 And i <= 100 Then
        SetIniFileProperty "Profile_Volume_VA_PERCENT", i, "ProfileOnChart", g.strIniFile   '6866
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.SaveProfileSettingsLastUsed", eGDRaiseError_Show

End Sub

Private Sub LoadProfileSettingsLastUsed()

    ' TLB 6/16/2016: user had an error in this routine, too hard to figure out, so this is easy solution
    If IsIDE Then
        On Error GoTo ErrSection:
    Else
        On Error Resume Next
    End If

    Dim i As Long
    
    m.nTicksPerRow = GetIniFileProperty("TicksPerRow", -1, "ProfileOnChart", g.strIniFile)
    m.nSessionsPerProfile = GetIniFileProperty("SessionPerProfile", 3, "ProfileOnChart", g.strIniFile)
    m.nNumProfilesIntraday = GetIniFileProperty("NumProfilesIntraDay", 10, "ProfileOnChart", g.strIniFile)
    m.nNumProfilesMultiday = GetIniFileProperty("NumProfilesMultiDay", 20, "ProfileOnChart", g.strIniFile)

    i = GetIniFileProperty("IntradayProfileMinutes", 30, "ProfileOnChart", g.strIniFile)
    If IsNumeric(i) Then
        txtIntradayMinutes.Text = Str(i)
    End If
    

'TPO properties
    'TPO font
    Me.Font.Name = GetIniFileProperty("FontName_TPO", "MS Sans Serif", "ProfileOnChart", g.strIniFile)
    Me.Font.Size = GetIniFileProperty("FontSize_TPO", 8, "ProfileOnChart", g.strIniFile)
    Me.Font.Bold = GetIniFileProperty("FontBold_TPO", -1, "ProfileOnChart", g.strIniFile)   '-1 = True
    
    'TPO color scheme
    i = GetIniFileProperty("TPO_ColorScheme", MktProf_Color_Gradient, "ProfileOnChart", g.strIniFile)
    If i = MktProf_Color_Gradient Or i = MktProf_Color_Rainbow Or i = MktProf_Color_OpenClose Then
        cboColorScheme.ListIndex = i
    Else
        cboColorScheme.ListIndex = MktProf_Color_Gradient
    End If
    
    'TPO letter style
    i = GetIniFileProperty("ProfileTPOStyle", 0, "ProfileOnChart", g.strIniFile)
    If i = 1 Then
        optClasicLettering.Value = True
    ElseIf i = 2 Then
        optBlocks.Value = True
    Else
        optAlphaNumeric.Value = True
    End If
    
    'TPO show lettering
    i = GetIniFileProperty("ProfileShow_TPO", 1, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkTPO.Value = i
    Else
        chkTPO.Value = vbChecked
    End If
    
    'TPO show POC
    i = GetIniFileProperty("ProfileShow_TPO_POC", 1, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkTPO_POC.Value = i
    Else
        chkTPO_POC.Value = vbChecked
    End If

    'TPO show value area
    i = GetIniFileProperty("ProfileShow_TPO_VA", 0, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkTPO_VA.Value = i
    Else
        chkTPO_VA.Value = vbUnchecked
    End If
    
    i = GetIniFileProperty("Profile_TPO_VA_PERCENT", 70, "ProfileOnChart", g.strIniFile)
    If i > 0 And i <= 100 Then
        txtPercentTPO_VA.Text = Str(i)
    Else
        txtPercentTPO_VA.Text = 70
    End If
    
    'show profile Open/Close price using triangles
    i = GetIniFileProperty("ProfileShow_OpenClose", 0, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkOpenClose.Value = i
    Else
        chkOpenClose.Value = vbUnchecked
    End If

    gdColorTPO_POC.Color = GetIniFileProperty("ProfileColor_TPO", 8388736, "ProfileOnChart", g.strIniFile)
    gdColorTPO_VA.Color = GetIniFileProperty("ProfileColor_TPO_VA", 32768, "ProfileOnChart", g.strIniFile)
    gdColorOpenClose.Color = GetIniFileProperty("ProfileColor_OpenClose", 0, "ProfileOnChart", g.strIniFile)

'volume properties
    'volume show
    i = GetIniFileProperty("ProfileShow_Volume", 0, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkVolume.Value = i
    Else
        chkVolume.Value = vbChecked
    End If
        
    'volume show POC
    i = GetIniFileProperty("ProfileShow_Volume_POC", 1, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkVolume_POC.Value = i
    Else
        chkVolume_POC.Value = vbChecked
    End If

    'volume show value area
    i = GetIniFileProperty("ProfileShow_Volume_VA", 0, "ProfileOnChart", g.strIniFile)
    If i = 0 Or i = 1 Then
        chkVolume_VA.Value = i
    Else
        chkVolume_VA.Value = vbUnchecked
    End If
    
    i = GetIniFileProperty("Profile_Volume_VA_PERCENT", 70, "ProfileOnChart", g.strIniFile)
    If i > 0 And i <= 100 Then
        txtPercentVolume_VA.Text = Str(i)
    Else
        txtPercentVolume_VA.Text = 70
    End If
        
    gdColorVolume.Color = GetIniFileProperty("ProfileColor_Volume", 12648447, "ProfileOnChart", g.strIniFile)
    gdColorVolume_POC.Color = GetIniFileProperty("ProfileColor_Volume_POC", vbCyan, "ProfileOnChart", g.strIniFile)
    gdColorVolume_VA.Color = GetIniFileProperty("ProfileColor_Volume_VA", vbGreen, "ProfileOnChart", g.strIniFile)
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmNewChart.LoadProfileSettingsLastUsed", eGDRaiseError_Show

End Sub


