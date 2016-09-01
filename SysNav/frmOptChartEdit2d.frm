VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOptChartEdit2d 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart Settings"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraGridLines 
      Height          =   600
      Left            =   135
      TabIndex        =   12
      Top             =   2265
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
      Caption         =   "frmOptChartEdit2d.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptChartEdit2d.frx":0034
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":0054
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   3
         Left            =   2730
         TabIndex        =   1
         Top             =   285
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
         Caption         =   "frmOptChartEdit2d.frx":0070
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit2d.frx":0098
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit2d.frx":00B8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   2
         Left            =   1865
         TabIndex        =   2
         Top             =   285
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
         Caption         =   "frmOptChartEdit2d.frx":00D4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit2d.frx":0100
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit2d.frx":0120
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   1
         Left            =   1000
         TabIndex        =   5
         Top             =   285
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
         Caption         =   "frmOptChartEdit2d.frx":013C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit2d.frx":0168
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit2d.frx":0188
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optGridLines 
         Height          =   220
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   285
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
         Caption         =   "frmOptChartEdit2d.frx":01A4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit2d.frx":01CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit2d.frx":01EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboShadow 
      Height          =   315
      Left            =   2925
      TabIndex        =   10
      Top             =   1770
      Width           =   945
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
      Tip             =   "frmOptChartEdit2d.frx":0208
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
      MouseIcon       =   "frmOptChartEdit2d.frx":0228
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboFontSize 
      Height          =   315
      Left            =   1065
      TabIndex        =   8
      Top             =   1785
      Width           =   945
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
      Tip             =   "frmOptChartEdit2d.frx":0244
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
      MouseIcon       =   "frmOptChartEdit2d.frx":0264
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboChartType 
      Height          =   315
      Left            =   1065
      TabIndex        =   6
      Top             =   1350
      Width           =   2805
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
      Tip             =   "frmOptChartEdit2d.frx":0280
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
      MouseIcon       =   "frmOptChartEdit2d.frx":02A0
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniComboImageXP cboDataY 
      Height          =   315
      Left            =   1065
      TabIndex        =   4
      Top             =   930
      Width           =   2805
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
      Tip             =   "frmOptChartEdit2d.frx":02BC
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
      MouseIcon       =   "frmOptChartEdit2d.frx":02DC
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPointsToGraph 
      Height          =   315
      Left            =   3075
      TabIndex        =   3
      Top             =   510
      Width           =   795
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOptChartEdit2d.frx":02F8
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
      Tip             =   "frmOptChartEdit2d.frx":032A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":034A
   End
   Begin gdOCX.gdSelectColor gdColor 
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      CustomColor     =   255
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   600
      Left            =   405
      TabIndex        =   7
      Top             =   2850
      Width           =   3240
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit2d.frx":0366
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOptChartEdit2d.frx":039A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":03BA
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   750
         TabIndex        =   9
         Top             =   195
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
         Caption         =   "frmOptChartEdit2d.frx":03D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit2d.frx":03FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit2d.frx":041C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   330
         Left            =   1650
         TabIndex        =   11
         Top             =   195
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
         Caption         =   "frmOptChartEdit2d.frx":0438
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOptChartEdit2d.frx":0466
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOptChartEdit2d.frx":0486
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label4 
      Height          =   225
      Left            =   2205
      Top             =   1815
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
      Caption         =   "frmOptChartEdit2d.frx":04A2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit2d.frx":04CE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":04EE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label3 
      Height          =   225
      Left            =   135
      Top             =   1830
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
      Caption         =   "frmOptChartEdit2d.frx":050A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit2d.frx":053C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":055C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   225
      Left            =   150
      Top             =   1395
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
      Caption         =   "frmOptChartEdit2d.frx":0578
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit2d.frx":05AE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":05CE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAxisY 
      Height          =   225
      Left            =   146
      Top             =   975
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
      Caption         =   "frmOptChartEdit2d.frx":05EA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit2d.frx":0626
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":0646
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblPointsToGraph 
      Height          =   255
      Left            =   153
      Top             =   540
      Width           =   2805
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmOptChartEdit2d.frx":0662
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit2d.frx":06C4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":06E4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   810
      Top             =   90
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
      Caption         =   "frmOptChartEdit2d.frx":0700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOptChartEdit2d.frx":072C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOptChartEdit2d.frx":074C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmOptChartEdit2d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    fmOptChart As Form
    Pe2d As Pegoa
    nPtsToGraph As Long
End Type
Private m As mPrivate

Public Sub ShowMe(fmForm As Form, Pe2d As Pegoa)
On Error GoTo ErrSection:

    Set m.fmOptChart = fmForm
    Set m.Pe2d = Pe2d
    
    If m.fmOptChart Is Nothing Or m.Pe2d Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
    InitControls
        
    m.nPtsToGraph = m.fmOptChart.PointsToGraph
    CenterTheForm Me
    Me.Show 1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.ShowMe", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    UpdateChart
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitControls()
On Error GoTo ErrSection:
    
    gdColor.Color = m.Pe2d.SubsetColors(0)
    txtPointsToGraph.Text = Str(m.fmOptChart.PointsToGraph)
    m.fmOptChart.InitComboY cboDataY
    InitChartType
    InitFontSize
    InitShadowStyle
    InitGridLines
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.InitControls", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub UpdateChart()
On Error GoTo ErrSection:
    
    Dim nPtsToGraph&, nIdxX&, nIdxZ&
    
    'color
    m.fmOptChart.Color = gdColor.Color
    'number of points per screen
    nPtsToGraph = Int(Val(txtPointsToGraph.Text))
    If nPtsToGraph > 0 And nPtsToGraph <> m.nPtsToGraph Then
        m.Pe2d.PointsToGraph = nPtsToGraph
        m.fmOptChart.PointsToGraph = nPtsToGraph
    End If
    'Data column to plot
    m.fmOptChart.IdxAxisY = cboDataY.ItemData(cboDataY.ListIndex)
    
    SetChartType
    SetFontSize
    SetShadowStyle
    SetGridLines
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.UpdateChart", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitChartType()
On Error GoTo ErrSection:
    
    cboChartType.Clear
    
    cboChartType.AddItem "Line"
    cboChartType.AddItem "Bar"
    cboChartType.AddItem "Point"
    cboChartType.AddItem "Area"
    cboChartType.AddItem "Points + Best Fit Line"
    cboChartType.AddItem "Points + Best Fit Curve"
    cboChartType.AddItem "Points + Line"
    cboChartType.AddItem "Spline"
    cboChartType.AddItem "Ribbon"
    
    Select Case m.Pe2d.PlottingMethod
        Case GPM_LINE
            cboChartType.ListIndex = 0
        Case GPM_BAR
            cboChartType.ListIndex = 1
        Case GPM_POINT
            cboChartType.ListIndex = 2
        Case GPM_AREA
            cboChartType.ListIndex = 3
        Case GPM_POINTSPLUSBFL
            cboChartType.ListIndex = 4
        Case GPM_POINTSPLUSBFC
            cboChartType.ListIndex = 5
        Case GPM_POINTSPLUSLINE
            cboChartType.ListIndex = 6
        Case GPM_SPLINE
            cboChartType.ListIndex = 7
        Case GPM_RIBBON
            cboChartType.ListIndex = 8
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.InitChartType", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitFontSize()
On Error GoTo ErrSection:

    cboFontSize.Clear
    
    cboFontSize.AddItem "Small"
    cboFontSize.AddItem "Medium"
    cboFontSize.AddItem "Large"

    Select Case m.Pe2d.Font.Size
        Case PEFS_SMALL
            cboFontSize.ListIndex = 0
        Case PEFS_MEDIUM
            cboFontSize.ListIndex = 1
        Case PEFS_LARGE
            cboFontSize.ListIndex = 2
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.InitFontSize", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitGridLines()
On Error GoTo ErrSection:

    Select Case m.Pe2d.GridLineControl
        Case PEGLC_BOTH
            optGridLines(0) = True
        Case PEGLC_XAXIS
            optGridLines(1) = True
        Case PEGLC_YAXIS
            optGridLines(2) = True
        Case PEGLC_NONE
            optGridLines(3) = True
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.InitGridLines", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitShadowStyle()
On Error GoTo ErrSection:

    cboShadow.Clear
    
    cboShadow.AddItem "Off"
    cboShadow.AddItem "Shadow"
    cboShadow.AddItem "3D"
    
    Select Case m.Pe2d.DataShadows
        Case PEDS_NONE
            cboShadow.ListIndex = 0
        Case PEDS_SHADOWS
            cboShadow.ListIndex = 1
        Case PEDS_3D
            cboShadow.ListIndex = 2
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.InitShadowStyle", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub SetChartType()
On Error GoTo ErrSection:
    
    Select Case cboChartType.ListIndex
        Case 0
            m.Pe2d.PlottingMethod = GPM_LINE
        Case 1
            m.Pe2d.PlottingMethod = GPM_BAR
        Case 2
            m.Pe2d.PlottingMethod = GPM_POINT
        Case 3
            m.Pe2d.PlottingMethod = GPM_AREA
        Case 4
            m.Pe2d.PlottingMethod = GPM_POINTSPLUSBFL
        Case 5
            m.Pe2d.PlottingMethod = GPM_POINTSPLUSBFC
        Case 6
            m.Pe2d.PlottingMethod = GPM_POINTSPLUSLINE
        Case 7
            m.Pe2d.PlottingMethod = GPM_SPLINE
        Case 8
            m.Pe2d.PlottingMethod = GPM_RIBBON
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.SetChartType", eGDRaiseError_Show
    Resume ErrExit

End Sub
    
Private Sub SetFontSize()
On Error GoTo ErrSection:

    Select Case cboFontSize.ListIndex
        Case 0
            m.fmOptChart.PeFontSize = PEFS_SMALL
        Case 1
            m.fmOptChart.PeFontSize = PEFS_MEDIUM
        Case 2
            m.fmOptChart.PeFontSize = PEFS_LARGE
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.SetFontSize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub
    
Private Sub SetShadowStyle()
On Error GoTo ErrSection:
    
    Select Case cboShadow.ListIndex
        Case 0
            m.Pe2d.DataShadows = PEDS_NONE
        Case 1
            m.Pe2d.DataShadows = PEDS_SHADOWS
        Case 2
            m.Pe2d.DataShadows = PEDS_3D
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.SetShadowStyle", eGDRaiseError_Show
    Resume ErrExit
    
End Sub
    
Private Sub SetGridLines()
On Error GoTo ErrSection:

    If optGridLines(0) = True Then
        m.Pe2d.GridLineControl = PEGLC_BOTH
    ElseIf optGridLines(1) = True Then
        m.Pe2d.GridLineControl = PEGLC_XAXIS
    ElseIf optGridLines(2) = True Then
        m.Pe2d.GridLineControl = PEGLC_YAXIS
    ElseIf optGridLines(3) = True Then
        m.Pe2d.GridLineControl = PEGLC_NONE
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.SetGridLines", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("ID_News"))
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.fmOptChart = Nothing
    Set m.Pe2d = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChartEdit2d.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

