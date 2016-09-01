VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmExportGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Symbol Group Export"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniCheckXP chkSplit 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3720
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
      Caption         =   "frmExportGroup.frx":0000
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmExportGroup.frx":0068
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":0088
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdAsciiOptions 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
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
      Caption         =   "frmExportGroup.frx":00A4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmExportGroup.frx":00E0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":0100
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraPath 
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   6375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmExportGroup.frx":011C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportGroup.frx":0148
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":0168
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPath 
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   0
         Width           =   4455
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmExportGroup.frx":0184
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
         Tip             =   "frmExportGroup.frx":01A4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":01C4
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   0
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
         Caption         =   "frmExportGroup.frx":01E0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportGroup.frx":020E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":022E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   0
         Top             =   60
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
         Caption         =   "frmExportGroup.frx":024A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportGroup.frx":0274
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":0294
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraDates 
      Height          =   1455
      Left            =   3120
      TabIndex        =   10
      Top             =   1440
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
      Caption         =   "frmExportGroup.frx":02B0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportGroup.frx":02E4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":0304
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optToEnd 
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1080
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
         Caption         =   "frmExportGroup.frx":0320
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmExportGroup.frx":0356
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":0376
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optToDate 
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   750
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
         Caption         =   "frmExportGroup.frx":0392
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmExportGroup.frx":03C0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":03E0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate FromDate 
         Height          =   315
         Left            =   840
         TabIndex        =   12
         Top             =   330
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
      End
      Begin gdOCX.gdSelectDate ToDate 
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   255
         Left            =   240
         Top             =   750
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
         Caption         =   "frmExportGroup.frx":03FC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportGroup.frx":0422
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":0442
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   240
         Top             =   360
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
         Caption         =   "frmExportGroup.frx":045E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportGroup.frx":0488
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":04A8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboPeriod 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
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
      Tip             =   "frmExportGroup.frx":04C4
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":04E4
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   4200
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
      Caption         =   "frmExportGroup.frx":0500
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportGroup.frx":052C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":054C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   465
         Left            =   0
         TabIndex        =   13
         Top             =   0
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
         Caption         =   "frmExportGroup.frx":0568
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportGroup.frx":058E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":05AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Default         =   -1  'True
         Height          =   465
         Left            =   1920
         TabIndex        =   18
         Top             =   0
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
         Caption         =   "frmExportGroup.frx":05CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportGroup.frx":05F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":0618
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboFormat 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1530
      Width           =   1695
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
      Tip             =   "frmExportGroup.frx":0634
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":0654
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraSymbolGroup 
      Height          =   975
      Left            =   248
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmExportGroup.frx":0670
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExportGroup.frx":06BC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":06DC
      RightToLeft     =   0   'False
      Begin MSComctlLib.ImageCombo cboSymbolGroup 
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Top             =   382
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   360
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
         Caption         =   "frmExportGroup.frx":06F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportGroup.frx":0720
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":0740
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
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
         Caption         =   "frmExportGroup.frx":075C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExportGroup.frx":0786
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":07A6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   120
         Top             =   420
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
         Caption         =   "frmExportGroup.frx":07C2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmExportGroup.frx":07FC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmExportGroup.frx":081C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label4 
      Height          =   255
      Left            =   360
      Top             =   2640
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
      Caption         =   "frmExportGroup.frx":0838
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExportGroup.frx":0866
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":0886
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label2 
      Height          =   255
      Left            =   360
      Top             =   1560
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
      Caption         =   "frmExportGroup.frx":08A2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExportGroup.frx":08D0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExportGroup.frx":08F0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmExportGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmExportGroup.frm
'' Description: Form to allow the user to edit an individual symbol group
''              export setup
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 06/22/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type mPrivate
    bOK As Boolean
    ExportGroup As cExportGroup
End Type
Private m As mPrivate

Private Const kCSI = 0
Private Const kMETASTOCK = 1
Private Const kASCII = 2
Private Const kGDB = 3

Private Sub FormatCheck()
On Error GoTo ErrSection:

    If IsIntraday(GetPeriodicity(cboPeriod.Text)) And (cboFormat.Text <> "ASCII") And (cboFormat.Text <> "GDB") Then
        InfBox "Intraday data can only be exported in|ASCII or GDB format", "i", , "Information"
        'cboFormat.Text = "ASCII"
    End If

    cmdAsciiOptions.Enabled = (cboFormat.ListIndex = kASCII)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.FormatCheck"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboFormat_Change
'' Description: If the user chooses ASCII format, enable the ASCII options
''              button, otherwise, disable it.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboFormat_Change()
On Error GoTo ErrSection:

    FormatCheck

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cboFormat_Change"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboFormat_Click
'' Description: If the user chooses ASCII format, enable the ASCII options
''              button, otherwise, disable it.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboFormat_Click()
On Error GoTo ErrSection:

    FormatCheck

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cboFormat_Click"
    Resume ErrExit
End Sub

Private Sub cboPeriod_Click()
On Error GoTo ErrSection:

    If IsIntraday(GetPeriodicity(cboPeriod.Text)) Then
        cboFormat.Text = "ASCII"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cboPeriod.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbolGroup_Change
'' Description: If the Symbol Group combo changes, only allow the edit button
''              to be enabled if the chosen symbol group is custom
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbolGroup_Change()
On Error GoTo ErrSection:

    With cboSymbolGroup
        cmdEdit.Enabled = g.SymbolPool.PoolObject(.SelectedItem.Key).Custom
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cboSymbolGroup.Change", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbolGroup_Click
'' Description: If the Symbol Group combo changes, only allow the edit button
''              to be enabled if the chosen symbol group is custom
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbolGroup_Click()
On Error GoTo ErrSection:

    With cboSymbolGroup
        cmdEdit.Enabled = g.SymbolPool.PoolObject(.SelectedItem.Key).Custom
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cboSymbolGroup.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAsciiOptions_Click
'' Description: If the user clicks on the ASCII options button, for now bring
''              up a message box saying that it will be available in a future
''              version.
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAsciiOptions_Click()
On Error GoTo ErrSection:

    m.ExportGroup.Path = Trim(txtPath.Text)
    frmExportAscii.ShowMe m.ExportGroup
    If AddSlash(UCase(m.ExportGroup.Path)) <> AddSlash(UCase(txtPath.Text)) Then
        If DirExist(m.ExportGroup.Path) Then txtPath.Text = AddSlash(m.ExportGroup.Path)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cmdAsciiOptions.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: If the user clicks on the Browse button, bring up a dialog
''              to allow them to choose a directory
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return from the browse form
    
    'strReturn = frmBrowseFolders.ShowMe(txtPath.Text, "Please select directory to export to...")
    strReturn = BrowseForFolder(txtPath.Text, "Please select directory to export to ...")
    
    If strReturn <> "" Then txtPath.Text = FileNameDisplay(strReturn)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmExportGroup.cmdBrowse.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user presses the Cancel button, hide the form and set
''              the OK to False
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: If the user clicks on the edit button, bring up the symbol
''              group form to allow the user to edit the symbol group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:
    
    Dim strKey As String                ' Key for the item selected in the combo
    Dim frm As Form                     ' Form object for the appropriate editor
    Dim strPrefix As String             ' Prefix to the ID
    
    strKey = cboSymbolGroup.SelectedItem.Key
    If g.SymbolPool.PoolObject(strKey).Custom Then
        ' Call the appropriate editor for the item selected...
        Select Case Parse(strKey, ":", 1)
            Case "GRP"
                Set frm = New frmSymbolGroup
                strKey = frm.ShowMe(AddSlash(App.Path) & "Custom\", Parse(strKey, ":", 2), , , , True)
                strPrefix = "GRP:"
            Case "FIL"
                Set frm = New frmFilter
                strKey = frm.ShowMe(AddSlash(App.Path) & "Custom\", Parse(strKey, ":", 2), True)
                strPrefix = "FIL:"
            Case "SCN"
                Set frm = New frmCriteria
                strKey = frm.ShowMe(AddSlash(App.Path) & "Custom\", Parse(strKey, ":", 2), True)
                strPrefix = "SCN:"
        End Select
        
        ' Reload the symbol group combo...
        cboSymbolGroup.ImageList = frmMain.img16
        cboSymbolGroup.Locked = True
        LoadCombo True
    
        ' Select the item in the combo that was selected before...
        If strKey <> "" Then
            cboSymbolGroup.ComboItems(strPrefix & strKey).Selected = True
        End If
    End If

ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: If the user clicks on the New button, bring up the symbol group
''              form to allow the user to create a new symbol group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    Dim strID As String                 ' ID of the new Symbol Group
    Dim frm As Form                     ' Form object for the appropriate editor

    ' Bring up the Symbol Group editor...
    Set frm = New frmSymbolGroup
    strID = frm.ShowMe(AddSlash(App.Path) & "Custom\", "", , , , True)
    
    ' Reload the combo box...
    cboSymbolGroup.ImageList = frmMain.img16
    cboSymbolGroup.Locked = True
    LoadCombo True
    
    ' Select the new group if it exists...
    If strID <> "" Then cboSymbolGroup.ComboItems("GRP:" & strID).Selected = True
    
ErrExit:
    Set frm = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user presses the OK button, hide the form and set the OK
''              to True
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    ' Force the user to enter in their ASCII options if they choose ASCII
    If cboFormat.ListIndex = kASCII Then
        If m.ExportGroup.Fields = "" Then
            cmdAsciiOptions_Click
            Exit Sub
        End If
    End If
    
    If Not TooManySymbols Then
        m.bOK = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, fill in the symbol group combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    'FillSymbolGroupCombo
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.Form.Activate", eGDRaiseError_Show
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
    RaiseError "frmExportGroup.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it and initialize the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    ' Center the form
    Me.Icon = Picture16(ToolbarIcon("ID_ExportData"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me

    cboSymbolGroup.ImageList = frmMain.img16
    cboSymbolGroup.Locked = True
    LoadCombo True
    'FillSymbolGroupCombo
    
    ' Load up the format combo
    With cboFormat
        .AddItem "CSI", kCSI
        .AddItem "MetaStock", kMETASTOCK
        .AddItem "ASCII", kASCII
        .AddItem "GDB", kGDB
    End With
    
    ' Load up the period combo
    With cboPeriod
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Quarterly"
        .AddItem "Yearly"
        
        ' If the user has Gold and is authorized for at least one type of intraday
        ' data, add the intraday export options...
        If HasGold(False) And (InStr(g.strAuthorizationString, "FT") Or InStr(g.strAuthorizationString, "IT") Or InStr(g.strAuthorizationString, "ST")) Then
            If FileExist(App.Path & "\AutoExport.flg") Then
                .AddItem "1 Minute"
            End If
            .AddItem "5 Minute"
            .AddItem "10 Minute"
            .AddItem "15 Minute"
            .AddItem "30 Minute"
            .AddItem "45 Minute"
            .AddItem "60 Minute"
            .AddItem "90 Minute"
            .AddItem "120 Minute"
            .AddItem "180 Minute"
            .AddItem "240 Minute"
            .AddItem "300 Minute"
            If IsIDE Then
                '.AddItem "300 Trades"
            End If
        End If
    End With
    
    ' Set up the minimum and maximum dates on the date control
    FromDate.MaxDateIsToday = True
    ToDate.MaxDateIsToday = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up the controls with the parameters passed in, show the
''              form, then pass back the information changed if the OK button
''              was hit.
'' Inputs:      Symbol Group Name, Format, Path, Symbol Group ID
'' Returns:     True if OK was hit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ExportGroup As cExportGroup) As Boolean
On Error GoTo ErrSection:
    
    If ExportGroup.SymbolGroup = "" Then
        cboSymbolGroup.ComboItems(1).Selected = True
        cboFormat.ListIndex = kCSI
        cboPeriod.ListIndex = 0
        FromDate.Value = 18264#
        ToDate.Value = Date
        optToEnd.Value = True
        txtPath.Text = "C:\"
        chkSplit.Value = vbChecked
    Else
        With ExportGroup
            If SelectInCombo(.SymbolGroupID) = False Then
                InfBox "The Symbol Group or Filter that this export group relies on no longer exists or is no longer active.", "!", , "Export Group Error"
            End If
            cboFormat.Text = .Format 'strFormat
            cboPeriod.Text = .Period
            FromDate.Value = .StartDate
            If .EndDate = 0 Then
                ToDate.Value = Date
                optToEnd.Value = True
            Else
                ToDate.Value = .EndDate
                optToDate.Value = True
            End If
            txtPath.Text = .Path 'strPath
            
            If .SplitAdjust Then
                chkSplit.Value = vbChecked
            Else
                chkSplit.Value = vbUnchecked
            End If
        End With
    End If
       
    Set m.ExportGroup = ExportGroup.MakeCopy
    ShowForm Me, True
    
    If m.bOK = True Then
        'Set ExportGroup = m.ExportGroup.MakeCopy
        With ExportGroup
            .SymbolGroupID = cboSymbolGroup.SelectedItem.Key 'g.SymbolPool.SymbolGroups(cboSymbolGroup.ItemData(cboSymbolGroup.ListIndex)).ID
            .SymbolGroup = cboSymbolGroup.Text
            .Format = cboFormat.Text
            .Period = cboPeriod.Text
            .StartDate = FromDate.Value
            If optToEnd.Value = True Then
                .EndDate = 0#
            Else
                .EndDate = ToDate.Value
            End If
            .Path = txtPath.Text
        
            .DateFormat = m.ExportGroup.DateFormat
            .Delimiter = m.ExportGroup.Delimiter
            .ShowHeader = m.ExportGroup.ShowHeader
            .CustomHeader = m.ExportGroup.CustomHeader
            .HeaderLine = m.ExportGroup.HeaderLine
            .DecimalChar = m.ExportGroup.DecimalChar
            .Extension = m.ExportGroup.Extension
            .SingleFile = m.ExportGroup.SingleFile
            .FileName = m.ExportGroup.FileName
            .Fields = m.ExportGroup.Fields
            
            If chkSplit.Value = vbChecked Then
                .SplitAdjust = True
            Else
                .SplitAdjust = False
            End If
        End With
    End If
    
    ShowMe = m.bOK
    Unload Me

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExportGroup.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hit the X in the control menu, hide the form and
''              set OK to false
'' Inputs:      Whether or not to Cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Me.Hide
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillSymbolGroupCombo
'' Description: Fill in the symbol group combo box with the available symbol
''              groups that are not built-in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If 0 Then
Private Sub FillSymbolGroupCombo()
On Error GoTo ErrSection:

    Dim Objects As cGdTree              ' Tree of symbol group objects
    Dim lIndex As Long                  ' Index for a for loop
    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim lOldListIndex As Long           ' List index before the save

    ' Load up the symbol grid combo
    With cboSymbolGroup
        lOldListIndex = .ListIndex
        .Clear
        Set Objects = g.SymbolPool.SymbolGroups
        For lIndex = 1 To Objects.Count
            If Objects(lIndex).GroupType <> eGROUP_Builtin Then
                lFieldNum = g.SymbolPool.FieldNumForID("GRP:" & Objects(lIndex).ID)
                If lFieldNum <> -1 Then
                    .AddItem Objects(lIndex).Name
                    .ItemData(.NewIndex) = lIndex 'lFieldNum
                End If
            End If
        Next lIndex
        
        If lOldListIndex < 0 Then .ListIndex = 0 Else .ListIndex = lOldListIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExportGroup.FillSymbolGroupCombo", eGDRaiseError_Raise
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TooManySymbols
'' Description: Check to make sure that the symbol group selected does not
''              exceed the maximum number of symbols for the format selected
'' Inputs:      None
'' Returns:     True if not too many symbols, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TooManySymbols() As Boolean
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message for a infbox
    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim aSymbols As cGdArray            ' Array of true/false in grid values
    Dim lCount As Long                  ' Number of symbols in the symbol group
    Dim lMax As Long
        
    With g.SymbolPool
'        lFieldNum = .FieldNumForID("GRP:" & .SymbolGroups(cboSymbolGroup.ItemData(cboSymbolGroup.ListIndex)).ID)
        lFieldNum = .FieldNumForID(cboSymbolGroup.SelectedItem.Key)
        Set aSymbols = .ArrayTable.FieldArray(lFieldNum)
        lCount = aSymbols.CountOf(1)
    End With
            
    strMessage = ""
    TooManySymbols = False
    Select Case cboFormat.ListIndex
        Case kCSI
            If lCount > 999 Then
                strMessage = "You can only have up to 999 symbols with CSI format"
                TooManySymbols = True
            End If
        Case kMETASTOCK
            If lCount > 255 Then
                strMessage = "You can only have up to 255 symbols with MetaStock format"
                TooManySymbols = True
            End If
        Case kASCII, kGDB
            ' if only symbol/desc then allow all symbols
            If m.ExportGroup.NoPrices Then
                lMax = 9999999
            Else
                lMax = 9999
            End If
            If lCount > lMax Then
                strMessage = "You can only have up to " & CStr(lMax) & " symbols"
                TooManySymbols = True
            End If
    End Select
    
    If strMessage <> "" Then InfBox "h=Error ; i=! ; " & strMessage

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExportGroup.TooManySymbols", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the filters combo box with the symbol groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(Optional ByVal bShowFilters As Boolean = False)
On Error Resume Next

    Dim lIndex As Long                  ' Index for a for loop
    Dim strID As String                 ' Symbol pool ID for the field
    Dim strType As String               ' Type of thing (i.e. Filter, Criteria, etc)
    Dim strPicture As String            ' Picture to use in the combo box
    Dim strSelID As String              ' ID of the currently selected item
    Dim bSelExists As Boolean           ' Old selection still exists
    Dim iSortStart As Long              ' Where to start the sort
    Dim strItem As String               ' Item to add to the combo box
    Dim aItems As New cGdArray          ' Items to add to the combo box
    Dim obj As Object                   ' Symbol Pool Object
    Dim bScans As Boolean               ' Are we doing scans?
   
    bScans = ScansEnabled
        
    If cboSymbolGroup.ComboItems.Count > 0 Then
        strSelID = cboSymbolGroup.SelectedItem.Key
        cboSymbolGroup.ComboItems.Clear
    End If
    
    ' get list of items to put into combo list
    With g.SymbolPool
        For lIndex = 0 To .ArrayTable.NumFields - 1
            strID = .FieldID(lIndex)
            If Len(strID) = 0 Then
                strType = "" '???
            Else
                strType = Left(strID, 3)
                strPicture = ""
                Set obj = .PoolObject(strID)
                Select Case UCase(strType)
                    Case "GRP"
                        If obj.GroupType <> eGROUP_Flag Then
                            strPicture = ToolbarIcon("ID_SymbolGroups")
                        End If
                    Case "FIL"
                        If bScans And bShowFilters Then
                            strPicture = ToolbarIcon("ID_Filters")
                        End If
                End Select
                If Len(strPicture) > 0 And obj.IsActive = True Then
                    If strID = strSelID Then
                        bSelExists = True
                    End If
                    
                    If iSortStart = 0 And lIndex >= g.SymbolPool.OtherFieldsStart Then
                        iSortStart = aItems.Size
                    End If
                    
                    aItems.Add .ArrayTable.FieldName(lIndex) & vbTab _
                            & strID & vbTab & strPicture
                End If
            End If
        Next
    End With
    If iSortStart > 0 Then
        aItems.Sort eGdSort_IgnoreCase, iSortStart
    End If

    For lIndex = 0 To aItems.Size - 1
        strItem = aItems(lIndex)
        cboSymbolGroup.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next


    If bSelExists Then
        cboSymbolGroup.ComboItems(strSelID).Selected = True
    Else
        cboSymbolGroup.ComboItems(1).Selected = True
    End If

    cboSymbolGroup.Refresh

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectInCombo
'' Description: Try to select the given id in the combo
'' Inputs:      ID of the group/filter to select
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectInCombo(ByVal strID As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Did we find the ID?
    
    With cboSymbolGroup
        .ComboItems(1).Selected = True
        bFound = False
        For lIndex = 1 To .ComboItems.Count
            If .ComboItems(lIndex).Key = strID Then
                bFound = True
                .ComboItems(lIndex).Selected = True
                Exit For
            End If
        Next lIndex
    End With
    
    SelectInCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExportGroup.SelectInCombo", eGDRaiseError_Raise
    
End Function

