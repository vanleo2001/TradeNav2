VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPatternProfit 
   Caption         =   "Patterns for Profit"
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin HexUniControls.ctlUniRadioXP optCorrSort 
      Height          =   255
      Left            =   11040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4545
      Visible         =   0   'False
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
      Caption         =   "frmPatternProfit.frx":0000
      Enabled         =   0   'False
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmPatternProfit.frx":0038
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":0058
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRadioXP optDaySort 
      Height          =   255
      Left            =   9990
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4545
      Visible         =   0   'False
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
      Caption         =   "frmPatternProfit.frx":0074
      Enabled         =   0   'False
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmPatternProfit.frx":009C
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":00BC
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chkDescending 
      Height          =   255
      Left            =   9750
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
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
      Caption         =   "frmPatternProfit.frx":00D8
      Enabled         =   0   'False
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmPatternProfit.frx":010E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":012E
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRadioXP optDateSort 
      Height          =   255
      Left            =   9000
      TabIndex        =   17
      Top             =   4545
      Visible         =   0   'False
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
      Caption         =   "frmPatternProfit.frx":014A
      Enabled         =   0   'False
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmPatternProfit.frx":0174
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":0194
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP cboPixPerBar 
      Height          =   315
      Left            =   7185
      TabIndex        =   18
      Top             =   1455
      Width           =   1575
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
      Tip             =   "frmPatternProfit.frx":01B0
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
      MouseIcon       =   "frmPatternProfit.frx":01D0
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
      MaxLength       =   0
      RightToLeft     =   0   'False
      LeftMargin      =   0
      RightMargin     =   0
      SelectOnFocus   =   0   'False
   End
   Begin VB.PictureBox pbPattern 
      BackColor       =   &H80000005&
      Height          =   1425
      Left            =   165
      ScaleHeight     =   1365
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   345
      Width           =   1575
      Begin VB.Line linDrag 
         Visible         =   0   'False
         X1              =   0
         X2              =   600
         Y1              =   0
         Y2              =   480
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7185
      TabIndex        =   15
      Top             =   840
      Width           =   1575
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
      Caption         =   "frmPatternProfit.frx":01EC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPatternProfit.frx":0218
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":0238
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdFind 
      Height          =   315
      Left            =   7185
      TabIndex        =   14
      Top             =   483
      Width           =   1575
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
      Caption         =   "frmPatternProfit.frx":0254
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPatternProfit.frx":0290
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":02B0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOptimize 
      Height          =   315
      Left            =   7185
      TabIndex        =   13
      Top             =   127
      Width           =   1575
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
      Caption         =   "frmPatternProfit.frx":02CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPatternProfit.frx":0306
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":0326
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraSource 
      Height          =   1740
      Left            =   1995
      TabIndex        =   0
      Tag             =   "1"
      Top             =   30
      Width           =   4785
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmPatternProfit.frx":0342
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPatternProfit.frx":03BE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":03DE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtMinCorr 
         Height          =   285
         Left            =   1740
         TabIndex        =   2
         Top             =   1080
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPatternProfit.frx":03FA
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
         Tip             =   "frmPatternProfit.frx":041E
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":043E
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSymbol 
         Height          =   255
         Left            =   2010
         TabIndex        =   11
         Top             =   300
         Width           =   300
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
         Caption         =   "frmPatternProfit.frx":045A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPatternProfit.frx":048C
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":04AC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboPeriod 
         Height          =   315
         Left            =   2460
         TabIndex        =   23
         Top             =   270
         Width           =   1440
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
         Tip             =   "frmPatternProfit.frx":04C8
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":04E8
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   315
         Left            =   825
         TabIndex        =   22
         Top             =   270
         Width           =   1500
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmPatternProfit.frx":0504
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
         Tip             =   "frmPatternProfit.frx":0538
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0558
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPtrnBars 
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   660
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPatternProfit.frx":0574
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
         Tip             =   "frmPatternProfit.frx":0596
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":05B6
      End
      Begin gdOCX.gdSelectDate gdDatePtrnTo 
         Height          =   345
         Left            =   2100
         TabIndex        =   9
         Top             =   660
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   609
      End
      Begin HexUniControls.ctlUniCheckXP chkClose 
         Height          =   255
         Left            =   3060
         TabIndex        =   6
         Top             =   1380
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
         Caption         =   "frmPatternProfit.frx":05D2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPatternProfit.frx":05FC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":061C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkLow 
         Height          =   255
         Left            =   2220
         TabIndex        =   5
         Top             =   1380
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
         Caption         =   "frmPatternProfit.frx":0638
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPatternProfit.frx":065E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":067E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkHigh 
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   1380
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
         Caption         =   "frmPatternProfit.frx":069A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPatternProfit.frx":06C2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":06E2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkOpen 
         Height          =   255
         Left            =   420
         TabIndex        =   3
         Top             =   1380
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
         Caption         =   "frmPatternProfit.frx":06FE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmPatternProfit.frx":0726
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0746
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label9 
         Height          =   255
         Left            =   180
         Top             =   720
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
         Caption         =   "frmPatternProfit.frx":0762
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":0792
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":07B2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblCorrelationFit 
         Height          =   225
         Left            =   165
         Top             =   1140
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
         Caption         =   "frmPatternProfit.frx":07CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":081C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":083C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPriceSymbol 
         Height          =   255
         Left            =   165
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
         Caption         =   "frmPatternProfit.frx":0858
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":0886
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":08A6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label29 
         Height          =   255
         Left            =   2220
         Top             =   1140
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
         Caption         =   "frmPatternProfit.frx":08C2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":0918
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0938
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label26 
         Height          =   255
         Left            =   1260
         Top             =   720
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
         Caption         =   "frmPatternProfit.frx":0954
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":0988
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":09A8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMatch 
      Height          =   6090
      Left            =   165
      TabIndex        =   12
      Top             =   1965
      Width           =   8595
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmPatternProfit.frx":09C4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPatternProfit.frx":09F2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":0A12
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraActual 
         Height          =   2760
         Left            =   3480
         TabIndex        =   19
         Top             =   390
         Width           =   4995
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPatternProfit.frx":0A2E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmPatternProfit.frx":0A60
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0A80
         RightToLeft     =   0   'False
         Begin VB.HScrollBar hsbActual 
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   2385
            Width           =   1695
         End
         Begin VB.PictureBox pbHit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2115
            Left            =   0
            ScaleHeight     =   2085
            ScaleWidth      =   4875
            TabIndex        =   21
            Top             =   195
            Width           =   4905
         End
         Begin HexUniControls.ctlUniLabelXP lblActual 
            Height          =   255
            Left            =   2040
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
            Caption         =   "frmPatternProfit.frx":0A9C
            BackColor       =   -2147483633
            ForeColor       =   12582912
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0AC8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0AE8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblData 
            Height          =   255
            Left            =   0
            Top             =   0
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
            Caption         =   "frmPatternProfit.frx":0B04
            BackColor       =   -2147483633
            ForeColor       =   8388608
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0B42
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0B62
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraComposite 
         Height          =   2760
         Left            =   3480
         TabIndex        =   24
         Top             =   3210
         Width           =   4995
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPatternProfit.frx":0B7E
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmPatternProfit.frx":0BB6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0BD6
         RightToLeft     =   0   'False
         Begin VB.HScrollBar hsbComposite 
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   2385
            Width           =   2460
         End
         Begin VB.PictureBox pbComposite 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2115
            Left            =   15
            ScaleHeight     =   2085
            ScaleWidth      =   4875
            TabIndex        =   30
            Top             =   195
            Width           =   4905
         End
         Begin HexUniControls.ctlUniLabelXP lblAvg 
            Height          =   255
            Left            =   15
            Top             =   15
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
            Caption         =   "frmPatternProfit.frx":0BF2
            BackColor       =   -2147483633
            ForeColor       =   8388736
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0C38
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0C58
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblForecast 
            Height          =   255
            Left            =   2055
            Top             =   15
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
            Caption         =   "frmPatternProfit.frx":0C74
            BackColor       =   -2147483633
            ForeColor       =   16711935
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0CA4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0CC4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraForecast 
         Height          =   315
         Left            =   3825
         TabIndex        =   26
         Top             =   0
         Width           =   4200
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPatternProfit.frx":0CE0
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmPatternProfit.frx":0D16
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0D36
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtStdDev 
            Height          =   285
            Left            =   2925
            TabIndex        =   28
            Top             =   15
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmPatternProfit.frx":0D52
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
            Tip             =   "frmPatternProfit.frx":0D74
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0D94
         End
         Begin HexUniControls.ctlUniTextBoxXP txtForecastBars 
            Height          =   285
            Left            =   1020
            TabIndex        =   27
            Top             =   15
            Width           =   375
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmPatternProfit.frx":0DB0
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
            Tip             =   "frmPatternProfit.frx":0DD2
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0DF2
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   255
            Left            =   75
            Top             =   30
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
            Caption         =   "frmPatternProfit.frx":0E0E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0E42
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0E62
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   1485
            Top             =   30
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
            Caption         =   "frmPatternProfit.frx":0E7E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0EC4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0EE4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   3390
            Top             =   30
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
            Caption         =   "frmPatternProfit.frx":0F00
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPatternProfit.frx":0F34
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPatternProfit.frx":0F54
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid fgHits 
         Height          =   5385
         Left            =   120
         TabIndex        =   25
         Top             =   390
         Width           =   2910
         _cx             =   5133
         _cy             =   9499
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
      Begin HexUniControls.ctlUniLabelXP lblOneArrow 
         Height          =   300
         Left            =   2985
         Top             =   1035
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
         Caption         =   "frmPatternProfit.frx":0F70
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":0F9C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":0FBC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAllArrow 
         Height          =   300
         Left            =   2985
         Top             =   2940
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
         Caption         =   "frmPatternProfit.frx":0FD8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":1004
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":1024
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAll 
         Height          =   300
         Left            =   2985
         Top             =   2775
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
         Caption         =   "frmPatternProfit.frx":1040
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":1068
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":1088
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblOne 
         Height          =   300
         Left            =   3030
         Top             =   840
         Width           =   450
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPatternProfit.frx":10A4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatternProfit.frx":10CC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatternProfit.frx":10EC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP Label12 
      Height          =   255
      Left            =   9000
      Top             =   4200
      Visible         =   0   'False
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
      Caption         =   "frmPatternProfit.frx":1108
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPatternProfit.frx":1138
      Style           =   0
      Enabled         =   0   'False
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":1158
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label8 
      Height          =   210
      Left            =   7185
      Top             =   1215
      Width           =   1440
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmPatternProfit.frx":1174
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPatternProfit.frx":11AA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":11CA
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP Label10 
      Height          =   255
      Left            =   165
      Top             =   562
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
      Caption         =   "frmPatternProfit.frx":11E6
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPatternProfit.frx":1224
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatternProfit.frx":1244
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "frmPatternProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_ADDR_REC = 16377      ' max # of items for
Private Const MAX_CORE_REC = MAX_ADDR_REC * 2

Private Const kHeight = 8505            'height with only one frame for displaying results
Private Const kHeightExt = 12450        'height with two frames for displaying results
Private Const kMinWidth = 9210

Private Const MATCH_OPEN = 1
Private Const MATCH_HIGH = 2
Private Const MATCH_LOW = 4
Private Const MATCH_CLOSE = 8

Private Enum eGDCols
    eGDCols_Use = 0
    eGDCols_Date
    eGDCols_Day
    eGDCols_CorrPercent
    eGDCols_Index
    eGDCols_DateDouble
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Globals for OLD DLL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim pfp_hits() As Long, pfp_corr() As Double, pfp_strength() As Double
'Search data
Dim gDate() As Long, gHourMin() As Long, gNumLoaded As Long
Dim gOpen() As Double, gHigh() As Double, gLow() As Double, gClose() As Double
Dim gVol() As Long, gOI() As Long, gTotVol() As Long, gTotOI() As Long
'PFP Pattern
Dim pDate() As Long, pHourMin() As Long, pNumLoaded As Long
Dim pOpen() As Double, pHigh() As Double, pLow() As Double, pClose() As Double
Dim pVol() As Long, pOI() As Long, pTotVol() As Long, pTotOI() As Long
'Composite data
Dim mDate() As Long, mHourMin() As Long, mNumLoaded As Double
Dim mOpen() As Double, mHigh() As Double, mLow() As Double, mClose() As Double
Dim mVol() As Long, mOI() As Long, mTotVol() As Long, mTotOI() As Long
'Bars
Dim PatrnCore As CoreBars
Dim SearchCore As CoreBars
Dim CompositeCore As CoreBars
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Enum enumGraphStyle
    eGrStyle_Line = 0
    eGrStyle_OHLC
    eGrStyle_Candles
End Enum

Private Type mPrivate
    Bars As cGdBars
    
    dPtrnDateFrom As Double
    dPtrnDateTo As Double
        
    nSymbolID As Long
    strSymbol As String
   
    nPtrnLen As Long
    nMaxBars As Long
    nLowestCorr As Long
    nMinHits As Long
    
    nValTxtPtrnBars As Long
    nValTxtForecast As Long
    nValTxtMinCorr As Long
    nValTxtStdDev As Long
    
    nPixPerBar As Long
    
    bOptimize As Boolean
End Type

Private m As mPrivate

Public Sub ShowMe()
On Error GoTo ErrSection:

    Dim Chart As cChart
    
    If Not ActiveChart Is Nothing Then Set Chart = ActiveChart.Chart
    
    If Not Chart Is Nothing Then
        If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then
            InfBox kSeasonalUnavail, "I", "Ok", "Pattern for Profit"
            Exit Sub
        End If
    End If
    
    LoadSettings
    InitControls
        
    mNumLoaded = 0
    If Chart Is Nothing Then
        cmdSymbol_Click
    Else
        m.strSymbol = ActiveChart.Chart.Symbol
        m.nSymbolID = ActiveChart.Chart.SymbolID
        txtSymbol.Text = m.strSymbol
    End If

    CenterTheForm Me
    ShowForm Me, , frmMain

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.ShowMe"

End Sub

Private Sub cboPeriod_Click()
On Error GoTo ErrSection:

    ClearControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.cboPeriod_Click"

End Sub

Private Sub cboPixPerBar_Click()
On Error GoTo ErrSection:

    Static bFirstTime As Boolean
    
    Dim iVal&, strText$
    
    If Not bFirstTime Then
        bFirstTime = True
        Exit Sub
    End If
    
    strText = cboPixPerBar.Text
    If strText = "Default" Then
        m.nPixPerBar = -1
        GraphHit
    Else
        strText = Parse(strText, " ", 1)
        iVal = ValOfText(strText)
        
        If iVal > 0 And iVal <= 50 Then
            m.nPixPerBar = iVal
            GraphHit
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.cboPixPerBar_Click"

End Sub

Private Sub chkClose_Click()
On Error GoTo ErrSection:

    ValidateCheckBox chkClose

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.chkClose_Click"

End Sub

Private Sub chkHigh_Click()
On Error GoTo ErrSection:

    ValidateCheckBox chkHigh

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.chkHigh_Click"

End Sub

Private Sub chkLow_Click()
On Error GoTo ErrSection:

    ValidateCheckBox chkLow

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.chkLow_Click"

End Sub

Private Sub chkOpen_Click()
On Error GoTo ErrSection:

    ValidateCheckBox chkOpen

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.chkOpen_Click"

End Sub

Private Sub ValidateCheckBox(chkBox As ctlUniCheckXP) 'RH was Checkbox
On Error GoTo ErrSection:
    
    Dim match_type%

    If chkBox.Caption = "Open" Or chkBox.Caption = "High" Or chkBox.Caption = "Low" Or chkBox.Caption = "Close" Then
        ClearControls
        
        match_type = 0
        If chkOpen Then match_type = match_type Or MATCH_OPEN
        If chkHigh Then match_type = match_type Or MATCH_HIGH
        If chkLow Then match_type = match_type Or MATCH_LOW
        If chkClose Then match_type = match_type Or MATCH_CLOSE
        
        If match_type = 0 Then chkBox.Value = vbChecked
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.ValidateCheckBox"

End Sub

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.cmdClose_Click"

End Sub

Private Function ForecastRank() As Double
On Error GoTo ErrSection:

    Dim b#, F#, i&, p&, l#, h#, max_f#

    max_f = -99999
    p = Val(txtPtrnBars)
    b = mClose(p) 'base
    For i = p + 1 To p + 3
        l = (mClose(i) - pfp_strength(i)) - b
        h = (mClose(i) + pfp_strength(i)) - b
        If Abs(l) < Abs(h) Then
            F = Abs(l)
        Else
            F = Abs(h)
        End If
        If h * l < 0 Then F = -F
        If F > max_f Then
            max_f = F
            'StatusMsg Str(i - p) + NumStr(f, 6, 2) + NumStr(l, 6, 2) + NumStr(h, 6, 2)
        End If
    Next i

    ForecastRank = max_f

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfit.ForecastRank"

End Function

Private Sub GetMatches()
On Error GoTo ErrSection:

    Dim rc&, Num&, temp$, i&, min_corr&, d, j&, p#
    Dim max_hits&, match_type&, fcast_len&, ptrn_len&
    Dim num_rules&, num_comp&, max_bars&
    Dim best_bars&, best_rank#, best_corr&, min_hits&, lowest_corr&

    Dim filtered&(), bResetIbars As Boolean
    Dim fh&, fName$, warning&

    Dim frame As Control
    Dim iRedrawSave As Long

    Set frame = fraMatch
    If Not FileExist(g.strAppPath & "\G32_PFP.dll") Then
        If Not FileExist("C:\development\GEN_32\G32_PFP\Debug\G32_PFP.dll") Then
            InfBox "File not found: G32_PFP.dll", "I"
            GoTo ErrExit
        End If
    End If

    Disable cmdFind
    Disable cmdOptimize
    Disable cmdClose
        
    best_bars = 0
    best_rank = -99999
    min_corr = Int(Abs(Val(txtMinCorr)))
    
    max_bars = Val(txtPtrnBars)
    ptrn_len = Val(txtPtrnBars)
    fcast_len = Val(txtForecastBars)
    
    If m.bOptimize Then
        ptrn_len = m.nPtrnLen
        If ptrn_len < 2 Then ptrn_len = 2
        max_bars = m.nMaxBars
        If max_bars < ptrn_len Then max_bars = ptrn_len
        min_hits = m.nMinHits
        lowest_corr = m.nLowestCorr
        If lowest_corr < 1 Or lowest_corr > 100 Then lowest_corr = 80
        min_corr = 98
    End If
    If ptrn_len < 1 Then
        InfBox "i=[] ; Invalid pattern length!"
        GoTo ErrExit
    End If
    If fcast_len < 1 Then
        InfBox "i=[] ; Invalid forecast length!"
        GoTo ErrExit
    End If

    LoadPatternCore
    If pNumLoaded <= 0 Then GoTo ErrExit
        
    LoadSearchCore
    If gNumLoaded <= 0 Then GoTo ErrExit
    
    max_hits = 1000
    ReDim pfp_hits&(max_hits), pfp_corr#(max_hits)
    ReDim filtered(gNumLoaded + 10) As Long

    match_type = 0
    If chkOpen Then match_type = match_type Or MATCH_OPEN
    If chkHigh Then match_type = match_type Or MATCH_HIGH
    If chkLow Then match_type = match_type Or MATCH_LOW
    If chkClose Then match_type = match_type Or MATCH_CLOSE

Do While min_corr <> 0
    txtPtrnBars.Text = Trim(Str(ptrn_len))
    txtMinCorr.Text = Trim(Str(min_corr))
        
    DoEvents
        
    LoadPatternCore
    If pNumLoaded <= 0 Then
        Exit Do
    End If

    GraphPattern eGrStyle_OHLC

    DoEvents
    
    rc = PFP_CorrelationMatches2(PatrnCore, _
                                 SearchCore, _
                                 match_type, _
                                 filtered(0), _
                                 min_corr, _
                                 max_hits, _
                                 pfp_hits(0), _
                                 pfp_corr(0))
    If m.bOptimize Then
        If pfp_hits(0) < min_hits Then
            min_corr = min_corr - 2
            rc = 0
        End If
    End If

    If rc > 0 Then
        mNumLoaded = ptrn_len + fcast_len
        LoadCompCore mNumLoaded + 10
        ReDim pfp_strength#(mNumLoaded + 10)
        num_comp = pfp_hits(0)
        'If chkExcludeLast <> 0 And num_comp > 0 Then num_comp = num_comp - 1
        rc = PFP_BuildComposite2(SearchCore, _
                                 match_type, _
                                 ptrn_len, _
                                 fcast_len, _
                                 num_comp, _
                                 pfp_hits(0), _
                                 CompositeCore, _
                                 pfp_strength(0))
        If m.bOptimize Then
            p = ForecastRank()
            If p > best_rank Then
                best_rank = p
                best_bars = ptrn_len
                best_corr = min_corr
            End If
            min_corr = -min_corr
        Else
            fName = "pfp.chk"
            KillFile fName
            
            fh = FreeFile
            If fh Then Open fName For Output As #fh
            If fh Then Print #fh, pNumLoaded, mNumLoaded
            For i = 1 To pNumLoaded
                If fh Then Print #fh, pDate(i), pOpen(i), pHigh(i), pLow(i), pClose(i)
            Next
            For i = 1 To mNumLoaded
                If fh Then Print #fh, i, mOpen(i), mHigh(i), mLow(i), mClose(i), pfp_strength(i)
                If mOpen(i) > mHigh(i) Or mLow(i) > mHigh(i) Or mClose(i) > mHigh(i) Then
                    warning = i
                End If
                If mOpen(i) < mLow(i) Or mHigh(i) < mLow(i) Or mClose(i) < mLow(i) Then
                    warning = i
                End If
            Next
            If fh Then
                Close #fh
                If warning Then InfBox "i=! ; h=Warning ; OHLC anomaly at" + Str(warning)
            End If
            Exit Do
        End If
    End If

    If Not m.bOptimize Then
        Exit Do
    ElseIf min_corr < lowest_corr Then
        min_corr = 100 'Abs(min_corr) + 2
        If min_corr > 98 Then min_corr = 98
        ptrn_len = ptrn_len + 1
        If ptrn_len > max_bars Then
            m.bOptimize = False
            ptrn_len = best_bars
            min_corr = best_corr
        End If
    End If
Loop

    iRedrawSave = fgHits.Redraw
    fgHits.Redraw = flexRDNone
    
    fgHits.Rows = fgHits.FixedRows
    
    If min_corr = 0 Then pfp_hits(0) = 0
    
    If pfp_hits(0) >= max_hits - 2 Then
        InfBox "i=! ; Too many matches -- please provide a greater restriction."
        pfp_hits(0) = 0
        GoTo ErrExit
    End If
    
    frame.Caption = "Found: " & Str(pfp_hits(0))
        
    Dim strText$, dDate#, iBars&
    
    iBars = Val(txtPtrnBars) - 1
    
    Do While pfp_hits(pfp_hits(0)) + iBars > gNumLoaded
        iBars = iBars - 1
        bResetIbars = True
    Loop

    fgHits.Cols = eGDCols_DateDouble + 1
    fgHits.ColHidden(eGDCols_DateDouble) = True
    For i = pfp_hits(0) To 1 Step -1
        dDate = gDate(pfp_hits(i) + iBars)
        If bResetIbars Then
            iBars = Val(txtPtrnBars) - 1
            bResetIbars = False
        End If
        strText = DateFormat(dDate, MM_DD_YYYY) + " " & WeekdayName(dDate) + Chr(9) + Format(pfp_corr(i), "00%") + Chr(9) + Str(i)
        
        With fgHits
            .Rows = .Rows + 1
            .Cell(flexcpChecked, .Rows - 1, eGDCols_Use) = flexChecked
            .Cell(flexcpPictureAlignment, .Rows - 1, eGDCols_Use) = flexAlignCenterCenter
            '.TextMatrix(.Rows - 1, eGDCols_Date) = DateFormat(dDate, MM_DD_YY) + " " & WeekdayName(dDate)
            .TextMatrix(.Rows - 1, eGDCols_Date) = DateFormat(dDate, MM_DD_YYYY)
            .TextMatrix(.Rows - 1, eGDCols_Day) = WeekdayName(dDate)
            .TextMatrix(.Rows - 1, eGDCols_CorrPercent) = Format(pfp_corr(i), "00%")
            .TextMatrix(.Rows - 1, eGDCols_Index) = Str(i)
            .TextMatrix(.Rows - 1, eGDCols_DateDouble) = dDate
        End With
    Next
    
    If fgHits.Rows > fgHits.FixedRows Then
        fgHits.Col = eGDCols_CorrPercent            '04-16-2009: Request from LW per Chad
        fgHits.Sort = flexSortGenericDescending
        fgHits.Row = fgHits.FixedRows
        GraphHit
        fgHits.Select fgHits.Row, eGDCols_Date
        fgHits.SetFocus
    End If
    
ErrExit:
    ' TLB: need redraw turned on regardless of previous setting (due to "Goto ErrExit" above)
    fgHits.Redraw = flexRDBuffered 'iRedrawSave
    
    Enable cmdFind
    Enable cmdOptimize
    Enable cmdClose
    Exit Sub

ErrSection:
    Enable cmdClose
    RaiseError "frmPatternProfit.GetMatchesNew"

End Sub

Private Sub ClearControls()
On Error GoTo ErrSection:

    Dim iRedrawSave&

    pbPattern.Cls
    pbHit.Cls
    pbComposite.Cls
    
    With fgHits
        iRedrawSave = .Redraw
        .Rows = .FixedRows
        .Redraw = iRedrawSave
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.ClearControls"

End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrSection:
    
    Dim rc&, i&, strPeriodicity$, strMsg$

    ClearControls

    If m.nSymbolID = 0 Then
        InfBox "Please select a symbol.", "I"
        cmdSymbol_Click
    End If

    If m.nSymbolID = 0 Then Exit Sub

    Set m.Bars = New cGdBars
    
    strPeriodicity = cboPeriod.Text
    If Len(strPeriodicity) = 0 Then strPeriodicity = "Daily"
    
    SetBarProperties m.Bars, m.nSymbolID
    rc = DM_GetBars(m.Bars, m.strSymbol, strPeriodicity, , , , False)
    AdjustBarPrices
        
    If m.Bars.Size = 0 Then Exit Sub
    
    If CDbl(gdDatePtrnTo.Value) > m.Bars.SessionDate(m.Bars.Size - 1) Then
        InfBox gdDatePtrnTo.Value & " is invalid." & vbCrLf & "Data is available only up to " & DateFormat(m.Bars.SessionDate(m.Bars.Size - 1))
        Exit Sub
    End If
    
    rc = m.Bars.FindDateTime(gdDatePtrnTo)
    If rc < 0 Or rc > m.Bars.Size - 1 Then
        InfBox gdDatePtrnTo.Value & " could not be found in the data." & vbCrLf & "Please select a different date.", "I"
        Exit Sub
    End If
    
    m.dPtrnDateTo = m.Bars(eBARS_DateTime, rc)
        
    If ValidateParams() Then
        i = Val(txtPtrnBars)
            
        If rc - i >= 0 Then
            m.dPtrnDateFrom = m.Bars(eBARS_DateTime, rc - i + 1)
        Else
            m.dPtrnDateFrom = m.Bars(eBARS_DateTime, 0)
        End If
        
        If m.dPtrnDateFrom > 0 And m.dPtrnDateTo > 0 Then
            GetMatches
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.cmdFind_Click"

End Sub

Private Sub LoadCompCore(ByVal Num&)
On Error GoTo ErrSection:

    ReDim mDate(Num) As Long
    ReDim mHourMin(Num) As Long
    ReDim mOpen(Num) As Double
    ReDim mHigh(Num) As Double
    ReDim mLow(Num) As Double
    ReDim mClose(Num) As Double
    ReDim mVol(Num) As Long
    ReDim mOI(Num) As Long
    ReDim mTotVol(Num) As Long
    ReDim mTotOI(Num) As Long

    CompositeCore.jdate_ptr = GetAddress(mDate(0))
    CompositeCore.hourmin_ptr = GetAddress(mHourMin(0))
    CompositeCore.open_ptr = GetAddress(mOpen(0))
    CompositeCore.high_ptr = GetAddress(mHigh(0))
    CompositeCore.low_ptr = GetAddress(mLow(0))
    CompositeCore.close_ptr = GetAddress(mClose(0))
    CompositeCore.vol_ptr = GetAddress(mVol(0))
    CompositeCore.oi_ptr = GetAddress(mOI(0))
    CompositeCore.tot_vol_ptr = GetAddress(mTotVol(0))
    CompositeCore.tot_oi_ptr = GetAddress(mTotOI(0))

    CompositeCore.max_bar = Num
    CompositeCore.first_bar = 1
    CompositeCore.last_bar = Num

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.LoadCompCore"

End Sub

Private Sub cmdOptimize_Click()
On Error GoTo ErrSection:
    
    frmPatternProfitOpt.ShowMe Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.cmdOptimize_Click"

End Sub

Private Sub cmdSymbol_Click()
On Error GoTo ErrSection:

    SelectSymbol

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmPatternProfit.cmdSymbol_Click"
End Sub

Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    ' if they click in the text box, make it act
    ' as if they clicked on the "..." command button
    MoveFocus cmdSymbol
    SelectSymbol

ErrExit:
    Exit Sub
ErrSection:
    RaiseError "frmPatternProfit.txtSymbol_Click"
End Sub

Private Sub SelectSymbol()
On Error GoTo ErrSection:

    Dim aStrings As cGdArray

    ClearControls
    
    Set aStrings = frmSymbolSelector.ShowMe(m.strSymbol, False, True, "Symbol for Pattern for Profit", True, , , , True)
    
    If aStrings.Size > 0 Then
        m.strSymbol = aStrings(0)
        m.nSymbolID = g.SymbolPool.SymbolIDforSymbol(m.strSymbol)
        txtSymbol.Text = m.strSymbol
    End If
    
ErrExit:
    Set aStrings = Nothing
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.SelectSymbol"
End Sub

Private Function LoadSearchCore() As Boolean
On Error GoTo ErrSection:

    Dim i&, j&, iSize&
        
    iSize = m.Bars.Size
    i = iSize + 1
        
    ReDim gDate(i) As Long
    ReDim gHourMin(i) As Long
    ReDim gOpen(i) As Double
    ReDim gHigh(i) As Double
    ReDim gLow(i) As Double
    ReDim gClose(i) As Double
    ReDim gVol(i) As Long
    ReDim gOI(i) As Long
    ReDim gTotVol(i) As Long
    ReDim gTotOI(i) As Long
    
    j = 1
    For i = 0 To iSize - 1
        'hourmin & vol info are not used by PFP dll
        'vol info arrays are currently arrays of "short"
        'can cause overflow for volume of bars > daily for certain symbols (e.g. AXP)
        gDate(j) = Int(m.Bars(eBARS_DateTime, i))
        gHourMin(j) = 0
        gOpen(j) = m.Bars(eBARS_Open, i)
        gHigh(j) = m.Bars(eBARS_High, i)
        gLow(j) = m.Bars(eBARS_Low, i)
        gClose(j) = m.Bars(eBARS_Close, i)
        gVol(j) = 0
        gOI(j) = 0
        gTotVol(j) = 0
        gTotOI(j) = 0
        j = j + 1
    Next
    j = j - 1

    SearchCore.jdate_ptr = GetAddress(gDate(0))
    SearchCore.hourmin_ptr = GetAddress(gHourMin(0))
    SearchCore.open_ptr = GetAddress(gOpen(0))
    SearchCore.high_ptr = GetAddress(gHigh(0))
    SearchCore.low_ptr = GetAddress(gLow(0))
    SearchCore.close_ptr = GetAddress(gClose(0))
    SearchCore.vol_ptr = GetAddress(gVol(0))
    SearchCore.oi_ptr = GetAddress(gOI(0))
    SearchCore.tot_vol_ptr = GetAddress(gTotVol(0))
    SearchCore.tot_oi_ptr = GetAddress(gTotOI(0))
    
    SearchCore.max_bar = j
    SearchCore.first_bar = 1
    SearchCore.last_bar = j
    gNumLoaded = j
    
    LoadSearchCore = True
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfit.LoadSearchCore"

End Function

Private Function LoadPatternCore() As Boolean
On Error GoTo ErrSection:

    Dim idxFirst&, idxLast&, iSize&, i&, j&
               
    If m.Bars Is Nothing Then Exit Function
               
    i = Val(txtPtrnBars.Text)
    If i <= 0 Then Exit Function
    
    idxLast = m.Bars.FindDateTime(m.dPtrnDateTo)
    m.dPtrnDateFrom = m.Bars(eBARS_DateTime, idxLast - i + 1)
    idxFirst = m.Bars.FindDateTime(m.dPtrnDateFrom)
    
    i = i + 2
    
    ReDim pDate(i) As Long
    ReDim pHourMin(i) As Long
    ReDim pOpen(i) As Double
    ReDim pHigh(i) As Double
    ReDim pLow(i) As Double
    ReDim pClose(i) As Double
    ReDim pVol(i) As Long
    ReDim pOI(i) As Long
    ReDim pTotVol(i) As Long
    ReDim pTotOI(i) As Long
        
    j = 1
    For i = idxFirst To idxLast
        'hourmin & vol info are not used by PFP dll
        'vol info arrays are currently arrays of "short"
        'can cause overflow for volume of bars > daily for certain symbols (e.g. AXP)
        pDate(j) = Int(m.Bars(eBARS_DateTime, i))
        pHourMin(j) = 0
        pOpen(j) = m.Bars(eBARS_Open, i)
        pHigh(j) = m.Bars(eBARS_High, i)
        pLow(j) = m.Bars(eBARS_Low, i)
        pClose(j) = m.Bars(eBARS_Close, i)
        pVol(j) = 0
        pOI(j) = 0
        pTotVol(j) = 0
        pTotOI(j) = 0
        j = j + 1
    Next
    j = j - 1
    
    PatrnCore.jdate_ptr = GetAddress(pDate(0))
    PatrnCore.hourmin_ptr = GetAddress(pHourMin(0))
    PatrnCore.open_ptr = GetAddress(pOpen(0))
    PatrnCore.high_ptr = GetAddress(pHigh(0))
    PatrnCore.low_ptr = GetAddress(pLow(0))
    PatrnCore.close_ptr = GetAddress(pClose(0))
    PatrnCore.vol_ptr = GetAddress(pVol(0))
    PatrnCore.oi_ptr = GetAddress(pOI(0))
    PatrnCore.tot_vol_ptr = GetAddress(pTotVol(0))
    PatrnCore.tot_oi_ptr = GetAddress(pTotOI(0))
    
    PatrnCore.max_bar = j
    PatrnCore.first_bar = 1
    PatrnCore.last_bar = j
    pNumLoaded = j
    
    LoadPatternCore = True
            
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfit.LoadPatternCore"

End Function

Private Sub GraphPattern(ByVal eStyle As enumGraphStyle)
On Error GoTo ErrSection:

    Dim i%, j%, Interval%, Num%, X%, Y%, Clr&
    Dim o%, h%, l%, c%, t%, b%, w%, last%, first%
    Dim highest#, lowest#, mult#
    
    Dim xcoord%(), ycoord%(), last_coord%

    If m.Bars Is Nothing Then Exit Sub

    Clr = vbBlue

    pbPattern.Cls
    pbPattern.AutoRedraw = True
    Num = Val(txtPtrnBars)
    If Num > m.Bars.Size Then Exit Sub

    ReDim xcoord%(Num, 3), ycoord%(Num, 3)
    For i = 0 To Num
        For j = 0 To 3
            xcoord(i, j) = 0
            ycoord(i, j) = 0
        Next
    Next

    Interval = Int(pbPattern.ScaleWidth / (Num + 1) / 15) * 15
    
    If Interval < 3 Then Interval = 3
    w = Interval / 3

    last = m.Bars.FindDateTime(m.dPtrnDateTo)
    first = last - Num + 1

    highest = -999999
    lowest = 999999
    For i = first To last
        If m.Bars(eBARS_High, i) > highest Then highest = m.Bars(eBARS_High, i)
        If m.Bars(eBARS_Low, i) < lowest Then lowest = m.Bars(eBARS_Low, i)
    Next
    
    If highest <= lowest Then Exit Sub
    
    t = pbPattern.ScaleHeight * 0.1
    b = pbPattern.ScaleHeight * 0.85
    mult = (b - t) / (highest - lowest)

    For i = first To last
        j = i - first + 1
        X = j * Interval
        h = b - (m.Bars(eBARS_High, i) - lowest) * mult
        l = b - (m.Bars(eBARS_Low, i) - lowest) * mult
        o = b - (m.Bars(eBARS_Open, i) - lowest) * mult
        c = b - (m.Bars(eBARS_Close, i) - lowest) * mult
        pbPattern.Line (X, h)-(X, l), Clr
        If eStyle = eGrStyle_OHLC Then
            pbPattern.Line (X - w, o)-(X + 15, o), Clr
            pbPattern.Line (X, c)-(X + w, c), Clr
        ElseIf eStyle = eGrStyle_Line Then
            pbPattern.Line (X - w, o)-(X + w, c), Clr, BF
        Else
            pbPattern.FillStyle = 0
            pbPattern.FillColor = pbPattern.BackColor
            pbPattern.Line (X - w, o)-(X + w, c), Clr, B
        End If
        pbPattern.CurrentX = X - 60
        pbPattern.CurrentY = b + 15
        pbPattern.ForeColor = QbClr("Gray") '+Black")
        
        If Num < 10 Then pbPattern.Print Trim(Str(Num - j + 1))

        ' save coordinates
        xcoord(j, 0) = X - w
        xcoord(j, 1) = X
        xcoord(j, 2) = X
        xcoord(j, 3) = X + w
        ycoord(j, 0) = o
        ycoord(j, 1) = h
        ycoord(j, 2) = l
        ycoord(j, 3) = c
        last_coord = j
    Next

    pbPattern.ZOrder
    pbPattern.Visible = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.GraphPattern"

End Sub

Private Sub fgHits_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    If Col = eGDCols_Use Then RebuildComposite

End Sub

Private Sub fgHits_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col <> eGDCols_Use Then Cancel = True

End Sub

Private Sub fgHits_BeforeSort(ByVal Col As Long, Order As Integer)
    
    If Order = flexSortGenericAscending Then
        chkDescending.Value = vbUnchecked
    Else
        chkDescending.Value = vbChecked
    End If
    
    If Col = eGDCols_Day Then
        optDaySort.Value = True
        SortHits
    ElseIf Col = eGDCols_Date Then
        optDateSort.Value = True
        SortHits
    ElseIf Col = eGDCols_CorrPercent Then
        optCorrSort.Value = True
    End If

End Sub

Private Sub fgHits_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Or vbKeyDown Then GraphHit

End Sub

Private Sub fgHits_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With fgHits
        If .Col = eGDCols_Date Then GraphHit
    End With

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText$
    
    g.Styler.StyleForm Me

    Me.Icon = Picture16(ToolbarIcon("ID_PatternProfit"), , True)
    cboPeriod.Clear
    cboPeriod.AddItem "Daily"
    cboPeriod.AddItem "Weekly"
    cboPeriod.AddItem "Monthly"
    cboPeriod.AddItem "Quarterly"
    cboPeriod.AddItem "Yearly"
    cboPeriod.ListIndex = 0
    
    'RH commented out fraActual.BorderStyle = 0
    'RH commented out fraComposite.BorderStyle = 0
            
    'initial defaults
    With hsbActual
        .Min = 0
        .Max = 100
        .LargeChange = 20
        .SmallChange = 1
        .Enabled = False
    End With
    
    With hsbComposite
        .Min = 0
        .Max = 100
        .LargeChange = 20
        .SmallChange = 1
        .Enabled = False
    End With
        
    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "P"
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.Form_Load"

End Sub

Private Sub Form_Resize()
On Error Resume Next:

    Dim iMinHeight&, dRatio#
        
    iMinHeight = kHeight
        
    If LimitFormSize(Me, kMinWidth, iMinHeight) Then Exit Sub

    fraMatch.Width = Me.Width - 450
    fraMatch.Height = Me.Height - fraSource.Height - 775
    
    fgHits.Height = fraMatch.Height - fraForecast.Height - 250
    
    fraActual.Width = fraMatch.Width - (fgHits.Width + lblOne.Width) - 375
    fraActual.Height = fgHits.Height / 2
    pbHit.Move 0, pbHit.Top, fraActual.Width, fraActual.Height - hsbActual.Height * 2
    hsbActual.Move 0, pbHit.Top + pbHit.Height + 3, pbHit.Width
    
    fraComposite.Move fraActual.Left, fraActual.Top + fraActual.Height + 50, fraActual.Width, fraActual.Height
    pbComposite.Move 0, pbComposite.Top, fraComposite.Width, fraComposite.Height - hsbComposite.Height * 2
    hsbComposite.Move 0, pbComposite.Top + pbComposite.Height + 3, pbComposite.Width
    
    lblOne.Move lblOne.Left, fraActual.Top + 500
    lblOneArrow.Move lblOneArrow.Left, lblOne.Top + lblOne.Height - 100
    
    lblAll.Move lblAll.Left, fraComposite.Top + 500
    lblAllArrow.Move lblAllArrow.Left, lblAll.Top + lblAll.Height - 100
    
    GraphHit

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    'if the close button is disabled then search is going on; disallow exit
    If cmdClose.Enabled Then
        SaveSettings
        Set m.Bars = Nothing
    Else
        InfBox "Please wait until pattern search is complete.", "I"
        Cancel = True
    End If

    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.Form_Unload"

End Sub

Private Function Normalized(ByVal Price#, ByVal norm_base#) As Double
On Error GoTo ErrSection:

    Normalized = (Price - norm_base) / norm_base * 100# + 100#

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfit.Normalized"

End Function

Private Sub GraphComposite()
On Error Resume Next

    Dim i&, Interval&, Num&, X&, Y&
    Dim o&, h&, l&, c&, t&, b&, w&, last&, first&
    Dim highest#, lowest#, mult#, p#
    Dim temp$, hit&, norm_base#, prev_x&, prev_h&, prev_l&
    Dim Clr&, data_clr&, comp_clr&, fcast_clr&, sdev_clr&, box_clr&

    Dim pic As Control
    Dim labelData As Control, labelForecast As Control, labelAvg As Control

    Set pic = pbComposite
    Set labelData = lblData
    Set labelAvg = lblAvg
    Set labelForecast = lblForecast

    data_clr = labelData.ForeColor
    comp_clr = labelAvg.ForeColor
    fcast_clr = labelForecast.ForeColor
    sdev_clr = QbClr("Gray")
    box_clr = QbClr("Gray")

    pic.Cls
    pic.AutoRedraw = True
    
    If mNumLoaded <= 0 Then Exit Sub
    
    Num = Val(txtPtrnBars) + Val(txtForecastBars)
        
    pic.Width = fraComposite.Width     'reset
    hsbComposite.Min = 0
    hsbComposite.Max = 0
    If m.nPixPerBar >= 1 Then
        Interval = m.nPixPerBar * Screen.TwipsPerPixelX
        If m.nPixPerBar <= 25 Then
            pic.Width = Interval * (Num + 3)
        Else
            pic.Width = Interval * (Num + 10)
        End If
        If pic.Width < fraComposite.Width Then
            pic.Width = fraComposite.Width
            hsbComposite.Enabled = False
        Else
            hsbComposite.Enabled = True
            hsbComposite.Max = ((pic.Width - fraComposite.Width) / pic.Width) * 100
        End If
    Else
        Interval = Int(pic.ScaleWidth / (Num + 1) / 15) * 15
        pic.Width = fraComposite.Width
        hsbComposite.Enabled = False
    End If
    
    If Interval < 3 Then Interval = 3
    w = Interval \ 3
    highest = -999999
    lowest = 999999
        
    ' set "100 base" = last close
    For i = 1 To mNumLoaded
        If mHigh(i) > highest Then highest = mHigh(i)
        If mLow(i) < lowest Then lowest = mLow(i)
        If i > Val(txtPtrnBars) Then
            p = mClose(i) + pfp_strength(i) * Val(txtStdDev)
            If p > highest Then highest = p
            p = mClose(i) - pfp_strength(i) * Val(txtStdDev)
            If p < lowest Then lowest = p
        End If
    Next
    If highest <= lowest Then Exit Sub
    t = pic.ScaleHeight * 0.1
    b = pic.ScaleHeight * 0.9
    mult = (b - t) / (highest - lowest)

    X = (Val(txtPtrnBars) + 0.5) * Interval
    pic.DrawStyle = 1
    pic.Line (0.5 * Interval, pic.ScaleHeight * 0.95)-(X, pic.ScaleHeight * 0.05), box_clr, B
    c = b - (100 - lowest) * mult
    pic.Line (0, c)-(pic.ScaleWidth, c), box_clr
    pic.DrawStyle = 0

    For i = 1 To mNumLoaded
        Clr = comp_clr
        If i > Val(txtPtrnBars) Then Clr = fcast_clr
        X = i * Interval + 15
        h = b - (mHigh(i) - lowest) * mult
        l = b - (mLow(i) - lowest) * mult
        o = b - (mOpen(i) - lowest) * mult
        c = b - (mClose(i) - lowest) * mult
        pic.Line (X, h)-(X, l), Clr
'        If Not candlesticks Then
            pic.Line (X - w, o)-(X + 15, o), Clr
            pic.Line (X, c)-(X + w, c), Clr
'        ElseIf o < c Then
'            pic.Line (X - w, o)-(X + w, c), Clr, BF
'        Else
'            pic.FillStyle = 0
'            pic.fillColor = pic.BackColor
'            pic.Line (X - w, o)-(X + w, c), Clr, B
'        End If
    Next

    If Val(txtStdDev) > 0 Then
        For i = 1 To mNumLoaded
            X = i * Interval + 15
            h = b - (mClose(i) - lowest + pfp_strength(i) * Val(txtStdDev)) * mult
            l = b - (mClose(i) - lowest - pfp_strength(i) * Val(txtStdDev)) * mult
            'If i > 1 Then
            If i > Val(txtPtrnBars) Then
                pic.Line (prev_x, prev_h)-(X, h), sdev_clr
                pic.Line (prev_x, prev_l)-(X, l), sdev_clr
            End If
            prev_x = X
            prev_h = h
            prev_l = l
        Next
    End If
    
    pic.Left = 0
    hsbComposite.Value = hsbComposite.Max
    
    pic.ZOrder
    pic.Visible = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.GraphComposite"

End Sub

Private Sub GraphHit()
On Error Resume Next

    Dim i&, Interval&, Num&, X&, Y&
    Dim o&, h&, l&, c&, t&, b&, w&, last&, first&
    Dim highest#, lowest#, mult#, p#
    Dim temp$, hit&, norm_base#, prev_x&, prev_h&, prev_l&
    Dim Clr&, data_clr&, comp_clr&, fcast_clr&, sdev_clr&, box_clr&
    Dim prev_c&

    Dim pic As Control
    Dim labelData As Control, lableActual As Control
    
    GraphComposite
        
    Set pic = pbHit
    Set labelData = lblData
    Set lableActual = lblActual
  
    If m.Bars Is Nothing Then Exit Sub
    If m.Bars.Size < 1 Then Exit Sub

    data_clr = labelData.ForeColor
    fcast_clr = lableActual.ForeColor
    box_clr = QbClr("Gray")

    pic.Cls
    pic.AutoRedraw = True
    Num = Val(txtPtrnBars) + Val(txtForecastBars)
    
    If fgHits.Row < fgHits.FixedRows Or fgHits.Row >= fgHits.Rows Then Exit Sub

    hit = Val(fgHits.TextMatrix(fgHits.Row, eGDCols_Index))
    'labelData = Parse(fgHits.TextMatrix(fgHits.Row, eGDCols_Date), " ", 1) + " data"
    labelData = DateFormat(fgHits.TextMatrix(fgHits.Row, eGDCols_DateDouble), MM_DD_YYYY) & " data"
    first = pfp_hits(hit)
    last = first + Num - 1
    If last > gNumLoaded Then last = gNumLoaded
    
    pic.Width = fraActual.Width     'reset
    hsbActual.Min = 0
    hsbActual.Max = 0
    If m.nPixPerBar >= 1 Then
        Interval = m.nPixPerBar * Screen.TwipsPerPixelX
        If m.nPixPerBar <= 25 Then
            pic.Width = Interval * ((last - first) + 3)
        Else
            pic.Width = Interval * ((last - first) + 10)
        End If
        If pic.Width < fraActual.Width Then
            pic.Width = fraActual.Width
            hsbActual.Enabled = False
        Else
            'pbComposite contains exact number of forecast bars requested; pbActual may contain less bars
            'by making width equal, we ensure bars will line up correctly
            pic.Width = pbComposite.Width
            hsbActual.Enabled = True
            hsbActual.Max = ((pic.Width - fraActual.Width) / pic.Width) * 100
        End If
    Else
        Interval = Int(pic.ScaleWidth / (Num + 1) / 15) * 15
        pic.Width = fraActual.Width
        hsbActual.Enabled = False
    End If
                    
    If Interval < 3 Then Interval = 3
    w = Interval \ 3
    highest = -999999
    lowest = 999999

    ' set "100 base" = last close
    If first + Val(txtPtrnBars) - 1 > gNumLoaded Then
        norm_base = gClose(gNumLoaded)
    Else
        norm_base = gClose(first + Val(txtPtrnBars) - 1)
    End If
    
    If norm_base <= 0 Then Exit Sub     'theoretically should never happen since we added AdjustBarPrices

    For i = first To last
        p = Normalized(gHigh(i), norm_base)
        If p > highest Then highest = p
        p = Normalized(gLow(i), norm_base)
        If p < lowest Then lowest = p
    Next
    If highest <= lowest Then Exit Sub
    t = pic.ScaleHeight * 0.1
    b = pic.ScaleHeight * 0.9
    mult = (b - t) / (highest - lowest)

    X = (Val(txtPtrnBars) + 0.5) * Interval
    pic.DrawStyle = 1
    pic.Line (0.5 * Interval, pic.ScaleHeight * 0.95)-(X, pic.ScaleHeight * 0.05), box_clr, B
    c = b - (100 - lowest) * mult
    pic.Line (0, c)-(pic.ScaleWidth, c), box_clr
    pic.DrawStyle = 0

    For i = first To last
        Clr = data_clr
        If i - first + 1 > Val(txtPtrnBars) Then Clr = fcast_clr
        X = (i - first + 1) * Interval
        h = b - (Normalized(gHigh(i), norm_base) - lowest) * mult
        l = b - (Normalized(gLow(i), norm_base) - lowest) * mult
        o = b - (Normalized(gOpen(i), norm_base) - lowest) * mult
        c = b - (Normalized(gClose(i), norm_base) - lowest) * mult
        pic.Line (X, h)-(X, l), Clr                     'vertical line
        'If Not candlesticks Then
            pic.Line (X - w, o)-(X + 15, o), Clr        'open line (horz)
            pic.Line (X, c)-(X + w, c), Clr             'close line (horz)
        'ElseIf o < c Then
        '    pic.Line (X - w, o)-(X + w, c), clr, BF
        'Else
        '    pic.FillStyle = 0
        '    pic.fillColor = pic.BackColor
        '    pic.Line (X - w, o)-(X + w, c), clr, B
        'End If
    Next

    pic.Left = 0
    hsbActual.Value = hsbActual.Max
    
    pic.ZOrder
    pic.Visible = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.GraphHit"

End Sub

Public Sub Optimize()
On Error GoTo ErrSection:
        
    m.nPtrnLen = Val(frmPatternProfitOpt.txtMinBars)
    m.nMaxBars = Val(frmPatternProfitOpt.txtMaxBars)
    m.nMinHits = Val(frmPatternProfitOpt.txtMinHits)
    m.nLowestCorr = Val(frmPatternProfitOpt.txtMinCorr)
    
    m.bOptimize = True
    cmdFind_Click
    m.bOptimize = False

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.Optimize"

End Sub

Private Sub SortHits()
On Error GoTo ErrSection:

    Dim i&, n&, ascending&, d, w&
    Dim temp$, chkBox$, strText$
    
    If fgHits.Rows <= fgHits.FixedRows Then Exit Sub

    With fgHits
        n = .Rows - 1
        ReDim hitstr(n)
        For i = .FixedRows To n
            temp = .TextMatrix(i, eGDCols_Date)
            w = 0
            d = DateOf(temp)
            If d > 0 Then w = Weekday(d)
            temp = Trim(Str(JulToLong(d, 1)))
            If optDaySort Then temp = Trim(Str(w)) + temp
            If .Cell(flexcpChecked, i, eGDCols_Use) = flexChecked Then
                chkBox = "1"
            Else
                chkBox = "0"
            End If
            strText = chkBox & vbTab & _
                    .TextMatrix(i, eGDCols_Date) & vbTab & _
                    .TextMatrix(i, eGDCols_Day) & vbTab & _
                    .TextMatrix(i, eGDCols_CorrPercent) & vbTab & _
                    .TextMatrix(i, eGDCols_Index) & vbTab & _
                    .TextMatrix(i, eGDCols_DateDouble)
            hitstr(i - 1) = temp + " |" + strText
        Next
    End With
    
    If chkDescending = 0 Then ascending = True
    Call QuickSortV(hitstr(), 0, n - 1, ascending)

    With fgHits
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To n - 1
            .AddItem Parse((hitstr(i)), "|", 2)
            .Cell(flexcpPictureAlignment, .Rows - 1, eGDCols_Use) = flexAlignCenterCenter
        Next
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.SortHits"

End Sub

Private Function ValidateParams() As Boolean
On Error GoTo ErrSection:

    Dim bOkay As Boolean

    m.nValTxtPtrnBars = Val(txtPtrnBars.Text)
    m.nValTxtMinCorr = Val(txtMinCorr.Text)
    m.nValTxtForecast = Val(txtForecastBars.Text)
    m.nValTxtStdDev = Val(txtStdDev.Text)
    
    bOkay = True
    
    'pattern length
    If m.nValTxtPtrnBars <= 0 Then
        InfBox "Pattern length cannot be less than 1.", "I"
        txtPtrnBars.Text = "1"
        m.nValTxtPtrnBars = 1
        bOkay = False
    ElseIf m.nValTxtPtrnBars > 260 Then
        InfBox "Pattern length cannot be greater than 260.", "I"
        txtPtrnBars.Text = "260"
        m.nValTxtPtrnBars = 260
        bOkay = False
    End If
    
    'forecast bars
    If m.nValTxtForecast <= 0 Then
        InfBox "Forecast bars cannot be less than 1.", "I"
        txtForecastBars.Text = "6"
        m.nValTxtForecast = 6
        bOkay = False
    ElseIf m.nValTxtForecast > 260 Then
        InfBox "Forecast bars cannot be greater than 260.", "I"
        txtForecastBars.Text = "260"
        m.nValTxtForecast = 260
        bOkay = False
    End If
    
    'min correlation
    If m.nValTxtMinCorr <= 0 Then
        InfBox "Percent correlation cannot be less than 1.", "I"
        txtMinCorr.Text = "90"
        m.nValTxtMinCorr = 90
        bOkay = False
    ElseIf m.nValTxtMinCorr > 100 Then
        InfBox "Percent correlation cannot be greater than 100.", "I"
        txtMinCorr.Text = "90"
        m.nValTxtPtrnBars = 90
        bOkay = False
    End If
    
    'standard deviation
    If m.nValTxtStdDev < 0 Then
        InfBox "Standard deviation cannot be less than zero.", "I"
        txtStdDev.Text = "1"
        m.nValTxtMinCorr = 1
        bOkay = False
    ElseIf m.nValTxtStdDev > 30 Then
        InfBox "Standard deviation cannot be greater than 30.", "I"
        txtStdDev.Text = "1"
        m.nValTxtMinCorr = 1
        bOkay = False
    End If
    
    ValidateParams = bOkay

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatternProfit.ValidateParams"

End Function

Private Sub hsbActual_Change()
    pbHit.Left = -(hsbActual.Value / 100) * pbHit.ScaleWidth
End Sub

Private Sub hsbActual_GotFocus()
    MoveFocus pbHit
End Sub

Private Sub hsbComposite_Change()
    pbComposite.Left = -(hsbComposite.Value / 100) * pbComposite.ScaleWidth
End Sub

Private Sub hsbComposite_GotFocus()
    MoveFocus pbComposite
End Sub

Private Sub mnuPrint_Click()
    frmPrintPreview.ShowMe "", Me, 0, 0.5, 0.5, 0.5, 0.5, True
End Sub

Private Sub pbHit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'04-30-2009: Started this, but will have to come back to it.
'For some reason, the frmPrintPreview does not work for picture box controls.

    Exit Sub
    
    If Button = vbRightButton Then
        mnuPrint.Visible = True
        mnuPrint.Enabled = True
    End If
    
    Me.PopupMenu mnuPopUp

End Sub

Private Sub txtForecastBars_LostFocus()
On Error GoTo ErrSection:

    If Val(txtForecastBars.Text) <> m.nValTxtForecast Then ClearControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.txtForecastBars_LostFocus"

End Sub

Private Sub txtMinCorr_LostFocus()
On Error GoTo ErrSection:

    If Val(txtMinCorr.Text) <> m.nValTxtMinCorr Then ClearControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.txtMinCorr_LostFocus"

End Sub

Private Sub txtPtrnBars_LostFocus()
On Error GoTo ErrSection:

    If Val(txtPtrnBars.Text) <> m.nValTxtPtrnBars Then ClearControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.txtPtrnBars_LostFocus"

End Sub

Private Sub txtStdDev_LostFocus()
On Error GoTo ErrSection:

    If Val(txtStdDev.Text) <> m.nValTxtStdDev Then ClearControls

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.txtStdDev_LostFocus"

End Sub

Private Sub InitGrid(fg As VSFlexGrid)

    If fg Is Nothing Then Exit Sub

    With fg
        .Redraw = flexRDNone
        SetupGrid fg, eGridMode_Grid
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarVertical
        .HighLight = flexHighlightNever
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .FixedCols = 0
        .FixedRows = 1
        .Rows = .FixedRows
        .Cols = 5
        'alignment
        .ColAlignment(eGDCols_Use) = flexAlignCenterCenter
        .ColAlignment(eGDCols_Date) = flexAlignLeftCenter
        .ColAlignment(eGDCols_Day) = flexAlignLeftCenter
        .ColAlignment(eGDCols_CorrPercent) = flexAlignRightCenter
        'column headers
        .TextMatrix(0, eGDCols_Use) = "Use"
        .TextMatrix(0, eGDCols_Date) = "Date"
        .TextMatrix(0, eGDCols_Day) = "Day"
        .TextMatrix(0, eGDCols_CorrPercent) = "Corr"
        'data type
        .ColDataType(eGDCols_Use) = flexDTBoolean
        .ColDataType(eGDCols_Day) = flexDTDate
        .ColDataType(eGDCols_Date) = flexDTDate
        .ColSort(eGDCols_Day) = flexSortNone
        .ColSort(eGDCols_Date) = flexSortNone
        .ColSort(eGDCols_Use) = flexSortNone
        .ColSort(eGDCols_Index) = flexSortNone
        'columns width
        .ColWidth(eGDCols_Use) = 500
        .ColWidth(eGDCols_Day) = 500
        .ColWidth(eGDCols_Date) = 1050 '900   '1400
        .ColWidth(eGDCols_CorrPercent) = 500
        .ColWidth(eGDCols_Index) = 100
        'hidden columns
        .ColHidden(eGDCols_Index) = True
        .Redraw = flexRDBuffered
    End With
    
End Sub


Private Sub InitControls()
On Error GoTo ErrSection:

    txtPtrnBars.Text = m.nValTxtPtrnBars
    txtMinCorr.Text = m.nValTxtMinCorr
    txtForecastBars.Text = m.nValTxtForecast
    txtStdDev.Text = m.nValTxtStdDev
    gdDatePtrnTo.Value = LastDailyDownload()
    
    Me.Height = kHeight
    
    cboPixPerBar.Clear
    cboPixPerBar.AddItem "Default"
    cboPixPerBar.AddItem "1 pixel"
    cboPixPerBar.AddItem "5 pixels"
    cboPixPerBar.AddItem "10 pixels"
    cboPixPerBar.AddItem "15 pixels"
    cboPixPerBar.AddItem "20 pixels"
    cboPixPerBar.AddItem "25 pixels"
    cboPixPerBar.AddItem "30 pixels"
    cboPixPerBar.AddItem "35 pixels"
    cboPixPerBar.AddItem "40 pixels"
    cboPixPerBar.AddItem "45 pixels"
    cboPixPerBar.AddItem "50 pixels"
    
    Select Case m.nPixPerBar
        Case 1:
            cboPixPerBar.ListIndex = 1
        Case 5:
            cboPixPerBar.ListIndex = 2
        Case 10:
            cboPixPerBar.ListIndex = 3
        Case 15:
            cboPixPerBar.ListIndex = 4
        Case 20:
            cboPixPerBar.ListIndex = 5
        Case 25:
            cboPixPerBar.ListIndex = 6
        Case 30:
            cboPixPerBar.ListIndex = 7
        Case 35:
            cboPixPerBar.ListIndex = 8
        Case 40:
            cboPixPerBar.ListIndex = 9
        Case 45:
            cboPixPerBar.ListIndex = 10
        Case 50:
            cboPixPerBar.ListIndex = 11
        Case Else
            cboPixPerBar.ListIndex = 0
    End Select
    
    pbHit.Cls
    pbComposite.Cls
    
    InitGrid fgHits

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.InitControls"

End Sub

Private Sub LoadSettings()
On Error GoTo ErrSection:

    m.nValTxtPtrnBars = GetIniFileProperty("PatternBars", 3, "PatternsForProfit", g.strIniFile)
    m.nValTxtMinCorr = GetIniFileProperty("MinCorrelation", 90, "PatternsForProfit", g.strIniFile)
    m.nValTxtForecast = GetIniFileProperty("ForecastBars", 6, "PatternsForProfit", g.strIniFile)
    m.nValTxtStdDev = GetIniFileProperty("StandardDev", 1, "PatternsForProfit", g.strIniFile)
    
    'optimization params
    m.nPtrnLen = GetIniFileProperty("PatternLen", 2, "PatternsForProfit", g.strIniFile)
    m.nMaxBars = GetIniFileProperty("MaxBars", 6, "PatternsForProfit", g.strIniFile)
    m.nLowestCorr = GetIniFileProperty("LowestCorr", 80, "PatternsForProfit", g.strIniFile)
    m.nMinHits = GetIniFileProperty("MinHits", 10, "PatternsForProfit", g.strIniFile)
    
    m.nPixPerBar = GetIniFileProperty("PixPerBar", -1, "PatternsForProfit", g.strIniFile)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.LoadSettings"

End Sub

Private Sub SaveSettings()
On Error GoTo ErrSection:

    SetIniFileProperty "PatternBars", m.nValTxtPtrnBars, "PatternsForProfit", g.strIniFile
    SetIniFileProperty "MinCorrelation", m.nValTxtMinCorr, "PatternsForProfit", g.strIniFile
    SetIniFileProperty "ForecastBars", m.nValTxtForecast, "PatternsForProfit", g.strIniFile
    SetIniFileProperty "StandardDev", m.nValTxtStdDev, "PatternsForProfit", g.strIniFile

    'optimization params
    SetIniFileProperty "PatternLen", m.nPtrnLen, "PatternsForProfit", g.strIniFile
    SetIniFileProperty "MaxBars", m.nMaxBars, "PatternsForProfit", g.strIniFile
    SetIniFileProperty "LowestCorr", m.nLowestCorr, "PatternsForProfit", g.strIniFile
    SetIniFileProperty "MinHits", m.nMinHits, "PatternsForProfit", g.strIniFile
    
    SetIniFileProperty "PixPerBar", m.nPixPerBar, "PatternsForProfit", g.strIniFile


ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.SaveSettings"

End Sub

Public Property Get OptimizePtrnLen() As Long
On Error GoTo ErrSection:

    OptimizePtrnLen = m.nPtrnLen

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmPatternProfit.OptimizePtrnLen"

End Property

Public Property Get OptimizeMaxBars() As Long
On Error GoTo ErrSection:

    OptimizeMaxBars = m.nMaxBars

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmPatternProfit.OptimizeMaxBars"

End Property

Public Property Get OptimizeLowestCorr() As Long
On Error GoTo ErrSection:

    OptimizeLowestCorr = m.nLowestCorr

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmPatternProfit.OptimizeLowestCorr"

End Property

Public Property Get OptimizeMinHits() As Long
On Error GoTo ErrSection:

    OptimizeMinHits = m.nMinHits

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmPatternProfit.OptimizeMinHits"

End Property

Private Sub RebuildComposite()
On Error GoTo ErrSection:

    Dim i&, j&, rc&, iCount&
    Dim ptrn_len&, fcast_len&, match_type&
    Dim filtered_hits() As Long

    If fgHits.Rows <= fgHits.FixedRows Then Exit Sub
    
    ReDim filtered_hits(pfp_hits(0)) As Long
    
    iCount = 1
    With fgHits
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, eGDCols_Use) = flexChecked Then
                j = Val(.TextMatrix(i, eGDCols_Index))
                filtered_hits(iCount) = pfp_hits(j)
                iCount = iCount + 1
            End If
        Next
    End With
    
    If iCount < 2 Then
        mNumLoaded = 0
        pbComposite.Cls
        Exit Sub
    End If
    
    filtered_hits(0) = iCount - 1
    
    ptrn_len = Val(txtPtrnBars)
    fcast_len = Val(txtForecastBars)
        
    mNumLoaded = ptrn_len + fcast_len
    LoadCompCore mNumLoaded + 10
    ReDim pfp_strength#(mNumLoaded + 10)
    iCount = filtered_hits(0)
        
    match_type = 0
    If chkOpen Then match_type = match_type Or MATCH_OPEN
    If chkHigh Then match_type = match_type Or MATCH_HIGH
    If chkLow Then match_type = match_type Or MATCH_LOW
    If chkClose Then match_type = match_type Or MATCH_CLOSE
    
    rc = PFP_BuildComposite2(SearchCore, _
                             match_type, _
                             ptrn_len, _
                             fcast_len, _
                             iCount, _
                             filtered_hits(0), _
                             CompositeCore, _
                             pfp_strength(0))
    
    If rc = 0 Then GraphComposite

ErrExit:
    Exit Sub

ErrSection:
    pbComposite.Cls
    mNumLoaded = 0
    RaiseError "frmPatternProfit.RebuildComposite"

End Sub

Public Sub PrintReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:


ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.PrintReport"
    
End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        .ZoomMode = zmPageWidth
        .RenderControl = pbHit.hWnd
        .EndDoc
    End With


ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.GenerateReport"

End Sub

Private Sub AdjustBarPrices()
On Error GoTo ErrSection:

    Dim i&, dLowest#, dHighest#, dAdjust#
    
    dLowest = gdMinValue(m.Bars.ArrayHandle(eBARS_Low), 0, m.Bars.Size)
    dHighest = gdMaxValue(m.Bars.ArrayHandle(eBARS_High), 0, m.Bars.Size)
    If dLowest < 0 Then
        ' TLB: if the lowest is negative, we are most likely dealing with a back-adjusted future
        ' so we should shift everything up so the lowest is about half the current highest
        ' in order to make all the percentage changes about right.
        If dHighest > 0 Then
            dAdjust = dHighest / 2 + Abs(dLowest)
        Else
            dAdjust = Abs(dLowest) * 2
        End If
        For i = 0 To m.Bars.Size - 1
            m.Bars(eBARS_Open, i) = m.Bars(eBARS_Open, i) + dAdjust
            m.Bars(eBARS_High, i) = m.Bars(eBARS_High, i) + dAdjust
            m.Bars(eBARS_Low, i) = m.Bars(eBARS_Low, i) + dAdjust
            m.Bars(eBARS_Close, i) = m.Bars(eBARS_Close, i) + dAdjust
        Next
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatternProfit.AdjustBarPrices"
End Sub


