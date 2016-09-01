VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPriceVolCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volume at Price Settings"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraBoxSettings 
      Height          =   1335
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmPriceVolCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPriceVolCfg.frx":003C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPriceVolCfg.frx":005C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkFillBox 
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1020
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
         Caption         =   "frmPriceVolCfg.frx":0078
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPriceVolCfg.frx":00D0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":00F0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboBoxThickness 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   570
         Width           =   795
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
         Tip             =   "frmPriceVolCfg.frx":010C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":012C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBoxCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   540
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
         Caption         =   "frmPriceVolCfg.frx":0148
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPriceVolCfg.frx":0174
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":0194
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBoxOk 
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   60
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
         Caption         =   "frmPriceVolCfg.frx":01B0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPriceVolCfg.frx":01D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":01F4
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdBoxColor 
         Height          =   375
         Left            =   1380
         TabIndex        =   17
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label24 
         Height          =   255
         Left            =   0
         Top             =   600
         Width           =   1245
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPriceVolCfg.frx":0210
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPriceVolCfg.frx":0242
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":0262
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label23 
         Height          =   255
         Left            =   -15
         Top             =   120
         Width           =   1245
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPriceVolCfg.frx":027E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPriceVolCfg.frx":02A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":02C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFrmSettings 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   100
      Width           =   3855
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmPriceVolCfg.frx":02E4
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPriceVolCfg.frx":0320
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPriceVolCfg.frx":0340
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtBlankRows 
         Height          =   315
         Left            =   2880
         TabIndex        =   18
         Top             =   960
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPriceVolCfg.frx":035C
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
         Tip             =   "frmPriceVolCfg.frx":037C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":039C
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   660
         TabIndex        =   19
         Top             =   4515
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
         Caption         =   "frmPriceVolCfg.frx":03B8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPriceVolCfg.frx":03DC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":03FC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   1950
         TabIndex        =   20
         Top             =   4515
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
         Caption         =   "frmPriceVolCfg.frx":0418
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPriceVolCfg.frx":0444
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":0464
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraAuctionBar 
         Height          =   1920
         Left            =   0
         TabIndex        =   8
         Top             =   2505
         Width           =   3765
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPriceVolCfg.frx":0480
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmPriceVolCfg.frx":04C0
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":04E0
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtTriangleWidth 
            Height          =   315
            Left            =   1635
            TabIndex        =   10
            Top             =   660
            Width           =   660
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmPriceVolCfg.frx":04FC
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
            Tip             =   "frmPriceVolCfg.frx":0526
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":0546
         End
         Begin HexUniControls.ctlUniTextBoxXP txtBarWidth 
            Height          =   315
            Left            =   1635
            TabIndex        =   9
            Top             =   300
            Width           =   660
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmPriceVolCfg.frx":0562
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
            Tip             =   "frmPriceVolCfg.frx":058C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":05AC
         End
         Begin gdOCX.gdSelectColor gdUnfairHigh 
            Height          =   375
            Left            =   1025
            TabIndex        =   11
            Top             =   1425
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdMode 
            Height          =   375
            Left            =   135
            TabIndex        =   12
            Top             =   1425
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdUnfairLow 
            Height          =   375
            Left            =   1915
            TabIndex        =   13
            Top             =   1425
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdTriangle 
            Height          =   375
            Left            =   2805
            TabIndex        =   14
            Top             =   1425
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label1 
            Height          =   255
            Left            =   360
            Top             =   330
            Width           =   1200
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmPriceVolCfg.frx":05C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":05FC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":061C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblTriWdPix 
            Height          =   255
            Left            =   2385
            Top             =   690
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
            Caption         =   "frmPriceVolCfg.frx":0638
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":0664
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":0684
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblBarWdPix 
            Height          =   255
            Left            =   2385
            Top             =   330
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
            Caption         =   "frmPriceVolCfg.frx":06A0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":06CC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":06EC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label3 
            Height          =   255
            Left            =   360
            Top             =   690
            Width           =   1200
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmPriceVolCfg.frx":0708
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":0746
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":0766
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label2 
            Height          =   255
            Left            =   1005
            Top             =   1170
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
            Caption         =   "frmPriceVolCfg.frx":0782
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":07B8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":07D8
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label5 
            Height          =   255
            Left            =   90
            Top             =   1170
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
            Caption         =   "frmPriceVolCfg.frx":07F4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":081C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":083C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label6 
            Height          =   255
            Left            =   1905
            Top             =   1170
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
            Caption         =   "frmPriceVolCfg.frx":0858
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":088C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":08AC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label4 
            Height          =   255
            Left            =   2790
            Top             =   1170
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
            Caption         =   "frmPriceVolCfg.frx":08C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":08F8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":0918
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   1410
         Width           =   3765
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPriceVolCfg.frx":0934
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmPriceVolCfg.frx":096E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":098E
         RightToLeft     =   0   'False
         Begin gdOCX.gdSelectColor gdAsk 
            Height          =   375
            Left            =   1485
            TabIndex        =   2
            Top             =   525
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdBid 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   525
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin gdOCX.gdSelectColor gdHistogram 
            Height          =   375
            Left            =   2730
            TabIndex        =   4
            Top             =   525
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   661
            CustomColor     =   255
         End
         Begin HexUniControls.ctlUniLabelXP Label9 
            Height          =   255
            Left            =   2730
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
            Caption         =   "frmPriceVolCfg.frx":09AA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":09DC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":09FC
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label10 
            Height          =   255
            Left            =   180
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
            Caption         =   "frmPriceVolCfg.frx":0A18
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":0A40
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":0A60
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP Label11 
            Height          =   255
            Left            =   1470
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
            Caption         =   "frmPriceVolCfg.frx":0A7C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmPriceVolCfg.frx":0AA2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmPriceVolCfg.frx":0AC2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin gdOCX.gdSelectColor gdMean 
         Height          =   375
         Left            =   1980
         TabIndex        =   21
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdValue 
         Height          =   375
         Left            =   1980
         TabIndex        =   22
         Top             =   420
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label14 
         Height          =   255
         Left            =   120
         Top             =   990
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
         Caption         =   "frmPriceVolCfg.frx":0ADE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPriceVolCfg.frx":0B3C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":0B5C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label7 
         Height          =   255
         Left            =   585
         Top             =   60
         Width           =   1245
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPriceVolCfg.frx":0B78
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPriceVolCfg.frx":0BA2
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":0BC2
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   255
         Left            =   600
         Top             =   480
         Width           =   1245
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmPriceVolCfg.frx":0BDE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPriceVolCfg.frx":0C14
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPriceVolCfg.frx":0C34
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmPriceVolCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kfrmWidth = 4155
Private Const kfrmSettingsHt = 5485
Private Const kfrmBoxHt = 1985

Private Type mPrivate
    frm As Form
    
    iPixTri As Long
    iPixBar As Long
    iPixMax As Long
    
    iAsk As Long                'colors
    IbID As Long
    iHist As Long
    iMean As Long
    iMode As Long
    iTriangle As Long
    iUnHigh As Long
    iUnLow As Long
    iValue As Long
    iBlankRows As Long
    
    iFill As Long
    
    bChanged As Boolean
End Type

Private m As mPrivate

Public Function ShowFormSettings(frmCaller As frmPriceVol) As Boolean

    If frmCaller Is Nothing Then Exit Function  'precautionary
    
    Set m.frm = frmCaller
    
    Me.Width = kfrmWidth
    Me.Height = kfrmSettingsHt
    Me.Icon = Picture16(ToolbarIcon("ID_VolumeAtPrice"))
    
    fraFrmSettings.Visible = True
    fraFrmSettings.Enabled = True
    
    fraBoxSettings.Visible = False
    fraBoxSettings.Visible = False
    
    InitFrmControls

    m.bChanged = False
    
    CenterTheForm Me
    ShowForm Me, eForm_Modal

    ShowFormSettings = m.bChanged
End Function

Public Function ShowBoxSettings(frmCaller As Form, ByVal iColor&, ByVal iThickness&, ByVal iFill&) As Boolean

    Dim iBoxPix&

    If frmCaller Is Nothing Then Exit Function  'precautionary
    
    Set m.frm = frmCaller
    
    Me.Width = kfrmWidth
    Me.Height = kfrmBoxHt
    
    If TypeOf frmCaller Is frmPriceVol Then
        Me.Caption = "Volume at Price"
        Me.Icon = Picture16(ToolbarIcon("ID_VolumeAtPrice"))
    Else
        Me.Caption = "Bid/Ask Directional Analysis"
        Me.Icon = Picture16("kBlank")
    End If
    
    fraFrmSettings.Visible = False
    fraFrmSettings.Enabled = False
    
    fraBoxSettings.Visible = True
    fraBoxSettings.Visible = True

    If iThickness > 0 And iThickness < 6 Then
        iBoxPix = iThickness
    Else
        iBoxPix = 1
    End If
    
    cboBoxThickness.ListIndex = iBoxPix - 1
    gdBoxColor.Color = iColor
    chkFillBox.Value = iFill

    m.bChanged = False
    
    CenterTheForm Me
    ShowForm Me, eForm_Modal

    ShowBoxSettings = m.bChanged
End Function

Private Sub InitFrmControls()

    m.frm.GetColors m.iAsk, m.IbID, m.iHist, m.iMean, m.iMode, m.iTriangle, m.iUnHigh, m.iUnLow, m.iValue
    m.frm.GetPixWidth m.iPixBar, m.iPixTri, m.iPixMax
    
    'pixels
    txtBarWidth = m.iPixBar
    txtTriangleWidth = m.iPixTri
    lblBarWdPix.Caption = "pixels (max " & Str(m.iPixMax) & ")"
    
    'colors
    gdAsk.Color = m.IbID            'IOAMT uses the opposite of Genesis's
    gdBid.Color = m.iAsk
    gdHistogram.Color = m.iHist
    gdMean.Color = m.iMean
    gdMode.Color = m.iMode
    gdTriangle.Color = m.iTriangle
    gdUnfairHigh.Color = m.iUnHigh
    gdUnfairLow.Color = m.iUnLow
    gdValue.Color = m.iValue
    
    'blankrows
    txtBlankRows.Text = ValOfText(m.frm.BlankRows)
    
End Sub

Private Sub cmdBoxCancel_Click()
    m.bChanged = False
    Unload Me
End Sub

Private Sub cmdBoxOk_Click()

    m.bChanged = True
    m.frm.BoxSettings gdBoxColor.Color, Val(cboBoxThickness.Text), chkFillBox.Value
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim iTemp&
    
'pixels
    iTemp = ValOfText(txtBarWidth)
    If iTemp > 0 And iTemp <> m.iPixBar And iTemp <= m.iPixMax Then
        m.iPixBar = iTemp
        m.bChanged = True
    End If
    
    iTemp = ValOfText(txtTriangleWidth)
    If iTemp > 0 And iTemp <> m.iPixTri Then
        m.iPixTri = iTemp
        m.bChanged = True
    End If
    
'colors
    iTemp = gdBid.Color
    If iTemp <> m.iAsk Then
        m.iAsk = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdAsk.Color
    If iTemp <> m.IbID Then
        m.IbID = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdHistogram.Color
    If iTemp <> m.iHist Then
        m.iHist = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdMean.Color
    If iTemp <> m.iMean Then
        m.iMean = iTemp
        m.bChanged = True
    End If
        
    iTemp = gdMode.Color
    If iTemp <> m.iMode Then
        m.iMode = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdTriangle.Color
    If iTemp <> m.iTriangle Then
        m.iTriangle = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdUnfairHigh.Color
    If iTemp <> m.iUnHigh Then
        m.iUnHigh = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdUnfairLow.Color
    If iTemp <> m.iUnLow Then
        m.iUnLow = iTemp
        m.bChanged = True
    End If
    
    iTemp = gdValue.Color
    If iTemp <> m.iValue Then
        m.iValue = iTemp
        m.bChanged = True
    End If
        
    If m.bChanged Then
        m.frm.SetPixWidth m.iPixBar, m.iPixTri
        m.frm.SetColors m.iAsk, m.IbID, m.iHist, m.iMean, m.iMode, m.iTriangle, m.iUnHigh, m.iUnLow, m.iValue
    End If
    
    iTemp = ValOfText(txtBlankRows.Text)
    If iTemp > 0 And iTemp <> m.frm.BlankRows Then
        m.frm.BlankRows = iTemp
        m.bChanged = True
    End If
        
    Unload Me

End Sub

Private Sub Form_Load()
    fraBoxSettings.Move fraFrmSettings.Left, fraFrmSettings.Top
    
    g.Styler.StyleForm Me
    
    'RH populate list
    With cboBoxThickness
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    
End Sub

