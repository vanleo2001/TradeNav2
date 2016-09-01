VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmLoginFix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtEnabledSymbols 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5940
      Width           =   7395
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLoginFix.frx":0000
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   2
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmLoginFix.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":0040
   End
   Begin HexUniControls.ctlUniFrameWL fraAccount 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   900
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
      Caption         =   "frmLoginFix.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginFix.frx":0088
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":00A8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   0
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
         Tip             =   "frmLoginFix.frx":00C4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":00E4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAccount 
         Height          =   255
         Left            =   0
         Top             =   30
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
         Caption         =   "frmLoginFix.frx":0100
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":0132
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0152
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraPriceServer 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   4560
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
      Caption         =   "frmLoginFix.frx":016E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginFix.frx":01BE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":01DE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPriceTarget 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginFix.frx":01FA
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
         Tip             =   "frmLoginFix.frx":021A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":023A
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPriceIP 
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1875
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginFix.frx":0256
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
         Tip             =   "frmLoginFix.frx":0276
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0296
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPricePort 
         Height          =   285
         Left            =   2940
         TabIndex        =   20
         Top             =   240
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginFix.frx":02B2
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
         Tip             =   "frmLoginFix.frx":02D2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":02F2
      End
      Begin HexUniControls.ctlUniLabelXP lblPriceTarget 
         Height          =   195
         Left            =   180
         Top             =   660
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
         Caption         =   "frmLoginFix.frx":030E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":0344
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0364
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPriceIP 
         Height          =   195
         Left            =   180
         Top             =   285
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
         Caption         =   "frmLoginFix.frx":0380
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":03A8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":03C8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPricePort 
         Height          =   195
         Left            =   2520
         Top             =   285
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
         Caption         =   "frmLoginFix.frx":03E4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":0410
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0430
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraLoginInfo 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   180
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
      Caption         =   "frmLoginFix.frx":044C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginFix.frx":0478
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":0498
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboUserName 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   0
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
         Tip             =   "frmLoginFix.frx":04B4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":04D4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmLoginFix.frx":04F0
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
         PasswordChar    =   "*"
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmLoginFix.frx":0510
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0530
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddLogin 
         Height          =   315
         Left            =   2940
         TabIndex        =   3
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLoginFix.frx":054C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginFix.frx":0582
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":05A2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveLogin 
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   0
         Width           =   375
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
         Caption         =   "frmLoginFix.frx":05BE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginFix.frx":05FA
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":061A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUserName 
         Height          =   255
         Left            =   0
         Top             =   60
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
         Caption         =   "frmLoginFix.frx":0636
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":066C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":068C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPassword 
         Height          =   255
         Left            =   0
         Top             =   420
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
         Caption         =   "frmLoginFix.frx":06A8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":06DC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":06FC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraServerInfo 
      Height          =   2115
      Left            =   120
      TabIndex        =   15
      Top             =   2340
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
      Caption         =   "frmLoginFix.frx":0718
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginFix.frx":075C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":077C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraRefreshServer 
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   1680
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
         Caption         =   "frmLoginFix.frx":0798
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginFix.frx":07C4
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":07E4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtRefreshServer 
            Height          =   285
            Left            =   960
            TabIndex        =   24
            Top             =   0
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginFix.frx":0800
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
            Tip             =   "frmLoginFix.frx":0820
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0840
         End
         Begin HexUniControls.ctlUniLabelXP lblRefreshServer 
            Height          =   225
            Left            =   0
            Top             =   30
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
            Caption         =   "frmLoginFix.frx":085C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginFix.frx":088E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":08AE
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSubsender 
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   1320
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
         Caption         =   "frmLoginFix.frx":08CA
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginFix.frx":08F6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0916
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtSubsender 
            Height          =   285
            Left            =   960
            TabIndex        =   28
            Top             =   0
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginFix.frx":0932
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
            Tip             =   "frmLoginFix.frx":0952
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0972
         End
         Begin HexUniControls.ctlUniLabelXP lblSubsender 
            Height          =   225
            Left            =   0
            Top             =   30
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
            Caption         =   "frmLoginFix.frx":098E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginFix.frx":09C4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":09E4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraServer 
         Height          =   615
         Left            =   180
         TabIndex        =   19
         Top             =   600
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
         Caption         =   "frmLoginFix.frx":0A00
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginFix.frx":0A2C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0A4C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtTargetID 
            Height          =   285
            Left            =   960
            TabIndex        =   25
            Top             =   360
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginFix.frx":0A68
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
            Tip             =   "frmLoginFix.frx":0A88
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0AA8
         End
         Begin HexUniControls.ctlUniTextBoxXP txtServerIP 
            Height          =   285
            Left            =   300
            TabIndex        =   21
            Top             =   0
            Width           =   1875
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginFix.frx":0AC4
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
            Tip             =   "frmLoginFix.frx":0AE4
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0B04
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPort 
            Height          =   285
            Left            =   2760
            TabIndex        =   23
            Top             =   0
            Width           =   675
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginFix.frx":0B20
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
            Tip             =   "frmLoginFix.frx":0B40
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0B60
         End
         Begin HexUniControls.ctlUniLabelXP lblTargetID 
            Height          =   225
            Left            =   0
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
            Caption         =   "frmLoginFix.frx":0B7C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginFix.frx":0BB2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0BD2
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblServerIP 
            Height          =   225
            Left            =   0
            Top             =   30
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
            Caption         =   "frmLoginFix.frx":0BEE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginFix.frx":0C16
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0C36
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPort 
            Height          =   225
            Left            =   2340
            Top             =   30
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
            Caption         =   "frmLoginFix.frx":0C52
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginFix.frx":0C7E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0C9E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraServerMode 
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   240
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
         Caption         =   "frmLoginFix.frx":0CBA
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginFix.frx":0CE6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0D06
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optDemo 
            Height          =   220
            Left            =   780
            TabIndex        =   18
            Top             =   0
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
            Caption         =   "frmLoginFix.frx":0D22
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmLoginFix.frx":0D4C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0D6C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optLive 
            Height          =   220
            Left            =   0
            TabIndex        =   17
            Top             =   0
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
            Caption         =   "frmLoginFix.frx":0D88
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmLoginFix.frx":0DB2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmLoginFix.frx":0DD2
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   915
      Left            =   120
      TabIndex        =   10
      Top             =   1320
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
      Caption         =   "frmLoginFix.frx":0DEE
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginFix.frx":0E1A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":0E3A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowIP 
         Height          =   435
         Left            =   2460
         TabIndex        =   14
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
         Caption         =   "frmLoginFix.frx":0E56
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginFix.frx":0EA6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0EC6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1080
         TabIndex        =   13
         Top             =   480
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
         Caption         =   "frmLoginFix.frx":0EE2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginFix.frx":0F10
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0F30
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdLogin 
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   12
         Top             =   480
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
         Caption         =   "frmLoginFix.frx":0F4C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginFix.frx":0F78
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":0F98
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAgree 
         Height          =   435
         Left            =   0
         Top             =   0
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
         Caption         =   "frmLoginFix.frx":0FB4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginFix.frx":1074
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginFix.frx":1094
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   2895
      Left            =   4020
      TabIndex        =   27
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5106
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLoginFix.frx":10B0
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
      Tip             =   "frmLoginFix.frx":10D0
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":10F0
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
   Begin HexUniControls.ctlUniLabelXP lblEnabledSymbols 
      Height          =   195
      Left            =   120
      Top             =   5700
      Width           =   2475
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLoginFix.frx":110C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmLoginFix.frx":114C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginFix.frx":116C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmLoginFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLoginFix.cls
'' Description: Form to get login information for fix servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/20/2011   DAJ         Added controls for price server information
'' 06/06/2012   DAJ         Added server mode and subsender ID
'' 07/10/2012   DAJ         Only require subsender ID if it is necessary
'' 08/17/2012   DAJ         Added account and refresh server controls
'' 09/26/2012   DAJ         Handle new user name with different case than existing one
'' 01/09/2013   DAJ         Make sure to set the live/demo options when default user name combo
'' 03/11/2013   DAJ         Added enabled symbols box
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK or Cancel?
    strIniFile As String                ' INI file for the broker
    strConnectIni As String             ' INI file for connection information
    strBrokerName As String             ' Display name for the broker
    nBroker As eTT_AccountType          ' Current broker
    bShowAccounts As Boolean            ' Show the account information?
    bShowPrice As Boolean               ' Show the price server information?
    bShowSubsender As Boolean           ' Show the subsender information?
    bShowRefresh As Boolean             ' Show the refresh server information?
    
    strUserName As String               ' User name that the user chose
    strPassword As String               ' Password from the user
    strAccount As String                ' Account the user selected
    bIsLive As Boolean                  ' Is this a live server?
    strIP As String                     ' IP address to connect to
    strPort As String                   ' Port to connect to
    strTargetID As String               ' ID of the target computer
    strSubsenderID As String            ' Subsender ID
    strRefreshServer As String          ' Refresh server information
    strPriceIP As String                ' IP address to connect to for the price server
    strPricePort As String              ' Port to connect to for the price server
    strPriceTargetID As String          ' ID of the target computer for the price server
    astrEnabledSymbols As cGdArray      ' List of enabled symbols
End Type
Private m As mPrivate

Public Property Get UserName() As String
    UserName = m.strUserName
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property

Public Property Get Account() As String
    Account = m.strAccount
End Property

Public Property Get IsLive() As Boolean
    IsLive = m.bIsLive
End Property

Public Property Get IP() As String
    IP = m.strIP
End Property

Public Property Get Port() As String
    Port = m.strPort
End Property

Public Property Get TargetID() As String
    TargetID = m.strTargetID
End Property

Public Property Get SubsenderID() As String
    SubsenderID = m.strSubsenderID
End Property

Public Property Get RefreshServer() As String
    RefreshServer = m.strRefreshServer
End Property

Public Property Get PriceIP() As String
    PriceIP = m.strPriceIP
End Property

Public Property Get PricePort() As String
    PricePort = m.strPricePort
End Property

Public Property Get PriceTargetID() As String
    PriceTargetID = m.strPriceTargetID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, UserID, Are we switching?, Show IP?, Show Price Info?,
''              Show Server Mode?, Show Subsender?, Show Accounts?, Enabled Symbols
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Broker As cBroker, Optional ByVal strUserName As String = "", Optional ByVal bSwitching As Boolean = False, Optional ByVal bShowIP As Boolean = False, Optional ByVal bShowPrice As Boolean = False, Optional ByVal bShowServerMode As Boolean = False, Optional ByVal bShowSubsender As Boolean = False, Optional ByVal bShowAccounts As Boolean = False, Optional ByVal bShowRefresh As Boolean = False, Optional ByVal astrEnabledSymbols As cGdArray = Nothing) As Boolean
On Error GoTo ErrSection:
    
    Dim strLastUserName As String       ' Last user name logged into
    Dim strLastAccount As String        ' Last account logged into
    Dim lIndex As Long                  ' Index into a for loop

    m.bOK = False
    m.strIniFile = Broker.IniFile
    m.strConnectIni = Broker.ConnectIni
    m.strBrokerName = Broker.BrokerName
    m.bShowAccounts = bShowAccounts
    m.bShowPrice = bShowPrice
    m.bShowSubsender = bShowSubsender
    m.bShowRefresh = bShowRefresh
    m.nBroker = Broker.Broker
    Set m.astrEnabledSymbols = astrEnabledSymbols
    Caption = m.strBrokerName & " Login Information"
    
    If Len(m.strIniFile) > 0 Then
        strLastUserName = GetIniFileProperty("LastUserName", "", "User", m.strIniFile)
        FixLogins
        LoadCombo
        If cboUserName.ListCount > 0 Then
            If SetCombo(strUserName) = False Then
                If SetCombo(strLastUserName) = False Then
                    strLastUserName = GetIniFileProperty("UserName", "", "User", m.strIniFile)
                    If SetCombo(strLastUserName) = False Then
                        cboUserName.ListIndex = 0
                        Select Case cboUserName.ItemData(0)
                            Case 0:
                                optLive.Value = True
                            Case 1:
                                optDemo.Value = True
                        End Select
                    End If
                End If
            End If
        End If
        
        If (cboUserName.ListCount = 0) Or ((cboUserName.ListCount = 1) And (bSwitching = True)) Then
            AddLogin
        End If
        
        If (cboUserName.ListCount > 1) Or ((cboUserName.ListCount = 1) And (bSwitching = False)) Then
            If bShowAccounts Then
                strLastAccount = GetIniFileProperty("LastAccount", "", "User", m.strIniFile)
                LoadAccountsCombo
                
                If cboAccounts.ListCount > 0 Then
                    If SetCombo(strLastAccount) = False Then
                        cboAccounts.ListIndex = 0
                    End If
                Else
                    AddAccount
                End If
            End If
            
            If (cboAccounts.ListCount > 0) Or (bShowAccounts = False) Then
                fraAccount.Visible = bShowAccounts
                CheckBoxValue(chkShowIP) = bShowIP
                fraServerInfo.Visible = bShowIP
                ShowServerControls bShowServerMode, bShowSubsender, bShowPrice, bShowRefresh
                
                If m.astrEnabledSymbols Is Nothing Then
                    txtEnabledSymbols.Text = ""
                Else
                    For lIndex = m.astrEnabledSymbols.Size - 1 To 0 Step -1
                        If (Left(m.astrEnabledSymbols(lIndex), 2) = "O:") Or (Left(m.astrEnabledSymbols(lIndex), 2) = "S:") Then
                            m.astrEnabledSymbols.Remove lIndex
                        End If
                    Next lIndex
                    
                    If m.astrEnabledSymbols.Size = 0 Then
                        txtEnabledSymbols.Text = "None"
                    Else
                        txtEnabledSymbols.Text = m.astrEnabledSymbols.JoinFields(", ")
                    End If
                End If
                            
                MoveFocus txtPassword
                
                ShowForm Me, eForm_Modal, frmMain
                
                If m.bOK = True Then
                    SetServerOverrides
                    SetPriceServerOverrides
                    
                    m.strUserName = cboUserName.Text
                    m.strPassword = Trim(txtPassword.Text)
                    m.bIsLive = optLive.Value
                    m.strIP = Trim(txtServerIP.Text)
                    m.strPort = Trim(txtPort.Text)
                    m.strTargetID = Trim(txtTargetID.Text)
                    m.strSubsenderID = Trim(txtSubsender.Text)
                    m.strRefreshServer = Trim(txtRefreshServer.Text)
                    m.strPriceIP = Trim(txtPriceIP.Text)
                    m.strPricePort = Trim(txtPricePort.Text)
                    m.strPriceTargetID = Trim(txtPriceTarget.Text)
                    
                    If bShowAccounts Then
                        m.strAccount = cboAccounts.Text
                    Else
                        m.strAccount = ""
                    End If
                    
                    SetIniFileProperty "LastUserName", cboUserName.Text, "User", m.strIniFile
                    SetIniFileProperty "LastAccount", cboAccounts.Text, "User", m.strIniFile
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
    RaiseError "frmLoginFix.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboUserName_Click
'' Description: Handle the user changing user names
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboUserName_Click()
On Error GoTo ErrSection:

    If Visible Then
        If cboUserName.ListIndex > -1& Then
            If cboUserName.ItemData(cboUserName.ListIndex) = 0 Then
                optLive.Value = True
            ElseIf cboUserName.ItemData(cboUserName.ListIndex) = 1 Then
                optDemo.Value = True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.cboUserName_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowIP_Click
'' Description: Show/Hide the server information as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowIP_Click()
On Error GoTo ErrSection:

    fraServerInfo.Visible = CheckBoxValue(chkShowIP)
    fraPriceServer.Visible = CheckBoxValue(chkShowIP) And m.bShowPrice

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.chkShowIP_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddLogin_Click
'' Description: Allow the user to add a login
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddLogin_Click()
On Error GoTo ErrSection:

    AddLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.cmdAddLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the dialog without logging in
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
    RaiseError "frmLoginFix.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLogin_Click
'' Description: Verify the user information and pass back to Rithmic object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLogin_Click()
On Error GoTo ErrSection:

    If cboUserName.ListIndex < 0 Then
        MoveFocus cboUserName
        InfBox "Please enter in a User Name", "!", , "Login Error"
    ElseIf Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in a Password", "!", , "Login Error"
    ElseIf Len(Trim(txtServerIP.Text)) = 0 Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtServerIP
        InfBox "Please enter in an IP address for the server", "!", , "Login Error"
    ElseIf Len(Trim(txtPort.Text)) = 0 Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtPort
        InfBox "Please enter in a Port for the server", "!", , "Login Error"
    ElseIf Len(Trim(txtTargetID.Text)) = 0 Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtTargetID
        InfBox "Please enter in a Target Computer ID for the server", "!", , "Login Error"
    ElseIf (m.bShowSubsender = True) And (Len(Trim(txtSubsender.Text)) = 0) Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtSubsender
        InfBox "Please enter in a Subsender ID for the server", "!", , "Login Error"
    ElseIf (m.bShowRefresh = True) And (Len(Trim(txtRefreshServer.Text)) = 0) Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtRefreshServer
        InfBox "Please enter in a the refresh server information", "!", , "Login Error"
    ElseIf (m.bShowPrice = True) And (Len(Trim(txtPriceIP.Text)) = 0) Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtPriceIP
        InfBox "Please enter in an IP address for the price server", "!", , "Login Error"
    ElseIf (m.bShowPrice = True) And (Len(Trim(txtPricePort.Text)) = 0) Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtPricePort
        InfBox "Please enter in a Port for the price server", "!", , "Login Error"
    ElseIf (m.bShowPrice = True) And (Len(Trim(txtPriceTarget.Text)) = 0) Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtPriceTarget
        InfBox "Please enter in a Target Computer ID for the price server", "!", , "Login Error"
    Else
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveLogin_Click
'' Description: Allow the user to remove one or more logins
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveLogin_Click()
On Error GoTo ErrSection:

    RemoveLogin

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.cmdRemoveLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Make sure when the form is activated that password gets focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Do some initialization when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kBlank")
    
    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    cmdAddLogin.ToolTipText = "Add a user name"
    cmdRemoveLogin.ToolTipText = "Remove user name(s)"

    txtEnabledSymbols.Enabled = True
    txtEnabledSymbols.BackColor = cmdAddLogin.BackColor
    txtEnabledSymbols.Locked = True

    m.strUserName = ""
    m.strPassword = ""
    m.bIsLive = True
    m.strIP = ""
    m.strPort = ""
    m.strTargetID = ""
    m.strSubsenderID = ""
    m.strRefreshServer = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hit the X, let ShowMe unload the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDemo_Click
'' Description: Handle the server mode being changed to demo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDemo_Click()
On Error GoTo ErrSection:

    UpdateServerMode

    SetServerControls
    SetPriceServerControls
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.optDemo_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLive_Click
'' Description: Handle the server mode being changed to live
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLive_Click()
On Error GoTo ErrSection:

    UpdateServerMode

    SetServerControls
    SetPriceServerControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.optLive_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load the accounts combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strUserName As String           ' User Name already in the INI file
    Dim strIP As String                 ' IP address from the INI file
    
    cboUserName.Clear
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            If Len(astrLogins(lIndex)) > 0 Then
                cboUserName.AddItem Parse(astrLogins(lIndex), "|", 1)
                cboUserName.ItemData(cboUserName.NewIndex) = CLng(Val(Parse(astrLogins(lIndex), "|", 2)))
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.LoadCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Set the user name combo box to the given user name if possible
'' Inputs:      User Name
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetCombo(ByVal strUserName As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboUserName.ListCount > 0) And (Len(strUserName) > 0) Then
        For lIndex = 0 To cboUserName.ListCount - 1
            If UCase(cboUserName.List(lIndex)) = UCase(strUserName) Then
                bFound = True
                cboUserName.ListIndex = lIndex
                
                Select Case cboUserName.ItemData(lIndex)
                    Case 0:
                        optLive.Value = True
                    Case 1:
                        optDemo.Value = True
                End Select
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginFix.SetCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLogin
'' Description: Allow the user to give us a new user name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLogin()
On Error GoTo ErrSection:

    Dim strUserName As String           ' User name from the user
    Dim strNewLogin As String           ' New login to save to INI file
    Dim strLogins As String             ' Login string from the INI file
    Dim iLoginExists As Integer         ' Does the login exist?
    Dim bAddLogin As Boolean            ' Add the login to the combo?
    Dim strReturn As String             ' Return value from an InfBox
    Dim strExistingLogin As String      ' Existing login
    
    strUserName = InfBox("What is your " & m.strBrokerName & " user name?", "?", , m.strBrokerName & " User Name", , , , , , "string")
    If Len(strUserName) > 0 Then
        iLoginExists = LoginExists(strUserName, strExistingLogin)
        Select Case iLoginExists
            Case 0:
                bAddLogin = True
            
            Case 1:
                SetCombo strUserName
                bAddLogin = False
            
            Case 2:
                strReturn = InfBox("'" & strUserName & "' already exists in the list as '" & strExistingLogin & "'||Do you want to rename it or add the new one?|", "?", "+Rename|Add|-Cancel")
                Select Case strReturn
                    Case "R"
                        RenameLogin strExistingLogin, strUserName
                        SetCombo strUserName
                        bAddLogin = False
                    Case "A"
                        bAddLogin = True
                    Case "C"
                        bAddLogin = False
                End Select
                
        End Select
                
        If bAddLogin Then
            cboUserName.AddItem strUserName
            cboUserName.ItemData(cboUserName.NewIndex) = 0
            
            SetCombo strUserName
            MoveFocus txtPassword
            
            SaveLogins
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.AddLogin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveLogin
'' Description: Allow the user to remove login information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLogin()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As cGdArray          ' Array of login information
    Dim astrList As cGdArray            ' List to send to the delete form
    Dim astrToDelete As cGdArray        ' List of logins to delete
    Dim strSelected As String           ' Currently selected login
    Dim lIndex As Long                  ' Index into a for loop

    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        Set astrLogins = New cGdArray
        Set astrList = New cGdArray
        astrList.Create eGDARRAY_Strings
        
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            astrList.Add Parse(astrLogins(lIndex), "|", 1) & vbTab & Str(lIndex)
        Next lIndex
        
        strSelected = cboUserName.Text
        
        Set astrToDelete = frmDelete.ShowMe(astrList, strSelected)
        If Not astrToDelete Is Nothing Then
            For lIndex = astrToDelete.Size - 1 To 0 Step -1
                astrLogins.Remove CLng(Val(astrToDelete(lIndex)))
            Next lIndex
            
            SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
            
            LoadCombo
            If SetCombo(strSelected) = False Then
                If cboUserName.ListCount > 0 Then
                    cboUserName.ListIndex = 0
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.RemoveLogin"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerControls
'' Description: Set the server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerControls()
On Error GoTo ErrSection:

    If optLive.Value = True Then
        SetServerControlsLive
    ElseIf optDemo.Value = True Then
        SetServerControlsDemo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetServerControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerControlsLive
'' Description: Set the server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerControlsLive()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target ID to override
    Dim strSubsender As String          ' Subsender ID to override
    Dim strRefresh As String            ' Refresh server information

    txtServerIP.Text = GetIniFileProperty("IP", "", "Server", m.strConnectIni)
    txtPort.Text = GetIniFileProperty("Port", "", "Server", m.strConnectIni)
    txtTargetID.Text = GetIniFileProperty("Target", "", "Server", m.strConnectIni)
    txtSubsender.Text = GetIniFileProperty("Subsender", "", "Server", m.strConnectIni)
    txtRefreshServer.Text = GetIniFileProperty("Refresh", "", "Server", m.strConnectIni)
    
    strIP = GetIniFileProperty("IP", "", "Override", m.strIniFile)
    If Len(strIP) > 0 Then
        If strIP = txtServerIP.Text Then
            SetIniFileProperty "IP", "", "Override", m.strIniFile
        Else
            txtServerIP.Text = strIP
        End If
    End If

    strPort = GetIniFileProperty("Port", "", "Override", m.strIniFile)
    If Len(strPort) > 0 Then
        If strPort = txtPort.Text Then
            SetIniFileProperty "Port", "", "Override", m.strIniFile
        Else
            txtPort.Text = strPort
        End If
    End If

    strTarget = GetIniFileProperty("Target", "", "Override", m.strIniFile)
    If Len(strTarget) > 0 Then
        If strTarget = txtTargetID.Text Then
            SetIniFileProperty "Target", "", "Override", m.strIniFile
        Else
            txtTargetID.Text = strTarget
        End If
    End If

    strSubsender = GetIniFileProperty("Subsender", "", "Override", m.strIniFile)
    If Len(strSubsender) > 0 Then
        If strSubsender = txtSubsender.Text Then
            SetIniFileProperty "Subsender", "", "Override", m.strIniFile
        Else
            txtSubsender.Text = strSubsender
        End If
    End If

    strRefresh = GetIniFileProperty("Refresh", "", "Override", m.strIniFile)
    If Len(strRefresh) > 0 Then
        If strRefresh = txtRefreshServer.Text Then
            SetIniFileProperty "Refresh", "", "Override", m.strIniFile
        Else
            txtRefreshServer.Text = strRefresh
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetServerControlsLive"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerControlsDemo
'' Description: Set the server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerControlsDemo()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target ID to override
    Dim strSubsender As String          ' Subsender ID to override
    Dim strRefresh As String            ' Refresh server information

    txtServerIP.Text = GetIniFileProperty("DemoIP", "", "Server", m.strConnectIni)
    txtPort.Text = GetIniFileProperty("DemoPort", "", "Server", m.strConnectIni)
    txtTargetID.Text = GetIniFileProperty("DemoTarget", "", "Server", m.strConnectIni)
    txtSubsender.Text = GetIniFileProperty("DemoSubsender", "", "Server", m.strConnectIni)
    txtRefreshServer.Text = GetIniFileProperty("DemoRefresh", "", "Server", m.strConnectIni)
    
    strIP = GetIniFileProperty("DemoIP", "", "Override", m.strIniFile)
    If Len(strIP) > 0 Then
        If strIP = txtServerIP.Text Then
            SetIniFileProperty "DemoIP", "", "Override", m.strIniFile
        Else
            txtServerIP.Text = strIP
        End If
    End If

    strPort = GetIniFileProperty("DemoPort", "", "Override", m.strIniFile)
    If Len(strPort) > 0 Then
        If strPort = txtPort.Text Then
            SetIniFileProperty "DemoPort", "", "Override", m.strIniFile
        Else
            txtPort.Text = strPort
        End If
    End If

    strTarget = GetIniFileProperty("DemoTarget", "", "Override", m.strIniFile)
    If Len(strTarget) > 0 Then
        If strTarget = txtTargetID.Text Then
            SetIniFileProperty "DemoTarget", "", "Override", m.strIniFile
        Else
            txtTargetID.Text = strTarget
        End If
    End If

    strSubsender = GetIniFileProperty("DemoSubsender", "", "Override", m.strIniFile)
    If Len(strSubsender) > 0 Then
        If strSubsender = txtSubsender.Text Then
            SetIniFileProperty "DemoSubsender", "", "Override", m.strIniFile
        Else
            txtSubsender.Text = strSubsender
        End If
    End If

    strRefresh = GetIniFileProperty("DemoRefresh", "", "Override", m.strIniFile)
    If Len(strRefresh) > 0 Then
        If strRefresh = txtRefreshServer.Text Then
            SetIniFileProperty "DemoRefresh", "", "Override", m.strIniFile
        Else
            txtRefreshServer.Text = strRefresh
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetServerControlsDemo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceServerControls
'' Description: Set the price server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceServerControls()
On Error GoTo ErrSection:

    If optLive.Value = True Then
        SetPriceServerControlsLive
    ElseIf optDemo.Value = True Then
        SetPriceServerControlsDemo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetPriceServerControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceServerControlsLive
'' Description: Set the price server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceServerControlsLive()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target ID to override

    txtPriceIP.Text = GetIniFileProperty("IP", "", "Price", m.strConnectIni)
    txtPricePort.Text = GetIniFileProperty("Port", "", "Price", m.strConnectIni)
    txtPriceTarget.Text = GetIniFileProperty("Target", "", "Price", m.strConnectIni)
    
    strIP = GetIniFileProperty("PriceIP", "", "Override", m.strIniFile)
    If Len(strIP) > 0 Then
        If strIP = txtPriceIP.Text Then
            SetIniFileProperty "PriceIP", "", "Override", m.strIniFile
        Else
            txtPriceIP.Text = strIP
        End If
    End If

    strPort = GetIniFileProperty("PricePort", "", "Override", m.strIniFile)
    If Len(strPort) > 0 Then
        If strPort = txtPricePort.Text Then
            SetIniFileProperty "PricePort", "", "Override", m.strIniFile
        Else
            txtPricePort.Text = strPort
        End If
    End If

    strTarget = GetIniFileProperty("PriceTarget", "", "Override", m.strIniFile)
    If Len(strTarget) > 0 Then
        If strTarget = txtPriceTarget.Text Then
            SetIniFileProperty "PriceTarget", "", "Override", m.strIniFile
        Else
            txtPriceTarget.Text = strTarget
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetPriceServerControlsLive"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceServerControlsDemo
'' Description: Set the price server controls from the INI files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceServerControlsDemo()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target ID to override

    txtPriceIP.Text = GetIniFileProperty("DemoIP", "", "Price", m.strConnectIni)
    txtPricePort.Text = GetIniFileProperty("DemoPort", "", "Price", m.strConnectIni)
    txtPriceTarget.Text = GetIniFileProperty("DemoTarget", "", "Price", m.strConnectIni)
    
    strIP = GetIniFileProperty("DemoPriceIP", "", "Override", m.strIniFile)
    If Len(strIP) > 0 Then
        If strIP = txtPriceIP.Text Then
            SetIniFileProperty "DemoPriceIP", "", "Override", m.strIniFile
        Else
            txtPriceIP.Text = strIP
        End If
    End If

    strPort = GetIniFileProperty("DemoPricePort", "", "Override", m.strIniFile)
    If Len(strPort) > 0 Then
        If strPort = txtPricePort.Text Then
            SetIniFileProperty "DemoPricePort", "", "Override", m.strIniFile
        Else
            txtPricePort.Text = strPort
        End If
    End If

    strTarget = GetIniFileProperty("DemoPriceTarget", "", "Override", m.strIniFile)
    If Len(strTarget) > 0 Then
        If strTarget = txtPriceTarget.Text Then
            SetIniFileProperty "DemoPriceTarget", "", "Override", m.strIniFile
        Else
            txtPriceTarget.Text = strTarget
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetPriceServerControlsDemo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerOverrides
'' Description: Set the server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerOverrides()
On Error GoTo ErrSection:

    If optLive.Value = True Then
        SetServerOverridesLive
    ElseIf optDemo.Value = True Then
        SetServerOverridesDemo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetServerOverrides"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerOverridesLive
'' Description: Set the server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerOverridesLive()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target Computer ID
    Dim strSubsender As String          ' Subsender ID
    Dim strRefresh As String            ' Refresh server information

    strIP = GetIniFileProperty("IP", "", "Server", m.strConnectIni)
    strPort = GetIniFileProperty("Port", "", "Server", m.strConnectIni)
    strTarget = GetIniFileProperty("Target", "", "Server", m.strConnectIni)
    strSubsender = GetIniFileProperty("Subsender", "", "Server", m.strConnectIni)
    strRefresh = GetIniFileProperty("Refresh", "", "Server", m.strConnectIni)
    
    If Len(Trim(txtServerIP.Text)) > 0 Then
        If Trim(txtServerIP.Text) = strIP Then
            SetIniFileProperty "IP", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "IP", Trim(txtServerIP.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtPort.Text)) > 0 Then
        If Trim(txtPort.Text) = strPort Then
            SetIniFileProperty "Port", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "Port", Trim(txtPort.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtTargetID.Text)) > 0 Then
        If Trim(txtTargetID.Text) = strTarget Then
            SetIniFileProperty "Target", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "Target", Trim(txtTargetID.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtSubsender.Text)) > 0 Then
        If Trim(txtSubsender.Text) = strSubsender Then
            SetIniFileProperty "Subsender", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "Subsender", Trim(txtSubsender.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtRefreshServer.Text)) > 0 Then
        If Trim(txtRefreshServer.Text) = strRefresh Then
            SetIniFileProperty "Refresh", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "Refresh", Trim(txtRefreshServer.Text), "Override", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetServerOverridesLive"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetServerOverridesDemo
'' Description: Set the server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetServerOverridesDemo()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target Computer ID
    Dim strSubsender As String          ' Subsender ID
    Dim strRefresh As String            ' Refresh server information

    strIP = GetIniFileProperty("DemoIP", "", "Server", m.strConnectIni)
    strPort = GetIniFileProperty("DemoPort", "", "Server", m.strConnectIni)
    strTarget = GetIniFileProperty("DemoTarget", "", "Server", m.strConnectIni)
    strSubsender = GetIniFileProperty("DemoSubsender", "", "Server", m.strConnectIni)
    strRefresh = GetIniFileProperty("DemoRefresh", "", "Server", m.strConnectIni)
    
    If Len(Trim(txtServerIP.Text)) > 0 Then
        If Trim(txtServerIP.Text) = strIP Then
            SetIniFileProperty "DemoIP", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoIP", Trim(txtServerIP.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtPort.Text)) > 0 Then
        If Trim(txtPort.Text) = strPort Then
            SetIniFileProperty "DemoPort", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoPort", Trim(txtPort.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtTargetID.Text)) > 0 Then
        If Trim(txtTargetID.Text) = strTarget Then
            SetIniFileProperty "DemoTarget", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoTarget", Trim(txtTargetID.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtSubsender.Text)) > 0 Then
        If Trim(txtSubsender.Text) = strSubsender Then
            SetIniFileProperty "DemoSubsender", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoSubsender", Trim(txtSubsender.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtRefreshServer.Text)) > 0 Then
        If Trim(txtRefreshServer.Text) = strRefresh Then
            SetIniFileProperty "DemoRefresh", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoRefresh", Trim(txtRefreshServer.Text), "Override", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetServerOverridesDemo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceServerOverrides
'' Description: Set the price server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceServerOverrides()
On Error GoTo ErrSection:

    If optLive.Value = True Then
        SetPriceServerOverridesLive
    ElseIf optDemo.Value = True Then
        SetPriceServerOverridesDemo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetPriceServerOverrides"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceServerOverridesLive
'' Description: Set the price server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceServerOverridesLive()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target Computer ID

    strIP = GetIniFileProperty("IP", "", "Price", m.strConnectIni)
    strPort = GetIniFileProperty("Port", "", "Price", m.strConnectIni)
    strTarget = GetIniFileProperty("Target", "", "Price", m.strConnectIni)
    
    If Len(Trim(txtPriceIP.Text)) > 0 Then
        If Trim(txtPriceIP.Text) = strIP Then
            SetIniFileProperty "PriceIP", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "PriceIP", Trim(txtPriceIP.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtPricePort.Text)) > 0 Then
        If Trim(txtPricePort.Text) = strPort Then
            SetIniFileProperty "PricePort", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "PricePort", Trim(txtPricePort.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtPriceTarget.Text)) > 0 Then
        If Trim(txtPriceTarget.Text) = strTarget Then
            SetIniFileProperty "PriceTarget", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "PriceTarget", Trim(txtPriceTarget.Text), "Override", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetPriceServerOverridesLive"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPriceServerOverridesDemo
'' Description: Set the price server overrides in the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetPriceServerOverridesDemo()
On Error GoTo ErrSection:

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strTarget As String             ' Target Computer ID

    strIP = GetIniFileProperty("DemoIP", "", "Price", m.strConnectIni)
    strPort = GetIniFileProperty("DemoPort", "", "Price", m.strConnectIni)
    strTarget = GetIniFileProperty("DemoTarget", "", "Price", m.strConnectIni)
    
    If Len(Trim(txtPriceIP.Text)) > 0 Then
        If Trim(txtPriceIP.Text) = strIP Then
            SetIniFileProperty "DemoPriceIP", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoPriceIP", Trim(txtPriceIP.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtPricePort.Text)) > 0 Then
        If Trim(txtPricePort.Text) = strPort Then
            SetIniFileProperty "DemoPricePort", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoPricePort", Trim(txtPricePort.Text), "Override", m.strIniFile
        End If
    End If

    If Len(Trim(txtPriceTarget.Text)) > 0 Then
        If Trim(txtPriceTarget.Text) = strTarget Then
            SetIniFileProperty "DemoPriceTarget", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "DemoPriceTarget", Trim(txtPriceTarget.Text), "Override", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SetPriceServerOverridesDemo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixLogins
'' Description: If the login information doesn't include Live vs Demo, add it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixLogins()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    
    cboUserName.Clear
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If (Len(strLogins) > 0) And (InStr(strLogins, "|") = 0) Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            astrLogins(lIndex) = astrLogins(lIndex) & "|0"
        Next lIndex
        
        SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.FixLogins"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowServerControls
'' Description: Show and move controls based on the given settings
'' Inputs:      Show Server Mode?, Show Subsender?, Show Price Server?,
''              Show Refresh Server?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowServerControls(ByVal bShowServerMode As Boolean, ByVal bShowSubsender As Boolean, ByVal bShowPrice As Boolean, ByVal bShowRefresh As Boolean)
On Error GoTo ErrSection:

    fraServerMode.Visible = bShowServerMode
    fraSubsender.Visible = bShowSubsender
    fraRefreshServer.Visible = bShowRefresh

    If (bShowServerMode = False) And (bShowSubsender = False) And (bShowRefresh = False) Then
        fraServerInfo.Height = 1035
        fraServer.Top = 240
    ElseIf (bShowServerMode = False) And (bShowSubsender = True) And (bShowRefresh = False) Then
        fraServerInfo.Height = 1335
        fraServer.Top = 240
        fraSubsender.Top = 960
    ElseIf (bShowServerMode = False) And (bShowSubsender = False) And (bShowRefresh = True) Then
        fraServerInfo.Height = 1335
        fraServer.Top = 240
        fraRefreshServer.Top = 960
    ElseIf (bShowServerMode = False) And (bShowSubsender = True) And (bShowRefresh = True) Then
        fraServerInfo.Height = 1755
        fraServer.Top = 240
        fraSubsender.Top = 960
        fraRefreshServer.Top = 1320
    
    ElseIf (bShowServerMode = True) And (bShowSubsender = False) And (bShowRefresh = False) Then
        fraServerInfo.Height = 1335
        fraServer.Top = 600
    ElseIf (bShowServerMode = True) And (bShowSubsender = True) And (bShowRefresh = False) Then
        fraServerInfo.Height = 1755
        fraServer.Top = 600
        fraSubsender.Top = 1320
    ElseIf (bShowServerMode = True) And (bShowSubsender = False) And (bShowRefresh = True) Then
        fraServerInfo.Height = 1755
        fraServer.Top = 600
        fraRefreshServer.Top = 1320
    ElseIf (bShowServerMode = True) And (bShowSubsender = True) And (bShowRefresh = True) Then
        fraServerInfo.Height = 2115
        fraServer.Top = 600
        fraSubsender.Top = 1320
        fraRefreshServer.Top = 1680
    End If
    
    fraPriceServer.Top = fraServerInfo.Top + fraServerInfo.Height + 60
    If bShowPrice Then
        fraPriceServer.Visible = CheckBoxValue(chkShowIP)
        Height = fraPriceServer.Top + fraPriceServer.Height + 555 + 60
    Else
        fraPriceServer.Visible = False
        Height = fraPriceServer.Top + 555
    End If
        
    If Not m.astrEnabledSymbols Is Nothing Then
        lblEnabledSymbols.Visible = True
        txtEnabledSymbols.Visible = True
        
        Height = Height + lblEnabledSymbols.Height + txtEnabledSymbols.Height + 60
        
        With txtEnabledSymbols
            .Move .Left, ScaleHeight - .Height - 60
        End With
        
        With lblEnabledSymbols
            .Move .Left, txtEnabledSymbols.Top - .Height
        End With
    
        With rtfDisclaimer
            .Move .Left, .Top, .Width, lblEnabledSymbols.Top - (.Top * 2)
        End With
    Else
        lblEnabledSymbols.Visible = False
        txtEnabledSymbols.Visible = False
    
        With rtfDisclaimer
            .Move .Left, .Top, .Width, ScaleHeight - (.Top * 2)
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.ShowServerControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateServerMode
'' Description: Update the server mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateServerMode()
On Error GoTo ErrSection:

    Dim lServerMode As Long             ' Server mode
    
    lServerMode = -1&
    If optLive.Value = True Then
        lServerMode = 0
    ElseIf optDemo.Value = True Then
        lServerMode = 1
    End If
    
    If (cboUserName.ItemData(cboUserName.ListIndex) <> lServerMode) And (lServerMode >= 0) Then
        cboUserName.ItemData(cboUserName.ListIndex) = lServerMode
        SaveLogins
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.UpdateServerMode"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveLogins
'' Description: Save the logins
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveLogins()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLogins As cGdArray          ' Array of login information
    
    Set astrLogins = New cGdArray
    astrLogins.Create eGDARRAY_Strings, cboUserName.ListCount
    
    For lIndex = 0 To cboUserName.ListCount - 1
        astrLogins(lIndex) = cboUserName.List(lIndex) & "|" & cboUserName.ItemData(lIndex)
    Next lIndex
    
    SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.SaveLogins"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadAccountsCombo
'' Description: Load the accounts combo box
'' Inputs:      None
'' Returns:     Number of Accounts Loaded
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadAccountsCombo() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    lReturn = 0&
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(m.nBroker) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        cboAccounts.AddItem rs!AccountNumber
        lReturn = lReturn + 1&
        
        rs.MoveNext
    Loop
    
    LoadAccountsCombo = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginFix.LoadAccountsCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountsCombo
'' Description: Set the accounts combo box to the given account if possible
'' Inputs:      Account
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetAccountsCombo(ByVal strAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboAccounts.ListCount > 0) And (Len(strAccount) > 0) Then
        For lIndex = 0 To cboAccounts.ListCount - 1
            If UCase(cboAccounts.List(lIndex)) = UCase(strAccount) Then
                bFound = True
                cboAccounts.ListIndex = lIndex
            End If
        Next lIndex
    End If
    
    SetAccountsCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginFix.SetAccountsCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddAccount
'' Description: Allow the user to give us a new account number
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddAccount()
On Error GoTo ErrSection:

    Dim strAccount As String            ' Account number
    Dim strDefault As String            ' Default account number
    Dim bAccountExists As Boolean       ' Does an account with that number already exist?
    
    Do
        If g.Broker.AccountExists(cboUserName.Text) Then
            strDefault = ""
        Else
            strDefault = cboUserName.Text
        End If
        
        strAccount = InfBox("What is your " & m.strBrokerName & " account number?", "?", , m.strBrokerName & " Account Number", , , , , , "string", strDefault)
        If Len(strAccount) > 0 Then
            bAccountExists = g.Broker.AccountExists(strAccount)
            If bAccountExists Then
                InfBox "Account '" & strAccount & "' already exists", "!", , "Error"
            Else
                If SetAccountsCombo(strAccount) = False Then
                    cboAccounts.AddItem strAccount
                    cboAccounts.ListIndex = cboAccounts.NewIndex
                    
                    mTradeTracker.CreateAccountFromNumber strAccount, m.nBroker
                End If
            End If
        End If
    Loop While ((Len(strAccount) > 0) And (bAccountExists = True))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.AddAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoginExists
'' Description: Does the given login exist?
'' Inputs:      Login
'' Returns:     0 = Doesn't Exist, 1 = Exists, 2 = Exists but different case
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoginExists(ByVal strLogin As String, Optional strExistingLogin As String) As Integer
On Error GoTo ErrSection:

    Dim iReturn As Integer              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    iReturn = 0
    
    For lIndex = 0 To cboUserName.ListCount - 1
        If cboUserName.List(lIndex) = strLogin Then
            iReturn = 1
            strExistingLogin = cboUserName.List(lIndex)
            Exit For
        ElseIf UCase(cboUserName.List(lIndex)) = UCase(strLogin) Then
            strExistingLogin = cboUserName.List(lIndex)
            iReturn = 2
        End If
    Next lIndex
    
    LoginExists = iReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginFix.LoginExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameLogin
'' Description: Rename the login
'' Inputs:      Old Login, New Login
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RenameLogin(ByVal strOldLogin As String, ByVal strNewLogin As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To cboUserName.ListCount - 1
        If cboUserName.List(lIndex) = strOldLogin Then
            cboUserName.List(lIndex) = strNewLogin
            Exit For
        End If
    Next lIndex
    
    SaveLogins

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginFix.RenameLogin"
    
End Sub

