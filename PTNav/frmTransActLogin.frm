VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTransActLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAction 
      Left            =   7020
      Top             =   4200
   End
   Begin HexUniControls.ctlUniFrameWL fraLoginInfo 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   780
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
      Caption         =   "frmTransActLogin.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTransActLogin.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTransActLogin.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraButtons 
         Height          =   915
         Left            =   0
         TabIndex        =   14
         Top             =   1140
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
         Caption         =   "frmTransActLogin.frx":0068
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTransActLogin.frx":0094
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTransActLogin.frx":00B4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkShowIP 
            Height          =   435
            Left            =   2220
            TabIndex        =   18
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
            Caption         =   "frmTransActLogin.frx":00D0
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmTransActLogin.frx":0120
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0140
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
            Cancel          =   -1  'True
            Height          =   435
            Left            =   1080
            TabIndex        =   17
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
            Caption         =   "frmTransActLogin.frx":015C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTransActLogin.frx":018A
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":01AA
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdLogin 
            Default         =   -1  'True
            Height          =   435
            Left            =   0
            TabIndex        =   16
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
            Caption         =   "frmTransActLogin.frx":01C6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTransActLogin.frx":01F2
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0212
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAgree 
            Height          =   435
            Left            =   0
            Top             =   0
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
            Caption         =   "frmTransActLogin.frx":022E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTransActLogin.frx":02EE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":030E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraServerInfo 
         Height          =   735
         Left            =   0
         TabIndex        =   19
         Top             =   2160
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
         Caption         =   "frmTransActLogin.frx":032A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTransActLogin.frx":036E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTransActLogin.frx":038E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtServerIP 
            Height          =   285
            Left            =   420
            TabIndex        =   6
            Top             =   300
            Width           =   2235
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTransActLogin.frx":03AA
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
            Tip             =   "frmTransActLogin.frx":03CA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":03EA
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdReset 
            Height          =   375
            Left            =   2760
            TabIndex        =   9
            Top             =   240
            Width           =   675
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
            Caption         =   "frmTransActLogin.frx":0406
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTransActLogin.frx":0432
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0452
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblServerIP 
            Height          =   195
            Left            =   120
            Top             =   300
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
            Caption         =   "frmTransActLogin.frx":046E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTransActLogin.frx":0496
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":04B6
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraLogin 
         CausesValidation=   0   'False
         Height          =   1035
         Left            =   0
         TabIndex        =   5
         Top             =   0
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
         Caption         =   "frmTransActLogin.frx":04D2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmTransActLogin.frx":04F2
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmTransActLogin.frx":0512
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdGet 
            Height          =   315
            Left            =   3000
            TabIndex        =   13
            Top             =   720
            Width           =   495
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
            Caption         =   "frmTransActLogin.frx":052E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTransActLogin.frx":0556
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0576
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboAccounts 
            Height          =   315
            Left            =   840
            TabIndex        =   12
            Top             =   720
            Width           =   2115
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
            Tip             =   "frmTransActLogin.frx":0592
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":05B2
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniComboImageXP cboUserIds 
            Height          =   315
            Left            =   840
            TabIndex        =   7
            Top             =   0
            Width           =   1455
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
            Tip             =   "frmTransActLogin.frx":05CE
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":05EE
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPassword 
            Height          =   315
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   2655
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmTransActLogin.frx":060A
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
            Tip             =   "frmTransActLogin.frx":062A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":064A
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdAddLogin 
            Height          =   315
            Left            =   2340
            TabIndex        =   8
            Top             =   0
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
            Caption         =   "frmTransActLogin.frx":0666
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmTransActLogin.frx":069A
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":06BA
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblAccount 
            Height          =   255
            Left            =   0
            Top             =   780
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
            Caption         =   "frmTransActLogin.frx":06D6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTransActLogin.frx":0708
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0728
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblUserID 
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
            Caption         =   "frmTransActLogin.frx":0744
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTransActLogin.frx":0776
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0796
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
            Caption         =   "frmTransActLogin.frx":07B2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmTransActLogin.frx":07E6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmTransActLogin.frx":0806
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMode 
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "frmTransActLogin.frx":0822
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTransActLogin.frx":084C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTransActLogin.frx":086C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optSimLiveMode 
         Height          =   220
         Left            =   2040
         TabIndex        =   3
         Top             =   240
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
         Caption         =   "frmTransActLogin.frx":0888
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTransActLogin.frx":08C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTransActLogin.frx":08E6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDemoMode 
         Height          =   220
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         Caption         =   "frmTransActLogin.frx":0902
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTransActLogin.frx":092C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTransActLogin.frx":094C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLiveMode 
         Height          =   220
         Left            =   1140
         TabIndex        =   2
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
         Caption         =   "frmTransActLogin.frx":0968
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmTransActLogin.frx":0992
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTransActLogin.frx":09B2
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtTransactSymbols 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTransActLogin.frx":09CE
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
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmTransActLogin.frx":09EE
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTransActLogin.frx":0A0E
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   3555
      Left            =   3840
      TabIndex        =   15
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6271
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTransActLogin.frx":0A2A
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
      Tip             =   "frmTransActLogin.frx":0A4A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTransActLogin.frx":0A6A
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
   Begin HexUniControls.ctlUniLabelXP lblTransactSymbols 
      Height          =   195
      Left            =   120
      Top             =   3840
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
      Caption         =   "frmTransActLogin.frx":0A86
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTransActLogin.frx":0AE8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTransActLogin.frx":0B08
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmTransActLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTransActLogin.frm
'' Description: Allow the user to choose their TransAct login information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/26/2009   DAJ         Added Get button to retrieve account information
'' 12/10/2010   DAJ         Changed over to the IsBrokerUser function
'' 02/21/2013   DAJ         Log when changing the UseSimUrl/Simulated registry keys
'' 07/19/2013   DAJ         New TransAct API
'' 04/22/2014   DAJ         No longer login to TransAct as "simuser"
'' 08/08/2014   DAJ         Reworked enabled symbols list for TransAct
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strIniFile As String                ' INI file for TransAct login information
    strGatewayURL As String             ' Gateway URL for TransAct
    
    strUserID As String                 ' Last User ID selected
    strLastPasswordTried As String      ' Last password tried
    
    lWaitForUnloaded As Long            ' Wait for app unloaded message?
    
    bOK As Boolean                      ' Did the user press OK or Cancel?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      User ID, Account, Show IP?, Are we switching?, New?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strUserID As String = "", Optional ByVal strAccount As String = "", Optional ByVal bShowIP As Boolean = False, Optional ByVal bSwitching As Boolean = False, Optional bNew As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastUserID As String         ' Last User ID the user logged in as
    Dim strGatewayURL As String         ' User defined gateway URL
    Dim strAccountList As String        ' Account list for the selected user ID
    Dim bReturn As Boolean              ' Return value for the function
    Dim bPassThrough As Boolean         ' Pass through as simulated
    Dim lUseSimUrl As Long              ' Value for the UseSimUrl registry key
    Dim lSimulation As Long             ' Value for the Simulated registry key
    Dim bUseSimUrl As Boolean           ' Was the registry key stored succesfully?
    Dim bSimulation As Boolean          ' Was the registry key stored succesfully?
    Static bInProgress As Boolean       ' Are we already in progress?

    fraMode.Visible = True
    fraLoginInfo.Top = 780
    lblTransactSymbols.Top = 3840
    txtTransactSymbols.Top = 4080
    rtfDisclaimer.Height = 3555
    Height = 5205

    bReturn = False
    bPassThrough = False
    
    If bInProgress = False Then
        bInProgress = True
        
        m.bOK = False
        m.strIniFile = AddSlash(App.Path) & "TransAct.INI"
        m.strLastPasswordTried = ""
        
        m.strGatewayURL = GetIniFileProperty("URL", "", "GateKeeper", AddSlash(App.Path) & "Provided\TaIps.INI")
        If Len(m.strGatewayURL) = 0 Then
            m.strGatewayURL = GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", "")
        End If
        
        strLastUserID = GetIniFileProperty("LastUserID", "", "User", m.strIniFile)
        LoadUserIdCombo
        If cboUserIds.ListCount > 0 Then
            If SetUserIdCombo(strUserID) = False Then
                If SetUserIdCombo(strLastUserID) = False Then
                    cboUserIds.ListIndex = 0
                End If
            End If
        End If
                
        optLiveMode.Value = True
        If cboUserIds.ListIndex >= 0 Then
            LoginModeToControls cboUserIds.ItemData(cboUserIds.ListIndex)
        End If
        SetGateKeeper
        
        ' DAJ 04/22/2014: Because of the new CME rules, TransAct is going to be required to
        ' get rid of the "simuser" demo account ( they will need to go to individual simulated
        ' accounts per user ).  Because of this, we are going to stop connecting to "simuser",
        ' but utilize the BRKRDEMO enablement to still give the user real-time data on the same
        ' handful of symbols for the first month ( or until they login with a real login )...
        'If (cboUserIds.ListCount = 0) Then
        '    If (g.Broker.IsBrokerSimUser(eTT_AccountType_TransAct) = True) And (bNew = False) Then
        '        bPassThrough = True
        '    End If
        'End If
        
        If bPassThrough Then
            If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", "") <> m.strGatewayURL Then
                SetRegistryValue rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", m.strGatewayURL
            End If
            
            g.Transact.LoginMode = eGDTransActLoginMode_Demo
            g.Transact.UserName = g.Transact.SimUserUserName
            g.Transact.Password = g.Transact.SimUserPassword
            g.Transact.Account = kTransActSimUserAccount
            
            bReturn = True
        Else
            If GetRegistryValue(rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", "") <> Trim(txtServerIP.Text) Then
                SetRegistryValue rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", Trim(txtServerIP.Text)
            End If
                    
            If (cboUserIds.ListCount = 0) Or ((cboUserIds.ListCount = 1) And (bSwitching = True)) Then
                NewUserID
                SetGateKeeper
            End If
            
            If (cboUserIds.ListCount > 1) Or ((cboUserIds.ListCount = 1) And (bSwitching = False)) Then
                If bShowIP = True Then chkShowIP.Value = vbChecked Else chkShowIP.Value = vbUnchecked
                
                m.strUserID = cboUserIds.Text
                
                strAccountList = GetIniFileProperty("AccountList", "", cboUserIds.Text, m.strIniFile)
                If Len(strAccountList) > 0 Then
                    SetAccounts Replace(strAccountList, ",", vbTab)
                End If
                
                MoveFocus txtPassword
                                
                If cboAccounts.ListCount = 0 Then
                    cmdLogin.Enabled = False
                End If
                EnableControls
                
                FillTransActSymbols
        
                ShowForm Me, eForm_ActModal, frmMain
                
                If m.bOK = True Then
                    SetIniFileProperty "LastUserID", cboUserIds.Text, "User", m.strIniFile
                    SetIniFileProperty "LastAccount", Parse(cboAccounts.Text, "(", 1), cboUserIds.Text, m.strIniFile
                    
                    g.Transact.LoginMode = LoginModeFromControls
                    g.Transact.UserName = Trim(cboUserIds.Text)
                    g.Transact.Password = Trim(txtPassword.Text)
                    g.Transact.Account = Parse(cboAccounts.Text, "(", 1)
                    
                    FixLogins
                    
                ElseIf (g.Transact.ConnectionStatus <> eGDConnectionStatus_Connected) And (g.Transact.AppLoaded = True) Then
                    g.Transact.SendTransactMessage eGDTransactMessageType_UnloadApp, ""
                End If
                
                bReturn = m.bOK
            End If
        End If
        
        bInProgress = False
    End If
    
    ShowMe = bReturn

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTransActLogin.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccounts
'' Description: Set the accounts combo box to the supplied account list
'' Inputs:      Account List
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAccounts(ByVal strAccountList As String)
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' Account list split out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLastAccount As String        ' Last Account the user logged in as
    Dim lAccountNumber As Long          ' Account number
    Dim lTimeOut As Long                ' Timeout variable
    
    Set astrAccounts = New cGdArray
    astrAccounts.SplitFields strAccountList, vbTab
    
    cboAccounts.Clear
    
    If UCase(astrAccounts(0)) = "ERROR" Then
        If Len(astrAccounts(2)) > 0 Then
            InfBox astrAccounts(1), "!", , "TransAct Error " & astrAccounts(2)
        Else
            InfBox astrAccounts(1), "!", , "TransAct Error"
        End If
    Else
        For lIndex = 0 To astrAccounts.Size - 1
            lAccountNumber = CLng(Val(Parse(astrAccounts(lIndex), " ", 1)))
            If lAccountNumber <> kTransActSimUserAccount Then
                If (lAccountNumber >= 20000) And (lAccountNumber < 30000) Then
                    cboAccounts.AddItem Str(lAccountNumber) & " (TransAct Demo)"
                Else
                    cboAccounts.AddItem astrAccounts(lIndex)
                End If
            End If
        Next lIndex
        
        strLastAccount = GetIniFileProperty("LastAccount", "", cboUserIds.Text, m.strIniFile)
        If cboAccounts.ListCount > 0 Then
            If SetAccountCombo(strLastAccount) = False Then
                cboAccounts.ListIndex = 0
            End If
        End If
        
        If (cboAccounts.ListCount > 0) And (cboAccounts.ListIndex >= 0) Then
            cmdLogin.Enabled = Not g.Transact.AppLoaded
        End If
        EnableControls
        
        SetIniFileProperty "AccountList", astrAccounts.JoinFields(","), cboUserIds.Text, m.strIniFile
        
        InfBox ""
        
        On Error Resume Next
        SetFocus
        MoveFocus cboAccounts
        On Error GoTo ErrSection
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.SetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AppUnloaded
'' Description: Receive an app unloaded message from TransAct object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AppUnloaded()
On Error GoTo ErrSection:

    Select Case m.lWaitForUnloaded
        Case 1:
            tmrAction.Tag = "LOGONFORACCOUNTLIST"
            tmrAction.Enabled = True
        
        Case 2:
            tmrAction.Tag = "ENABLEBUTTON"
            tmrAction.Enabled = True
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.AppUnloaded"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboUserIds_Click
'' Description: When the user changes users, load the accounts combo and clear
''              the password field
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboUserIds_Click()
On Error GoTo ErrSection:

    Dim strAccountList As String        ' Account list from the INI file
    
    If (Visible = True) And (cboUserIds.Text <> m.strUserID) Then
        txtPassword.Text = ""
        LoginModeToControls cboUserIds.ItemData(cboUserIds.ListIndex)
        
        strAccountList = GetIniFileProperty("AccountList", "", cboUserIds.Text, m.strIniFile)
        If Len(strAccountList) > 0 Then
            SetAccounts Replace(strAccountList, ",", vbTab)
        Else
            cboAccounts.Clear
            cmdLogin.Enabled = False
        End If
        
        EnableControls
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.cboUserIds_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowIP_Click
'' Description: Show/Hide the Server IP information as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowIP_Click()
On Error GoTo ErrSection:

    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.chkShowIP_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddLogin_Click
'' Description: Allow the user to add a new login
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddLogin_Click()
On Error GoTo ErrSection:

    NewUserID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.cmdAddLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel out of the form
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
    RaiseError "frmTransActLogin.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdGet_Click
'' Description: Allow the user to request accounts from the TransAct servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdGet_Click()
On Error GoTo ErrSection:

    MoveFocus cmdGet
    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.cmdGet_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLogin_Click
'' Description: Allow the user to login to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLogin_Click()
On Error GoTo ErrSection:

    MoveFocus cmdLogin

    If Len(Trim(txtServerIP.Text)) = 0 Then
        chkShowIP.Value = vbChecked
        MoveFocus txtServerIP
        InfBox "Please enter in an IP address for the TransAct server", "!", , "TransAct Login Error"
        GoTo ErrExit
    End If

    m.bOK = True
    Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdReset_Click
'' Description: Allow the user to reset the IP to the default
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdReset_Click()
On Error GoTo ErrSection:

    txtServerIP.Text = m.strGatewayURL
    MoveFocus txtServerIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.cmdReset_Click"
    
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

    If Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.Form_Activate"
    
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

    Icon = Picture16("kBlank")
    Caption = "TransAct Login Information"
    
    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    CenterTheForm Me

    g.Styler.StyleForm Me
    
    ' Lock the TransAct symbol box and make it look disabled...
    txtTransactSymbols.Enabled = True
    txtTransactSymbols.BackColor = cmdLogin.BackColor
    txtTransactSymbols.Locked = True
    
    m.lWaitForUnloaded = 0&
    
    ' Default the timer to an interval of 1 second, but leave it off until needed...
    tmrAction.Enabled = False
    tmrAction.Interval = 1000
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Cancel out of the form if the user clicks on the X
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
    RaiseError "frmTransActLogin.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If chkShowIP.Value = vbChecked Then
        fraServerInfo.Visible = True
    Else
        fraServerInfo.Visible = False
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrAction.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optDemoMode_Click
'' Description: Update the IP address with the demo server IP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optDemoMode_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetGateKeeper
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.optDemoMode_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optLiveMode_Click
'' Description: Update the IP address with the live server IP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLiveMode_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetGateKeeper
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.optLiveMode_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSimLiveMode_Click
'' Description: Update the IP address with the live server IP
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSimLiveMode_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetGateKeeper
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.optSimLiveMode_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrAction_Timer
'' Description: When the timer goes off, do something based on the state
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrAction_Timer()
On Error GoTo ErrSection:

    Select Case UCase(tmrAction.Tag)
        Case "LOGONFORACCOUNTLIST"
            tmrAction.Tag = ""
            tmrAction.Enabled = False
            
            g.Transact.LoginMode = LoginModeFromControls
            
            m.lWaitForUnloaded = 2&
            
            If g.Transact.LogonForAccountList(cboUserIds.Text, Trim(txtPassword.Text)) Then
                InfBox "Requesting list of accounts for " & cboUserIds.Text & "|from the TransAct servers.||Please wait...|", , , "TransAct Account List", True
            Else
                InfBox ""
            End If
        
        Case "ENABLEBUTTON"
            tmrAction.Tag = ""
            tmrAction.Enabled = False
            
            m.lWaitForUnloaded = 0&
            cmdLogin.Enabled = True
            EnableControls
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.tmrAction_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_Change
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_Change()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.txtPassword_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: Select all of the text in the control when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPassword

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.txtPasword_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_LostFocus
'' Description: Ask the TransAct servers for the account list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_LostFocus()
On Error GoTo ErrSection:
    
    Dim lTimeOut As Long                ' Time out variable
    
    If (Len(Trim(txtPassword.Text)) > 0) And (Trim(txtPassword.Text) <> m.strLastPasswordTried) Then
        If cboAccounts.ListCount = 0 Then
            GetAccounts
            m.strLastPasswordTried = Trim(txtPassword.Text)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.txtPassword_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_GotFocus
'' Description: Select all of the text in the control when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtServerIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtServerIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.txtServerIP_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_LostFocus
'' Description: Save the server IP as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtServerIP_LostFocus()
On Error GoTo ErrSection:

    If Visible Then
        If Trim(txtServerIP.Text) <> m.strGatewayURL Then
            SetIniFileProperty "Gateway", Trim(txtServerIP.Text), "User", m.strIniFile
            SetRegistryValue rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", Trim(txtServerIP.Text)
        Else
            SetIniFileProperty "Gateway", "", "User", m.strIniFile
            SetRegistryValue rkLocalMachine, kTransActRegistryKey, "GateKeeperURL", m.strGatewayURL
        End If
    End If
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.txtServerIP_LostFocus"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUserIdCombo
'' Description: Load the User ID combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUserIdCombo()
On Error GoTo ErrSection:

    Dim strLogins As String             ' Login information string from INI file
    Dim astrLogins As New cGdArray      ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strUserID As String             ' User ID already in the INI file
    Dim strIP As String                 ' IP address from the INI file
    Dim bFixLogins As Boolean           ' Do we want to fix the logins?
    
    bFixLogins = False
    strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
    If Len(strLogins) > 0 Then
        astrLogins.SplitFields strLogins, ","
        For lIndex = 0 To astrLogins.Size - 1
            If (Len(astrLogins(lIndex)) > 0) And (astrLogins(lIndex) <> g.Transact.SimUserUserName) Then
                cboUserIds.AddItem Parse(astrLogins(lIndex), ";", 1)
                cboUserIds.ItemData(cboUserIds.NewIndex) = CLng(Val(Parse(astrLogins(lIndex), ";", 2)))
                If Len(Parse(astrLogins(lIndex), ";", 2)) = 0 Then
                    bFixLogins = True
                End If
            End If
        Next lIndex
    End If
    
    If cboUserIds.ListCount = 0 Then
        strUserID = GetIniFileProperty("UserID", "", "User", m.strIniFile)
        If (Len(strUserID) > 0) And (strUserID <> g.Transact.SimUserUserName) Then
            cboUserIds.AddItem strUserID
            cboUserIds.ItemData(cboUserIds.NewIndex) = 0
            SetIniFileProperty "Logins", strUserID & ";0", "User", m.strIniFile
        End If
    End If
    
    If (bFixLogins = True) Then
        strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
        If Len(strLogins) > 0 Then
            astrLogins.SplitFields strLogins, ","
            For lIndex = 0 To astrLogins.Size - 1
                If Len(Parse(astrLogins(lIndex), ";", 2)) = 0 Then
                    astrLogins(lIndex) = Parse(astrLogins(lIndex), ";", 1) & ";0"
                End If
            Next lIndex
            
            SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.LoadUserIdCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUserIdCombo
'' Description: Set the User ID combo box to the given User ID if possible
'' Inputs:      User ID
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetUserIdCombo(ByVal strUserID As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboUserIds.ListCount > 0) And (Len(strUserID) > 0) Then
        For lIndex = 0 To cboUserIds.ListCount - 1
            If cboUserIds.List(lIndex) = strUserID Then
                bFound = True
                cboUserIds.ListIndex = lIndex
                Exit For
            End If
        Next lIndex
    End If
    
    SetUserIdCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransActLogin.SetUserIdCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAccountCombo
'' Description: Set the Account combo box to the given Account if possible
'' Inputs:      Account
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetAccountCombo(ByVal strAccount As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboAccounts.ListCount > 0) And (Len(strAccount) > 0) Then
        For lIndex = 0 To cboAccounts.ListCount - 1
            If Parse(cboAccounts.List(lIndex), "(", 1) = strAccount Then
                bFound = True
                cboAccounts.ListIndex = lIndex
            End If
        Next lIndex
    End If
    
    SetAccountCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransActLogin.SetAccountCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewUserID
'' Description: Allow the user to give us a new User ID
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewUserID()
On Error GoTo ErrSection:

    Dim strUserID As String             ' User ID from the user
    Dim strLogins As String             ' Login string from the INI file
    Dim strMode As String               ' Mode of the new user ID from user
    
    strUserID = InfBox("What is your TransAct User ID?", "?", , "TransAct User ID", , , , , , "string")
    If Len(strUserID) > 0 Then
        If SetUserIdCombo(strUserID) = False Then
            strMode = InfBox("What connection mode does this ID require?", "?", "+-Live|Demo", "TransAct Connection Mode")
            
            cboUserIds.AddItem strUserID
            Select Case strMode
                Case "L"
                    cboUserIds.ItemData(cboUserIds.NewIndex) = eGDTransActLoginMode_Live
                    optLiveMode.Value = True
                Case "D"
                    cboUserIds.ItemData(cboUserIds.NewIndex) = eGDTransActLoginMode_Demo
                    optDemoMode.Value = True
            End Select
            
            SetUserIdCombo strUserID
            MoveFocus txtPassword
            
            strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
            If Len(strLogins) = 0 Then
                strLogins = strUserID
            Else
                strLogins = strLogins & "," & strUserID
            End If
            Select Case strMode
                Case "L"
                    strLogins = strLogins & ";0"
                Case "D"
                    strLogins = strLogins & ";1"
            End Select
            SetIniFileProperty "Logins", strLogins, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.NewUserID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillTransActSymbols
'' Description: Fill in the TransAct symbols text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillTransActSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols to display
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbols As String            ' String to display
    
    Set astrSymbols = g.Transact.EnabledSymbols
    If astrSymbols.Size > 0 Then
        strSymbols = astrSymbols.JoinFields(",")
        txtTransactSymbols.Text = strSymbols
    Else
        txtTransactSymbols.Text = "You are not activated for any markets.  Please call TransAct to get enabled for symbols you wish to trade."
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.FillTransActSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetGateKeeper
'' Description: Set the Gate Keeper URL based on the login mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGateKeeper()
On Error GoTo ErrSection:

    Dim strIP As String                 ' Override if there is one
    
    strIP = GetIniFileProperty("Gateway", "", "User", m.strIniFile)
    If Len(strIP) > 0 Then
        txtServerIP.Text = strIP
    Else
        txtServerIP.Text = m.strGatewayURL
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.SetGateKeeper"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixLogins
'' Description: Ensure that the login info is correct for current selection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixLogins()
On Error GoTo ErrSection:

    Dim astrLogins As cGdArray          ' Array of login information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strUserID As String             ' User ID
    Dim bFound As Boolean               ' Was the user ID found?
    
    Set astrLogins = New cGdArray
    astrLogins.SplitFields GetIniFileProperty("Logins", "", "User", m.strIniFile), ","
    
    bFound = False
    For lIndex = 0 To astrLogins.Size - 1
        strUserID = Parse(astrLogins(lIndex), ";", 1)
        If strUserID = Trim(cboUserIds.Text) Then
            Select Case True
                Case optDemoMode
                    astrLogins(lIndex) = strUserID & ";" & Str(eGDTransActLoginMode_Demo)
                Case optLiveMode
                    astrLogins(lIndex) = strUserID & ";" & Str(eGDTransActLoginMode_Live)
                Case optSimLiveMode
                    astrLogins(lIndex) = strUserID & ";" & Str(eGDTransActLoginMode_SimLive)
            End Select
            
            bFound = True
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        Select Case True
            Case optDemoMode
                astrLogins.Add Trim(cboUserIds.Text) & ";" & Str(eGDTransActLoginMode_Demo)
            Case optLiveMode
                astrLogins.Add Trim(cboUserIds.Text) & ";" & Str(eGDTransActLoginMode_Live)
            Case optSimLiveMode
                astrLogins.Add Trim(cboUserIds.Text) & ";" & Str(eGDTransActLoginMode_SimLive)
        End Select
    End If
    
    SetIniFileProperty "Logins", astrLogins.JoinFields(","), "User", m.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.FixLogins"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Retrieve account list from the TransAct servers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetAccounts()
On Error GoTo ErrSection:

    If Not g.Transact Is Nothing Then
        If g.Transact.ConnectionStatus <> eGDConnectionStatus_Disconnected Then
            m.lWaitForUnloaded = 1&
            InfBox "Please wait while disconnecting from TransAct...", , , "TransAct Account List", True
            g.Transact.Disconnect "", False, "Logging in for account list"
        Else
            g.Transact.LoginMode = LoginModeFromControls
            
            m.lWaitForUnloaded = 2&
            
            If g.Transact.LogonForAccountList(cboUserIds.Text, Trim(txtPassword.Text)) Then
                InfBox "Requesting list of accounts for " & cboUserIds.Text & "|from the TransAct servers.||Please wait...|", , , "TransAct Account List", True
            Else
                InfBox ""
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    ' Only enable the Get button if the user has selected an account and typed
    ' in a password...
    cmdGet.Enabled = ((Len(txtPassword.Text) > 0) And (cboUserIds.ListIndex >= 0))
    
    ' If the Login button is disabled, then make the Get button the default
    ' button, otherwise make the Login button the default button...
    If cmdLogin.Enabled = True Then
        cmdLogin.Default = True
        cmdGet.Default = False
    ElseIf cmdGet.Enabled = True Then
        cmdLogin.Default = False
        cmdGet.Default = True
    Else
        cmdLogin.Default = False
        cmdGet.Default = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTransActLogin.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoginModeFromControls
'' Description: Get the login mode from the controls
'' Inputs:      None
'' Returns:     Login Mode
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoginModeFromControls() As eGDTransActLoginModes
On Error GoTo ErrSection:

    Dim nReturn As eGDTransActLoginModes ' Return value for the function
    
    If optLiveMode.Value = True Then
        nReturn = eGDTransActLoginMode_Live
    ElseIf optDemoMode.Value = True Then
        nReturn = eGDTransActLoginMode_Demo
    ElseIf optSimLiveMode.Value = True Then
        nReturn = eGDTransActLoginMode_SimLive
    End If
    
    LoginModeFromControls = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTransActLogin.LoginModeFromControls"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoginModeToControls
'' Description: Set the login mode controls based on the given value
'' Inputs:      Login Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoginModeToControls(ByVal nLoginMode As eGDTransActLoginModes)
On Error GoTo ErrSection:

    Select Case nLoginMode
        Case eGDTransActLoginMode_Live
            optLiveMode.Value = True
        Case eGDTransActLoginMode_Demo
            optDemoMode.Value = True
        Case eGDTransActLoginMode_SimLive
            optSimLiveMode.Value = True
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTransActLogin.LoginModeToControls"
    
End Sub

