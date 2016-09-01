VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmLoginPats 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   915
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
      Caption         =   "frmLoginPats.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginPats.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginPats.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdLogin 
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   9
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
         Caption         =   "frmLoginPats.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginPats.frx":0094
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":00B4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1080
         TabIndex        =   10
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
         Caption         =   "frmLoginPats.frx":00D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginPats.frx":00FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":011E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkShowIP 
         Height          =   435
         Left            =   2460
         TabIndex        =   11
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
         Caption         =   "frmLoginPats.frx":013A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmLoginPats.frx":018A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":01AA
         ShowFocus       =   -1  'True
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
         Caption         =   "frmLoginPats.frx":01C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginPats.frx":0286
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":02A6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraServerInfo 
      Height          =   1395
      Left            =   120
      TabIndex        =   12
      Top             =   1920
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
      Caption         =   "frmLoginPats.frx":02C2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginPats.frx":0306
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginPats.frx":0326
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraHost 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   240
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
         Caption         =   "frmLoginPats.frx":0342
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginPats.frx":036E
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":038E
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtServerIP 
            Height          =   285
            Left            =   660
            TabIndex        =   15
            Top             =   0
            Width           =   1575
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginPats.frx":03AA
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
            Tip             =   "frmLoginPats.frx":03CA
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":03EA
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPort 
            Height          =   285
            Left            =   2760
            TabIndex        =   17
            Top             =   0
            Width           =   675
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginPats.frx":0406
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
            Tip             =   "frmLoginPats.frx":0426
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":0446
         End
         Begin HexUniControls.ctlUniLabelXP lblServerIP 
            Height          =   195
            Left            =   0
            Top             =   0
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
            Caption         =   "frmLoginPats.frx":0462
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginPats.frx":0494
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":04B4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPort 
            Height          =   195
            Left            =   2340
            Top             =   0
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
            Caption         =   "frmLoginPats.frx":04D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginPats.frx":04FC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":051C
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraPrice 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
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
         Caption         =   "frmLoginPats.frx":0538
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginPats.frx":0564
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":0584
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniTextBoxXP txtPricePort 
            Height          =   285
            Left            =   2760
            TabIndex        =   1
            Top             =   0
            Width           =   675
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginPats.frx":05A0
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
            Tip             =   "frmLoginPats.frx":05C0
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":05E0
         End
         Begin HexUniControls.ctlUniTextBoxXP txtPriceIP 
            Height          =   285
            Left            =   660
            TabIndex        =   5
            Top             =   0
            Width           =   1575
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLoginPats.frx":05FC
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
            Tip             =   "frmLoginPats.frx":061C
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":063C
         End
         Begin HexUniControls.ctlUniLabelXP lblPricePort 
            Height          =   195
            Left            =   2340
            Top             =   0
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
            Caption         =   "frmLoginPats.frx":0658
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginPats.frx":0684
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":06A4
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblPriceIP 
            Height          =   195
            Left            =   0
            Top             =   0
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
            Caption         =   "frmLoginPats.frx":06C0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginPats.frx":06F4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":0714
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraPats 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   960
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
         Caption         =   "frmLoginPats.frx":0730
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmLoginPats.frx":075C
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":077C
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniComboImageXP cboEnvironment 
            Height          =   315
            Left            =   1020
            TabIndex        =   14
            Top             =   0
            Width           =   1215
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
            Tip             =   "frmLoginPats.frx":0798
            Sorted          =   0   'False
            HScroll         =   0   'False
            RoundedBorders  =   -1  'True
            IconDim         =   16
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":07B8
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSuperTAS 
            Height          =   220
            Left            =   2340
            TabIndex        =   16
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "frmLoginPats.frx":07D4
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmLoginPats.frx":0808
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":0828
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblEnvironment 
            Height          =   225
            Left            =   0
            Top             =   45
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
            Caption         =   "frmLoginPats.frx":0844
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmLoginPats.frx":087E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLoginPats.frx":089E
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraLoginInfo 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "frmLoginPats.frx":08BA
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmLoginPats.frx":08E6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginPats.frx":0906
      RightToLeft     =   0   'False
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
         Caption         =   "frmLoginPats.frx":0922
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginPats.frx":095E
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":097E
         RightToLeft     =   0   'False
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
         Caption         =   "frmLoginPats.frx":099A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmLoginPats.frx":09D0
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":09F0
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
         Text            =   "frmLoginPats.frx":0A0C
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
         Tip             =   "frmLoginPats.frx":0A2C
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":0A4C
      End
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
         Tip             =   "frmLoginPats.frx":0A68
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":0A88
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
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
         Caption         =   "frmLoginPats.frx":0AA4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginPats.frx":0AD8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":0AF8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
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
         Caption         =   "frmLoginPats.frx":0B14
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLoginPats.frx":0B4A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLoginPats.frx":0B6A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   3195
      Left            =   4020
      TabIndex        =   19
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5636
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLoginPats.frx":0B86
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
      Tip             =   "frmLoginPats.frx":0BA6
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLoginPats.frx":0BC6
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
End
Attribute VB_Name = "frmLoginPats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLoginPats.frm
'' Description: Allow the user to login to PATS
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK or Cancel?
    strIniFile As String                ' INI file for the broker
    strConnectIni As String             ' INI file for connection information
    strBrokerName As String             ' Display name for the broker
    
    strUserName As String               ' User name that the user chose
    strPassword As String               ' Password from the user
    strIP As String                     ' IP address to connect to
    strPort As String                   ' Port to connect to
    strPriceIP As String                ' IP address to connect to for the price server
    strPricePort As String              ' Port to connect to for the price server
    strEnvironment As String            ' Environment for the login
    bSuperTAS As Boolean                ' Use SuperTAS?
End Type
Private m As mPrivate

Public Property Get UserName() As String
    UserName = m.strUserName
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property

Public Property Get IP() As String
    IP = m.strIP
End Property

Public Property Get Port() As String
    Port = m.strPort
End Property

Public Property Get PriceIP() As String
    PriceIP = m.strPriceIP
End Property

Public Property Get PricePort() As String
    PricePort = m.strPricePort
End Property

Public Property Get Environment() As String
    Environment = m.strEnvironment
End Property

Public Property Get SuperTAS() As Boolean
    SuperTAS = m.bSuperTAS
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, UserID, Are we switching?, Show IP?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Broker As cBroker, Optional ByVal strUserName As String = "", Optional ByVal bSwitching As Boolean = False, Optional ByVal bShowIP As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastUserName As String       ' Last user name logged into

    m.bOK = False
    m.strIniFile = Broker.IniFile
    m.strConnectIni = Broker.ConnectIni
    m.strBrokerName = Broker.BrokerName
    Caption = m.strBrokerName & " Login Information"
    
    If Len(m.strIniFile) > 0 Then
        strLastUserName = GetIniFileProperty("LastUserName", "", "User", m.strIniFile)
        LoadCombo
        If cboUserName.ListCount > 0 Then
            If SetCombo(strUserName) = False Then
                If SetCombo(strLastUserName) = False Then
                    strLastUserName = GetIniFileProperty("UserName", "", "User", m.strIniFile)
                    If SetCombo(strLastUserName) = False Then
                        cboUserName.ListIndex = 0
                    End If
                End If
            End If
        End If
        
        If (cboUserName.ListCount = 0) Or ((cboUserName.ListCount = 1) And (bSwitching = True)) Then
            AddLogin
        End If
        
        SetServerControls
        
        If (cboUserName.ListCount > 1) Or ((cboUserName.ListCount = 1) And (bSwitching = False)) Then
            CheckBoxValue(chkShowIP) = bShowIP
            fraServerInfo.Visible = bShowIP
            
            With rtfDisclaimer
                .Move .Left, .Top, .Width, ScaleHeight - (.Top * 2)
            End With
                        
            MoveFocus txtPassword
    
            ShowForm Me, eForm_Modal, frmMain
            
            If m.bOK = True Then
                SetServerOverrides
                
                m.strUserName = cboUserName.Text
                m.strPassword = Trim(txtPassword.Text)
                m.strIP = Trim(txtServerIP.Text)
                m.strPort = Trim(txtPort.Text)
                m.strPriceIP = Trim(txtPriceIP.Text)
                m.strPricePort = Trim(txtPricePort.Text)
                m.strEnvironment = Chr(cboEnvironment.ItemData(cboEnvironment.ListIndex))
                m.bSuperTAS = CheckBoxValue(chkSuperTAS)
                
                SetIniFileProperty "LastUserName", cboUserName.Text, "User", m.strIniFile
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLoginPats.ShowMe"
    
End Function

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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.chkShowIP_Click"
    
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
    RaiseError "frmLoginPats.cmdAddLogin_Click"
    
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
    RaiseError "frmLoginPats.cmdCancel_Click"
    
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
    ElseIf Len(Trim(txtPriceIP.Text)) = 0 Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtPriceIP
        InfBox "Please enter in an IP address for the price server", "!", , "Login Error"
    ElseIf Len(Trim(txtPricePort.Text)) = 0 Then
        CheckBoxValue(chkShowIP) = True
        MoveFocus txtPricePort
        InfBox "Please enter in a Port for the price server", "!", , "Login Error"
    Else
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.cmdLogin_Click"
    
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
    RaiseError "frmLoginPats.cmdRemoveLogin_Click"
    
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
    RaiseError "frmLoginPats.Form_Activate"
    
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

    m.strUserName = ""
    m.strPassword = ""
    m.strIP = ""
    m.strPort = ""
    m.strPriceIP = ""
    m.strPricePort = ""
    m.strEnvironment = ""
    m.bSuperTAS = False

    ' Environment = (G) ptGateway, (C) ptClient, (T) ptTestClient, (g) ptTestGateway, (D) ptDemoClient
    With cboEnvironment
        .AddItem "Gateway"
        .ItemData(.NewIndex) = Asc("G")
        .AddItem "Client"
        .ItemData(.NewIndex) = Asc("C")
        .AddItem "Test Client"
        .ItemData(.NewIndex) = Asc("T")
        .AddItem "Test Gateway"
        .ItemData(.NewIndex) = Asc("g")
        .AddItem "Demo Client"
        .ItemData(.NewIndex) = Asc("D")
        
        .ListIndex = 1
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.Form_Load"
    
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
    RaiseError "frmLoginPats.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPort_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.txtPort.GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceIP_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceIP_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPriceIP

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.txtPriceIP.GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPricePort_GotFocus
'' Description: When the control gets the focus, select all the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPricePort_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtPricePort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.txtPricePort.GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtServerIP_GotFocus
'' Description: When the control gets the focus, select all the text
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
    RaiseError "frmLoginPats.txtServerIP.GotFocus"
    
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
                cboUserName.AddItem astrLogins(lIndex)
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.LoadCombo"
    
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
            End If
        Next lIndex
    End If
    
    SetCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginPats.SetCombo"
    
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
    
    strUserName = InfBox("What is your " & m.strBrokerName & " user name?", "?", , m.strBrokerName & " User Name", , , , , , "string")
    If Len(strUserName) > 0 Then
        If SetCombo(strUserName) = False Then
            cboUserName.AddItem strUserName
            
            SetCombo strUserName
            MoveFocus txtPassword
            
            strLogins = GetIniFileProperty("Logins", "", "User", m.strIniFile)
            If Len(strLogins) = 0 Then
                strLogins = strUserName
            Else
                strLogins = strLogins & "," & strUserName
            End If
            SetIniFileProperty "Logins", strLogins, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.AddLogin"
    
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
            astrList.Add astrLogins(lIndex) & vbTab & Str(lIndex)
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
    RaiseError "frmLoginPats.RemoveLogin"
    
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

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strEnvironment As String        ' Environment to override
    Dim strSuperTAS As String           ' SuperTAS to override

    txtServerIP.Text = GetIniFileProperty("IP", "", "Server", m.strConnectIni)
    txtPort.Text = GetIniFileProperty("Port", "", "Server", m.strConnectIni)
    txtPriceIP.Text = GetIniFileProperty("PriceIP", "", "Server", m.strConnectIni)
    txtPricePort.Text = GetIniFileProperty("PricePort", "", "Server", m.strConnectIni)
    cboEnvironment.Text = GetIniFileProperty("Environment", "Client", "Server", m.strConnectIni)
    strSuperTAS = GetIniFileProperty("SuperTAS", "Y", "Server", m.strConnectIni)
    CheckBoxValue(chkSuperTAS) = (strSuperTAS = "Y")
    
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
    
    strEnvironment = GetIniFileProperty("Environment", "", "Override", m.strIniFile)
    If Len(strEnvironment) > 0 Then
        If strEnvironment = cboEnvironment.Text Then
            SetIniFileProperty "Environment", "", "Override", m.strIniFile
        Else
            cboEnvironment.Text = strEnvironment
        End If
    End If
    
    strSuperTAS = GetIniFileProperty("SuperTAS", "", "Override", m.strIniFile)
    If Len(strSuperTAS) > 0 Then
        If strSuperTAS = "Y" Then
            If CheckBoxValue(chkSuperTAS) = True Then
                SetIniFileProperty "SuperTAS", "", "Override", m.strIniFile
            Else
                CheckBoxValue(chkSuperTAS) = False
            End If
        Else
            If CheckBoxValue(chkSuperTAS) = False Then
                SetIniFileProperty "SuperTAS", "", "Override", m.strIniFile
            Else
                CheckBoxValue(chkSuperTAS) = True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.SetServerControls"
    
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

    Dim strIP As String                 ' IP Address to override
    Dim strPort As String               ' Port to override
    Dim strEnvironment As String        ' Environment to override
    Dim strSuperTAS As String           ' SuperTAS to override

    strIP = GetIniFileProperty("IP", "", "Server", m.strConnectIni)
    strPort = GetIniFileProperty("Port", "", "Server", m.strConnectIni)
    
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

    strIP = GetIniFileProperty("PriceIP", "", "Server", m.strConnectIni)
    strPort = GetIniFileProperty("PricePort", "", "Server", m.strConnectIni)
    
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

    strEnvironment = GetIniFileProperty("Environment", "Client", "Server", m.strConnectIni)
    If Len(Trim(cboEnvironment.Text)) > 0 Then
        If Trim(cboEnvironment.Text) = strEnvironment Then
            SetIniFileProperty "Environment", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "Environment", Trim(cboEnvironment.Text), "Override", m.strIniFile
        End If
    End If

    strSuperTAS = GetIniFileProperty("SuperTAS", "Y", "Server", m.strConnectIni)
    If CheckBoxValue(chkSuperTAS) = True Then
        If strSuperTAS = "Y" Then
            SetIniFileProperty "SuperTAS", "", "Override", m.strIniFile
        Else
            SetIniFileProperty "SuperTAS", "N", "Override", m.strIniFile
        End If
    Else
        If strSuperTAS = "Y" Then
            SetIniFileProperty "SuperTAS", "Y", "Override", m.strIniFile
        Else
            SetIniFileProperty "SuperTAS", "", "Override", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginPats.SetServerOverrides"
    
End Sub

