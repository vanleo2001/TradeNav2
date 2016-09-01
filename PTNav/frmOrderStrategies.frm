VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOrderStrategies 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   5235
      Left            =   3240
      TabIndex        =   8
      Top             =   660
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
      Caption         =   "frmOrderStrategies.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOrderStrategies.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderStrategies.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraFavorites 
         Height          =   1335
         Left            =   0
         TabIndex        =   15
         Top             =   3600
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
         Caption         =   "frmOrderStrategies.frx":0068
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmOrderStrategies.frx":009A
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":00BA
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniRadioXP optSymSpecific 
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   1020
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
            Caption         =   "frmOrderStrategies.frx":00D6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmOrderStrategies.frx":010C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmOrderStrategies.frx":012C
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optGlobal 
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   690
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
            Caption         =   "frmOrderStrategies.frx":0148
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmOrderStrategies.frx":017E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmOrderStrategies.frx":019E
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExitA 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   305
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
            Caption         =   "frmOrderStrategies.frx":01BA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmOrderStrategies.frx":01DC
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmOrderStrategies.frx":01FC
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExitB 
            Height          =   375
            Left            =   480
            TabIndex        =   18
            Top             =   240
            Width           =   305
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
            Caption         =   "frmOrderStrategies.frx":0218
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmOrderStrategies.frx":023A
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmOrderStrategies.frx":025A
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExitC 
            Height          =   375
            Left            =   840
            TabIndex        =   17
            Top             =   240
            Width           =   305
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
            Caption         =   "frmOrderStrategies.frx":0276
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmOrderStrategies.frx":0298
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmOrderStrategies.frx":02B8
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdExitD 
            Height          =   375
            Left            =   1200
            TabIndex        =   16
            Top             =   240
            Width           =   305
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
            Caption         =   "frmOrderStrategies.frx":02D4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmOrderStrategies.frx":02F6
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmOrderStrategies.frx":0316
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSettings 
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   3000
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
         Caption         =   "frmOrderStrategies.frx":0332
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":0364
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0384
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2280
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
         Caption         =   "frmOrderStrategies.frx":03A0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":03CE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":03EE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1740
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
         Caption         =   "frmOrderStrategies.frx":040A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":0434
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0454
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1200
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
         Caption         =   "frmOrderStrategies.frx":0470
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":0498
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":04B8
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   540
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
         Caption         =   "frmOrderStrategies.frx":04D4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":0502
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0522
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   240
         TabIndex        =   9
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
         Caption         =   "frmOrderStrategies.frx":053E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":0564
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0584
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraAccount 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4875
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
      Caption         =   "frmOrderStrategies.frx":05A0
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOrderStrategies.frx":05CC
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderStrategies.frx":05EC
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdLookup 
         Height          =   255
         Left            =   4590
         TabIndex        =   5
         Top             =   30
         Width           =   240
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
         Caption         =   "frmOrderStrategies.frx":0608
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderStrategies.frx":063A
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":065A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   1755
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
         Tip             =   "frmOrderStrategies.frx":0676
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0696
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtSymbol 
         Height          =   315
         Left            =   3420
         TabIndex        =   4
         Top             =   0
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmOrderStrategies.frx":06B2
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
         Tip             =   "frmOrderStrategies.frx":06E6
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0706
      End
      Begin HexUniControls.ctlUniLabelXP lblSymbol 
         Height          =   195
         Left            =   2820
         Top             =   60
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
         Caption         =   "frmOrderStrategies.frx":0722
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOrderStrategies.frx":0750
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":0770
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAccount 
         Height          =   165
         Left            =   60
         Top             =   60
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
         Caption         =   "frmOrderStrategies.frx":078C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmOrderStrategies.frx":07BC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmOrderStrategies.frx":07DC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtDescription 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   4335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmOrderStrategies.frx":07F8
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
      Tip             =   "frmOrderStrategies.frx":0818
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderStrategies.frx":0838
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrderStrategies 
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   2895
      _cx             =   5106
      _cy             =   6165
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNew 
         Caption         =   "New Order Strategy"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Order Strategy"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Order Strategy"
      End
   End
End
Attribute VB_Name = "frmOrderStrategies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOrderStrategies.frm
'' Description: Allow the user to select or manage their order strategies
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 11/07/2011   DAJ         Utilize required module flag for auto exits
'' 01/18/2013   DAJ         Don't allow automated trading for spreads
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kFraFavHtNormal = 735
Private Const kFraFavHtRevert = 1335

Private Enum eGDOrderStrategyModes
    eGDOrderStrategyMode_Select = 0
    eGDOrderStrategyMode_Manage
End Enum

Private Enum eGDCols
    eGDCol_Favorite = 0
    eGDCol_Name
    eGDCol_Type
    eGDCol_Filename
    eGDCol_Description
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK or Cancel?
    nMode As eGDOrderStrategyModes      ' Mode that the form is being shown in
    eFavoritesMode As eExitFavoritesMode
    bAllowAccountAndSymbol As Boolean   ' Allow user to change the account and symbol?

    oSymExits As cSymExitFavorites      'object for autoexits by symbol
    strBaseSym As String
    
    ExitGlobalA As cExitStrategy        'global exit favorites
    ExitGlobalB As cExitStrategy
    ExitGlobalC As cExitStrategy
    ExitGlobalD As cExitStrategy
    
    ExitSymA As cExitStrategy           'symbol specific exit favorites
    ExitSymB As cExitStrategy
    ExitSymC As cExitStrategy
    ExitSymD As cExitStrategy
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

Private Property Get AccountID() As Long
    If cboAccounts.ListIndex >= 0 Then
        AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    Else
        AccountID = -1&
    End If
End Property

Private Property Get SymbolOrSymbolID() As Variant
    If GetSymbolID(txtSymbol.Text) = 0 Then
        SymbolOrSymbolID = txtSymbol.Text
    Else
        SymbolOrSymbolID = GetSymbolID(txtSymbol.Text)
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Account ID, Symbol or Symbol ID, Strategy to Select
'' Returns:     Order Strategy File or Blank if none selected
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(lAccountID As Long, vSymbolOrSymbolID As Variant, Optional ByVal strSelect As String = "", Optional ByVal bAllowAccountAndSymbol As Boolean = False) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol passed in
        
    m.nMode = eGDOrderStrategyMode_Select
    m.eFavoritesMode = g.ChartGlobals.eAEFMode
    m.bAllowAccountAndSymbol = bAllowAccountAndSymbol
    
    fraAccount.Visible = True
    cmdOK.Caption = "&OK"
    cmdCancel.Visible = True
    
    ' Populate the accounts combo-box...
    PopulateAccountsCbo cboAccounts, lAccountID
    
    ' Fill in the symbol...
    strSymbol = GetSymbol(vSymbolOrSymbolID)
    If Len(strSymbol) > 0 Then
        txtSymbol.Text = ConvertToTradeSymbol(strSymbol)
    ElseIf Not (ActiveChart Is Nothing) Then
        txtSymbol.Text = ConvertToTradeSymbol(GetSymbol(ActiveChart.SymbolID))
    Else
        txtSymbol.Text = ""
    End If
    
    m.strBaseSym = BaseForAutoExitFavorites(txtSymbol.Text)
    
    ExitObjectsGet m.ExitGlobalA, m.ExitGlobalB, m.ExitGlobalC, m.ExitGlobalD, eAEFMode_Global
    ExitObjectsGet m.ExitSymA, m.ExitSymB, m.ExitSymC, m.ExitSymD, eAEFMode_Symbol
    
    ' Initialize and load the order strategy grid...
    InitGrid
    LoadGrid
    
    If Len(strSelect) > 0 Then
        With fgOrderStrategies
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDCol(eGDCol_Name)) = strSelect Then
                    .Row = lIndex
                    .RowSel = lIndex
                    Exit For
                End If
            Next lIndex
        End With
    End If

    EnableControls
    cboAccounts.Enabled = bAllowAccountAndSymbol
    txtSymbol.Enabled = bAllowAccountAndSymbol
    cmdLookup.Enabled = bAllowAccountAndSymbol
    
    Form_Resize
    SetFavoritesCaption
    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        If Validate Then
            ShowMe = fgOrderStrategies.TextMatrix(fgOrderStrategies.Row, GDCol(eGDCol_Filename))
        End If
        
        lAccountID = AccountID
        vSymbolOrSymbolID = SymbolOrSymbolID
    End If
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmOrderStrategies.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeManage
'' Description: Initialize and show the form in manage mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMeManage()
On Error GoTo ErrSection:

    m.strBaseSym = ""
    m.nMode = eGDOrderStrategyMode_Manage
    m.eFavoritesMode = eAEFMode_Global
    m.bAllowAccountAndSymbol = False
    
    fraAccount.Visible = False
    cmdOK.Caption = "&Close"
    cmdCancel.Visible = False

    ExitObjectsGet m.ExitGlobalA, m.ExitGlobalB, m.ExitGlobalC, m.ExitGlobalD, eAEFMode_Global
    
    ' Initialize and load the order strategy grid...
    InitGrid
    LoadGrid
    
    EnableControls
    Form_Resize
    SetFavoritesCaption
    ShowForm Me, eForm_Modal, frmMain

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmOrderStrategies.ShowMeManage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on Cancel, exit without selection
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
    RaiseError "frmOrderStrategies.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete an order strategy (if not provided)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    DeleteStrategy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to bring up an order strategy in the editor
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditStrategy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdEdit_Click"
    
End Sub

Private Sub cmdExitA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites "A"
    Else
        AssignFavorites "A"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdExitA_MouseUp"

End Sub

Private Sub cmdExitB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites "B"
    Else
        AssignFavorites "B"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdExitB_MouseUp"

End Sub

Private Sub cmdExitC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites "C"
    Else
        AssignFavorites "C"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdExitC_MouseUp"

End Sub

Private Sub cmdExitD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        ClearFavorites "D"
    Else
        AssignFavorites "D"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdExitD_MouseUp"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdLookup_Click
'' Description: Allow the user to lookup a symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdLookup_Click()
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdLookup_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a new order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    NewStrategy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdNew_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on OK, select the current row and exit
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If Validate Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSettings_Click
'' Description: Allow the user to view the trading settings
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSettings_Click()
On Error GoTo ErrSection:

    frmTTSummaryCfg.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdSettings_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderStrategies_AfterRowColChange
'' Description: After a row change, enable/disable controls appropriately
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderStrategies_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If ValidRow Then
        DisplayDescription NewRow
    End If
    
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.fgOrderStrategies_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderStrategies_DblClick
'' Description: If the user double clicks on a valid row, click OK for them
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderStrategies_DblClick()
On Error GoTo ErrSection:

    Select Case m.nMode
        Case eGDOrderStrategyMode_Select
            If Validate Then
                m.bOK = True
                Hide
            End If
            
        Case eGDOrderStrategyMode_Manage
            If ValidRow Then
                EditStrategy
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.fgOrderStrategies_DblClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderStrategies_KeyPress
'' Description: If the user hits Enter on a valid row, click OK for them
'' Inputs:      Ascii version of Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderStrategies_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    Select Case KeyAscii
        Case vbKeyReturn
            Select Case m.nMode
                Case eGDOrderStrategyMode_Select
                    If Validate Then
                        m.bOK = True
                        Hide
                    End If
                    
                Case eGDOrderStrategyMode_Manage
                    If ValidRow Then
                        EditStrategy
                    End If
                    
            End Select
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.fgOrderStrategies_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderStrategies_KeyUp
'' Description: If the user hits Delete on a valid row, delete the strategy
'' Inputs:      Key Code Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderStrategies_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            If ValidRow Then
                DeleteStrategy
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.fgOrderStrategies_KeyUp"
    
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

    Set m.ExitGlobalA = Nothing
    Set m.ExitGlobalB = Nothing
    Set m.ExitGlobalC = Nothing
    Set m.ExitGlobalD = Nothing
    
    Set m.ExitSymA = Nothing
    Set m.ExitSymB = Nothing
    Set m.ExitSymC = Nothing
    Set m.ExitSymD = Nothing
    
    Caption = "Order Strategies"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    CenterTheForm Me
    mnuPopUp.Visible = False
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', allow the ShowMe to do the exit
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim i&, j&
    Dim bChanged As Boolean
    Dim aExitsInfo As cGdArray
    
    If g.bUnloading Then GoTo ErrExit       'let form unload when application is exiting
    
    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If
    
    If m.bOK Then
        If m.nMode = eGDOrderStrategyMode_Select Then
            If m.eFavoritesMode <> g.ChartGlobals.eAEFMode Then
                bChanged = True
                g.ChartGlobals.eAEFMode = m.eFavoritesMode
            End If
        End If
        
        If FavoritesChanged() Or bChanged Then ExitFavoritesNotify True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long
    Dim lMinScaleHeight As Long

    Select Case m.nMode
        Case eGDOrderStrategyMode_Select
            lMinScaleWidth = fraAccount.Width + 120
            lMinScaleHeight = fraAccount.Height + fraButtons.Height + txtDescription.Height + 240
            
            If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then Exit Sub
            
            With fraAccount
                .Move 60, 60
            End With
            
            With fgOrderStrategies
                .Move 60, fraAccount.Height + 120, ScaleWidth - fraButtons.Width - 180, ScaleHeight - txtDescription.Height - fraAccount.Height - 240
            End With
            
            With fraButtons
                .Move ScaleWidth - .Width - 60, fraAccount.Height + 120
            End With
            
            With txtDescription
                .Move 60, ScaleHeight - .Height - 60, ScaleWidth - 120
            End With
            
        Case eGDOrderStrategyMode_Manage
            lMinScaleWidth = fraAccount.Width + 120
            lMinScaleHeight = fraButtons.Height + txtDescription.Height + 240
        
            If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then Exit Sub
            
            With fgOrderStrategies
                .Move 60, 60, ScaleWidth - fraButtons.Width - 180, ScaleHeight - txtDescription.Height - 240
            End With
            
            With fraButtons
                .Move ScaleWidth - .Width - 60, 60
            End With
            
            With txtDescription
                .Move 60, ScaleHeight - .Height - 60, ScaleWidth - 120
            End With
    End Select

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDelete_Click
'' Description: Allow the user to delete an order strategy (if not provided)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDelete_Click()
On Error GoTo ErrSection:

    DeleteStrategy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.mnuDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEdit_Click
'' Description: Allow the user to edit an order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    EditStrategy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.mnuEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNew_Click
'' Description: Allow the user to create a new order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNew_Click()
On Error GoTo ErrSection:

    NewStrategy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.mnuNew_Click"
    
End Sub

Private Sub optGlobal_Click()
On Error GoTo ErrSection:

    Dim strMsg$

    If Not Me.Visible Then Exit Sub
    
    If optGlobal.Value = True And m.eFavoritesMode = eAEFMode_Global Then Exit Sub
    
    strMsg = "Switching to 'Global Mode' will use the same set of favorite auto-exits for all markets.||(NOTE: the mode selected will apply to| all charts and price ladders.)|"
    If InfBox(strMsg, "i", "+OK|-Cancel", "Auto-exit Favorites Mode") = "C" Then
        optSymSpecific.Value = True
        m.eFavoritesMode = eAEFMode_Symbol
    Else
        m.eFavoritesMode = eAEFMode_Global
        Form_Resize
        LoadGrid
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.optGlobal_Click"
    
End Sub

Private Sub optSymSpecific_Click()
On Error GoTo ErrSection:

    Dim strMsg$

    If Not Me.Visible Then Exit Sub

    If optSymSpecific.Value = True And m.eFavoritesMode = eAEFMode_Symbol Then Exit Sub
    
    strMsg = "Switching to 'Symbol Mode' allows selecting a different set of favorite auto-exits for each market.||(NOTE: the mode selected will apply to| all charts and price ladders.)|"
    If InfBox(strMsg, "i", "+OK|-Cancel", "Auto-exit Favorites Mode") = "C" Then
        optGlobal.Value = True
        m.eFavoritesMode = eAEFMode_Global
    Else
        m.eFavoritesMode = eAEFMode_Symbol
        Form_Resize
        LoadGrid
    End If
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.optSymSpecific_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_Click
'' Description: If the user clicks on the text box, bring up the symbol selector
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_Click(Button As Integer)
On Error GoTo ErrSection:

    LookupSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.txtSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_GotFocus
'' Description: Select the text of the text box when it gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.txtQty_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_KeyPress
'' Description: If the user types in the text box, bring up the symbol selector
'' Inputs:      Ascii Value of Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_KeyPress(KeyAscii As Integer, Shift As Integer)
On Error GoTo ErrSection:

    LookupSymbol KeyAscii
    KeyAscii = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.txtSymbol_KeyPress", 0
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSymbol_LostFocus
'' Description: Do some setup when the symbol changes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSymbol_LostFocus()
On Error GoTo ErrSection:
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.txtSymbol_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgOrderStrategies
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
'        .BackColorBkg = vbWindowBackground
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        '.GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        '.SheetBorder = RGB(128, 128, 128)
        
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 0
        .Cols = GDCol(eGDCol_NumCols)
        
        .ColHidden(GDCol(eGDCol_Type)) = True
        .ColHidden(GDCol(eGDCol_Filename)) = True
        .ColHidden(GDCol(eGDCol_Description)) = True
        
        .ColWidth(eGDCol_Favorite) = 305
        .ColAlignment(eGDCol_Favorite) = flexAlignCenterCenter
        
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim astrExits As New cGdArray       ' List of order strategy files
    Dim lIndex As Long                  ' Index into a for loop
    Dim ExitStrategy As cExitStrategy   ' Exit strategy object
    
    Dim ExitA As cExitStrategy          'exit favorites
    Dim ExitB As cExitStrategy
    Dim ExitC As cExitStrategy
    Dim ExitD As cExitStrategy
    
    'set local exit objects to get current exit info
    If m.eFavoritesMode = eAEFMode_Global Then
        Set ExitA = m.ExitGlobalA
        Set ExitB = m.ExitGlobalB
        Set ExitC = m.ExitGlobalC
        Set ExitD = m.ExitGlobalD
    Else
        Set ExitA = m.ExitSymA
        Set ExitB = m.ExitSymB
        Set ExitC = m.ExitSymC
        Set ExitD = m.ExitSymD
    End If
    
    If Not ExitA Is Nothing Then
        If Len(ExitA.FileName) = 0 Then Set ExitA = Nothing
    End If
    If Not ExitB Is Nothing Then
        If Len(ExitB.FileName) = 0 Then Set ExitB = Nothing
    End If
    If Not ExitC Is Nothing Then
        If Len(ExitC.FileName) = 0 Then Set ExitC = Nothing
    End If
    If Not ExitD Is Nothing Then
        If Len(ExitD.FileName) = 0 Then Set ExitD = Nothing
    End If
    
    With fgOrderStrategies
        .Redraw = flexRDNone
        
        .Rows = 0
        
        Set astrExits = GetExitOrderStrategies
        For lIndex = 0 To astrExits.Size - 1
            Set ExitStrategy = New cExitStrategy
            ExitStrategy.Load Parse(astrExits(lIndex), vbTab, 2)
            
            If HasModule(ExitStrategy.Required) Then
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = ExitStrategy
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = ExitStrategy.StrategyName
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = "Exit"
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Filename)) = ExitStrategy.FileName
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Description)) = ExitStrategy.Description
                
                'check if this is an exit favorite
                If Not ExitA Is Nothing Then
                    If ExitA.FileName = ExitStrategy.FileName Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Favorite)) = "A"
                        Set ExitA = Nothing     'so won't keep checking
                    End If
                End If
                
                If Not ExitB Is Nothing Then
                    If ExitB.FileName = ExitStrategy.FileName Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Favorite)) = "B"
                        Set ExitB = Nothing
                    End If
                End If
            
                If Not ExitC Is Nothing Then
                    If ExitC.FileName = ExitStrategy.FileName Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Favorite)) = "C"
                        Set ExitC = Nothing
                    End If
                End If
            
                If Not ExitD Is Nothing Then
                    If ExitD.FileName = ExitStrategy.FileName Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Favorite)) = "D"
                        Set ExitD = Nothing
                    End If
                End If
            End If
        Next lIndex
        
        If .Rows > 0 Then
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortStringAscending
            
            .Row = 0
            .RowSel = 0
        End If
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bValidRow As Boolean            ' Is this a valid row?
    Dim bCustom As Boolean              ' Is this a custom order strategy?

    With fgOrderStrategies
        bValidRow = ValidRow
        If bValidRow Then bCustom = Not .RowData(.Row).Provided
    End With
    
    Enable cmdOK, bValidRow
    Enable cmdCancel, True
    
    Enable cmdNew, True
    Enable mnuNew, True
    Enable cmdEdit, bValidRow
    Enable mnuEdit, bValidRow
    Enable cmdDelete, bValidRow And bCustom
    Enable mnuDelete, bValidRow And bCustom
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.EnableControls"
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRow
'' Description: Is the currently selected grid row a valid row?
'' Inputs:      None
'' Returns:     True if Valid Row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRow() As Boolean
On Error GoTo ErrSection:

    With fgOrderStrategies
        ValidRow = (.Row >= 0 And .Row < .Rows)
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderStrategies.ValidRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteStrategy
'' Description: Allow the user to delete an order strategy (if not provided)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteStrategy()
On Error GoTo ErrSection:

    Dim ExitStrategy As cExitStrategy   ' Exit strategy object

    If ValidRow Then
        With fgOrderStrategies
            If UCase(.TextMatrix(.Row, GDCol(eGDCol_Type))) = "EXIT" Then
                Set ExitStrategy = .RowData(.Row)
                If ExitStrategy.Provided Then
                    InfBox "You cannot delete a provided exit order strategy", "!", , "Order Strategy Error"
                Else
                    If InfBox("Are you sure you want to delete|" & ExitStrategy.StrategyName & "?", "?", "+Yes|-No", "Order Strategy Delete Confirm") = "Y" Then
                        ClearFavorites .TextMatrix(.Row, eGDCol_Favorite)
                        KillFile AddSlash(App.Path) & ExitStrategy.FileName
                        .RemoveItem .Row
                    End If
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.DeleteStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditStrategy
'' Description: Allow the user to edit an order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditStrategy()
On Error GoTo ErrSection:

    Dim ExitStrategy As cExitStrategy   ' Exit strategy object
    Dim bReload As Boolean              ' If the user did a save as, reload the grid
    
    Dim strFavorite$

    If ValidRow Then
        With fgOrderStrategies
            If UCase(.TextMatrix(.Row, GDCol(eGDCol_Type))) = "EXIT" Then
                Set ExitStrategy = .RowData(.Row)
                If frmExitStrategy.ShowMe(ExitStrategy, False, bReload) Then
                    If bReload Then
                        'user chose "Save As" to make copy of strategy, do not make changes to favorites
                        LoadGrid
                        SelectStrategy ExitStrategy.StrategyName
                    Else
                        strFavorite = .TextMatrix(.Row, GDCol(eGDCol_Favorite))
                        DisplayExitStrategy ExitStrategy, .Row
                    End If
                End If
            End If
        End With
    End If
    
    'chekc if edited strategy was a favorite
    Select Case strFavorite
        Case "A"
            If m.eFavoritesMode = eAEFMode_Symbol Then
                Set m.ExitSymA = Nothing
                Set m.ExitSymA = ExitStrategy
            Else
                Set m.ExitGlobalA = Nothing
                Set m.ExitGlobalA = ExitStrategy
            End If
        Case "B"
            If m.eFavoritesMode = eAEFMode_Symbol Then
                Set m.ExitSymB = Nothing
                Set m.ExitSymB = ExitStrategy
            Else
                Set m.ExitGlobalB = Nothing
                Set m.ExitGlobalB = ExitStrategy
            End If
        Case "C"
            If m.eFavoritesMode = eAEFMode_Symbol Then
                Set m.ExitSymC = Nothing
                Set m.ExitSymC = ExitStrategy
            Else
                Set m.ExitGlobalC = Nothing
                Set m.ExitGlobalC = ExitStrategy
            End If
        Case "D"
            If m.eFavoritesMode = eAEFMode_Symbol Then
                Set m.ExitSymD = Nothing
                Set m.ExitSymD = ExitStrategy
            Else
                Set m.ExitGlobalD = Nothing
                Set m.ExitGlobalD = ExitStrategy
            End If
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.EditStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewStrategy
'' Description: Allow the user to create a new order strategy
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewStrategy()
On Error GoTo ErrSection:

    Dim ExitStrategy As cExitStrategy   ' Exit strategy object
    
    Set ExitStrategy = New cExitStrategy
    If frmExitStrategy.ShowMe(ExitStrategy, False) Then
        If Len(ExitStrategy.StrategyName) > 0 Then
            With fgOrderStrategies
                .Redraw = flexRDNone
                
                .Rows = .Rows + 1
                DisplayExitStrategy ExitStrategy, .Rows - 1
                
                .Redraw = flexRDBuffered
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.NewStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectStrategy
'' Description: Select a specific strategy in the list
'' Inputs:      Strategy to select
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SelectStrategy(ByVal strStrategyName As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgOrderStrategies
        For lIndex = 0 To .Rows - 1
            If (.TextMatrix(lIndex, GDCol(eGDCol_Name)) = strStrategyName) Then
                .Row = lIndex
                .RowSel = lIndex
                
                DisplayDescription lIndex
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.SelectStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayExitStrategy
'' Description: Display an exit strategy in the grid
'' Inputs:      Strategy to display, Row to change
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayExitStrategy(ExitStrategy As cExitStrategy, ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the redraw

    With fgOrderStrategies
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .RowData(lRow) = ExitStrategy
        .TextMatrix(lRow, GDCol(eGDCol_Name)) = ExitStrategy.StrategyName
        .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Exit"
        .TextMatrix(lRow, GDCol(eGDCol_Filename)) = ExitStrategy.FileName
        .TextMatrix(lRow, GDCol(eGDCol_Description)) = ExitStrategy.Description
        
        .Col = GDCol(eGDCol_Name)
        .Sort = flexSortStringAscending
        
        SelectStrategy ExitStrategy.StrategyName
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.DisplayExitStrategy"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayDescription
'' Description: Display the description for the given row in the grid
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayDescription(ByVal lRow As Long)
On Error GoTo ErrSection:

    txtDescription.Text = fgOrderStrategies.TextMatrix(lRow, GDCol(eGDCol_Description))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.DisplayDescription"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Lookup a symbol for the user to trade
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LookupSymbol(Optional ByVal KeyAscii As Long = 0&)
On Error GoTo ErrSection:

    Dim astrSymbol As New cGdArray      ' Array to get lookup symbol from
    Dim strSymbol As String
    
    If KeyAscii = 0& Then
        Set astrSymbol = frmSymbolSelector.ShowMe(txtSymbol.Text, False, True, "Symbol for Order Strategy", , , True)
    Else
        Set astrSymbol = frmSymbolSelector.ShowMe(Chr(KeyAscii), False, True, "Symbol for Order Strategy", False, False, True)
    End If
    If astrSymbol.Size > 0 Then
        strSymbol = ConvertToTradeSymbol(astrSymbol(0), Date)
        
        If strSymbol <> UCase(Trim(txtSymbol.Text)) Then
            If ValidAutomatedSymbol(AccountID, strSymbol, "Auto Exit", "Order Strategies") Then
                txtSymbol.Text = strSymbol
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.LookupSymbol"
    
End Sub

Private Sub ClearFavorites(ByVal strFavorite$)
On Error GoTo ErrSection:

    Dim globalIdx&, i&
    Dim strGridText$

    If Not ValidRow Then Exit Sub
    If Len(strFavorite) = 0 Then Exit Sub       'no need to waste time processing
    
    If m.eFavoritesMode = eAEFMode_Global Then
        Select Case strFavorite
            Case "A"
                Set m.ExitGlobalA = Nothing
            Case "B"
                Set m.ExitGlobalB = Nothing
            Case "C"
                Set m.ExitGlobalC = Nothing
            Case "D"
                Set m.ExitGlobalD = Nothing
        End Select
    Else
        Select Case strFavorite
            Case "A"
                Set m.ExitSymA = Nothing
            Case "B"
                Set m.ExitSymB = Nothing
            Case "C"
                Set m.ExitSymC = Nothing
            Case "D"
                Set m.ExitSymD = Nothing
        End Select
    End If
    
    With fgOrderStrategies
        'walk through grid and clear text in grid if this button currently assigned to something else
        For i = .FixedRows To .Rows - 1
            strGridText = .TextMatrix(i, eGDCol_Favorite)
            If strGridText = strFavorite Then
                .TextMatrix(i, eGDCol_Favorite) = ""
                Exit For
            End If
        Next
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.ClearFavorites"
    
End Sub

Private Sub AssignFavorites(ByVal strFavorite$)
On Error GoTo ErrSection:

    Dim globalIdx&, i&
    Dim strGridText$, strMsg$
    
    Dim NewExit As cExitStrategy
    
    globalIdx = -1
    
    If Not ValidRow Then Exit Sub
    
    With fgOrderStrategies
    
        strGridText = .TextMatrix(.Row, eGDCol_Favorite)
        
        If strGridText = strFavorite Then
            'button already assigned to this auto exit, prompt to clear
            strMsg = "Clear assignment for button '" & strFavorite & "'?"
            If InfBox(strMsg, "?", "Yes|No") = "Y" Then
                .TextMatrix(.Row, eGDCol_Favorite) = ""
                ClearFavorites strFavorite
            End If
        Else
            'call clear in case strategy is assigned to some other button
            ClearFavorites strFavorite
            ClearFavorites strGridText
            
            Set NewExit = New cExitStrategy
            
            If NewExit.Load(fgOrderStrategies.TextMatrix(fgOrderStrategies.Row, GDCol(eGDCol_Filename))) Then
                If m.eFavoritesMode = eAEFMode_Global Then
                    Select Case strFavorite
                        Case "A"
                            Set m.ExitGlobalA = Nothing
                            Set m.ExitGlobalA = NewExit
                        Case "B"
                            Set m.ExitGlobalB = Nothing
                            Set m.ExitGlobalB = NewExit
                        Case "C"
                            Set m.ExitGlobalC = Nothing
                            Set m.ExitGlobalC = NewExit
                        Case "D"
                            Set m.ExitGlobalD = Nothing
                            Set m.ExitGlobalD = NewExit
                    End Select
                Else
                    Select Case strFavorite
                        Case "A"
                            Set m.ExitSymA = Nothing
                            Set m.ExitSymA = NewExit
                        Case "B"
                            Set m.ExitSymB = Nothing
                            Set m.ExitSymB = NewExit
                        Case "C"
                            Set m.ExitSymC = Nothing
                            Set m.ExitSymC = NewExit
                        Case "D"
                            Set m.ExitSymD = Nothing
                            Set m.ExitSymD = NewExit
                    End Select
                End If
                
                .TextMatrix(.Row, eGDCol_Favorite) = strFavorite
            End If
        End If
    
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.AssignFavorites"
    
End Sub

Private Sub ExitObjectsGet(a As cExitStrategy, b As cExitStrategy, _
    c As cExitStrategy, d As cExitStrategy, ByVal eMode As eExitFavoritesMode)
On Error GoTo ErrExit

    'this function is meant to be called only from ShowMe and ShowMeManage
    'to set the local exit strategy objects on initial show of form

    If eMode = eAEFMode_Symbol Then
        Set m.oSymExits = g.ChartGlobals.treeExitFavorites.Item(m.strBaseSym)
    End If

    If m.oSymExits Is Nothing Or eMode = eAEFMode_Global Then
        If Not g.ChartGlobals.aExitFavorites Is Nothing Then
            Set a = g.ChartGlobals.aExitFavorites(0)
            Set b = g.ChartGlobals.aExitFavorites(1)
            Set c = g.ChartGlobals.aExitFavorites(2)
            Set d = g.ChartGlobals.aExitFavorites(3)
            
            If Len(a.FileName) = 0 Then Set a = Nothing
            If Len(b.FileName) = 0 Then Set b = Nothing
            If Len(c.FileName) = 0 Then Set c = Nothing
            If Len(d.FileName) = 0 Then Set d = Nothing
            
        End If
    Else
        Set a = m.oSymExits.ExitObjectGet(m.strBaseSym, "A")
        Set b = m.oSymExits.ExitObjectGet(m.strBaseSym, "B")
        Set c = m.oSymExits.ExitObjectGet(m.strBaseSym, "C")
        Set d = m.oSymExits.ExitObjectGet(m.strBaseSym, "D")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.ExitObjectsGet"
    
End Sub

Private Sub SetFavoritesCaption()
On Error GoTo ErrSection

    Dim strCaption$

    If m.nMode = eGDOrderStrategyMode_Manage Then
        Me.optGlobal.Visible = False
        Me.optSymSpecific.Visible = False
        fraFavorites.Height = kFraFavHtNormal
    Else
        Me.optGlobal.Visible = True
        Me.optSymSpecific.Visible = True
        fraFavorites.Height = kFraFavHtRevert
        
        If m.eFavoritesMode = eAEFMode_Global Then
            Me.optGlobal.Value = True
        Else
            Me.optSymSpecific.Value = True
        End If
        
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.SetFavoritesCaption"
    
End Sub

Private Function FavoritesChanged() As Boolean
On Error GoTo ErrExit:

    Dim oExit As cExitStrategy
    Dim bChanged As Boolean
    
    If m.eFavoritesMode = eAEFMode_Global Then
        'Global A
        If Not m.ExitGlobalA Is g.ChartGlobals.aExitFavorites(0) Then
            bChanged = True
            Set g.ChartGlobals.aExitFavorites(0) = Nothing
            
            If m.ExitGlobalA Is Nothing Then
                Set g.ChartGlobals.aExitFavorites(0) = New cExitStrategy        'needs to have valid object
            ElseIf Len(m.ExitGlobalA.FileName) > 0 Then
                Set g.ChartGlobals.aExitFavorites(0) = m.ExitGlobalA
            End If
        End If
    
        'Global B
        If Not m.ExitGlobalB Is g.ChartGlobals.aExitFavorites(1) Then
            bChanged = True
            Set g.ChartGlobals.aExitFavorites(1) = Nothing
            
            If m.ExitGlobalB Is Nothing Then
                Set g.ChartGlobals.aExitFavorites(1) = New cExitStrategy        'needs to have valid object
            ElseIf Len(m.ExitGlobalB.FileName) > 0 Then
                Set g.ChartGlobals.aExitFavorites(1) = m.ExitGlobalB
            End If
        End If
    
        'Global C
        If Not m.ExitGlobalC Is g.ChartGlobals.aExitFavorites(2) Then
            bChanged = True
            Set g.ChartGlobals.aExitFavorites(2) = Nothing
            
            If m.ExitGlobalC Is Nothing Then
                Set g.ChartGlobals.aExitFavorites(2) = New cExitStrategy        'needs to have valid object
            ElseIf Len(m.ExitGlobalC.FileName) > 0 Then
                Set g.ChartGlobals.aExitFavorites(2) = m.ExitGlobalC
            End If
        End If
    
        'Global D
        If Not m.ExitGlobalD Is g.ChartGlobals.aExitFavorites(3) Then
            bChanged = True
            Set g.ChartGlobals.aExitFavorites(3) = Nothing
            
            If m.ExitGlobalD Is Nothing Then
                Set g.ChartGlobals.aExitFavorites(3) = New cExitStrategy        'needs to have valid object
            ElseIf Len(m.ExitGlobalA.FileName) > 0 Then
                Set g.ChartGlobals.aExitFavorites(3) = m.ExitGlobalD
            End If
        End If
    Else
        
        Set m.oSymExits = g.ChartGlobals.treeExitFavorites.Item(m.strBaseSym)
    
        If m.oSymExits Is Nothing Then
            'add symbol specific favorites for first time
            SaveFavoritesForSym
        Else
            'update existing symbol specific info
            If Not m.ExitSymA Is m.oSymExits.ExitObjectGet(m.strBaseSym, "A") Then
                bChanged = True
                If m.ExitSymA Is Nothing Then
                    m.oSymExits.ExitObjectClear m.strBaseSym, "A"
                Else
                    m.oSymExits.ExitObjectSet m.strBaseSym, "A", m.ExitSymA
                End If
            End If
    
            If Not m.ExitSymB Is m.oSymExits.ExitObjectGet(m.strBaseSym, "B") Then
                bChanged = True
                If m.ExitSymA Is Nothing Then
                    m.oSymExits.ExitObjectClear m.strBaseSym, "B"
                Else
                    m.oSymExits.ExitObjectSet m.strBaseSym, "B", m.ExitSymB
                End If
            End If
    
            If Not m.ExitSymC Is m.oSymExits.ExitObjectGet(m.strBaseSym, "C") Then
                bChanged = True
                If m.ExitSymC Is Nothing Then
                    m.oSymExits.ExitObjectClear m.strBaseSym, "C"
                Else
                    m.oSymExits.ExitObjectSet m.strBaseSym, "C", m.ExitSymC
                End If
            End If
    
            If Not m.ExitSymD Is m.oSymExits.ExitObjectGet(m.strBaseSym, "D") Then
                bChanged = True
                If m.ExitSymD Is Nothing Then
                    m.oSymExits.ExitObjectClear m.strBaseSym, "D"
                Else
                    m.oSymExits.ExitObjectSet m.strBaseSym, "D", m.ExitSymD
                End If
            End If
        End If
    End If
    
    FavoritesChanged = bChanged

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderStrategies.FavoritesChanged"
    
End Function

Private Sub SaveFavoritesForSym()
On Error GoTo ErrExit

    Dim i&
    
    Dim a As cExitStrategy
    Dim b As cExitStrategy
    Dim c As cExitStrategy
    Dim d As cExitStrategy
    
    Dim aExitsInfo As New cGdArray
    
    If m.ExitGlobalA Is m.ExitSymA And m.ExitGlobalB Is m.ExitSymB Then
        If m.ExitGlobalC Is m.ExitSymC And m.ExitGlobalD Is m.ExitSymD Then
            Exit Sub
        End If
    End If
    
    If Not m.ExitSymA Is Nothing Then
        If Len(m.ExitSymA.FileName) > 0 Then aExitsInfo.Add "A=" & m.ExitSymA.FileName
    End If
    If Not m.ExitSymB Is Nothing Then
        If Len(m.ExitSymB.FileName) > 0 Then aExitsInfo.Add "B=" & m.ExitSymB.FileName
    End If
    If Not m.ExitSymC Is Nothing Then
        If Len(m.ExitSymC.FileName) > 0 Then aExitsInfo.Add "C=" & m.ExitSymC.FileName
    End If
    If Not m.ExitSymD Is Nothing Then
        If Len(m.ExitSymD.FileName) > 0 Then aExitsInfo.Add "D=" & m.ExitSymD.FileName
    End If
    
    If aExitsInfo.Size <= 0 Then Exit Sub
    
    Set m.oSymExits = New cSymExitFavorites
    m.oSymExits.ExitFavoritesSet m.strBaseSym, aExitsInfo
    
    i = g.ChartGlobals.treeExitFavorites.Add(m.oSymExits, m.strBaseSym)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderStrategies.SaveFavoritesForSym"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Determine if everything is valid
'' Inputs:      None
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Validate() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Are the account and symbol combination valid?
    Dim ExitStrategy As cExitStrategy   ' Exit strategy object
    Dim strTarget As String             ' Which target has an issue?
    Dim lQuantity As Long               ' Quantity

    bReturn = False
    If ValidRow Then
        bReturn = True
        If m.bAllowAccountAndSymbol Then
            bReturn = CanActivateAutomatedItem(AccountID, SymbolOrSymbolID, "Auto Exit", "Order Strategies")
        End If
        If (bReturn = True) And (m.nMode = eGDOrderStrategyMode_Select) Then
            Set ExitStrategy = fgOrderStrategies.RowData(fgOrderStrategies.Row)
            If ExitStrategy.SpecifyLots Then
                If ExitStrategy.UseTarget1 Then
                    bReturn = g.Broker.ValidQuantity(AccountID, SymbolOrSymbolID, ExitStrategy.Target1Quantity)
                    If bReturn = False Then
                        strTarget = "first"
                        lQuantity = ExitStrategy.Target1Quantity
                    End If
                End If
                If (bReturn = True) And (ExitStrategy.UseTarget2 = True) Then
                    bReturn = g.Broker.ValidQuantity(AccountID, SymbolOrSymbolID, ExitStrategy.Target2Quantity)
                    If bReturn = False Then
                        strTarget = "second"
                        lQuantity = ExitStrategy.Target2Quantity
                    End If
                End If
                
                If bReturn = False Then
                    InfBox "The " & strTarget & " target quantity (" & Str(lQuantity) & ") is invalid for '" & GetSymbol(SymbolOrSymbolID) & "'", "!", , "Order Strategy Selection Error"
                End If
            End If
        End If
    Else
        InfBox "You must select a valid item in the list", "!", , "Order Strategy Selection Error"
    End If
    
    Validate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderStrategies.cmdOK_Click"
    
End Function


