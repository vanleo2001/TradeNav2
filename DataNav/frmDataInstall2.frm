VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDataInstall2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Install Historical Data"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmDataInstall2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vsOcx6LibCtl.vsElastic vsWeb 
      Height          =   975
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2100
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   $"frmDataInstall2.frx":0442
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   5
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin HexUniControls.ctlUniFrameWL fraHistory 
      Height          =   4035
      Left            =   240
      TabIndex        =   27
      Top             =   1680
      Width           =   5355
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
      Caption         =   "frmDataInstall2.frx":04D9
      Enabled         =   -1  'True
      ForeColor       =   8388608
      BackColor       =   -2147483633
      Tip             =   "frmDataInstall2.frx":0547
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":0567
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraSecType 
         Height          =   615
         Index           =   3
         Left            =   180
         TabIndex        =   35
         Top             =   3240
         Width           =   4935
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
         Caption         =   "frmDataInstall2.frx":0583
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDataInstall2.frx":05AF
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":05CF
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkAll 
            Height          =   375
            Index           =   3
            Left            =   3900
            TabIndex        =   17
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":05EB
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0625
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0645
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSome 
            Height          =   375
            Index           =   3
            Left            =   2820
            TabIndex        =   16
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":0661
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":06A1
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":06C1
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkDaily 
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   15
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":06DD
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0711
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0731
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSecType 
            Height          =   435
            Index           =   3
            Left            =   120
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":074D
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDataInstall2.frx":0795
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":07B5
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin VB.Line lineSep 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   0
            X2              =   5820
            Y1              =   0
            Y2              =   0
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSecType 
         Height          =   675
         Index           =   2
         Left            =   180
         TabIndex        =   33
         Top             =   2520
         Width           =   4935
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
         Caption         =   "frmDataInstall2.frx":07D1
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDataInstall2.frx":07FD
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":081D
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkAll 
            Height          =   375
            Index           =   2
            Left            =   3900
            TabIndex        =   14
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":0839
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0873
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0893
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSome 
            Height          =   375
            Index           =   2
            Left            =   2820
            TabIndex        =   13
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":08AF
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":08F1
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0911
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkDaily 
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   12
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":092D
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0961
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0981
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSecType 
            Height          =   435
            Index           =   2
            Left            =   120
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":099D
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDataInstall2.frx":09F1
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0A11
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin VB.Line lineSep 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   0
            X2              =   5820
            Y1              =   0
            Y2              =   0
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSecType 
         Height          =   615
         Index           =   1
         Left            =   180
         TabIndex        =   31
         Top             =   1800
         Width           =   4935
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
         Caption         =   "frmDataInstall2.frx":0A2D
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDataInstall2.frx":0A59
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":0A79
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkDaily 
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   9
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":0A95
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0AC9
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0AE9
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSome 
            Height          =   375
            Index           =   1
            Left            =   2820
            TabIndex        =   10
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":0B05
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0B49
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0B69
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkAll 
            Height          =   375
            Index           =   1
            Left            =   3900
            TabIndex        =   11
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":0B85
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0BBF
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0BDF
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSecType 
            Height          =   435
            Index           =   1
            Left            =   120
            Top             =   180
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
            Caption         =   "frmDataInstall2.frx":0BFB
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDataInstall2.frx":0C3D
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0C5D
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin VB.Line lineSep 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   0
            X2              =   5820
            Y1              =   0
            Y2              =   0
         End
      End
      Begin HexUniControls.ctlUniFrameWL fraSecType 
         Height          =   435
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   1260
         Width           =   4935
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
         Caption         =   "frmDataInstall2.frx":0C79
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDataInstall2.frx":0CA5
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":0CC5
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniCheckXP chkAll 
            Height          =   375
            Index           =   0
            Left            =   3900
            TabIndex        =   8
            Top             =   0
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
            Caption         =   "frmDataInstall2.frx":0CE1
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0D1B
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0D3B
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkSome 
            Height          =   375
            Index           =   0
            Left            =   2820
            TabIndex        =   7
            Top             =   0
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
            Caption         =   "frmDataInstall2.frx":0D57
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0D9B
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0DBB
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chkDaily 
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   6
            Top             =   0
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
            Caption         =   "frmDataInstall2.frx":0DD7
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frmDataInstall2.frx":0E0B
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0E2B
            ShowFocus       =   -1  'True
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblSecType 
            Height          =   435
            Index           =   0
            Left            =   120
            Top             =   0
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
            Caption         =   "frmDataInstall2.frx":0E47
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   2
            VAlignment      =   0
            BackStyle       =   0
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDataInstall2.frx":0E93
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDataInstall2.frx":0EB3
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
         Begin VB.Line lineSep 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   0
            X2              =   5820
            Y1              =   -60
            Y2              =   -60
         End
      End
      Begin HexUniControls.ctlUniLabelXP lblNote 
         Height          =   195
         Left            =   180
         Top             =   300
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
         Caption         =   "frmDataInstall2.frx":0ECF
         BackColor       =   -2147483633
         ForeColor       =   4210752
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":0F5D
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":0F7D
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblNote2 
         Height          =   195
         Left            =   180
         Top             =   510
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
         Caption         =   "frmDataInstall2.frx":0F99
         BackColor       =   -2147483633
         ForeColor       =   4210752
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":1045
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1065
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblIntraday 
         Height          =   195
         Left            =   2940
         Top             =   840
         Visible         =   0   'False
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
         Caption         =   "frmDataInstall2.frx":1081
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":10C1
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":10E1
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin VB.Line lineIntraday 
         BorderColor     =   &H80000012&
         Visible         =   0   'False
         X1              =   2940
         X2              =   4980
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin vsOcx6LibCtl.vsElastic vsNotDisk1 
      Height          =   735
      Left            =   3300
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   255
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   $"frmDataInstall2.frx":10FD
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   5
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin vsOcx6LibCtl.vsElastic vsInvalid 
      Height          =   735
      Left            =   3360
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   255
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   "Could not find installation data on the specifed drive (requires a valid DataInst.CFG file)"
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   5
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin HexUniControls.ctlUniFrameWL fraPaths 
      Height          =   1185
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   5355
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
      Caption         =   "frmDataInstall2.frx":119F
      Enabled         =   -1  'True
      ForeColor       =   8388608
      BackColor       =   -2147483633
      Tip             =   "frmDataInstall2.frx":11ED
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":120D
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optFTP 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   330
         Width           =   3195
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
         Caption         =   "frmDataInstall2.frx":1229
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":128F
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":12AF
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdBrowse 
         Height          =   270
         Left            =   5025
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
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
         Caption         =   "frmDataInstall2.frx":12CB
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":12F1
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1311
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniDriveBoxXP drvFrom 
         Height          =   315
         Left            =   3180
         TabIndex        =   5
         Top             =   690
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         BorderColor     =   16711680
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   16777215
         ButtonForeColor =   0
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
         Tip             =   "frmDataInstall2.frx":132D
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         BackColorOut    =   16777215
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":134D
         DropDownWidth   =   -1
      End
      Begin HexUniControls.ctlUniRadioXP optWeb 
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   3075
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
         Caption         =   "frmDataInstall2.frx":1369
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmDataInstall2.frx":13C7
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":13E7
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFrom 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3015
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
         Caption         =   "frmDataInstall2.frx":1403
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":145D
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":147D
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   195
         Left            =   3480
         Top             =   345
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
         Caption         =   "frmDataInstall2.frx":1499
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":14D3
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":14F3
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   195
         Left            =   3300
         Top             =   1080
         Visible         =   0   'False
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
         Caption         =   "frmDataInstall2.frx":150F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":1563
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1583
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFrom 
         Height          =   255
         Left            =   2760
         Top             =   660
         Visible         =   0   'False
         Width           =   2505
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
         Caption         =   "frmDataInstall2.frx":159F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":15F9
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1619
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdInstall 
      Height          =   435
      Left            =   3360
      TabIndex        =   0
      Top             =   6150
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmDataInstall2.frx":1635
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmDataInstall2.frx":166D
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":168D
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraSpace 
      Height          =   705
      Left            =   420
      TabIndex        =   20
      Top             =   5940
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
      Caption         =   "frmDataInstall2.frx":16A9
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDataInstall2.frx":16D9
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":16F9
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblSpaceAvailable 
         Height          =   225
         Left            =   1140
         Top             =   420
         Width           =   1395
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
         Caption         =   "frmDataInstall2.frx":1715
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":1749
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1769
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label8 
         Height          =   225
         Left            =   180
         Top             =   420
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
         Caption         =   "frmDataInstall2.frx":1785
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":17C5
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":17E5
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblSpaceRequired 
         Height          =   225
         Left            =   1200
         Top             =   180
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
         Caption         =   "frmDataInstall2.frx":1801
         BackColor       =   -2147483633
         ForeColor       =   12583104
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":1831
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1851
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   225
         Left            =   180
         Top             =   180
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
         Caption         =   "frmDataInstall2.frx":186D
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":18AB
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":18CB
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txtHwnd 
      Height          =   285
      Left            =   3420
      TabIndex        =   26
      Top             =   -120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmDataInstall2.frx":18E7
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
      Tip             =   "frmDataInstall2.frx":1911
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":1931
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4560
      TabIndex        =   1
      Top             =   6150
      Width           =   855
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
      Caption         =   "frmDataInstall2.frx":194D
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmDataInstall2.frx":197B
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":199B
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsElastic vsProgress 
      Height          =   315
      Left            =   240
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6180
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   3
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   8421504
      ForeColor       =   16777215
      FloodColor      =   12582912
      ForeColorDisabled=   -2147483631
      Caption         =   "50%"
      Align           =   0
      Appearance      =   3
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   1
      FloodPercent    =   50
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
   End
   Begin HexUniControls.ctlUniFrameWL fraWeb 
      Height          =   4155
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   5355
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
      Caption         =   "frmDataInstall2.frx":19B7
      Enabled         =   -1  'True
      ForeColor       =   8388608
      BackColor       =   -2147483633
      Tip             =   "frmDataInstall2.frx":1A27
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDataInstall2.frx":1A47
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optWebSet 
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   28
         Top             =   3300
         Visible         =   0   'False
         Width           =   4635
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
         Caption         =   "frmDataInstall2.frx":1A63
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":1ABF
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1ADF
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optWebSet 
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   4635
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
         Caption         =   "frmDataInstall2.frx":1AFB
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":1B57
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1B77
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optWebSet 
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   32
         Top             =   1740
         Visible         =   0   'False
         Width           =   4635
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
         Caption         =   "frmDataInstall2.frx":1B93
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":1BEF
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1C0F
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optWebSet 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   900
         Visible         =   0   'False
         Width           =   4935
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
         Caption         =   "frmDataInstall2.frx":1C2B
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDataInstall2.frx":1C87
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1CA7
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblWebSet 
         Height          =   390
         Index           =   3
         Left            =   840
         Top             =   3540
         Visible         =   0   'False
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
         Caption         =   "frmDataInstall2.frx":1CC3
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmDataInstall2.frx":1DCD
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1DED
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
      Begin HexUniControls.ctlUniLabelXP lblWebHeader 
         Height          =   495
         Left            =   240
         Top             =   300
         Width           =   4935
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
         Caption         =   "frmDataInstall2.frx":1E09
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDataInstall2.frx":1F17
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":1F37
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblWebSet 
         Height          =   390
         Index           =   2
         Left            =   840
         Top             =   2760
         Visible         =   0   'False
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
         Caption         =   "frmDataInstall2.frx":1F53
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmDataInstall2.frx":205D
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":207D
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
      Begin HexUniControls.ctlUniLabelXP lblWebSet 
         Height          =   390
         Index           =   1
         Left            =   840
         Top             =   1980
         Visible         =   0   'False
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
         Caption         =   "frmDataInstall2.frx":2099
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmDataInstall2.frx":21A3
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":21C3
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
      Begin HexUniControls.ctlUniLabelXP lblWebSet 
         Height          =   390
         Index           =   0
         Left            =   540
         Top             =   1140
         Visible         =   0   'False
         Width           =   4635
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
         Caption         =   "frmDataInstall2.frx":21DF
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Tip             =   "frmDataInstall2.frx":22E9
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDataInstall2.frx":2309
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmDataInstall2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDataInstall2.frm
'' Description: Performs a data install according to what the user chooses
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/21/2015   DAJ         Send URL's through FixURL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kTempFolder As String = "\_Temp_\"

Private Type mPrivate
    bInstall As Boolean
    bAborted As Boolean
    bAddMode As Boolean
    strFromPath As String
    strUnzipPath As String
    bSkipShowSizes As Boolean
    bMoreDisks As Boolean
    strFtpLogin As String
    strFtpDataset As String
    
    ' used when unzipping all the zip files
    strZipFiles As String
    dZipFileSize As Double
    dTotalBytes As Double
    dBytesSoFar As Double
End Type
Private m As mPrivate

Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    Dim s$, strDrive$, bPromptForNextDisk As Boolean
    Static bInProgress As Boolean
    
    ' setup an empty temporary folder to unzip things into
    m.strUnzipPath = App.Path & kTempFolder
    MakeDir m.strUnzipPath, False
    
    ' see if an FTP dataset is downloading or needs to be linked in
    If g.FtpDownloader.DownloaderIsRunning Then
        InfBox "Please wait until the FTP Downloader has finished.", "!", , "Data Install"
        Exit Function
    End If
    
    ' see if FTP dataset available
    If Not GetFtpDataset Then
        optFTP.Enabled = False
    End If
    
DoNextDisk:
    If g.bUnloading Or bInProgress Then Exit Function
    If Not g.bStarting Then
        ' can't check for ProcessIsBusy when need to auto-install data during startup
        If ProcessIsBusy Then Exit Function
    End If
    If optFTP And bPromptForNextDisk Then
        s = "Select more historical data to install now?||(or you can install more later by selecting|'Install Data' under the 'File' menu)"
        If InfBox(s, "?", "+Yes|-No", "Data Install") = "N" Then
            Exit Function
        End If
    End If
    If g.RealTime.Active Then
        InfBox "Please turn realtime streaming OFF before installing data.", "!", , "Data Install"
        Exit Function
    End If

    bInProgress = True
    frmStatus.IsBusy = True
    bPromptForNextDisk = False
    
    cmdInstall.Enabled = True
    cmdInstall.Visible = True
    cmdCancel.Enabled = True
    cmdCancel.Caption = "&Cancel"
    fraHistory.Enabled = True
    fraPaths.Enabled = True
    vsProgress.FloodPercent = 0
    vsProgress.Caption = ""
    vsProgress.Visible = False
    fraSpace.Visible = True

    ' find CD/DVD (if it's in one of their drives)
    m.bSkipShowSizes = True
    strDrive = GenesisCDInDrive(True)
#If 1 Then ' TLB 6/16/2015: removed the obsolete "starter dataset" option, and FTP is now always the default
    If Len(strDrive) > 0 Then
        drvFrom.Drive = strDrive
    End If
    optFTP.Value = True
#Else
    If Len(strDrive) > 0 Then
        optFrom.Value = True
        drvFrom.Drive = strDrive
        ' if we don't have the true enablement codes (i.e. still the default), try to get them now
        ' so ReadCfg can set good defaults for the Data History CD/DVD
        AskForActivate
    ElseIf Len(m.strFtpDataset) > 0 Then
        optFTP.Value = True
    Else
        optWeb.Value = True
    End If
#End If
    m.bSkipShowSizes = False

    ''m.strFromPath = "C:\DI\"
    If Not ReadCfg Then
'        InfBox "Could not find DataInst.cfg", "e", , "Install Data"
        cmdBrowse.Visible = True
    ElseIf FileExist("c:\common\files.exe") Then
        cmdBrowse.Visible = True
    End If

    ' if they have no data currently installed, just start by getting the starter set from the FTP server
    m.bAborted = False
    m.bInstall = False
    If optFTP Then
        If g.SymbolPool.NumRecords = 0 And Not m.bAddMode Then
            s = "We recommend downloading a small set of recent historical data from our FTP server| to start with (you can add more later)."
            If InfBox(s, "?", "+OK|-Cancel", "Install Data") = "O" Then
                If StartFtpDownloader Then
                    ' in this case, wait for it to finish ...
                    frmStatus.ShowDetails False
                    frmStatus.SetTitle "Installing Data"
                    ShowWaitMessage
                    ' after a few seconds, show the quick start info
                    ' (this allows getting the downloading started)
                    frmMain.tmrQuickStart.Interval = 1000
                    frmMain.tmrQuickStart.Enabled = True
                    ' and while we're waiting, this is a good time to get a refresh of enablements
                    ' (which are needed if this is a first-time startup)
                    GetNYTime ' just to connect to our servers
                    frmStatus.AddDetail "Downloading Files"
                    frmStatus.UpdateProgress "Downloading Files"
                    frmStatus.Status = eStatus_Running
                    Do While True
                        Sleep 1
                        s = ""
                        Select Case g.FtpDownloader.Status(s)
                        Case eGDDownloadStatus_Done
                            If InstallFiles Then
                                m.bInstall = True
                            End If
                            Exit Do
                        Case eGDDownloadStatus_Error
                            If Len(s) = 0 Then s = "Error downloading from the FTP server."
                            InfBox s, "e", , "FTP Download Error"
                            Exit Do
                        Case eGDDownloadStatus_Downloading
                            ' allow user to abort from within TradeNav
                            If frmStatus.Status = eStatus_Aborting Then
                                frmStatus.Status = eStatus_Aborted
                                Exit Do
                            End If
                        Case Else
                            Exit Do
                        End Select
                    Loop
                    If frmStatus.Status = eStatus_Aborted Then
                        ' if user hit Abort button in TradeNav, then just exit out of the install
                        bPromptForNextDisk = False
                        GoTo ErrExit
                    ElseIf frmStatus.Status <> eStatus_Error Then
                        frmStatus.Status = eStatus_Completed
                        frmStatus.AddDetail "Finished"
                    End If
                    If m.bInstall Then
                        ' now see if they want to add more data
                        ShowMe = True
                        m.bInstall = False
                        frmStatus.IsBusy = False
                        bInProgress = False
                        bPromptForNextDisk = True
                        GoTo DoNextDisk
                    End If
                End If
            End If
        ElseIf ReadyToLink Then
            ' if data had been downloaded, then link it in now
            frmStatus.IsBusy = True
            If InstallFiles Then
                bPromptForNextDisk = True
            End If
            frmStatus.Status = eStatus_Completed
            frmStatus.AddDetail "Finished"
            frmStatus.IsBusy = False
            bInProgress = False
            If m.bMoreDisks And bPromptForNextDisk Then
                GoTo DoNextDisk
            End If
            Exit Function
        End If
    End If

    KillFile m.strUnzipPath & "*.*", True

    ' allow user to select options and unzip the files
    ShowForm Me, eForm_Modal
    
    If m.bInstall And Not m.bAborted Then
        If optFTP.Value Then
            s = "After the historical data has finished downloading, it will automatically be linked into the program (but not while streaming is on)."
            'InfBox s, "i", , "Data Install"
            If Not StartFtpDownloader Then
                s = "Could not start the Historical Data Downloader."
                InfBox s, "e", , "Data Install"
            End If
        Else
            frmStatus.Status = eStatus_Initialized
            
            If optWeb.Value Then
                ' download the starter data set
                m.bAddMode = False
                If Not DownloadStarterData Then
                    m.bAborted = True
                End If
            End If
    
            If Not m.bAborted Then
                ' install the files
                If InstallFiles Then
                    ShowMe = True
                End If
            End If
            
            If frmStatus.Status <> eStatus_Aborted And frmStatus.Status <> eStatus_Error Then
                frmStatus.Status = eStatus_Completed
                frmStatus.AddDetail "Finished"
                If ShowMe = True And m.bMoreDisks = True Then
                    bPromptForNextDisk = True
                End If
            End If
        End If
    End If
    
    ' get rid of the temp folder
    If Not g.FtpDownloader.Status = eGDDownloadStatus_Downloading Then
        KillFolder m.strUnzipPath, True
    End If

ErrExit:
    frmStatus.IsBusy = False
    bInProgress = False
    
    If bPromptForNextDisk Then
        bPromptForNextDisk = False
        s = "If you have another data disk to install from, please put it in now.  Then select 'Next Disk' ..."
        If InfBox(s, "?", "+Next Disk|-Finished", "Install Next Disk?") = "N" Then
            GoTo DoNextDisk
        End If
    End If
    
    Unload Me
    Exit Function
    
ErrSection:
    frmStatus.IsBusy = False
    frmStatus.Status = eStatus_Error
    bInProgress = False
    RaiseError "frmDataInstall2.ShowMe"
    Resume ErrExit
End Function

Private Sub chkAll_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    m.bSkipShowSizes = True
    
    If chkAll(Index).Value = 1 Then
        chkSome(Index).Value = 1
        chkDaily(Index).Value = 1
    End If
    
    m.bSkipShowSizes = False
    ShowSizes

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.chkAll_Click"
    Resume ErrExit
End Sub

Private Sub chkDaily_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    m.bSkipShowSizes = True
    
    If chkDaily(Index).Value = 0 Then
        chkSome(Index).Value = 0
        chkAll(Index).Value = 0
    End If
    
    m.bSkipShowSizes = False
    ShowSizes

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.chkDaily_Click"
    Resume ErrExit
End Sub

Private Sub chkSome_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    m.bSkipShowSizes = True
    
    If chkSome(Index).Value = 0 Then
        chkAll(Index).Value = 0
    Else
        chkDaily(Index).Value = 1
    End If
    
    m.bSkipShowSizes = False
    ShowSizes

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.chkSome_Click"
    Resume ErrExit
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    Dim i&, s$
    
    m.bSkipShowSizes = True
    optFrom = True
    m.bSkipShowSizes = False
    
    s = CommonDialogFile(frmMain.CommonDialog1, False, "DataInst.CFG", "C:\", "Location of Data to Install")
    If Len(s) > 0 Then
        drvFrom.Visible = False
        lblFrom.Caption = LCase(Trim(FilePath(s)))
        lblFrom.Visible = True
    End If
        
    ReadCfg

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.cmdBrowse_Click"
    Resume ErrExit
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bAborted = True
    If InStr(UCase(cmdCancel.Caption), "ABORT") > 0 Then
        cmdCancel.Enabled = False
        GenZipAbort True
    Else
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.cmdCancel_Click"
    Resume ErrExit
End Sub

Private Sub cmdInstall_Click()
On Error GoTo ErrSection:

    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdInstall
    DoEvents

    ' make sure an option has been selected
    If optWeb.Value <> 0 And Len(optWebSet(0).Tag) > 0 Then
        If optWebSet(0).Value + optWebSet(1).Value + optWebSet(2).Value + optWebSet(3).Value = 0 Then
            InfBox "One of the downloading options must be selected.", "e", , "Error"
            Exit Sub
        End If
    End If

    m.bInstall = True
    m.bAborted = False
    cmdInstall.Enabled = False
    cmdInstall.Visible = False
    fraSpace.Visible = False
    cmdCancel.Caption = "&ABORT"
    fraHistory.Enabled = False ' so can't change options while installing
    fraPaths.Enabled = False
    
    'fraAmount.Enabled = False
    'fraMode.Enabled = False
    'fraPaths.Enabled = False
    vsProgress.Caption = ""
    vsProgress.Visible = True
    
    If optFTP.Value <> 0 Then
        Me.Hide
    ElseIf optWeb.Value <> 0 Then
        Me.Hide
    ElseIf UnzipFiles Then
        Me.Hide
    Else
        m.bAborted = True
        cmdCancel.Enabled = True
        cmdCancel.Caption = "Exit"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.cmdInstall_Click"
    Resume ErrExit
End Sub

Private Sub drvFrom_Change()
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    
    m.bSkipShowSizes = True
    optFrom.Value = True
    m.bSkipShowSizes = False
    
    ReadCfg

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.drvFrom_Change"
    Resume ErrExit
End Sub

#If 0 Then
Private Sub drvFrom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Dim s$
    s = Data.Files.Item(1)
    If Len(s) > 0 Then
        
    End If

End Sub
#End If

Private Sub Form_Activate()

    Dim s$
    Static bAlreadyDone As Boolean
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
        'DoEvents
        
    End If

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.Form_Load"
    Resume ErrExit
End Sub

Private Sub optFrom_Click()
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    
    ReadCfg

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.optFrom_Click"
    Resume ErrExit
End Sub

Private Sub optFrom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If Button = 2 Then cmdBrowse.Visible = True
    
End Sub

' History options are only visible for what's supported in the config file.
' History options are enabled only if:
' - file exists on the current CD/DVD
' - user is enabled for the required code
' - and is not "less" than what has already been installed
' (the .DMA file in each .GZP gets renamed to .DON when successfuly linked in)
Private Function ReadCfg() As Boolean
On Error GoTo ErrSection:

    Dim i&, iLine&, nSecType&, s$, dUnzipSize#
    Dim strFiles$, strDON$, strSecName$, strStarter$, strStarterSize$
    Dim bInvalidDisk As Boolean, bShowIntraday As Boolean, bDefaultOn As Boolean
    Dim aStrings As New cGdArray

    m.bSkipShowSizes = True
    m.bMoreDisks = False
    
    ' hide all controls and clear the tags and values
    bInvalidDisk = False
    m.bAddMode = False
    vsInvalid.Visible = False
    vsNotDisk1.Visible = False
    vsWeb.Visible = False
    fraWeb.Visible = False
    optFrom.Tag = ""
    lblNote2.Visible = False '(visible if at least one Daily is available to check)
    For i = chkDaily.LBound To chkDaily.UBound
        fraSecType(i).Visible = False
        chkDaily(i).Visible = False
        chkSome(i).Visible = False
        chkAll(i).Visible = False
        chkDaily(i).Tag = ""
        chkSome(i).Tag = ""
        chkAll(i).Tag = ""
        chkDaily(i).Value = False
        chkSome(i).Value = False
        chkAll(i).Value = False
        chkDaily(i).Enabled = False
        chkSome(i).Enabled = False
        chkAll(i).Enabled = False
    Next
    
    m.strFromPath = ""
    If optWeb Then
        ' TLB 12/12/2010: new options for downloading a starter data set
        'DatasetHdr=If you don't have ...
        'Dataset1=ID#   Name    Desc    Default
        'Dataset2=ID#   Name    Desc
        'Dataset3=ID#   Name    Desc
        'Dataset4=ID#   Name    Desc
        For i = 0 To 3
            s = GetProvidedProperty("Dataset" & Str(i + 1), "")
            If Len(s) = 0 Then
                optWebSet(i).Visible = False
                lblWebSet(i).Visible = False
                optWebSet(i).Tag = ""
                optWebSet(i).Value = False
            Else
                optWebSet(i).Visible = True
                lblWebSet(i).Visible = True
                optWebSet(i).Tag = Parse(s, vbTab, 1)
                optWebSet(i).Caption = Parse(s, vbTab, 2)
                lblWebSet(i).Caption = Replace(Parse(s, vbTab, 3), "|", vbCrLf)
                If Left(Parse(s, vbTab, 4), 1) = "x" Then
                    optWebSet(i).Value = True
                End If
                If i > 0 Then
                    optWebSet(i).Move optWebSet(0).Left, lblWebSet(i - 1).Top + lblWebSet(i - 1).Height + 150, optWebSet(0).Width, optWebSet(0).Height
                    lblWebSet(i).Move lblWebSet(0).Left, optWebSet(i).Top + (lblWebSet(0).Top - optWebSet(0).Top), lblWebSet(0).Width
                End If
            End If
        Next
        If Len(optWebSet(0).Tag) > 0 Then
            ' new method
            s = GetProvidedProperty("DatasetHdr", "")
            If Len(s) > 0 Then
                'lblWebHeader.Caption = s
            End If
            fraWeb.Move fraHistory.Left, fraHistory.Top
            fraWeb.Visible = True
        Else
            ' just use old method (no options given)
            vsWeb.Left = (Me.ScaleWidth - vsWeb.Width) / 2
            vsWeb.Visible = True
        End If
        fraHistory.Visible = False
    Else
        If optFTP Then
            m.strFromPath = AddSlash(App.Path) & "Downloader\CFG\"
        ElseIf lblFrom.Visible Then
            m.strFromPath = AddSlash(lblFrom.Caption)
        Else
            m.strFromPath = AddSlash(Parse(drvFrom.Drive, " ", 1)) & "DataInst\"
        End If

        ' read config file and check version# (first non-commented line of file)
        aStrings.FromFile m.strFromPath & "DataInst.cfg", , , "'"
        i = 0
        s = UCase(aStrings(0))
        If Left(s, 8) = "VERSION=" Then
            i = Val(Mid(s, 9))
        End If
        If i < 1 Or i > 1 Then '(min and max version allowed)
            ' incorrect version of data install
            vsInvalid.Left = (Me.ScaleWidth - vsInvalid.Width) / 2
            vsInvalid.Visible = True
            bInvalidDisk = True
        ElseIf optFTP Then
            ' determine if in "Add" mode (if this is the same .CFG file)
            s = UCase(Trim(FileToString(m.strFromPath & "DataInst.CFG")))
            If Len(s) > 0 Then
                If s = UCase(Trim(FileToString(DataPath & "DataInst.DON"))) Then
                    m.bAddMode = True
                End If
            End If
        Else
            ' determine if in "Add" mode (if this is the same .CFG file)
            i = FileLength(m.strFromPath & "DataInst.CFG")
            If i > 0 And i = FileLength(DataPath & "DataInst.DON") Then
                ' allow 2 hour difference in timestamp to accomodate potential timezone issues (e.g. Aussie vs. US)
                If Abs(FileDate(m.strFromPath & "DataInst.CFG") - FileDate(DataPath & "DataInst.DON")) < 0.1 Then
                    m.bAddMode = True
                End If
            End If
            ' if not adding data, make sure Disk #1 is in
            If Not m.bAddMode And (InStr(m.strFromPath, ":\") > 0) Then
                If Not FileExist(m.strFromPath & "..\Setup.exe") Then
                    vsNotDisk1.Left = (Me.ScaleWidth - vsNotDisk1.Width) / 2
                    vsNotDisk1.Visible = True
                    bInvalidDisk = True
                End If
            End If
        End If
        
        If bInvalidDisk Then
            fraHistory.Visible = False
        Else
            ReadCfg = True
            fraHistory.Visible = True

            ' read rest of file
            Screen.MousePointer = vbHourglass
            nSecType = -1
            For iLine = 1 To aStrings.Size - 1
                s = Trim(aStrings(iLine))
                ' see if a valid line
                If Len(s) > 0 And Left(s, 1) <> "'" Then
                    If Left(s, 2) = ">>" Then
                        nSecType = -999 ' done with security types
                    ElseIf Left(s, 1) = ">" Then
                        ' new security type
                        nSecType = nSecType + 1
                        If nSecType <= chkDaily.UBound Then
                            strSecName = Trim(Mid(s, 2))
                            lblSecType(nSecType).Caption = strSecName
                            lblSecType(nSecType).Visible = True
                            fraSecType(nSecType).Visible = True
                        End If
                    ElseIf nSecType >= chkDaily.LBound And nSecType <= chkDaily.UBound Then
                        'AmtType  ReqCode  Label    Zipfile1|Zipfile2|etc
                        'Some     ST       2 years  tck_s5.gzp|tck_s.gzp
                        strFiles = Parse(s, vbTab, 4)
                        strDON = DataPath & Parse(strFiles, ".", 1) & ".DON"
                        If m.bAddMode And FileExist(strDON) Then
                            dUnzipSize = 0
                        Else
                            dUnzipSize = GetUnzipSize(strFiles, Parse(s, vbTab, 6))
                            strDON = ""
                        End If
                        
                        ' set default according to the 5th field
                        bDefaultOn = False
                        Select Case UCase(Parse(s, vbTab, 5))
                        Case "T" ' True
                            bDefaultOn = True
                        Case "W" ' if Western Hemisphere (North/South America)
                            If IsInWesternHemisphere Then
                                bDefaultOn = True
                            End If
                        Case "E" ' if Eastern Hemisphere (Europe/Asia/Africa)
                            If Not IsInWesternHemisphere Then
                                bDefaultOn = True
                            End If
                        End Select
                        
                        ' set properties for this particular check box
                        Select Case UCase(Parse(s, vbTab, 1))
                        Case "DAILY"
                            chkDaily(nSecType).Tag = Str(dUnzipSize) & vbTab & strFiles
                            chkDaily(nSecType).Caption = Parse(s, vbTab, 3)
                            chkDaily(nSecType).Visible = True
                            If Len(strDON) > 0 Then
                                chkDaily(nSecType).Value = 1
                                chkDaily(nSecType).Enabled = False
                            ElseIf dUnzipSize > 0 And HasModule(Parse(s, vbTab, 2)) Then
                                lblNote2.Visible = True
                                chkDaily(nSecType).Enabled = True
                                If bDefaultOn Then
                                    chkDaily(nSecType).Value = 1
                                End If
                            End If
                        Case "SOME"
                            bShowIntraday = True
                            chkSome(nSecType).Tag = Str(dUnzipSize) & vbTab & strFiles
                            chkSome(nSecType).Caption = Parse(s, vbTab, 3)
                            chkSome(nSecType).Visible = True
                            If Len(strDON) > 0 Then
                                chkSome(nSecType).Value = 1
                                chkSome(nSecType).Enabled = False
                            ElseIf dUnzipSize > 0 And HasModule(Parse(s, vbTab, 2)) Then
                                chkSome(nSecType).Enabled = True
                                If bDefaultOn Then
                                    chkSome(nSecType).Value = 1
                                End If
                            End If
                        Case "ALL"
                            bShowIntraday = True
                            chkAll(nSecType).Tag = Str(dUnzipSize) & vbTab & strFiles
                            chkAll(nSecType).Caption = Parse(s, vbTab, 3)
                            chkAll(nSecType).Visible = True
                            If Len(strDON) > 0 Then
                                chkAll(nSecType).Value = 1
                                chkAll(nSecType).Enabled = False
                            ElseIf dUnzipSize > 0 And HasModule(Parse(s, vbTab, 2)) Then
                                chkAll(nSecType).Enabled = True
                                If bDefaultOn Then
                                    chkAll(nSecType).Value = 1
                                End If
                            Else
                                ' if an "All" is enabled and not checked, then there
                                ' may be more data disks to install after this
                                m.bMoreDisks = True
                            End If
                        End Select
                    Else
                        ' misc properties
                        Select Case UCase(Parse(s, "=", 1))
                        Case "STARTER"
                            strStarter = Parse(s, "=", 2)
                        Case "STARTERSIZE"
                            strStarterSize = Parse(s, "=", 2)
                        Case "X"
                            i = Val(Parse(s, "=", 2))
                        Case "NOTE"
                            lblNote.Caption = Parse(s, "=", 2)
                        Case "NOTE2"
                            lblNote2.Caption = Parse(s, "=", 2)
                        End Select
                    End If
                End If
            Next
            If Len(strStarter) > 0 Then
                If m.bAddMode Then
                    dUnzipSize = 0
                ElseIf Len(strStarterSize) > 0 Then
                    dUnzipSize = GetUnzipSize(strStarter, strStarterSize)
                Else
                    dUnzipSize = GetUnzipSize(strStarter, "300mb") ' default starter size
                End If
                optFrom.Tag = Str(dUnzipSize) & vbTab & strStarter
            End If
            Screen.MousePointer = vbNormal
        End If
    End If
    
    lblIntraday.Visible = bShowIntraday
    lineIntraday.Visible = bShowIntraday
    m.bSkipShowSizes = False
    ShowSizes

    ChangePath App.Path

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall2.ReadCfg"
    Resume ErrExit
End Function

Private Function GetUnzipSize(ByVal strFiles$, ByVal strSize$) As Double
On Error GoTo ErrSection:

    Dim i&, s$, dSize#, dTotalSize#
    Dim aFiles As New cGdArray
    
    If optFTP And Len(strSize) > 0 Then
        dTotalSize = ValOfText(strSize)
    Else
        aFiles.SplitFields strFiles, "|"
        For i = 0 To aFiles.Size - 1
            If Len(aFiles(i)) > 0 Then
                s = m.strFromPath & aFiles(i)
                If FileExist(s) Then
                    dSize = ZipExecute("S", s, "")
                    If dSize > 0 Then dTotalSize = dTotalSize + dSize
                End If
            End If
        Next
        Set aFiles = Nothing
    End If
    GetUnzipSize = dTotalSize

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall2.GetUnzipSize"
    Resume ErrExit
End Function

Private Sub ShowSizes()
On Error GoTo ErrSection:

    Dim i&, dSize#, d#
               
    ' don't run this routine if skip flag set (from in here or from other routines)
    If m.bSkipShowSizes Then Exit Sub
    m.bSkipShowSizes = True
               
    m.strZipFiles = ""
    dSize = 0
    If optWeb Then
        dSize = 150 * 1024# * 1024#
    Else
        If Not m.bAddMode Then
            ' starter file
            dSize = Val(Parse(optFrom.Tag, vbTab, 1))
            m.strZipFiles = Parse(optFrom.Tag, vbTab, 2) & "|"
        End If
        ' additional files (based on selected options)
        For i = chkDaily.LBound To chkDaily.UBound
            If chkAll(i).Value <> 0 Then
                d = Val(Parse(chkAll(i).Tag, vbTab, 1))
                If d > 0 Then
                    dSize = dSize + d
                    m.strZipFiles = m.strZipFiles & Parse(chkAll(i).Tag, vbTab, 2) & "|"
                End If
                chkSome(i).Value = 1
                chkDaily(i).Value = 1
            End If
            If chkSome(i).Value <> 0 Then
                d = Val(Parse(chkSome(i).Tag, vbTab, 1))
                If d > 0 Then
                    dSize = dSize + d
                    m.strZipFiles = m.strZipFiles & Parse(chkSome(i).Tag, vbTab, 2) & "|"
                End If
                chkDaily(i).Value = 1
            End If
            If chkDaily(i).Value <> 0 Then
                d = Val(Parse(chkDaily(i).Tag, vbTab, 1))
                If d > 0 Then
                    dSize = dSize + d
                    m.strZipFiles = m.strZipFiles & Parse(chkDaily(i).Tag, vbTab, 2) & "|"
                End If
            End If
        Next
    End If
    
    If dSize > 0 Then
        cmdInstall.Enabled = True
    Else
        cmdInstall.Enabled = False
    End If
    lblSpaceRequired.Caption = Format(dSize / 1024# / 1024#, "#,##0") & " MB"
    lblSpaceRequired.Tag = Str(dSize)
    d = dSize
    
    On Error Resume Next
    dSize = -1
    'dSize = fs.Drives(Left(App.Path, 2)).AvailableSpace
    dSize = GetDiskFreeSpace(Left(App.Path, 2))
    If dSize < 0 Then
        lblSpaceAvailable.Caption = "(unknown)"
        lblSpaceRequired.ForeColor = lblSpaceAvailable.ForeColor
    Else
        lblSpaceAvailable.Caption = Format(dSize / 1024# / 1024#, "#,##0") & " MB"
        If d > dSize Then
            lblSpaceRequired.ForeColor = vbRed
        ElseIf d > dSize - 1000000000 Then
            lblSpaceRequired.ForeColor = &HC000C0
        Else
            lblSpaceRequired.ForeColor = lblSpaceAvailable.ForeColor
        End If
    End If
    
    m.bSkipShowSizes = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.ShowSizes"
    Resume ErrExit
End Sub

Private Sub optFTP_Click()
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    
    ReadCfg

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.optFTP_Click"
    Resume ErrExit
End Sub

Private Sub optWeb_Click()
On Error GoTo ErrSection:

    If m.bSkipShowSizes Then Exit Sub
    
    ReadCfg

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDataInstall2.optWeb_Click"
    Resume ErrExit
End Sub

Private Sub txtHwnd_Change()

    Dim iPercent&, dBytes#
    Static dPrevDisplay As Double
    
    On Error Resume Next ' (since this is from a DLL callback)
    
    ' calculate percent of entire process (all files together)
    iPercent = Val(StripStr(txtHwnd.Text, "%"))
    dBytes = m.dZipFileSize * iPercent / 100# + m.dBytesSoFar
    iPercent = Int(100# * dBytes / m.dTotalBytes)
    
    ' display whenever percent changes
    ' (but also allow for capturing Abort click at least once every second)
    If iPercent <> vsProgress.FloodPercent Or gdTickCount > dPrevDisplay + 1000 Then
        'vsProgress.Caption = iPercent & "%  (" & Format(dBytes / 1024# / 1024#, "#0") & " MB)"
        vsProgress.Caption = Format(dBytes / 1024# / 1024#, "#0") & " MB  =  " & Str(iPercent) & "%"
        vsProgress.FloodPercent = iPercent
        DoEvents
        'vsProgress.Refresh
        dPrevDisplay = gdTickCount
    End If

End Sub

Private Function UnzipFiles() As Boolean
On Error GoTo ErrSection:

    Dim i&, rc&, d#, strErr$
    Dim aFiles As New cGdArray
       
    KillFile m.strUnzipPath & "*.*", True
         
    ' copy DataInst.CFG over as DataInst.DON
    If Not m.bAddMode Then
        FileCopy m.strFromPath & "DataInst.CFG", m.strUnzipPath & "DataInst.DON", True
    End If
       
    vsProgress.Caption = "0%"
    DoEvents
       
'If IsIDE Then m.strFromPath = "d:\"
       
    aFiles.SplitFields m.strZipFiles, "|"
    m.dTotalBytes = 0
    For i = aFiles.Size - 1 To 0 Step -1
        If Len(Trim(aFiles(i))) = 0 Then
            aFiles.Remove i
        Else
            d = ZipExecute("S", m.strFromPath & aFiles(i), "")
            m.dTotalBytes = m.dTotalBytes + d
            aFiles(i) = m.strFromPath & aFiles(i) & vbTab & Str(d)
        End If
    Next
    
'If IsIDE Then Me.Caption = "Total Bytes = " & Str(m.dTotalBytes)
   
    strErr = ""
    m.dBytesSoFar = 0
    For i = 0 To aFiles.Size - 1
        m.dZipFileSize = Val(Parse(aFiles(i), vbTab, 2))
        strErr = Space(250)
        rc = ZipExecute("U", Parse(aFiles(i), vbTab, 1), m.strUnzipPath, "", , True, , , , txtHwnd.hWnd, strErr)
        FixNullTermStr strErr
        strErr = Trim(strErr)
        DoEvents
        If Len(strErr) > 0 Or m.bAborted Then
            Exit For
        End If
        m.dBytesSoFar = m.dBytesSoFar + m.dZipFileSize
    Next
    
    If m.bAborted Then
        vsProgress.Caption = "ABORTED"
    ElseIf Len(strErr) > 0 Then
        vsProgress.Caption = "ERROR Unzipping Files"
        m.bAborted = True
        vsProgress.ToolTipText = strErr
    Else
        vsProgress.FloodPercent = 100
        cmdCancel.Enabled = False
        UnzipFiles = True
    End If
    
ErrExit:
    Set aFiles = Nothing
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall2.UnzipFiles"
    Resume ErrExit
End Function

Private Function StartFtpDownloader() As Boolean
On Error GoTo ErrSection:

    Dim i&, d#, s$, strPath$, strZipPath$, strServerPath$, strFile$
    Dim aFiles As New cGdArray
    Dim FtpFile As cDownloaderFile
       
    strServerPath = ".\DATASETS\" & m.strFtpDataset
    strPath = App.Path & "\Downloader\"
    strZipPath = strPath & "Zipped"
    MakeDir strZipPath
    
    KillFile strZipPath & "\*.*", True
    KillFile m.strUnzipPath & "*.*", True
    
    ' init the Downloader object (from first line of the Dataset.TXT file)
    s = FileToString(strPath & "CFG\Datasets.txt", , True)
    If Len(s) = 0 Then Exit Function
    g.FtpDownloader.UserName = Parse(s, vbTab, 1)
    g.FtpDownloader.Password = Parse(s, vbTab, 2)
    g.FtpDownloader.IP = Parse(s, vbTab, 3)
    g.FtpDownloader.Port = Val(Parse(s, vbTab, 4))
    If g.SymbolPool.NumRecords = 0 Then
        g.FtpDownloader.Note = "NOTE: after the historical data has finished downloading, it will automatically be linked into the program."
    Else
        g.FtpDownloader.Note = "NOTE: after the historical data has finished downloading, it will automatically be linked into the program (but only when data streaming is turned off)."
    End If
        
    ' add files
    aFiles.SplitFields m.strZipFiles, "|"
    g.FtpDownloader.Files.Clear
    For i = 0 To aFiles.Size - 1
        strFile = UCase(Trim(aFiles(i)))
        If Len(strFile) > 0 Then
            Set FtpFile = New cDownloaderFile
            FtpFile.ServerPath = strServerPath
            FtpFile.ServerFilename = strFile
            FtpFile.LocalFilename = strFile
            Select Case UCase(FileExt(strFile))
            Case "ZIP", "GZP"
                FtpFile.IsZipFile = True
                FtpFile.LocalPath = strZipPath
            Case Else
                FtpFile.IsZipFile = False
                FtpFile.LocalPath = m.strUnzipPath
            End Select
            FtpFile.ZipPath = m.strUnzipPath
            g.FtpDownloader.Files.Add FtpFile
        End If
    Next

    ' start the Downloader
    If HasDotNet Then
        If g.FtpDownloader.Download(strPath & "Request.TXT") Then
            ' copy DataInst.CFG over as DataInst.DON
            If Not m.bAddMode Then
                FileCopy m.strFromPath & "DataInst.CFG", m.strUnzipPath & "DataInst.DON", True
            End If
            StartFtpDownloader = True
        End If
    End If
       
ErrExit:
    Set aFiles = Nothing
    Exit Function
    
ErrSection:
    RaiseError "frmDataInstall2.StartFtpDownloader"
    Resume ErrExit
End Function

Private Function DownloadStarterData() As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim astrRequest As New cGdArray     ' Array of requests
    Dim bSuccess As Boolean             ' Success from the download
    Dim lValid As Long                  ' Return from Unzip command
    Dim strButtons As String            ' Buttons to display to the user
  
'    If g.SymbolPool.NumRecords = 0 Then
'        strButtons = "+OK|-Exit"
'    Else
'        strButtons = "+OK|-Cancel"
'    End If
    
'Do While InfBox("We will need to connect to Genesis and download your starting data set.|(This will require an internet connection)", "i", strButtons, "Starting Data Set") = "O"
    
    KillFile App.Path & "\Ftp\*.*", True
    
    frmStatus.SetTitle "Downloading Data Set"
    frmStatus.UpdateProgress "Requesting Data Set"
    
    astrRequest.Size = 0
    For i = 0 To 3
        If optWebSet(i).Value <> 0 Then
            astrRequest.Add "/Dataset:" & optWebSet(i).Tag
            Exit For
        End If
    Next
    astrRequest.Add "%Download Data Set"
    If App.Major >= 4 Then
        astrRequest.Add "+FULLTICK:TRUE"
    Else
        astrRequest.Add "+FULLTICK:FALSE"
    End If
    If frmStatus.Status < eStatus_Aborting Then
        ' after 10 seconds, show the quick start info
        ' (this allows getting the downloading started)
        frmMain.tmrQuickStart.Interval = 10000
        frmMain.tmrQuickStart.Enabled = True
        bSuccess = FtpRequest(astrRequest, , , True)
        frmMain.tmrQuickStart.Enabled = False
    End If
    
    If frmStatus.Status = eStatus_Aborted Then
        '(nothing more to do)
    ElseIf frmStatus.Status = eStatus_Error Or bSuccess = False Then
        frmStatus.AddDetail "ERROR downloading data set"
        frmStatus.Status = eStatus_Error
    ElseIf frmStatus.Status = eStatus_Completed Then
        ' backup file, or use the backup (if didn't need to be redownloaded)
        If FileExist(App.Path & "\FTP\DataSet.GZP") Then
            FileCopy App.Path & "\FTP\DataSet.GZP", App.Path & "\FTP\Backup\", True
        ElseIf FileExist(App.Path & "\FTP\Backup\DataSet.GZP") Then
            FileCopy App.Path & "\FTP\Backup\DataSet.GZP", App.Path & "\FTP\", True
        End If
            
        If FileExist(App.Path & "\FTP\DataSet.GZP") Then
            lValid = ZipExecute("U", App.Path & "\FTP\DataSet.GZP", m.strUnzipPath, "", False, False)
        End If
        
        If lValid > 0 Then
            DownloadStarterData = True
            ''InfBox "Please wait while Trade Navigator|sets up the data", , , , True
        Else
            frmStatus.AddDetail "Data set not found"
            frmStatus.Status = eStatus_Error
        End If
    End If
        
    'If frmStatus.Status = eStatus_Completed Then Exit Do
'Loop

ErrExit:
    Exit Function
    
ErrSection:
    frmStatus.IsBusy = False
    frmStatus.Status = eStatus_Error
    RaiseError "frmDataInstall2.DownloadStarterData"
End Function

' Move files into the Data folder, "Install" them, and reload all data/forms
Private Function InstallFiles() As Boolean
On Error GoTo ErrSection:

    Dim i&, strMsg$
    Dim aFiles As New cGdArray
    Dim frmActive As Form
       
    ChangePath App.Path
    KillFile g.FtpDownloader.DoneFile
       
    frmStatus.ShowDetails False
    frmStatus.SetTitle "Installing Data"
    frmStatus.AddDetail "Installing Files"
    frmStatus.UpdateProgress "Installing Files"
    
    If g.SymbolPool.NumRecords > 0 Or m.bAddMode Then
        ShowWaitMessage
    End If
    
    If Not m.bAddMode Then
        ' close down the data manager before replacing the files
        DM_Init False
        g.Universe.CloseDb
        DoEvents
        
        ' delete all files
        CleanOutDataFolder
    End If
        
    ' move the new files from the temp folder to the Data folder
    aFiles.GetMatchingFiles m.strUnzipPath & "*.*", False
    For i = 0 To aFiles.Size - 1
        KillFile DataPath & aFiles(i)
        Name m.strUnzipPath & aFiles(i) As DataPath & aFiles(i)
    Next
        
    If Not m.bAddMode Then
        ' now re-open the data manager with the replaced files
        DM_Close g.DMS
        DM_Init True
        g.Universe.OpenDb
    End If
        
    ' and link the new files in
    If Not UpdateDBConfig Then
        ' error!
        frmStatus.AddDetail "ERROR: Linking files"
    End If
       
    ' reload various data
    frmStatus.AddDetail "Reloading Symbols"
    frmStatus.UpdateProgress "Reloading Symbols"
    g.SymbolPool.Load False
    
    frmStatus.AddDetail "Reloading Data"
    frmStatus.UpdateProgress "Reloading Data"
    frmStatus.Status = eStatus_Running
        
    If Not m.bAddMode Then
        frmSymbolGrid.InitForm
        If (DockState(frmSymbolGrid) = eHidden) And ExtremeCharts >= 1 Then
            DockState(frmSymbolGrid) = eShowAsPrevious
        End If
    End If
        
    Set frmActive = ActiveChart
    If frmActive Is Nothing Then
        Set frmActive = New frmChart
        frmActive.Chart.SetSymbol g.SymbolPool.SymbolIDforSymbol("$DJIA")
        frmActive.WindowState = 2
        frmActive.Show
    End If
    Set frmActive = Nothing
    
    frmQuotes.LoadTable
    frmSymbolGrid.RefreshGrid
    frmSymbolGrid.ShowInitialSymbol

    g.RealTime.RefreshAllFormData True
    
    InstallFiles = True
        
ErrExit:
    InfBox ""
    Set aFiles = Nothing
    Exit Function
    
ErrSection:
    frmStatus.IsBusy = False
    frmStatus.Status = eStatus_Error
    RaiseError "frmDataInstall2.InstallFiles"
    Resume ErrExit
End Function


Private Function GetFtpDataset(Optional ByVal bRefresh As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim i&, s$, strPath$, strListFile$, strUrl$, strDataset$, strInstalled$
    Dim aList As New cGdArray
    Static nPrevLDD&
        
    m.strFtpDataset = ""
        
    ' this is invalid if Downloader program is not there
    If Not FileExist(App.Path & "\Downloader\Downloader.*") Then Exit Function
        
    ' get contents of installed CFG file (to see if any matches)
    strInstalled = UCase(Trim(FileToString(DataPath & "DataInst.DON")))
        
    ' get the list of available FTP Datasets from our web server
    strPath = App.Path & "\Downloader\CFG\"
    strListFile = strPath & "DataSets.txt"
    MakeDir strPath
    If LastDailyDownload <> nPrevLDD Or nPrevLDD = 0 Or bRefresh Then
        KillFile strListFile
    End If
    If Not FileExist(strListFile) Then
        bRefresh = True
        nPrevLDD = LastDailyDownload
        KillFile strPath & "*.*", True
        strUrl = FixURL("www.TradeNavigator.com/DataInst/DataSets.txt")
        s = GetWebPageData(strUrl)
        ' make sure we got a text file and not some goofy default HTML file
        If Len(s) > 0 And InStr(s, "//") = 0 Then
            FileFromString strListFile, s
        End If
    End If
    
    ' and get the FTP DataInst.CFG files from our web server
    aList.FromFile strListFile
    m.strFtpLogin = aList(0) ' save off the FTP login info (first line of file)
    aList.Remove 0, 1
    aList.Sort eGdSort_Descending ' so newest dataset is first
    For i = 0 To aList.Size - 1
        strDataset = Trim(aList(i))
        ' if dataset ends in asterisk, it has not yet been "released"
        If Right(strDataset, 1) = "*" Then
            ' only use this dataset if we are doing the testing
            If FileExist("c:\TestDataset.flg") Then
                strDataset = Trim(Left(strDataset, Len(strDataset) - 1))
            Else
                strDataset = ""
            End If
        End If
        If Len(strDataset) > 0 Then
            If Not FileExist(strPath & strDataset & ".CFG") Or bRefresh Then
                strUrl = FixURL("www.TradeNavigator.com/DataInst/" & strDataset & ".TXT")
                s = GetWebPageData(strUrl)
                ' make sure we got a text file and not some goofy default HTML file
                If Len(s) > 0 And InStr(s, "//") = 0 Then
                    FileFromString strPath & strDataset & ".CFG", s
                End If
            End If
            If FileExist(strPath & strDataset & ".CFG") Then
                ' default to the newest dataset
                If Len(m.strFtpDataset) = 0 Then
                    m.strFtpDataset = strDataset
                End If
                If Len(strInstalled) = 0 Then
                    Exit For ' no need to keep checking the rest of the datasets
                Else
                    ' but see if this CFG file matches with what's already been installed
                    s = UCase(Trim(FileToString(strPath & strDataset & ".CFG")))
                    If s = strInstalled Then
                        m.strFtpDataset = strDataset
                    End If
                End If
            End If
        End If
    Next
    
    ' copy to DataInst.CFG
    KillFile strPath & "DataInst.CFG", True
    If Len(m.strFtpDataset) > 0 Then
        FileCopy strPath & m.strFtpDataset & ".CFG", strPath & "DataInst.CFG"
        GetFtpDataset = True
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmDataInstall2.GetFtpDataset"
    Resume ErrExit
End Function

Public Function ReadyToLink() As Boolean
On Error GoTo ErrSection:

    Dim s$
    
    s = g.FtpDownloader.DoneFile
    If FileExist(s) Then
        s = UCase(Trim(FileToString(s, , True)))
        If s = "DONE" Then
            s = App.Path & kTempFolder & "*.dma"
            If FileExist(s) Then
                ReadyToLink = True
            End If
        End If
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmDataInstall2.ReadyToLink"
    Resume ErrExit
End Function

Private Sub ShowWaitMessage()

    On Error Resume Next
    Dim strMsg$
    
    If ExtremeCharts >= 1 Then
        strMsg = "Please wait while Extreme Charts|reloads from the new data files"
    ElseIf m.bMoreDisks And optFTP.Value = 0 Then
        strMsg = "Please wait while Trade Navigator|reloads from the new data files.||If there is another data disk to install,|you may insert it into the drive now."
    Else
        strMsg = "Please wait while Trade Navigator|reloads from the new data files"
    End If
    InfBox strMsg, , , , True

End Sub

