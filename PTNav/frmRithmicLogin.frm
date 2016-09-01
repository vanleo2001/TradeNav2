VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmRithmicLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraRithmic 
      Height          =   615
      Left            =   960
      TabIndex        =   13
      Top             =   2400
      Width           =   2055
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
      Caption         =   "frmRithmicLogin.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmRithmicLogin.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRithmicLogin.frx":004C
      RightToLeft     =   0   'False
      Begin VB.PictureBox picPbo 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   60
         Picture         =   "frmRithmicLogin.frx":0068
         ScaleHeight     =   255
         ScaleWidth      =   1875
         TabIndex        =   1
         Top             =   420
         Width           =   1875
      End
      Begin VB.PictureBox picRithmic 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         Picture         =   "frmRithmicLogin.frx":0576
         ScaleHeight     =   315
         ScaleWidth      =   2055
         TabIndex        =   14
         Top             =   0
         Width           =   2055
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraCopyright 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   7515
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
      Caption         =   "frmRithmicLogin.frx":0805
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmRithmicLogin.frx":0831
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRithmicLogin.frx":0851
      RightToLeft     =   0   'False
      Begin VB.PictureBox picPboSmall 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   0
         Picture         =   "frmRithmicLogin.frx":086D
         ScaleHeight     =   225
         ScaleWidth      =   1830
         TabIndex        =   7
         Top             =   840
         Width           =   1830
      End
      Begin HexUniControls.ctlUniLabelXP lblOmne 
         Height          =   195
         Left            =   0
         Top             =   480
         Width           =   6795
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
         Caption         =   "frmRithmicLogin.frx":0D7B
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":0E71
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":0E91
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPbo 
         Height          =   435
         Left            =   1980
         Top             =   720
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
         Caption         =   "frmRithmicLogin.frx":0EAD
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":0F73
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":0F93
         RightToLeft     =   0   'False
         WordWrap        =   -1  'True
      End
      Begin HexUniControls.ctlUniLabelXP lblRithmic 
         Height          =   195
         Left            =   0
         Top             =   240
         Width           =   6795
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
         Caption         =   "frmRithmicLogin.frx":0FAF
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":107D
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":109D
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRapi 
         Height          =   195
         Left            =   0
         Top             =   0
         Width           =   6795
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
         Caption         =   "frmRithmicLogin.frx":10B9
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":1179
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":1199
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   915
      Left            =   120
      TabIndex        =   9
      Top             =   1320
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
      Caption         =   "frmRithmicLogin.frx":11B5
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmRithmicLogin.frx":11E1
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRithmicLogin.frx":1201
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdLogin 
         Default         =   -1  'True
         Height          =   435
         Left            =   780
         TabIndex        =   11
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
         Caption         =   "frmRithmicLogin.frx":121D
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmRithmicLogin.frx":1249
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":1269
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1860
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
         Caption         =   "frmRithmicLogin.frx":1285
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmRithmicLogin.frx":12B3
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":12D3
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
         Caption         =   "frmRithmicLogin.frx":12EF
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":13AF
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":13CF
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraLoginInfo 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   180
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
      Caption         =   "frmRithmicLogin.frx":13EB
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmRithmicLogin.frx":1417
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRithmicLogin.frx":1437
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboUserIds 
         Height          =   315
         Left            =   900
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
         Tip             =   "frmRithmicLogin.frx":1453
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":1473
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddUserId 
         Height          =   315
         Left            =   2880
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
         Caption         =   "frmRithmicLogin.frx":148F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmRithmicLogin.frx":14C7
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":14E7
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveUserId 
         Height          =   315
         Left            =   3300
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
         Caption         =   "frmRithmicLogin.frx":1503
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmRithmicLogin.frx":1541
         Style           =   1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":1561
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboSystems 
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Top             =   720
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
         Tip             =   "frmRithmicLogin.frx":157D
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":159D
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmRithmicLogin.frx":15B9
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
         Tip             =   "frmRithmicLogin.frx":15D9
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":15F9
      End
      Begin HexUniControls.ctlUniLabelXP lblSystem 
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
         Caption         =   "frmRithmicLogin.frx":1615
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":1645
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":1665
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
         Caption         =   "frmRithmicLogin.frx":1681
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":16B5
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":16D5
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblUserID 
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
         Caption         =   "frmRithmicLogin.frx":16F1
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmRithmicLogin.frx":1723
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmRithmicLogin.frx":1743
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtfDisclaimer 
      Height          =   3075
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5424
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmRithmicLogin.frx":175F
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
      Tip             =   "frmRithmicLogin.frx":177F
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmRithmicLogin.frx":179F
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
Attribute VB_Name = "frmRithmicLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmRithmicLogin.frm
'' Description: Form to assist user in logging into the Rithmic servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/24/2010   DAJ         Fixed some of the Rithmic artwork
'' 10/05/2010   DAJ         Changed the Rithmic image
'' 10/26/2010   DAJ         Changed the Rithmic image
'' 10/27/2010   DAJ         More mods to the Rithmic image
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 12/10/2010   DAJ         Added Zen-Fire
'' 03/07/2011   DAJ         Moved Rithmic to Broker Base Class
'' 07/18/2012   DAJ         Added Alpari (Zen-Fire), Zaner (Rithmic), and Zaner (Zen-Fire)
'' 07/24/2012   DAJ         Changed the Rithmic login stuff
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK or Cancel?
    Broker As cBroker                   ' Broker object
    strIniFile As String                ' INI file for the broker
    strConnectIniFile As String         ' INI file for the broker connection information
    
    strUserID As String                 ' User ID for the login
    strPassword As String               ' Password for the login
    lSystemID As Long                   ' ID for the system chosen
End Type
Private m As mPrivate

Public Property Get UserID() As String
    UserID = m.strUserID
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property

Public Property Get SystemID() As Long
    SystemID = m.lSystemID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Broker, UserID, Are we switching?
'' Returns:     True if login, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Broker As cBroker, Optional ByVal strUserID As String = "", Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim strLastUserID As String         ' Last User ID logged into
    
    m.bOK = False
    
    Set m.Broker = Broker
    Caption = "Login to " & Broker.BrokerName
    m.strIniFile = Broker.IniFile
    m.strConnectIniFile = Broker.ConnectIni
    
    LoadUserIdsCombo
    LoadSystemsCombo

    strLastUserID = GetIniFileProperty("LastUserID", "", "User", m.strIniFile)
    If cboUserIds.ListCount > 0 Then
        If SetUserIdsCombo(strUserID) = False Then
            If SetUserIdsCombo(strLastUserID) = False Then
                cboUserIds.ListIndex = 0
            End If
        End If
    End If
    If (cboUserIds.ListCount = 0) Then
        AddUserId
    End If

    MoveFocus txtPassword
    
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        m.strUserID = Trim(cboUserIds.Text)
        m.strPassword = Trim(txtPassword.Text)
        m.lSystemID = cboSystems.ItemData(cboSystems.ListIndex)
        
        SetIniFileProperty "LastUserID", m.strUserID, "User", m.strIniFile
        UpdateCurrentUserId
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmRithmicLogin.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddUserId_Click
'' Description: Allow the user to add a user id
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddUserId_Click()
On Error GoTo ErrSection:

    AddUserId

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.cmdAddUserId_Click"

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
    RaiseError "frmRithmicLogin.cmdCancel_Click"
    
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

    If cboUserIds.ListIndex < 0 Then
        MoveFocus cboUserIds
        InfBox "Please select a User ID", "!", , "Login Error"
    ElseIf Len(Trim(txtPassword.Text)) = 0 Then
        MoveFocus txtPassword
        InfBox "Please enter in a password", "!", , "Login Error"
    Else
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.cmdLogin_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveUserId_Click
'' Description: Allow the user to remove user id(s)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveUserId_Click()
On Error GoTo ErrSection:

    RemoveUserId

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.cmdRemoveUserId_Click"

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
    RaiseError "frmRithmicLogin.Form_Activate"
    
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
    
    g.Styler.StyleForm Me
    
    rtfDisclaimer.Locked = True
    rtfDisclaimer.Font.Name = "Arial"
    rtfDisclaimer.BackColor = vbButtonFace
    rtfDisclaimer.TextRTF = FileToString(AddSlash(App.Path) & "Info\BrokerDisclaimer.rtf")
    
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.Form_Load"
    
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
    RaiseError "frmRithmicLogin.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lHorzSpace As Long              ' Horizontal space between controls
    Dim lVertSpace As Long              ' Vertical space between controls
    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height

    lHorzSpace = 120
    lVertSpace = 120
    lMinScaleWidth = fraButtons.Width + rtfDisclaimer.Width + (lHorzSpace * 3)
    lMinScaleHeight = picPbo.Height + picRithmic.Height + fraLoginInfo.Height + fraButtons.Height + fraCopyright.Height + (lVertSpace * 4)

    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With fraCopyright
            .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
        End With
        
        With rtfDisclaimer
            .Move ScaleWidth - .Width - lHorzSpace, lVertSpace, , ScaleHeight - fraCopyright.Height - (lVertSpace * 3)
        End With
        
        With fraButtons
            .Move lHorzSpace
        End With
        
        With fraLoginInfo
            .Move lHorzSpace
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUserIdsCombo
'' Description: Load the user ids combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUserIdsCombo()
On Error GoTo ErrSection:

    Dim strUserIds As String            ' User ID information string from INI file
    Dim astrUserIds As New cGdArray     ' Array of user id information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLastUserID As String         ' Last user id attempted
    
    cboUserIds.Clear
    
    strUserIds = GetIniFileProperty("UserIds", "", "User", m.strIniFile)
    If Len(strUserIds) = 0 Then
        strLastUserID = GetIniFileProperty("LastUserID", "", "User", m.strIniFile)
        If Len(strLastUserID) > 0 Then
            strUserIds = strLastUserID & "|" & GetIniFileProperty("Default", "1", "Environments", m.strConnectIniFile)
            SetIniFileProperty "UserIds", strUserIds, "User", m.strIniFile
        End If
    End If
    
    If Len(strUserIds) > 0 Then
        astrUserIds.SplitFields strUserIds, ","
        For lIndex = 0 To astrUserIds.Size - 1
            If Len(astrUserIds(lIndex)) > 0 Then
                cboUserIds.AddItem Parse(astrUserIds(lIndex), "|", 1)
                cboUserIds.ItemData(cboUserIds.NewIndex) = CLng(Val(Parse(astrUserIds(lIndex), "|", 2)))
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.LoadUserIdsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUserIdsCombo
'' Description: Set the user ids combo box to the given user id if possible
'' Inputs:      User ID
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetUserIdsCombo(ByVal strUserID As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboUserIds.ListCount > 0) And (Len(strUserID) > 0) Then
        For lIndex = 0 To cboUserIds.ListCount - 1
            If UCase(cboUserIds.List(lIndex)) = UCase(strUserID) Then
                bFound = True
                cboUserIds.ListIndex = lIndex
                
                SetSystemsCombo cboUserIds.ItemData(lIndex)
            End If
        Next lIndex
    End If
    
    SetUserIdsCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRithmicLogin.SetUserIdsCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSystemsCombo
'' Description: Load up the systems combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSystemsCombo()
On Error GoTo ErrSection:

    Dim lNumSystems As Long             ' Number of systems
    Dim lIndex As Long                  ' Index into a for loop
    
    cboSystems.Clear
    
    lNumSystems = GetIniFileProperty("NumEnvironments", 0&, "Environments", m.strConnectIniFile)
    For lIndex = 1 To lNumSystems
        cboSystems.AddItem GetIniFileProperty("Name", "", "Environment" & Str(lIndex), m.strConnectIniFile)
        cboSystems.ItemData(cboSystems.NewIndex) = lIndex
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.LoadSystemsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetSystemsCombo
'' Description: Set the systems combo box to the given system if possible
'' Inputs:      System
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetSystemsCombo(ByVal lSystem As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the item found?
    
    bFound = False
    If (cboSystems.ListCount > 0) And (lSystem >= 0) Then
        For lIndex = 0 To cboSystems.ListCount - 1
            If cboSystems.ItemData(lIndex) = lSystem Then
                bFound = True
                cboSystems.ListIndex = lIndex
            End If
        Next lIndex
    End If
    
    SetSystemsCombo = bFound

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmRithmicLogin.SetSystemsCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddUserId
'' Description: Allow the user to give us a new user id
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddUserId()
On Error GoTo ErrSection:

    Dim strUserID As String             ' User id from the user
    Dim strNewUserID As String          ' New user id to save to INI file
    Dim strUserIds As String            ' User ids string from the INI file
    Dim lSystem As Long                 ' System for the user id
    
    strUserID = InfBox("What is your " & m.Broker.BrokerName & " user id?", "?", , m.Broker.BrokerName & " User ID", , , , , , "string")
    If Len(strUserID) > 0 Then
        If SetUserIdsCombo(strUserID) = False Then
            lSystem = GetIniFileProperty("Default", 1&, "Environments", m.strConnectIniFile)
            
            cboUserIds.AddItem strUserID
            cboUserIds.ItemData(cboUserIds.NewIndex) = lSystem
            
            SetUserIdsCombo strUserID
            MoveFocus txtPassword
            
            strUserIds = GetIniFileProperty("UserIds", "", "User", m.strIniFile)
            If Len(strUserIds) = 0 Then
                strUserIds = strUserID & "|" & Str(lSystem)
            Else
                strUserIds = strUserIds & "," & strUserID & "|" & Str(lSystem)
            End If
            SetIniFileProperty "UserIds", strUserIds, "User", m.strIniFile
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.AddUserId"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveUserId
'' Description: Allow the user to remove user id information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveUserId()
On Error GoTo ErrSection:

    Dim strUserIds As String            ' User id information string from INI file
    Dim astrUserIds As cGdArray         ' Array of user id information
    Dim astrList As cGdArray            ' List to send to the delete form
    Dim astrToDelete As cGdArray        ' List of user ids to delete
    Dim strSelected As String           ' Currently selected user id
    Dim lIndex As Long                  ' Index into a for loop

    strUserIds = GetIniFileProperty("UserIds", "", "User", m.strIniFile)
    If Len(strUserIds) > 0 Then
        Set astrUserIds = New cGdArray
        Set astrList = New cGdArray
        astrList.Create eGDARRAY_Strings
        
        astrUserIds.SplitFields strUserIds, ","
        For lIndex = 0 To astrUserIds.Size - 1
            astrList.Add Parse(astrUserIds(lIndex), "|", 1) & vbTab & Str(lIndex)
        Next lIndex
        
        strSelected = cboUserIds.Text
        
        Set astrToDelete = frmDelete.ShowMe(astrList, strSelected)
        If Not astrToDelete Is Nothing Then
            For lIndex = astrToDelete.Size - 1 To 0 Step -1
                astrUserIds.Remove CLng(Val(astrToDelete(lIndex)))
            Next lIndex
            
            SetIniFileProperty "UserIds", astrUserIds.JoinFields(","), "User", m.strIniFile
            
            LoadUserIdsCombo
            If SetUserIdsCombo(strSelected) = False Then
                If cboUserIds.ListCount > 0 Then
                    cboUserIds.ListIndex = 0
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.RemoveUserId"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateCurrentUserId
'' Description: Update the current user id in the user ids list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateCurrentUserId()
On Error GoTo ErrSection:

    Dim strUserIds As String            ' User id information string from INI file
    Dim astrUserIds As cGdArray         ' Array of user id information
    Dim lIndex As Long                  ' Index into a for loop
    
    strUserIds = GetIniFileProperty("UserIds", "", "User", m.strIniFile)
    
    Set astrUserIds = New cGdArray
    astrUserIds.SplitFields strUserIds, ","
        
    For lIndex = 0 To astrUserIds.Size - 1
        If Parse(astrUserIds(lIndex), "|", 1) = cboUserIds.Text Then
            astrUserIds(lIndex) = Parse(astrUserIds(lIndex), "|", 1) & "|" & Str(cboSystems.ItemData(cboSystems.ListIndex))
        End If
    Next lIndex

    SetIniFileProperty "UserIds", astrUserIds.JoinFields(","), "User", m.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmRithmicLogin.UpdateCurrentUserId"
    
End Sub

