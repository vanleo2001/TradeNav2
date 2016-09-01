VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmHTTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection Settings"
   ClientHeight    =   3315
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   1275
      Left            =   4500
      TabIndex        =   5
      Top             =   2700
      Visible         =   0   'False
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
      Caption         =   "frmHTTP.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmHTTP.frx":0030
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmHTTP.frx":0050
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optFtpAlways 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   900
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
         Caption         =   "frmHTTP.frx":006C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmHTTP.frx":00DA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":00FA
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optFTP 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
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
         Caption         =   "frmHTTP.frx":0116
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmHTTP.frx":0176
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":0196
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optHTTP 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   2955
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
         Caption         =   "frmHTTP.frx":01B2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmHTTP.frx":020E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":022E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraHTTP 
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5415
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
      Caption         =   "frmHTTP.frx":024A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmHTTP.frx":0282
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmHTTP.frx":02A2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPassword2 
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   1980
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmHTTP.frx":02BE
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
         Tip             =   "frmHTTP.frx":02DE
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":02FE
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   285
         Left            =   1500
         TabIndex        =   13
         Top             =   1980
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmHTTP.frx":031A
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
         Tip             =   "frmHTTP.frx":033A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":035A
      End
      Begin HexUniControls.ctlUniTextBoxXP txtName 
         Height          =   285
         Left            =   1500
         TabIndex        =   11
         Top             =   1620
         Width           =   2595
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmHTTP.frx":0376
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
         Tip             =   "frmHTTP.frx":0396
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":03B6
      End
      Begin HexUniControls.ctlUniCheckXP chkLogin 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
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
         Caption         =   "frmHTTP.frx":03D2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmHTTP.frx":043A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":045A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPort 
         Height          =   285
         Left            =   4440
         TabIndex        =   8
         Top             =   900
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmHTTP.frx":0476
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
         Tip             =   "frmHTTP.frx":0496
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":04B6
      End
      Begin HexUniControls.ctlUniTextBoxXP txtAddress 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Top             =   900
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmHTTP.frx":04D2
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
         Tip             =   "frmHTTP.frx":04F2
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":0512
      End
      Begin HexUniControls.ctlUniCheckXP chkUseProxy 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   4815
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
         Caption         =   "frmHTTP.frx":052E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmHTTP.frx":05B4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":05D4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label4 
         Height          =   255
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
         Caption         =   "frmHTTP.frx":05F0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHTTP.frx":069E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":06BE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPassword2 
         Height          =   255
         Left            =   2700
         Top             =   2010
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
         Caption         =   "frmHTTP.frx":06DA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHTTP.frx":071A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":073A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPassword 
         Height          =   255
         Left            =   600
         Top             =   2010
         Width           =   915
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
         Caption         =   "frmHTTP.frx":0756
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHTTP.frx":0786
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":07A6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblName 
         Height          =   255
         Left            =   600
         Top             =   1650
         Width           =   915
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
         Caption         =   "frmHTTP.frx":07C2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHTTP.frx":07F6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":0816
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPort 
         Height          =   255
         Left            =   3900
         Top             =   930
         Width           =   555
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
         Caption         =   "frmHTTP.frx":0832
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHTTP.frx":0860
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":0880
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblAddress 
         Height          =   255
         Left            =   600
         Top             =   930
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
         Caption         =   "frmHTTP.frx":089C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHTTP.frx":08CC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":08EC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   2700
      Width           =   2835
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
      Caption         =   "frmHTTP.frx":0908
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmHTTP.frx":0934
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmHTTP.frx":0954
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   0
         TabIndex        =   0
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
         Caption         =   "frmHTTP.frx":0970
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmHTTP.frx":099A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":09BA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1560
         TabIndex        =   1
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
         Caption         =   "frmHTTP.frx":09D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmHTTP.frx":0A04
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmHTTP.frx":0A24
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmHTTP.frm
'' Description: Allow the user to enter in connection information
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    Dim iFtpMode As Long                ' Use FTP or HTTP?
    Dim strIniFile As String            ' Path and Name of the INI file
    Dim strPassword As String           ' Password for proxy logon

    strIniFile = AddSlash(g.strAppPath) & "GClient.INI"
    txtAddress.Text = GetIniFileProperty("ProxyServer", "", "Proxy", strIniFile)
    txtPort.Text = GetIniFileProperty("ProxyPort", "", "Proxy", strIniFile)
    chkUseProxy = GetIniFileProperty("UseProxy", vbUnchecked, "Proxy", strIniFile)
    chkLogin = GetIniFileProperty("SendLogin", vbUnchecked, "Proxy", strIniFile)
    txtName = GetIniFileProperty("LoginUser", "", "Proxy", strIniFile)
    
    strPassword = DecryptFromHex(GetIniFileProperty("LoginPassword", "", "Proxy", strIniFile), PasswordKey)
    txtPassword = strPassword
    txtPassword2 = strPassword
    
    SyncGclient '(to check if need to revert back to HTTP)
    Select Case GetIniFileProperty("UseFTP", 0, "Mode", strIniFile)
    Case 1
        optFTP = True
    Case 2
        optFtpAlways = True
    Case Else
        optHTTP = True
    End Select
    
    EnableControls
    ShowForm Me, eForm_Modal, frmMain
    
    If m.bOK = True Then
        If optFTP Then
            iFtpMode = 1
        ElseIf optFtpAlways Then
            iFtpMode = 2
        Else
            iFtpMode = 0
        End If
        SetIniFileProperty "UseFTP", iFtpMode, "Mode", strIniFile
        SetIniFileProperty "WhenSetFTP", CDbl(Now), "Mode", strIniFile
        SyncGclient
        
        SetIniFileProperty "UseProxy", chkUseProxy, "Proxy", strIniFile
        SetIniFileProperty "ProxyServer", Trim(txtAddress.Text), "Proxy", strIniFile
        SetIniFileProperty "ProxyPort", Trim(txtPort.Text), "Proxy", strIniFile
        
        SetIniFileProperty "SendLogin", chkLogin, "Proxy", strIniFile
        SetIniFileProperty "LoginUser", Trim(txtName.Text), "Proxy", strIniFile
        
        strPassword = Trim(txtPassword)
        SetIniFileProperty "LoginPassword", EncryptToHex(strPassword, PasswordKey), "Proxy", strIniFile
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    If FormIsLoaded("frmHTTP") Then Unload Me
    RaiseError "frmHTTP.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkLogin_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkLogin_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.chkLogin.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkUseProxy_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkUseProxy_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.chkUseProxy.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving information
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
    RaiseError "frmHTTP.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Verify Information, then unload the form and save information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If chkUseProxy.Value = vbChecked Then
        If Len(Trim(txtAddress.Text)) > 255 Or Len(Trim(txtAddress.Text)) = 0 Then
            MoveFocus txtAddress
            Err.Raise vbObjectError + 1000, , "Address must be between 1 and 255 characters in length."
        End If
        
        If ValOfText(txtPort.Text) < 1 Or ValOfText(txtPort.Text) > 65535 Then
            MoveFocus txtPort
            Err.Raise vbObjectError + 1000, , "Port must be from 1 to 65535."
        End If
        
        If txtPassword <> txtPassword2 Then
            MoveFocus txtPassword
            Err.Raise vbObjectError + 1000, , "The passwords do not match."
        End If
    End If

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help if the user presses F1
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
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

    Me.Icon = Picture16(ToolbarIcon("ID_Download"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks the X, unload the form without saving
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
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

    ' Only need the proxy stuff enabled if the user chose HTTP
    Enable chkUseProxy, optHTTP.Value
    
    Enable lblAddress, optHTTP.Value And chkUseProxy.Value
    Enable txtAddress, optHTTP.Value And chkUseProxy.Value
    Enable lblPort, optHTTP.Value And chkUseProxy.Value
    Enable txtPort, optHTTP.Value And chkUseProxy.Value
    
    Enable chkLogin, optHTTP.Value And chkUseProxy.Value
    Enable lblName, optHTTP.Value And chkUseProxy.Value And chkLogin.Value
    Enable txtName, optHTTP.Value And chkUseProxy.Value And chkLogin.Value
    Enable lblPassword, optHTTP.Value And chkUseProxy.Value And chkLogin.Value
    Enable txtPassword, optHTTP.Value And chkUseProxy.Value And chkLogin.Value
    Enable lblPassword2, optHTTP.Value And chkUseProxy.Value And chkLogin.Value
    Enable txtPassword2, optHTTP.Value And chkUseProxy.Value And chkLogin.Value

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optFTP_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optFTP_Click()
On Error GoTo ErrSection:

    Dim strMsg$

    If Me.Visible Then
        EnableControls
        strMsg = "FTP will be used for the rest of the day.|At midnight, the setting will revert back to| 'Try HTTP first' (the preferred protocol)."
        InfBox strMsg, "i", , "Protocol"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.optFTP_Click"
    Resume ErrExit
End Sub

Private Sub optFtpAlways_Click()
On Error GoTo ErrSection:

    Dim strMsg$

    If Me.Visible Then
        EnableControls
        strMsg = "This option should only be selected if there is a known issue with your computer, proxy server, or Internet Service Provider which is causing the preferred HTTP protocol to be unsuccessful in connecting to Genesis on an on-going basis."
        InfBox strMsg, "!", , "Warning"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.optFtpAlways_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optHTTP_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optHTTP_Click()
On Error GoTo ErrSection:

    If Me.Visible Then EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmHTTP.optHTTP_Click"
    Resume ErrExit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PasswordKey
'' Description: Determine the key for the password encryption
'' Inputs:      None
'' Returns:     Key for the Password Encryption
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PasswordKey() As String
On Error GoTo ErrSection:

    Dim strKey$, i&
    Dim mb As New cMemBuffer
       
    mb.PutByte 42, 0
    mb.PutByte 87, 1
    For i = 2 To 31
        mb.PutByte ((mb.GetByte(i - 1) * 7) Xor (mb.GetByte(i - 2) * 3)) And &HFF, i
    Next
       
ErrExit:
    PasswordKey = mb.Buffer
    Exit Function
    
ErrSection:
    RaiseError "frmHTTP.PasswordKey", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword_GotFocus
'' Description: When the box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_GotFocus()
    
    SelectAll txtPassword

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPassword2_GotFocus
'' Description: When the box gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword2_GotFocus()
    
    SelectAll txtPassword2

End Sub

