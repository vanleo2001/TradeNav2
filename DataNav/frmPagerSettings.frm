VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPagerSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4530
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5145
   Icon            =   "frmPagerSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL Frame3 
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   3960
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
      Caption         =   "frmPagerSettings.frx":030A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPagerSettings.frx":0336
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPagerSettings.frx":0356
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSendPage 
         Height          =   435
         Left            =   3180
         TabIndex        =   1
         Top             =   0
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
         Caption         =   "frmPagerSettings.frx":0372
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPagerSettings.frx":03B0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":03D0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   915
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
         Caption         =   "frmPagerSettings.frx":03EC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPagerSettings.frx":0416
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0436
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1020
         TabIndex        =   8
         Top             =   0
         Width           =   915
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
         Caption         =   "frmPagerSettings.frx":0452
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPagerSettings.frx":0480
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":04A0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdHelp 
         Height          =   435
         Left            =   2100
         TabIndex        =   0
         Top             =   0
         Width           =   915
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
         Caption         =   "frmPagerSettings.frx":04BC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPagerSettings.frx":04E6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0506
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   1905
      Left            =   180
      TabIndex        =   10
      Top             =   1860
      Width           =   4755
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
      Caption         =   "frmPagerSettings.frx":0522
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPagerSettings.frx":055E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPagerSettings.frx":057E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtMessage 
         Height          =   315
         Left            =   3300
         TabIndex        =   12
         Top             =   1440
         Width           =   1275
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPagerSettings.frx":059A
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
         Tip             =   "frmPagerSettings.frx":05C4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":05E4
      End
      Begin HexUniControls.ctlUniTextBoxXP txtPhone 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   540
         Width           =   4395
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPagerSettings.frx":0600
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
         Tip             =   "frmPagerSettings.frx":0620
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0640
      End
      Begin HexUniControls.ctlUniComboImageXP cboWait 
         Height          =   315
         Left            =   630
         TabIndex        =   11
         Top             =   990
         Width           =   585
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
         Tip             =   "frmPagerSettings.frx":065C
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":067C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label6 
         Height          =   255
         Left            =   180
         Top             =   1470
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
         Caption         =   "frmPagerSettings.frx":0698
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPagerSettings.frx":0708
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0728
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   1260
         Top             =   1050
         Width           =   3315
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
         Caption         =   "frmPagerSettings.frx":0744
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPagerSettings.frx":07BC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":07DC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   180
         Top             =   300
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
         Caption         =   "frmPagerSettings.frx":07F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPagerSettings.frx":0876
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0896
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   180
         Top             =   1050
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
         Caption         =   "frmPagerSettings.frx":08B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPagerSettings.frx":08DE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":08FE
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL Frame2 
      Height          =   1530
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   4755
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
      Caption         =   "frmPagerSettings.frx":091A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPagerSettings.frx":0956
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPagerSettings.frx":0976
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDefault 
         Height          =   375
         Left            =   2940
         TabIndex        =   7
         Top             =   990
         Width           =   1635
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
         Caption         =   "frmPagerSettings.frx":0992
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPagerSettings.frx":09DE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":09FE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtModemInit 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1020
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPagerSettings.frx":0A1A
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
         Tip             =   "frmPagerSettings.frx":0A3A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0A5A
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAutoDetect 
         Height          =   375
         Left            =   2220
         TabIndex        =   4
         Top             =   300
         Width           =   2355
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
         Caption         =   "frmPagerSettings.frx":0A76
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPagerSettings.frx":0AD2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0AF2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboComPort 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   330
         Width           =   855
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
         Tip             =   "frmPagerSettings.frx":0B0E
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0B2E
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label5 
         Height          =   255
         Left            =   180
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
         Caption         =   "frmPagerSettings.frx":0B4A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPagerSettings.frx":0BBC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0BDC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblModem 
         Height          =   255
         Left            =   180
         Top             =   360
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
         Caption         =   "frmPagerSettings.frx":0BF8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPagerSettings.frx":0C32
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPagerSettings.frx":0C52
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmPagerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAutoDetect_Click()
On Error Resume Next

    Dim strTemp$, i&, nComPort&, strTitle$

    strTitle = "Auto-Detect ComPort"
    Screen.MousePointer = 11
    EnableControls False
                    
    With frmMain.MSComm1
        ' try each Com port
        For i = 1 To 9
            Me.Caption = "Trying COM" & CStr(i) & " ..."
            ' Open com port
            .CommPort = i
            ' 9600 baud, no parity, 8 data, and 1 stop bit.
            .Settings = "9600,N,8,1"
            ' Tell the control to read entire buffer when Input is used.
            .InputLen = 0
            ' Open the port.
            .PortOpen = True
            ' Send the attention command to the modem.
            .Output = "AT" + Chr$(13)
            ' Wait for data to come back to the serial port.
            Sleep 2
            ' Read the "OK" response data in the serial port.
            strTemp = .Input
            ' Close port
            .PortOpen = False
            DoEvents
            
            ' see if got a good response
            If InStr(strTemp, "OK") > 0 Then
                nComPort = i
                Exit For
            End If
        Next
    End With
    Screen.MousePointer = 0
    EnableControls True
    
    If nComPort > 0 Then
        cboComPort.ListIndex = nComPort - 1
        InfBox "Modem found on COM" & CStr(nComPort), , , strTitle
    Else
        InfBox "No modem was detected.", "e", , strTitle
    End If
    
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDefault_Click()
On Error GoTo ErrSection:

    txtModemInit = MODEM_INIT_DEFAULT

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.cmdDefault.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdHelp_Click()
On Error GoTo ErrSection:

    Dim strFile$
    strFile = App.Path & "\Info\PagerHelp.rtf"
    frmMessage.ShowMe "Help for Pager Settings", "@" & strFile, eModalMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.cmdHelp.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrSection:
  
    If Len(Trim(txtPhone)) = 0 Then
        InfBox "Phone number to dial is required.", "e"
        MoveFocus txtPhone
        Exit Sub
    End If

    ' save properties
    Call SetIniFileProperty("Pager", Trim(txtPhone), "PagerSettings", g.strIniFile)
    Call SetIniFileProperty("ComPort", cboComPort.ListIndex + 1, "PagerSettings", g.strIniFile)
    Call SetIniFileProperty("WaitSeconds", ValOfText(cboWait.Text), "PagerSettings", g.strIniFile)
    Call SetIniFileProperty("ModemInit", Trim(txtModemInit), "PagerSettings", g.strIniFile)
    
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdSendPage_Click()
On Error GoTo ErrSection:

    EnableControls False
    Me.Caption = "Sending Test Page ..."
    If Not DialPager(cboComPort.ListIndex + 1, txtPhone, ValOfText(cboWait.Text), txtMessage, txtModemInit, False) Then
        InfBox "Error sending page.|(the modem may be on a different Com Port)", "e", , "Send Test Page"
    End If
    EnableControls True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.cmdSendPage.Click", eGDRaiseError_Show
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
    RaiseError "frmPagerSettings.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim i&

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    With cboWait
        For i = 2 To 60 Step 2
            .AddItem CStr(i)
        Next
        .Text = "10"
    End With
    
    With cboComPort
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
        .AddItem "COM9"
        .ListIndex = 0
    End With
    
    ' get properties
    On Error Resume Next
    'txtPhone = "575-3898"
    txtPhone = GetIniFileProperty("Pager", "", "PagerSettings", g.strIniFile)
    i = GetIniFileProperty("ComPort", 1, "PagerSettings", g.strIniFile)
    If i >= 1 Then cboComPort.ListIndex = i - 1
    i = GetIniFileProperty("WaitSeconds", 10, "PagerSettings", g.strIniFile)
    If i > 0 Then cboWait.Text = CStr(i)
    txtModemInit = GetIniFileProperty("ModemInit", MODEM_INIT_DEFAULT, "PagerSettings", g.strIniFile)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Sub ShowMe()
On Error GoTo ErrSection:

    ' if the first time, show help
    If GetIniFileProperty("HelpShown", 0, "PagerSettings", g.strIniFile) = 0 Then
        Call SetIniFileProperty("HelpShown", 1, "PagerSettings", g.strIniFile)
        cmdHelp_Click
    End If

    ShowForm Me, True
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.ShowMe", eGDRaiseError_Raise
    
End Sub

Private Sub EnableControls(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    Dim i&
    Static strOrigCaption$
    
    If Len(strOrigCaption) = 0 Then strOrigCaption = Me.Caption
    
    On Error Resume Next
    For i = 0 To Me.Controls.Count - 1
        Me.Controls(i).Enabled = bEnable
    Next
    
    If bEnable Then
        Me.Caption = strOrigCaption
        MoveFocus cmdSave
    End If
    DoEvents
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPagerSettings.EnableControls", eGDRaiseError_Raise
    
End Sub

